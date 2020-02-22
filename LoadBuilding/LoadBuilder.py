"""

Created by Nicolas Raymond on 2019-11-15.

This file provides a LoadBuilder. This object manage all trailer's loading from one plant to another.

Last update : 2019-10-31
By : Nicolas Raymond

"""
import numpy as np
import LoadingObjects as LoadObj
import pandas as pd
from collections import Counter
from packer import newPacker
from math import floor
from copy import deepcopy as dc


class LoadBuilder:

    """
    Object that deals with loadings construction from one plant to another
    """
    validate_with_ref = False
    trailer_reference = None
    patching_activated = False
    score_multiplication_base = 1.20  # used to boost the score of a load when there's mandatory crates
    overhang_authorized = 51.5   # Maximum of overhang authorized for a trailer (in inches)
    max_trailer_length = 636  # Maximum load length possible
    plc_lb = 0.80  # Lowest percentage of trailer's length that must be covered (using validation length)
    individual_width_tolerance = 68  # Smallest width tolerated for a lonely crate (without anything by his side)

    def __init__(self, trailers_data):
        """

        :param trailers_data: Pandas data frame containing details on trailers available
        """

        self.trailers_data = trailers_data
        self.model_names = []
        self.warehouse, self.remaining_crates = LoadObj.Warehouse(), LoadObj.CratesManager()
        self.metal_warehouse, self.metal_remaining_crates = LoadObj.Warehouse(), LoadObj.CratesManager()
        self.trailers, self.trailers_done, self.unused_models = [], [], []
        self.all_size_codes = set()

    def __len__(self):
        return len(self.trailers_done)

    def __warehouse_init(self, models_data, ranking={}):

        """
        Initializes a warehouse according to the models available in model data

        :param: models_data : pandas dataframe with the following columns
        [QTY | MODEL | LENGTH | WIDTH | HEIGHT | NUMBER_PER_CRATE | CRATE_TYPE | STACK_LIMIT | OVERHANG ]

        :param ranking: dictionary with size code as keys and lists of integers as value
        """

        # For all lines of the data frame
        for i in models_data.index:

            # We save the quantity of the model and the plant_to
            qty = models_data['QTY'][i]

            if qty > 0:

                # We save the crate type
                crate_type = models_data['CRATE_TYPE'][i]

                # We save the name of the model
                for j in range(qty):
                    self.model_names.append((models_data['MODEL'][i], crate_type))

                # We save the stack limit
                stack_limit = models_data['STACK_LIMIT'][i]

                # We save the number of models per crate
                nbr_per_crate = models_data['NBR_PER_CRATE'][i]

                # We save the overhang permission indicator
                overhang = bool(models_data['OVERHANG'][i])

                # We compute the number of models per stack
                items_per_stack = stack_limit * nbr_per_crate

                # We compute the number of stacks that we can build
                nbr_stacks = int(np.floor(qty / items_per_stack))

                # We save the crate type
                crate_type = models_data['CRATE_TYPE'][i]

                # We save the number of MANDATORY crates if it's available in the input data frame
                if 'NB_OF_X' in models_data.columns:
                    total_of_mandatory = int(models_data['NB_OF_X'][i])
                else:
                    total_of_mandatory = 0

                # We save the number of individual crates to build and convert it into
                # integer to avoid conflict with range function. Also, with int(), every number in [0,1[ will
                # be convert as 0. This way, no individual crate of SP2 will be build if there's less than 2 SP2 left
                nbr_individual_crates = int((qty - (items_per_stack * nbr_stacks)) / nbr_per_crate)

                crates_component = [[models_data['MODEL'][i]] * nbr_per_crate,
                                    max(models_data['LENGTH'][i], models_data['WIDTH'][i]),
                                    min(models_data['WIDTH'][i], models_data['LENGTH'][i]),
                                    models_data['HEIGHT'][i],
                                    stack_limit, overhang]

                # We select the good type of storage of the stacks and crates that will be build
                if crate_type == 'W':
                    warehouse = self.warehouse
                    crates_manager = self.remaining_crates

                else:  # elif crate_type == 'M'
                    warehouse = self.metal_warehouse
                    crates_manager = self.metal_remaining_crates

                # We take the list of ranking associate with the size code or create one filled with 0
                # and start an index to navigate through ranking list
                r = ranking.get(models_data['MODEL'][i], [0]*qty)
                index = 0

                for j in range(nbr_stacks):

                    # We initialize a list that will contain all the crates needed to build the stack
                    stack_crates = []

                    for k in range(stack_limit):
                        crate_temp_components = dc(crates_component) + [total_of_mandatory > 0] + [r[index], crate_type]
                        total_of_mandatory -= 1
                        index += 1
                        stack_crates.append(LoadObj.Crate(*crate_temp_components))

                    # We build the stack and send it into the warehouse
                    warehouse.add_stack(LoadObj.Stack(stack_crates))

                for j in range(nbr_individual_crates):

                    # We add the missing number of mandatory crates to the stacks component list
                    crate_temp_components = dc(crates_component) + [total_of_mandatory > 0] + [r[index], crate_type]

                    # We build the crate and send it to the crates manager
                    crates_manager.add_crate(LoadObj.Crate(*crate_temp_components))

                    # We update the total number of mandatory left and increment the index
                    total_of_mandatory -= 1
                    index += 1

    def __trailers_init(self):

        """
        Initializes the list with all the trailers available for the loading

        """

        # For every lines of the data frame
        for i in self.trailers_data.index:

            # We save the quantity, the plant_from and the plant _to
            qty = self.trailers_data['QTY'][i]

            if qty > 0:

                # We save trailer's length
                t_length = self.trailers_data['LENGTH'][i]

                # We compute overhanging measure allowed for the trailer
                if bool(self.trailers_data['OVERHANG'][i]):
                    trailer_oh = min(self.max_trailer_length - t_length, self.overhang_authorized)
                else:
                    trailer_oh = 0

                # We set the priority rank of the trailer
                if 'PRIORITY_RANK' in self.trailers_data.columns:
                    rank = self.trailers_data['PRIORITY_RANK'][i]
                else:
                    rank = 1

                # We build "qty" trailer that we add to the trailers list
                for j in range(0, qty):
                    self.trailers.append(LoadObj.Trailer(self.trailers_data['CATEGORY'][i], t_length,
                                                         self.trailers_data['WIDTH'][i],
                                                         self.trailers_data['HEIGHT'][i], rank, trailer_oh))

    def __prepare_warehouse(self):

        """
        Finishes stacking procedure with crates that weren't stacked at first

        """

        # Finishing stacking procedure for wooden crates
        if len(self.remaining_crates.crates) > 0:
            self.remaining_crates.create_stacks(self.warehouse)

            if self.patching_activated and len(self.remaining_crates.stand_by_crates) > 0:
                self.remaining_crates.create_incomplete_stacks(self.warehouse)

            else:
                self.remaining_crates.save_unused_crates(self.unused_models)

        # Finishing stacking for metal crates
        if len(self.metal_remaining_crates.crates) > 0:
            self.metal_remaining_crates.create_stacks(self.metal_warehouse)

            if self.patching_activated and len(self.metal_remaining_crates.stand_by_crates) > 0:
                self.metal_remaining_crates.create_incomplete_stacks(self.metal_warehouse)

            else:
                self.metal_remaining_crates.save_unused_crates(self.unused_models)

    def __trailer_packing(self, initial_lb=1.00, decreasing_step=0.02):
        """

        Using a modified version of Skyline 2D bin packing algorithms provided by the rectpack library,
        this function performs trailer loading by considering the bottom surface of every stack
        as a rectangle that needs to be placed in a bin.

        Check https://github.com/secnot/rectpack for more informations on source code and
        http://citeseerx.ist.psu.edu/viewdoc/download;jsessionid=3A00D5E102A95EF7C941408817666342?doi=10.1.1.695.2918&rep=rep1&type=pdf
        for more information on algorithms implemented themselves.

        """

        # We sort trailer by the area of their surface
        self.trailers.sort(key=lambda s: (s.priority, s.area()), reverse=True)

        # We initialize lower bounds of length and area coverage that need to be satistied by the trailer done
        lower_bound = initial_lb

        # While we have not reached the lower bound of percentage covered and there's is still item available
        while lower_bound >= self.plc_lb and (len(self.warehouse) != 0 or len(self.metal_warehouse) != 0):

            # We initialize a variable that will contain the name of the last category that didn't satisfied lb
            last_category = ''

            # We initialize a list of tuple with potential trailers and sort functions associated to them
            potential_trailers = []

            # For all trailer that hasn't been used yet
            for t in [trailer for trailer in self.trailers if not trailer.packed]:

                if t.category != last_category:

                    # We initialize a list that will contain tuples of crate type
                    # and "packer" loading strategies that were tried.
                    packers = []

                    # We try to fill the trailer the best way as possible considering best configurations between
                    # wooden warehouse and metal warehouse
                    for crate_type, warehouse in [('W', self.warehouse), ('M', self.metal_warehouse)]:

                        # We update the packers tuple list
                        self.__test_all_config(packers, crate_type, warehouse, t, lower_bound)

                    # We save the index of the best loading configuration that respected the constraint of plc_lb
                    best_packer, crate_type, score = self.__select_best_packer(packers)

                    # If an index is found (at least one load satisfies the constraint for this category of trailer)
                    if best_packer is not None:

                        # We save the crate type, the score and the packer associated to the trailer
                        t.packer = best_packer
                        t.score = score
                        t.crate_type = crate_type

                        # We add this potential trailer to the list
                        potential_trailers.append(t)

                    last_category = t.category

            # If we found trailers that were respecting the constraint of plc
            if len(potential_trailers) != 0:

                selected_trailer = self.__select_best_trailer(potential_trailers)

                # We select the warehouse that will be used to pack the trailer
                if selected_trailer.crate_type == 'W':
                    warehouse = self.warehouse
                else:
                    warehouse = self.metal_warehouse

                # We pack the trailer and print the trailer
                selected_trailer.pack(warehouse)

            else:
                lower_bound = round(lower_bound - decreasing_step, 2)

        # We remove trailer that were not used during the loading process
        self.__remove_leftover_trailers()

        # We save unused models from both warehouses
        self.warehouse.save_unused_crates(self.unused_models)
        self.metal_warehouse.save_unused_crates(self.unused_models)

    @staticmethod
    def __select_best_trailer(potential_trailers):
        """
        Select the trailer with the highest score among the list of potential trailers
        and reset all the others
        :param potential_trailers: list of trailers
        :return: best trailer
        """
        # We sort trailer by the area of their surface
        potential_trailers.sort(key=lambda t: t.score, reverse=True)

        for i in range(1, len(potential_trailers)):
            potential_trailers[i].reset()

        return potential_trailers[0]

    @staticmethod
    def __select_best_packer(packers_list):

        """
        Pick the best loading configuration done (the best packer) among the list according to the number of units
        placed in the trailer.

        :param packers_list: List containing tuples with packers object, crate_type and score associated to it
        :return: best packer in terms of score
        """
        if len(packers_list) == 0:
            return None, None, 0

        # We sort packers in descending order with their score
        packers_list.sort(key=lambda p: p[2], reverse=True)

        return packers_list[0][0], packers_list[0][1], packers_list[0][2]

    def __test_all_config(self, packers, crate_type, warehouse, trailer, lower_bound, sort_choice=0):

        """
        Tests efficiently different load configurations possible for the trailer and selected warehouse

        :param packers: list of tuples with crate types and packers object
        :param crate_type: One type among 'W' and 'M'
        :param warehouse: object of class Warehouse from which we'll pull the stacks
        :param trailer: object of class Trailer
        :param lower_bound: actual lower bound of coverage that must be satisfied
        :param sort_choice: index indicating the sort option to take from sort_options list
        """
        # We define our sort options
        sort_options = [(sort_by_area, True), (sort_by_length, True), (sort_by_width, True), (sort_by_ratio, True),
                        (sort_by_area, False), (sort_by_length, False), (sort_by_width, False), (sort_by_ratio, False)]

        if sort_choice < len(sort_options):

            # We save sort option chosen
            sort_function = sort_options[sort_choice][0]
            ranking_effectiveness = sort_options[sort_choice][1]

            # We save actual packers list length
            nb_configs_already_found = len(packers)

            # We save the number of actual stacks available in the warehouse
            nb_stacks = len(warehouse)

            # We compute all possible configurations of loading (efficiently)
            if nb_stacks != 0 and \
                    sum([stack.length for stack in warehouse.stacks_to_ship]) >= self.plc_lb*trailer.length:

                sort_function(warehouse, ranking_effective=ranking_effectiveness)
                all_configs = self.__create_all_configs(warehouse, trailer)
            else:
                all_configs = []

            # If there's possible configurations
            if len(all_configs) != 0:

                for config in all_configs:

                    # We initialize a packer with default parameter (except rotation)
                    packer = newPacker(rotation=False)

                    # We add stacks to load in the trailer (the rectangles)
                    for i in range(len(config)):

                        # If the rectangle is rotated
                        if config[i]:
                            packer.add_rect(warehouse[i].length, warehouse[i].width, rid=i,
                                            overhang=warehouse[i].overhang)

                        else:
                            packer.add_rect(warehouse[i].width, warehouse[i].length, rid=i,
                                            overhang=warehouse[i].overhang)

                    # We add other dummy bins to store rectangles that do not enter in our trailer (1st bin)
                    for i in range(2):
                        packer.add_bin(trailer.width, trailer.length, bid=None, overhang=trailer.oh)

                    # We execute the packing
                    packer.pack()

                    # We complete the packing (look if some unconsidered rectangles could enter at the end)
                    self.__complete_packing(warehouse, trailer, packer, len(config))

                    # We check if packing respect plc lower bound and how many items it contains
                    qualified, score = self.__validate_packing(trailer, crate_type, packer, lower_bound)

                    # We save the loading configuration (the packer) if the loading is qualified
                    if qualified:
                        packers.append((packer, crate_type, score))

                # We check if some loading configurations were found
                new_configuration_found = len(packers) - nb_configs_already_found

                # If nothing was found we retry the same steps with a new sort function
                if new_configuration_found == 0:
                    self.__test_all_config(packers, crate_type, warehouse, trailer,
                                           lower_bound, sort_choice=sort_choice+1)

    def __validate_packing(self, trailer, crate_type, packer, lower_bound):

        """
        Verifies if the packing satisfies plc_lb constraint (Lower bound of percentage of length that must be covered)

        :param trailer: trailer for which we're testing configurations possible
        :param crate_type: 'W' for wood, 'M' for metal
        :param packer: Packer object
        :param lower_bound: actual lower bound of coverage that must be satisfied
        :returns: Boolean indicating if the loading satisfies constraint and a score for the load
        """

        # We set few useful variable values
        mandatory_crates = 0
        score = 0
        qualified = True
        bin = packer[0]
        if crate_type == 'W':
            warehouse = self.warehouse
        else:
            warehouse = self.metal_warehouse

        if not bin.valid_length(lower_bound, self.individual_width_tolerance):
            qualified = False

        elif trailer.category == 'DRYBOX' and self.validate_with_ref and\
                not self.__sanity_check(crate_type, warehouse, [rect.rid for rect in bin]):
            qualified = False
        else:
            used_area = bin.used_area()
            mandatory_crates += sum([warehouse[rect.rid].nb_of_mandatory for rect in bin])
            score_boost = self.score_multiplication_base**mandatory_crates
            score = used_area*score_boost

        return qualified, score

    def __max_rect_upperbound(self, warehouse, trailer, last_upper_bound):

        """
        Recursive function that approximates a maximum number of rectangle that can fit in the trailer,
        according to rectangles available that are going to enter in the trailer.

        :param warehouse: Object of class Warehouse
        :param trailer: Object of class Trailer
        :param last_upper_bound: Last upper bound found (int)
        :return: Approximation of the maximal number (int)
        """

        # We build a set containing all pairs of objects' width and length in a certain range in the warehouse
        unique_tuples = set((warehouse[i].width, warehouse[i].length)
                            for i in range(min(last_upper_bound, len(warehouse))))

        # We initialize the shortest length found with the maximal length found
        shortest_length = max(max(dimensions) for dimensions in unique_tuples)

        # We find the real minimum height looking at all item in a certain range in the warehouse
        for dimensions in unique_tuples:

            # We pretend that the item doesn't fit in the trailer
            fit = False

            # We initialize two variables "length" and "length_if_rotated"
            length_if_rotated, length = None, None

            # We create two variables containing width and length of item to facilitate comprehension of code
            item_width, item_length = dimensions[0], dimensions[1]

            # If the item is less large than long and it's possible to rotate it
            if item_width < item_length <= trailer.width:
                # We save his width divided by the number of times it fits side by side once rotated
                length_if_rotated = item_width / (floor(trailer.width / item_length))
                fit = True

            # If the item fits with the original positioning
            if item_length <= trailer.length and item_width <= trailer.width:
                # We save is length divided by the number of times it fits side by side
                length = item_length / (floor(trailer.width / item_width))
                fit = True

            # We update min_shortest_length value
            if fit:
                lengths_list = [l for l in [length_if_rotated, length] if l is not None]
                shortest_length = min([shortest_length] + lengths_list)

        # We compute the upper bound
        new_upper_bound = floor((trailer.length + trailer.oh) / shortest_length)

        # If the upper bound found equals the upper bound found in the last iteration we stop the process
        if new_upper_bound == last_upper_bound:
            return new_upper_bound

        else:
            return self.__max_rect_upperbound(warehouse, trailer, new_upper_bound)

    def __create_all_configs(self, warehouse, trailer):

        """
        Creates all configurations of loading (using warehouse units) that can be done for the trailer.
        To avoid considering a large number of bad configurations and enhance the efficiency of the algorithm,
        we will pre-set wisely the positions of rectangles for a certain range of the trailer and THEN consider
        all possible configurations for the end of the loading of the trailer.

        :param warehouse: Object of class warehouse
        :param trailer: Object of class Trailer
        :return: List of list of boolean indicating permission of rotation
        """

        # We initialize a numpy array of configurations (only one configuration for now)
        configs, nb_oversize = warehouse.merge_for_trailer(trailer, self.individual_width_tolerance)
        configs = np.array([configs])

        # We compute an upper bound for the maximal number of rectangles that can fit in our trailer
        nb_rect_to_consider = len(warehouse) - nb_oversize

        if nb_rect_to_consider == 0:
            return []

        ub = self.__max_rect_upperbound(warehouse, trailer, nb_rect_to_consider)

        # We save the number of stack pre-rotated
        nb_of_pre_rotated = configs.shape[1]

        # Initialization of list that will contain index of stack that cannot fit in the trailer
        leftover = []

        # We set the start index and the end index of research to build possible configurations
        i = nb_of_pre_rotated
        end_index = min(len(warehouse), ub)

        while i < end_index:

            # We pretend that the item doesn't fit in the trailer
            fit = False

            # We initialize an empty numpy array of new configurations
            new_configs = np.array([[]])

            # If it's possible to rotate the i-th item in the warehouse for this trailer
            if trailer.fit(warehouse[i], rotated=True):
                # We add the rotated rectangle indicator (True) to all configs found until now
                true_vec = [[True]] * len(configs)
                new_configs = np.append(np.copy(configs), true_vec, axis=1)
                fit = True

            # If the i-th item fit not rotated in this trailer
            if trailer.fit(warehouse[i]):

                # We add the rotated rectangle indicator (False) to all configs found until now
                false_vec = [[False]] * len(configs)
                configs = np.append(configs, false_vec, axis=1)
                fit = True

            if not fit:

                # We add the stack index in the leftover list
                leftover.append(i)

                # (If possible) We extend our research space cause we didn't enter the current stack in our configs
                end_index = min(len(warehouse), end_index + 1)

            else:
                # We update configurations
                if new_configs.shape[1] == configs.shape[1]:
                    configs = np.append(configs, new_configs, axis=0)

                elif new_configs.shape[1] > configs.shape[1]:
                    configs = new_configs

            i += 1

        # We push leftover at the end of the warehouse to avoid conflict during loading process
        if len(leftover) > 0:
            for j in leftover:
                warehouse.add_stack(warehouse[j])
            warehouse.remove_stacks(leftover)

        return configs

    def __sanity_check(self, crate_type, warehouse, warehouse_used_indexes):

        # Initialization of the reference trailer
        t = dc(LoadBuilder.trailer_reference)

        # We save the actual length of the trailer
        original_length = t.length

        # We change the length of the trailer to avoid wrong sanity check results because of overflow problem
        t.length = self.max_trailer_length

        # Initialization of an empty list that will contain tuples of crate_type and packer
        packers = []

        # Construction of a copy of the warehouse
        w = dc(warehouse)

        # We build a list with all indexes unused
        all_indexes = set(list(range(len(warehouse))))
        unused_indexes = list(all_indexes-set(warehouse_used_indexes))

        # We remove unused index from the warehouse copy
        w.remove_stacks(unused_indexes)

        # We compute the plc_lb to satisfy based on original_length and select the best packer with our function
        lowerbound = round((self.plc_lb*original_length)/t.length, 4)

        # We test all configurations possible
        self.__test_all_config(packers, crate_type, w, t, lowerbound)

        # False would indicate that no satisfying load could be done with the reference trailer
        return len(packers) != 0

    @staticmethod
    def __complete_packing(warehouse, trailer, packer, start_index):

        """
        Verifies if one (or multiple) item unconsidered in the first part of packing fits at the end of the trailer

        :param warehouse: Object of class warehouse
        :param trailer: Object of class Trailer
        :param packer: Packer object
        :param start_index: If i is the index of the last item we considered in first part, then start_index = i + 1
        :returns : number of stacks added
        """

        # We look if there are items remaining in the warehouse (that were not considered in the first phase of packing)
        # and if there's still place in the trailer.
        if len(warehouse) != start_index and max([rect.top for rect in packer[0]]) < trailer.length:

            # We save the current number of stack in the trailer
            last_res = len(packer[0])

            # We initialize a new packer with rotation not allowed to simply computation and save time
            new_packer = newPacker(rotation=False)

            # We add rectangles unconsidered in the first phase of packing
            for i in range(start_index, len(warehouse)):
                new_packer.add_rect(warehouse[i].width, warehouse[i].length, rid=i, overhang=warehouse[i].overhang)

            # We add a large number of dummy bins
            for j in range(len(warehouse) - start_index + 1):
                new_packer.add_bin(trailer.width, trailer.length, overhang=packer[0].overhang_measure)

            # We open the first bin
            new_packer._open_bins.append(packer[0])

            # We allow unlock rotation in this first bin.
            new_packer[0].rot = True

            # We pack the new packer
            new_packer.pack(reset_opened_bins=False)

            # We return the number of stacks added
            return len(new_packer[0]) - last_res

        else:
            return 0

    def __remove_leftover_trailers(self):

        """
        Removes trailers that contain no units after the loading
        """

        # We initialize an index at the end of the list containing trailers
        i = len(self.trailers) - 1

        # We remove trailer that we're not used during the loading process
        while i >= 0:
            if self.trailers[i].nbr_of_units() == 0:
                self.trailers.pop(i)
            i -= 1

    def __select_top_n(self, n):

        """
        Selects the n best trailers in terms of number of mandatory crates and score
        """
        # We sort trailer in decreasing order by their score
        self.trailers.sort(key=lambda t: (t.nb_of_mandatory(), t.score), reverse=True)

        # We initialize an index at the end of the list containing trailers
        i = len(self.trailers) - 1

        # We unload all trailers that exceed maximum number of trailer
        while i >= n:
            self.trailers[i].unload_trailer(self.unused_models)
            self.trailers.pop(i)
            i -= 1

    def __size_code_used(self):

        """
        Save all size code used in the new loads

        :return : list of size codes (also called model names in other part of code)
        """
        # We counts all the models that were introduced in loads
        counts = Counter(self.model_names)  # Initial counts
        counts_of_unused = Counter(self.unused_models)
        counts.subtract(counts_of_unused)

        # We erase model names in memory
        self.model_names.clear()
        self.unused_models.clear()

        # We save the size codes used for the loads done in this iteration
        size_codes = list(counts.elements())
        self.all_size_codes.update([size_code[0] for size_code in size_codes])

        return size_codes

    def __update_trailers_data(self):

        """
        Updates quantities in original trailers data frame

        """
        # We counts all different trailers used
        counts = Counter([(t.category, t.length, t.width) for t in self.trailers])

        # We update the trailers data
        for key, item in counts.items():
            row_to_change = self.trailers_data.index[(self.trailers_data['CATEGORY'] == key[0]) &
                                                     (self.trailers_data['LENGTH'] == key[1]) &
                                                     (self.trailers_data['WIDTH'] == key[2])].tolist()

            self.trailers_data.loc[row_to_change[0], 'QTY'] -= item

        # We save trailers done and clear the current trailers list
        self.trailers_done += self.trailers.copy()
        self.trailers.clear()

    def get_loading_summary(self):

        """
        Create a Pandas data frame with a summary of all loads done by the LoadBuilder

        :return: Pandas data frame
        """
        size_codes = list(self.all_size_codes)

        # We initialize a data frame with column names needed
        data_frame = pd.DataFrame(columns=(["TRAILER", "TRAILER LENGTH", "LOAD LENGTH"] + size_codes))

        # We initialize an index
        i = 0

        # We sort trailer with their average ranking
        self.trailers_done.sort(key=lambda t: (t.nb_of_mandatory(), t.average_ranking()))

        # We add a line in the dataframe for every trailer used
        for trailer in self.trailers_done:

            # We save the quantities of every models inside the trailer
            s = Counter(trailer.load_summary())

            # Every line of data frame has the category of trailer, his length, his remaining_length (in feets)
            # and the quantities of every models in it.
            data_frame.loc[i] = [trailer.category] + [round(trailer.length / 12, 1)] + \
                                [round(trailer.length_used / 12, 1)] + \
                                [s[model] if s[model] > 0 else '' for model in size_codes]
            i += 1

        # We execute a groupby with trailer in the same category
        data_frame = data_frame.groupby(data_frame.columns.tolist()).size().to_frame('QTY').reset_index()

        # We rearrange columns
        cols = data_frame.columns.tolist()
        cols = cols[0:1] + cols[-1:] + cols[1:-1]
        data_frame = data_frame[cols]

        # We set indexes
        # data_frame.set_index("TRAILER", inplace=True)

        return data_frame

    def number_of_units(self):
        """
        Returns the total number of units in the trailers done
        """
        return sum([trailer.nbr_of_units() for trailer in self.trailers_done])

    def __complete_packed_stacks(self):

        if len(self.trailers_done) != 0:

            # We look if it's possible to complete some stacks in already packed trailer
            self.remaining_crates.complete_trailers_stack(self.trailers_done)
            self.metal_remaining_crates.complete_trailers_stack(self.trailers_done)

    def __fill_trailers_empty_spaces(self):

        """
        Looks if the new inputs items could fit in an already packed trailer
        """
        if len(self.trailers_done) != 0:

            # We sort trailers done by their used length (to fill trailer with more
            self.trailers_done.sort(key=lambda t: t.used_area())

            # We also try to fill empty spaces in already packed trailer
            for trailer in self.trailers_done:

                # We save the warehouse we can use to fill the rest of the trailer
                if trailer.crate_type == 'W':
                    warehouse = self.warehouse

                else:  # elif crate_type == 'M'
                    warehouse = self.metal_warehouse

                nb_stacks_added = self.__complete_packing(warehouse, trailer, trailer.packer, start_index=0)

                if nb_stacks_added > 0:

                    trailer.pack(warehouse, nb_stacks_added)

    def build(self, models_data, max_load, plot_load_done=False, ranking={}):

        """
        This is the core of the object.
        It contains the principal steps of the loading process.

        :param models_data: Pandas data frame containing details on models to load
        :param max_load: maximum number of loads
        :param plot_load_done: boolean that indicates if plots of loads are going to be shown
        :param ranking: dictionnary with ranking lists associated with each size code
        :return: list of tuples with size code used and crate types
        """
        # We look if models_data is empty
        if models_data.empty:
            return []

        # We init the warehouse
        self.__warehouse_init(models_data, ranking)

        # We init the list of trailers
        self.__trailers_init()

        # We try to complete stacks that we're already done with the new inputs
        # self.__complete_packed_stacks()

        # We finish the stacking process with leftover crates
        self.__prepare_warehouse()

        # We try to fill empty spaces in loads that we're already done with the new inputs
        if self.patching_activated:
            self.__fill_trailers_empty_spaces()

        # We execute the loading of the trailers
        self.__trailer_packing()

        # We consider the max
        nb_new_loads = len(self.trailers)
        total_nb_loads = len(self.trailers_done) + nb_new_loads

        if max_load < total_nb_loads:
            self.__select_top_n(max(nb_new_loads - (total_nb_loads - max_load), 0))

        # We plot every new loads if the user wants to
        if plot_load_done:
            for trailer in self.trailers:
                trailer.plot_load()

        # We update all data
        self.__update_trailers_data()

        return self.__size_code_used()


def set_trailer_reference(ref):

    """
    Set the trailer_reference static attribute of the Trailer class

    :param ref: pandas dataframe containing one row with trailer's data
    """
    LoadBuilder.trailer_reference = LoadObj.Trailer(cat='FLATBED_48', l=ref['LENGTH'][0], w=ref['WIDTH'][0],
                                                    h=ref['HEIGHT'][0], p=0, oh=ref['OVERHANG'][0])


def sort_by_volume(warehouse, ranking_effective=False):
    """
    Sorts stacks to ship by their volumes (and their ranking if True)
    """
    warehouse.sort_by_volume(ranking_effective=ranking_effective)


def sort_by_area(warehouse, ranking_effective=False):
    """
    Sorts stacks by their area (and their ranking if True)
    """
    warehouse.sort_by_area(ranking_effective=ranking_effective)


def sort_by_width(warehouse, ranking_effective=False):
    """
    Sorts stacks by their width (and their ranking if True)
    """
    warehouse.sort_by_width(ranking_effective=ranking_effective)


def sort_by_length(warehouse, ranking_effective=False):
    """
    Sorts stacks by their length (and their ranking if True)
    """
    warehouse.sort_by_length(ranking_effective=ranking_effective)


def sort_by_ratio(warehouse, ranking_effective=False):
    """
    Sorts stacks by their ratio length on width (and their ranking if True)
    """
    warehouse.sort_by_ratio(ranking_effective=ranking_effective)

