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


class LoadBuilder:

    score_multiplicator_basis = 1.02  # used to boost the score of a load when there's mandatory crates

    def __init__(self, trailers_data,
                 overhang_authorized=40, maximum_trailer_length=636, plc_lb=0.75):

        """
        :param trailers_data: Pandas data frame containing details on trailers available
        :param overhang_authorized: maximum overhanging measure authorized by law for a trailer
        :param maximum_trailer_length: maximum length authorized by law for a trailer
        :param plc_lb: lower bound of length percentage covered that must be satisfied for all trailer
        """

        self.trailers_data = trailers_data
        self.overhang_authorized = overhang_authorized  # In inches
        self.max_trailer_length = maximum_trailer_length  # In inches
        self.plc_lb = plc_lb
        self.model_names = []
        self.warehouse, self.remaining_crates = LoadObj.Warehouse(), LoadObj.CratesManager('W')
        self.metal_warehouse, self.metal_remaining_crates = LoadObj.Warehouse(), LoadObj.CratesManager('M')
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
        print(models_data)
        tot = 0

        # For all lines of the data frame
        for i in models_data.index:

            # We save the quantity of the model and the plant_to
            qty = models_data['QTY'][i]

            if qty > 0:

                # We save the name of the model
                self.model_names.append([models_data['MODEL'][i]]*qty)

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

                tot += total_of_mandatory

                # We save the number of individual crates to build and convert it into
                # integer to avoid conflict with range function. Also, with int(), every number in [0,1[ will
                # be convert as 0. This way, no individual crate of SP2 will be build if there's less than 2 SP2 left
                nbr_individual_crates = int((qty - (items_per_stack * nbr_stacks)) / nbr_per_crate)

                stacks_component = [max(models_data['LENGTH'][i], models_data['WIDTH'][i]),
                                    min(models_data['WIDTH'][i], models_data['LENGTH'][i]),
                                    models_data['HEIGHT'][i] * stack_limit,
                                    [models_data['MODEL'][i]] * items_per_stack, overhang]

                crates_component = [[models_data['MODEL'][i]] * nbr_per_crate,
                                    max(models_data['LENGTH'][i],
                                    models_data['WIDTH'][i]),
                                    min(models_data['WIDTH'][i], models_data['LENGTH'][i]),
                                    models_data['HEIGHT'][i],
                                    stack_limit, overhang]

                # We select the good type of storage of the stacks and crates that will be build
                if crate_type == 'W':
                    warehouse = self.warehouse
                    crates_manager = self.remaining_crates

                elif crate_type == 'M':
                    warehouse = self.metal_warehouse
                    crates_manager = self.metal_remaining_crates

                # We take the list of ranking associate with the size code or create one filled with 0
                # and start an index to navigate through ranking list
                r = ranking.get(models_data['MODEL'][i], [0]*qty)
                index = 0

                for j in range(nbr_stacks):

                    # We compute the number of mandatory crates in the stack
                    mandatory_crates = min(total_of_mandatory, stack_limit)

                    # We add the missing number of mandatory crates and avg ranking to the stacks component list
                    temp_components = stacks_component + [mandatory_crates] + [np.mean(r[index:(index+stack_limit)])]

                    # We build the stack and send it into the warehouse
                    warehouse.add_stack(LoadObj.Stack(*temp_components))

                    # We update the total number of mandatory left and move our index
                    total_of_mandatory -= mandatory_crates
                    index += stack_limit

                for j in range(nbr_individual_crates):

                    # We add the missing number of mandatory crates to the stacks component list
                    temp_components = crates_component + [total_of_mandatory > 0] + [r[index]]

                    # We build the crate and send it to the crates manager
                    crates_manager.add_crate(LoadObj.Crate(*temp_components))

                    # We update the total number of mandatory left and increment the index
                    total_of_mandatory -= 1
                    index += 1

        # We flatten the model_names list
        self.model_names = [item for sublist in self.model_names for item in sublist]
        print('TOTAL OF MANDATORY :', tot)

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

            if len(self.remaining_crates.stand_by_crates) > 0:
                self.remaining_crates.create_incomplete_stacks(self.warehouse)

        # Finishing stacking for metal crates
        if len(self.metal_remaining_crates.crates) > 0:
            self.metal_remaining_crates.create_stacks(self.metal_warehouse)

            if len(self.metal_remaining_crates.stand_by_crates) > 0:
                self.metal_remaining_crates.create_incomplete_stacks(self.metal_warehouse)

        print('NB OF X IN WOOD STACKS :', sum([stack.nb_of_mandatory for stack in self.warehouse.stacks_to_ship]))
        print('NB OF X IN METAL STACKS :', sum([stack.nb_of_mandatory for stack in self.metal_warehouse.stacks_to_ship]))

    def __trailer_packing(self):
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

        for t in self.trailers:

            # We initialize a list that will contain stacks stored in the trailer, and a list of different
            # containing tuples of crate type and "packer" loading strategies that were tried.
            stacks_used, packers = [], []

            # We try to fill the trailer the best way as possible considering best configurations between
            # wooden warehouse and metal warehouse
            for crate_type, warehouse in [('W', self.warehouse), ('M', self.metal_warehouse)]:

                # We compute all possible configurations of loading (efficiently) if there's still stacks available
                if len(warehouse) != 0:
                    warehouse.sort_by_ranking_and_volume()
                    print(crate_type + 'STACKS POSITIONNING :',
                          [(stack.average_ranking, stack.volume) for stack in warehouse])
                    all_configs = self.__create_all_configs(warehouse, t)

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

                        # We add two other dummy bins to store rectangles that do not enter in our trailer (first bin)
                        for i in range(2):
                            packer.add_bin(t.width, t.length, bid=None, overhang=t.oh)

                        # We execute the packing
                        packer.pack()

                        # We complete the packing (look if some unconsidered rectangles could enter at the end)
                        self.__complete_packing(warehouse, t, packer, len(config))

                        # We save the loading configuration (the packer)
                        packers.append((crate_type, packer))

            # We save the index of the best loading configuration that respected the constraint of plc_lb
            best_packer_index, score = self.__select_best_packer(packers)

            # If an index is found (at least one load satisfies the constraint)
            if best_packer_index is not None:

                # We save the specified packer and the crate_type concerned
                best_packer = packers[best_packer_index][1]
                crate_type = packers[best_packer_index][0]

                # We determine the warehouse concerned with the crate_type
                if crate_type == 'W':
                    warehouse = self.warehouse
                elif crate_type == 'M':
                    warehouse = self.metal_warehouse

                # For every stack concerned by this loading configuration of the trailer
                for stack in best_packer[0]:

                    # We concretely assign the stack to the trailer and note his location index
                    t.add_stack(warehouse[stack.rid])
                    stacks_used.append(stack.rid)

                # We update the length_used of the trailer
                # (using the top of the rectangle that is the most at the edge)
                t.length_used = max([rect.top for rect in best_packer[0]])

                # We set the score associated to the trailer and save the packer object
                t.score = score
                t.packer = best_packer

                # We remove stacks used from the warehouse concerned
                warehouse.remove_stacks(stacks_used)

        # We remove trailer that we're not used during the loading process
        self.__remove_leftover_trailers()

        # We save unused models from both warehouses
        self.warehouse.save_unused_crates(self.unused_models)
        self.metal_warehouse.save_unused_crates(self.unused_models)

    def __select_best_packer(self, packers_list):

        """
        Pick the best loading configuration done (the best packer) among the list according to the number of units
        placed in the trailer.

        :param packers_list: List containing tuples with crate_types and packers object
        :return: Index of the location of the best packer and the best score
        """

        i = 0
        best_packer_index = None
        best_score = 0

        for crate_type, packer in packers_list:

            # We check if packing respect plc lower bound and how many items it contains
            qualified, score = self.__validate_packing(crate_type, packer)

            # If the packing respect constraints and has more items than the best one yet,
            # we change our best packer for this one.
            if qualified and score > best_score:
                best_score = score
                best_packer_index = i

            i += 1

        return best_packer_index, best_score

    def __validate_packing(self, crate_type, packer):

        """
        Verifies if the packing satisfies plc_lb constraint (Lower bound of percentage of length that must be covered)

        :param crate_type: 'W' for wood, 'M' for metal
        :param packer: Packer object
        :returns: Boolean indicating if the loading satisfies constraint and a score for the load
        """

        mandatory_crates = 0
        score = 0
        qualified = True
        trailer = packer[0]

        if max([rect.top for rect in trailer]) / trailer.height < self.plc_lb:
            qualified = False
        else:
            used_area = trailer.used_area()
            if crate_type == 'W':
                mandatory_crates += sum([self.warehouse[rect.rid].nb_of_mandatory for rect in trailer])
                score_boost = self.score_multiplicator_basis**mandatory_crates
                score = used_area*score_boost

            elif crate_type == 'M':
                mandatory_crates += sum([self.metal_warehouse[rect.rid].nb_of_mandatory for rect in trailer])
                score_boost = self.score_multiplicator_basis**mandatory_crates
                score = used_area * score_boost

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
        configs, nb_oversize = warehouse.merge_for_trailer(trailer)
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

    @staticmethod
    def __complete_packing(warehouse, trailer, packer, start_index):

        """
        Verifies if one (or multiple) item unconsidered in the first part of packing fits at the end of the trailer

        :param warehouse: Object of class warehouse
        :param trailer: Object of class Trailer
        :param packer: Packer object
        :param start_index: If i is the index of the last item we considered in first part, then start_index = i + 1
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

            # We send a message if the second phase of packing was effective
            # if last_res < len(new_packer[0]):
            #     print('COMPLETION EFFECTIVE')

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
        Selects the n best trailers in terms of units in the load
        """
        # We sort trailer in decreasing order by their score
        self.trailers.sort(key=lambda t: t.score, reverse=True)

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
        self.all_size_codes.update(size_codes)

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

    def build(self, models_data, max_load, plot_load_done=False, ranking={}):

        """
        This is the core of the object.
        It contains the principal steps of the loading process.

        :param models_data: Pandas data frame containing details on models to load
        :param max_load: maximum number of loads
        :param plot_load_done: boolean that indicates if plots of loads are going to be shown
        :return: list of size code used
        """
        # We look if models_data is empty
        if models_data.empty:
            return []

        # We init the warehouse
        self.__warehouse_init(models_data, ranking)

        # We init the list of trailers
        self.__trailers_init()

        # We finish the stacking process with leftover crates
        self.__prepare_warehouse()

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

