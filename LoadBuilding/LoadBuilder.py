"""

Created by Nicolas Raymond on 2019-11-15.

This file provides a LoadBuilder. This object manage all trailer's loading from one plant to another.

Last update : 2019-10-31
By : Nicolas Raymond

"""
import LoadBuilding.LoadingObjects as LoadObj
import numpy as np
from LoadBuilding.packer import newPacker


class LoadBuilder:

    def __init__(self, plant_from, plant_to, models_data, trailers_data, minimum_trailer, maximum_trailer,
                 overhang_authorized=72, maximum_trailer_length=636, plc_lb=0.75):

        """
        :param plant_from: name of the plant from where the item are shipped
        :param plant_to: name of the plant where the item are shipped
        :param models_data: Pandas data frame containing details on models to load
        :param trailers_data: Pandas data frame containing details on trailers available
        :param minimum_trailer: minimum number of trailer
        :param maximum_trailer: maximum number of trailer
        :param overhang_authorized: maximum overhanging measure authorized by law for a trailer
        :param maximum_trailer_length: maximum length authorized by law for a trailer
        :param plc_lb: lower bound of length percentage covered that must be satisfied for all trailer
        """
        self.overhang_authorized = overhang_authorized  # In inches
        self.max_trailer_length = maximum_trailer_length  # In inches
        self.plc_lb = plc_lb
        self.plant_from = plant_from
        self.plant_to = plant_to
        self.model_names, self.warehouse, self.remaining_crates = self.warehouse_init(models_data)
        self.trailers = self.trailers_init(trailers_data, overhang_authorized, maximum_trailer_length)
        self.minimum_trailer = minimum_trailer
        self.maximum_trailer = maximum_trailer
        self.unused_models = []
        self.second_phase_activated = False

    @staticmethod
    def warehouse_init(models_data):

        """
        Initializes a warehouse according to the models available in model data

        :param models_data: Pandas data frame containing details on models to load
        :return: List of all the model names, Warehouse with the stacks created and a CratesManager with leftover crates
        """
        # Creation of a warehouse and a crate manager that will be returned by the function
        warehouse = LoadObj.Warehouse()
        remaining_crates = LoadObj.CratesManager()

        # Creation of a list that will contain model names that will also be returned
        model_names = []

        # For all lines of the data frame
        for i in models_data.index:

            # We save the quantity of the model
            qty = models_data['QTY'][i]

            if qty > 0:

                # We save the name of the model
                model_names.append(models_data['MODEL'][i])

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

                for j in range(nbr_stacks):
                    # We build the stack and send it into the warehouse
                    warehouse.add_stack(LoadObj.Stack(max(models_data['LENGTH'][i], models_data['WIDTH'][i]),
                                                      min(models_data['WIDTH'][i], models_data['LENGTH'][i]),
                                                      models_data['HEIGHT'][i] * stack_limit,
                                                      [models_data['MODEL'][i]] * items_per_stack, overhang))

                # We save the number of individual crates to build and convert it into
                # integer to avoid conflict with range function
                nbr_individual_crates = int((qty - (items_per_stack * nbr_stacks)) / nbr_per_crate)

                for j in range(nbr_individual_crates):
                    # We build the crate and send it to the crates manager
                    remaining_crates.add_crate(LoadObj.Crate([models_data['MODEL'][i]] * nbr_per_crate,
                                                             max(models_data['LENGTH'][i],
                                                             models_data['WIDTH'][i]),
                                                             min(models_data['WIDTH'][i],
                                                             models_data['LENGTH'][i]),
                                                             models_data['HEIGHT'][i], stack_limit, overhang))

        return model_names, warehouse, remaining_crates

    @staticmethod
    def trailers_init(trailers_data, overhang_authorized, maximum_trailer_length):

        """
        Initializes a list with all the trailers available for the loading

        :param trailers_data: Pandas data frame containing details on trailers available
        :param overhang_authorized: maximum overhanging measure authorized by law for a trailer
        :param maximum_trailer_length: maximum length authorized by law for a trailer
        :return:List with all the trailers
        """
        # We initialize a list of trailers
        trailers = []

        # For every lines of the data frame
        for i in trailers_data.index:

            # We save the quantity
            qty = trailers_data['QTY'][i]

            if qty > 0:

                # We save trailer's length
                t_length = trailers_data['LENGTH'][i]

                # We compute overhanging measure allowed for the trailer
                trailer_oh = min(maximum_trailer_length - t_length, overhang_authorized )

                # We build "qty" trailer that we add to the trailers list
                for j in range(0, qty):
                    trailers.append(LoadObj.Trailer(trailers_data['CATEGORY'][i], t_length,
                                                    trailers_data['WIDTH'][i], trailers_data['HEIGHT'][i],
                                                    trailers_data['PRIORITY_RANK'][i], trailer_oh))
        return trailers

    def __prepare_warehouse(self):

        """
        Finishes stacking procedure with crates that weren't stacked at first
        """

        if len(self.remaining_crates.crates) > 0:
            self.remaining_crates.create_stacks(self.warehouse)

            if len(self.remaining_crates.stand_by_crates) > 0:
                self.remaining_crates.create_incomplete_stacks(self.warehouse)

    def __trailer_packing(self, plot_enabled=False):

        """

        Using a modified version of Skyline 2D bin packing algorithms provided by the rectpack library,
        this function performs trailer loading by considering the bottom surface of every stack
        as a rectangle that needs to be placed in a bin.

        Check https://github.com/secnot/rectpack for more informations on source code and
        http://citeseerx.ist.psu.edu/viewdoc/download;jsessionid=3A00D5E102A95EF7C941408817666342?doi=10.1.1.695.2918&rep=rep1&type=pdf
        for more information on algorithms implemented themselves.

        :param plot_enabled: Bool indicating if plotting is enable to visualize every load

        """

        # We sort trailer by the area of their surface
        self.trailers.sort(key=lambda s: (s.priority, s.area()), reverse=True)

        for t in self.trailers:

            # We initialize a list that will contain stacks stored in the trailer, and a list of different
            # "packer" loading strategies that were tried.
            stacks_used, packers = [], []

            # We compute all possible configurations of loading (efficiently) if there's still stacks available
            if len(self.warehouse) != 0:
                self.warehouse.sort_by_volume()
                all_configs = self.__create_all_configs(t)

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
                            packer.add_rect(self.warehouse[i].length, self.warehouse[i].width, rid=i,
                                            overhang=self.warehouse[i].overhang)

                        else:
                            packer.add_rect(self.warehouse[i].width, self.warehouse[i].length, rid=i,
                                            overhang=self.warehouse[i].overhang)

                    # We add two other dummy bins to store rectangles that do not enter in our trailer (first bin)
                    for i in range(2):
                        packer.add_bin(t.width, t.length, bid=None, overhang=t.oh)

                    # We execute the packing
                    packer.pack()

                    # We complete the packing (look if some unconsidered rectangles could enter at the end)
                    self.__complete_packing(t, packer, len(config))

                    # We save the loading configuration (the packer)
                    packers.append(packer)

                # We save the index of the best loading configuration that respected the constraint of plc_lb
                best_packer_index = self.__select_best_packer(packers)

                # If an index is found (at least one load satisfies the constraint)
                if best_packer_index is not None:

                    # We save the specified packer
                    best_packer = packers[best_packer_index]

                    # For every stack concerned by this loading configuration of the trailer
                    for stack in best_packer[0]:
                        # We concretely assign the stack to the trailer and note his location (index) in the warehouse
                        t.add_stack(self.warehouse[stack.rid])
                        stacks_used.append(stack.rid)

                    # We update the length_used of the trailer
                    # (using the top of the rectangle that is the most at the edge)
                    t.length_used = max([rect.top for rect in best_packer[0]])

                    # We remove stacks used from the warehouse
                    self.warehouse.remove_stacks(stacks_used)

                    # We print the loading configuration of the trailer to visualize the result
                    if plot_enabled:
                        self.__print_load(best_packer[0])

        # We remove trailer that we're not used during the loading process
        self.__remove_leftover_trailers()

        # We save unused stacks
        self.warehouse.save_unused_crates()

    def __select_best_packer(self, packers_list):

        """
        Pick the best loading configuration done (the best packer) among the list according to the number of units
        placed in the trailer.

        :param packers_list: List containing packers object
        :return: Index of the location of the best packer
        """

        i = 0
        best_packer_index = None
        best_nb_items_used = 0

        for packer in packers_list:

            # We check if packing respect plc lower bound and how many items it contains
            qualified, items_used = self.__validate_packing(packer)

            # If the packing respect constraints and has more items than the best one yet,
            # we change our best packer for this one.
            if qualified and items_used > best_nb_items_used:
                best_nb_items_used = items_used
                best_packer_index = i

            i += 1

        return best_packer_index

    def __validate_packing(self, packer):

        """
        Verifies if the packing satisfies plc_lb constraint (Lower bound of percentage of length that must be covered)

        :param packer: Packer object
        :returns: Boolean indicating if the loading satisfies constraint and number of units in the load
        """

        items_used = 0
        qualified = True
        trailer = packer[0]

        if max([rect.top for rect in trailer]) / trailer.height < self.plc_lb:
            qualified = False
        else:
            items_used += sum([self.warehouse[rect.rid].nbr_of_models() for rect in trailer])

        return qualified, items_used

    def __max_rect_upperbound(self, trailer, last_upper_bound):

        """
        Recursive function that approximates a maximum number of rectangle that can fit in the trailer,
        according to rectangles available that are going to enter in the trailer.

        :param trailer: Object of class Trailer
        :param last_upper_bound: Last upper bound found (int)
        :return: Approximation of the maximal number (int)
        """

        # We build a set containing all pairs of objects' width and length in a certain range in the warehouse
        unique_tuples = set((self.warehouse[i].width, self.warehouse[i].length)
                            for i in range(min(last_upper_bound, len(self.warehouse))))

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
                length_if_rotated = item_width / (np.floor(trailer.width / item_length))
                fit = True

            # If the item fits with the original positioning
            if item_length <= trailer.length and item_width <= trailer.width:
                # We save is length divided by the number of times it fits side by side
                length = item_length / (np.floor(trailer.width / item_width))
                fit = True

            # We update min_shortest_length value
            if fit:
                lengths_list = [l for l in [length_if_rotated, length] if l is not None]
                shortest_length = min([shortest_length] + lengths_list)

        # We compute the upper bound
        new_upper_bound = np.floor((trailer.length + trailer.oh) / shortest_length)

        # If the upper bound found equals the upper bound found in the last iteration we stop the process
        if new_upper_bound == last_upper_bound:
            return new_upper_bound

        else:
            return self.max_rect_upperbound(trailer, new_upper_bound)

    def __create_all_configs(self, trailer):

        """
        Creates all configurations of loading that can be done for the trailer. To avoid considering
        a large number of bad configurations and enhance the efficiency of the algorithm, we will pre-set
        wisely the positions of rectangles for a certain range of the trailer and THEN consider
        all possible configurations for the end of the loading of the trailer.

        :param trailer: Object of class Trailer
        :return: List of list of boolean indicating permission of rotation
        """

        # We initialize a numpy array of configurations (only one configuration for now)
        configs, nb_oversize = self.warehouse.merge_for_trailer(trailer)
        configs = np.array([configs])

        # We compute an upper bound for the maximal number of rectangles that can fit in our trailer
        ub = self.__max_rect_upperbound(trailer, len(self.warehouse) - nb_oversize)

        # We save the number of stack pre-rotated
        nb_of_pre_rotated = configs.shape[1]

        # Initialization of list that will contain index of stack that cannot fit in the trailer
        leftover = []

        # We set the start index and the end index of research to build possible configurations
        i = nb_of_pre_rotated
        end_index = min(len(self.warehouse), ub)

        while i < end_index:

            # We pretend that the item doesn't fit in the trailer
            fit = False

            # We initialize an empty numpy array of new configurations
            new_configs = np.array([[]])

            # If it's possible to rotate the i-th item in the warehouse for this trailer
            if trailer.fit(self.warehouse[i], rotated=True):
                # We add the rotated rectangle indicator (True) to all configs found until now
                true_vec = [[True]] * len(configs)
                new_configs = np.append(np.copy(configs), true_vec, axis=1)
                fit = True

            # If the i-th item fit not rotated in this trailer
            if trailer.fit(self.warehouse[i]):
                # We add the rotated rectangle indicator (False) to all configs found until now
                false_vec = [[False]] * len(configs)
                configs = np.append(configs, false_vec, axis=1)
                fit = True

            if not fit:

                # We add the stack index in the leftover list
                leftover.append(i)

                # (If possible) We extend our research space cause we didn't enter the current stack in our configs
                end_index = min(len(self.warehouse), end_index + 1)

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
                self.warehouse.add_stack(self.warehouse[j])
            self.warehouse.remove_stacks(leftover)

        return configs