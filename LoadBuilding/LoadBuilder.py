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
                all_configs = self.create_all_configs(t)

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
                    self.complete_packing(t, packer, len(config))

                    # We save the loading configuration (the packer)
                    packers.append(packer)

                # We save the index of the best loading configuration that respected the constraint of plc_lb
                best_packer_index = self.select_best_packer(packers)

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
                        self.print_load(best_packer[0])

        # We remove trailer that we're not used during the loading process
        self.remove_leftover_trailers()

        # We save unused stacks
        self.warehouse.save_unused_crates()

    def select_best_packer(self, packers_list):

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
            qualified, items_used = self.validate_packing(packer)

            # If the packing respect constraints and has more items than the best one yet,
            # we change our best packer for this one.
            if qualified and items_used > best_nb_items_used:
                best_nb_items_used = items_used
                best_packer_index = i

            i += 1

        return best_packer_index