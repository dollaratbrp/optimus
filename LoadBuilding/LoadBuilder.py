"""

Created by Nicolas Raymond on 2019-11-15.

This file provides a LoadBuilder. This object manage all trailer's loading from one plant to another.

Last update : 2019-10-31
By : Nicolas Raymond

"""
import LoadBuilding.LoadingObjects as LoadObj
import numpy as np


class LoadBuilder:

    def __init__(self, plant_from, plant_to, models_data, trailers_data, minimum_trailer, maximum_trailer,
                 overhang_authorized=72, maximum_trailer_length=636):

        """
        :param plant_from: name of the plant from where the item are shipped
        :param plant_to: name of the plant where the item are shipped
        :param models_data: Pandas data frame containing details on models to load
        :param trailers_data: Pandas data frame containing details on trailers available
        :param minimum_trailer: minimum number of trailer
        :param maximum_trailer: maximum number of trailer
        :param overhang_authorized: maximum overhanging measure authorized by law for a trailer
        :param maximum_trailer_length: maximum length authorized by law for a trailer
        """
        self.overhang_authorized = overhang_authorized  # In inches
        self.max_trailer_length = maximum_trailer_length  # In inches
        self.plant_from = plant_from
        self.plant_to = plant_to
        self.model_names, self.warehouse, self.remaining_crates = self.warehouse_ignition(models_data)
        self.trailers = self.trailers_ignition(trailers_data, overhang_authorized, maximum_trailer_length)
        self.minimum_trailer = minimum_trailer
        self.maximum_trailer = maximum_trailer
        self.second_phase_activated = False

    @staticmethod
    def warehouse_ignition(models_data):

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
    def trailers_ignition(trailers_data, overhang_authorized, maximum_trailer_length):

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