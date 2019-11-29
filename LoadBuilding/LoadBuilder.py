"""

Created by Nicolas Raymond on 2019-11-15.

This file provides a LoadBuilder. This object manage all trailer's loading from one plant to another.

Last update : 2019-10-31
By : Nicolas Raymond

"""
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.patches as patches
import LoadingObjects as LoadObj
import pandas as pd
import os
import time
from collections import Counter
from packer import newPacker
from matplotlib.path import Path
from math import floor
from datetime import date


class LoadBuilder:

    def __init__(self, plant_from, plant_to, trailers_data, shipping_date,
                 overhang_authorized=40, maximum_trailer_length=636, plc_lb=0.75):

        """
        :param plant_from: name of the plant from where the item are shipped
        :param plant_to: name of the plant where the item are shipped
        :param trailers_data: Pandas data frame containing details on trailers available
        :param shipping_date: date associated to the shipping of the load that will be built
        :param overhang_authorized: maximum overhanging measure authorized by law for a trailer
        :param maximum_trailer_length: maximum length authorized by law for a trailer
        :param plc_lb: lower bound of length percentage covered that must be satisfied for all trailer
        """
        self.trailers_data = trailers_data
        self.overhang_authorized = overhang_authorized  # In inches
        self.max_trailer_length = maximum_trailer_length  # In inches
        self.plc_lb = plc_lb
        self.plant_from = plant_from
        self.plant_to = plant_to
        self.model_names, self.warehouse, self.remaining_crates = [], LoadObj.Warehouse(), LoadObj.CratesManager()
        self.trailers = []
        self.trailers_done = None
        self.shipping_date = shipping_date
        self.unused_models = []
        self.second_phase_activated = False

    def __len__(self):
        return len(self.trailers_done)

    def __warehouse_init(self, models_data):

        """
        Initializes a warehouse according to the models available in model data

        """

        # For all lines of the data frame
        for i in models_data.index:

            # We save the quantity of the model and the plant_to
            qty = models_data['QTY'][i]
            plant_to = models_data['PLANT_TO'][i]

            if qty > 0 and plant_to == plant_to:

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

                for j in range(nbr_stacks):
                    # We build the stack and send it into the warehouse
                    self.warehouse.add_stack(LoadObj.Stack(max(models_data['LENGTH'][i], models_data['WIDTH'][i]),
                                                           min(models_data['WIDTH'][i], models_data['LENGTH'][i]),
                                                           models_data['HEIGHT'][i] * stack_limit,
                                                           [models_data['MODEL'][i]] * items_per_stack, overhang))

                # We save the number of individual crates to build and convert it into
                # integer to avoid conflict with range function
                nbr_individual_crates = int((qty - (items_per_stack * nbr_stacks)) / nbr_per_crate)

                for j in range(nbr_individual_crates):
                    # We build the crate and send it to the crates manager
                    self.remaining_crates.add_crate(LoadObj.Crate([models_data['MODEL'][i]] * nbr_per_crate,
                                                                  max(models_data['LENGTH'][i],
                                                                      models_data['WIDTH'][i]),
                                                                  min(models_data['WIDTH'][i],
                                                                      models_data['LENGTH'][i]),
                                                                  models_data['HEIGHT'][i],
                                                                  stack_limit, overhang))

        # We flatten the model_names list
        self.model_names = [item for sublist in self.model_names for item in sublist]

    def __trailers_init(self):

        """
        Initializes the list with all the trailers available for the loading

        """

        # For every lines of the data frame
        for i in self.trailers_data.index:

            # We save the quantity, the plant_from and the plant _to
            qty = self.trailers_data['QTY'][i]
            plant_from = self.trailers_data['PLANT_FROM'][i]
            plant_to = self.trailers_data['PLANT_TO'][i]

            if qty > 0 and plant_from == self.plant_from and plant_to == self.plant_to:

                # We save trailer's length
                t_length = self.trailers_data['LENGTH'][i]

                # We compute overhanging measure allowed for the trailer
                if bool(self.trailers_data['OVERHANG'][i]):
                    trailer_oh = min(self.max_trailer_length - t_length, self.overhang_authorized)
                else:
                    trailer_oh = 0

                # We build "qty" trailer that we add to the trailers list
                for j in range(0, qty):
                    self.trailers.append(LoadObj.Trailer(self.trailers_data['CATEGORY'][i], t_length,
                                                         self.trailers_data['WIDTH'][i],
                                                         self.trailers_data['HEIGHT'][i],
                                                         self.trailers_data['PRIORITY_RANK'][i], trailer_oh))

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

        # We save unused models
        self.warehouse.save_unused_crates(self.unused_models)

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
            return self.__max_rect_upperbound(trailer, new_upper_bound)

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

    def __complete_packing(self, trailer, packer, start_index):

        """
        Verifies if one (or multiple) item unconsidered in the first part of packing fits at the end of the trailer

        :param trailer: Object of class Trailer
        :param packer: Packer object
        :param start_index: If i is the index of the last item we considered in first part, then start_index = i + 1
        """

        # We look if there are items remaining in the warehouse (that were not considered in the first phase of packing)
        # and if there's still place in the trailer.
        if len(self.warehouse) != start_index and max([rect.top for rect in packer[0]]) < trailer.length:

            # We save the current number of stack in the trailer
            last_res = len(packer[0])

            # We initialize a new packer with rotation not allowed to simply computation and save time
            new_packer = newPacker(rotation=False)

            # We add rectangles unconsidered in the first phase of packing
            for i in range(start_index, len(self.warehouse)):
                new_packer.add_rect(self.warehouse[i].width, self.warehouse[i].length,
                                    rid=i, overhang=self.warehouse[i].overhang)

            # We add a large number of dummy bins
            for j in range(len(self.warehouse) - start_index + 1):
                new_packer.add_bin(trailer.width, trailer.length, overhang=packer[0].overhang_measure)

            # We open the first bin
            new_packer._open_bins.append(packer[0])

            # We allow unlock rotation in this first bin.
            new_packer[0].rot = True

            # We pack the new packer
            new_packer.pack(reset_opened_bins=False)

            # We send a message if the second phase of packing was effective
            if last_res < len(new_packer[0]):
                print('COMPLETION EFFECTIVE')

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
        # We sort trailer in decreasing order with their number of units
        self.trailers.sort(key=lambda t: t.nbr_of_units(), reverse=True)

        # We initialize an index at the end of the list containing trailers
        i = len(self.trailers) - 1

        # We unload all trailers that exceed maximum number of trailer
        while i >= n:
            self.trailers[i].unload_trailer(self.unused_models)
            self.trailers.pop(i)
            i -= 1

    def __unpack_trailers(self):

        """
        Unpack all trailer while saving their contents

        """
        self.__select_top_n(0)

    def __update_models_data(self):

        """
        Updates quantities in original models data frame

        """

        # We counts all the models that were introduced in loads
        counts = Counter(self.model_names)  # Initial counts
        counts_of_unused = Counter(self.unused_models)
        counts.subtract(counts_of_unused)

        # We update the models data
        for key, item in counts.items():
            row_to_change = self.models_data.index[(self.models_data['MODEL'] == key) &
                                                   (self.models_data['PLANT_TO'] == self.plant_to)].tolist()

            self.models_data.loc[row_to_change[0], 'QTY'] -= item

        # We erase model names in memory
        self.model_names.clear()

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
                                                     (self.trailers_data['WIDTH'] == key[2]) &
                                                     (self.trailers_data['PLANT_FROM'] == self.plant_from) &
                                                     (self.trailers_data['PLANT_TO'] == self.plant_to)].tolist()

            self.trailers_data.loc[row_to_change[0], 'QTY'] -= item

        # We save trailers done and clear the current trailers list
        self.trailers_done = self.trailers.copy()
        self.trailers.clear()

    @staticmethod
    def __print_load(trailer):

        """
        Plots a the loading configuration of the trailer

        :param trailer: Object of class Trailer
        """

        fig, ax = plt.subplots()
        codes = [Path.MOVETO, Path.LINETO, Path.LINETO, Path.LINETO, Path.CLOSEPOLY]
        rect_list = [trailer[i] for i in range(len(trailer))]

        for rect in rect_list:
            vertices = [
                (rect.left, rect.bottom),  # Left, bottom
                (rect.left, rect.top),  # Left, top
                (rect.right, rect.top),  # Right, top
                (rect.right, rect.bottom),  # Right, bottom
                (rect.left, rect.bottom),  # Ignored
            ]

            path = Path(vertices, codes)
            patch = patches.PathPatch(path, facecolor="yellow", lw=2)
            ax.add_patch(patch)

        plt.axis('scaled')
        ax.set_xlim(0, trailer.width)
        ax.set_ylim(0, trailer.height + trailer.overhang_measure)

        if trailer.overhang_measure != 0:
            line = plt.axhline(trailer.height, color='black', ls='--')

        plt.show()
        plt.close()

    def write_summarized_data(self, directory):

        """
        Writes a loading summary in a .xlsx file and create a folder named "P2P - <date>"
        at the directory mentioned and save the file in it.

        :param directory: String that mention path where the created folder is going to be saved

        """

        # We initialize a data frame with column names needed
        data_frame = pd.DataFrame(columns=(["TRAILER", "TRAILER LENGTH", "LOAD LENGTH"] + list(set(self.model_names))))

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
                                [s[model] if s[model] > 0 else '' for model in self.model_names]
            i += 1

        # We add a line for unused model
        unused_summary = Counter(self.unused_models)
        data_frame.loc[i] = ["REMAINING", '', ''] + \
                            [unused_summary[model] if unused_summary[model] > 0 else '' for model in self.model_names]

        # We execute a groupby with trailer in the same category
        data_frame = data_frame.groupby(data_frame.columns.tolist()).size().to_frame('QTY').reset_index()

        # We rearrange columns
        cols = data_frame.columns.tolist()
        cols = cols[0:1] + cols[-1:] + cols[1:-1]
        data_frame = data_frame[cols]

        # We set indexes
        data_frame.set_index("TRAILER", inplace=True)

        # We erase the quantity of the QTY column
        data_frame.loc[["REMAINING"], ['QTY']] = ''

        # We generate today's date
        folder_date = date.today()
        folder_date = str(folder_date.strftime("%m-%d-%y"))

        # We save the folder name
        folder = folder_date + '/'

        # We save the path
        path = directory + folder

        # We create the folder in which the file will be stored (if the folder doesn't exist)
        create_folder(path)

        # We save the complete title of our future file
        title = path + "P2P from " + self.plant_from + " to " + self.plant_to + " " + self.shipping_date + ".xlsx"

        # We initialize a "writer"
        writer = pd.ExcelWriter(title, engine='xlsxwriter')

        # We export results in the file created
        data_frame.to_excel(writer, sheet_name='Loads', index=True)

        # We initialize a workbook
        workbook = writer.book

        # We initialize a worksheet
        worksheet = writer.sheets['Loads']

        # We set the widths of the columns
        worksheet.set_column('A:B', 15)
        worksheet.set_column('B:C', 4.5)
        worksheet.set_column('C:D', 15)
        worksheet.set_column('E:AK', 4.5)

        writer.save()

    def build(self, models_data, max_load, min_load=0, plot_load_done=False):

        """
        This is the core of the object.
        It contains the principal steps of the loading process.

        :param models_data: Pandas data frame containing details on models to load
        :param max_load: maximum number of loads
        :param min_load: minimum number of loads
        :param plot_load_done: boolean that indicates if plots of loads are going to be shown
        :return: list of the models unused and time of execution
        """
        # We save start time
        start_time = time.time()

        # We init the warehouse
        self.__warehouse_init(models_data)

        # We init the list of trailers
        self.__trailers_init()

        # We finish the stacking process with leftover crates
        self.__prepare_warehouse()

        # We execute the loading of the trailers
        self.__trailer_packing(plot_enabled=plot_load_done)

        # We consider the min and the max
        nbr_of_load = len(self.trailers)

        if self.second_phase_activated and min_load > nbr_of_load:

            # We unpack trailers
            self.__unpack_trailers()

        else:
            if max_load < nbr_of_load:
                self.__select_top_n(max_load)

            # We update all data
            self.__update_models_data()
            self.__update_trailers_data()

        # We activate phase 2
        self.second_phase_activated = True

        # Copy the unused models list
        unused_copy = self.unused_models.copy()
        self.unused_models.clear()

        # We save the end of execution time
        end_time = time.time()

        return unused_copy, end_time - start_time


def create_folder(directory):

    """
    Creates a folder with the directory mentioned

    :param directory: String that mention path where the folder is going to be created
    """
    try:
        if not os.path.exists(directory):
            os.makedirs(directory)

    except OSError:
        print('Error while creating directory : ' + directory)

