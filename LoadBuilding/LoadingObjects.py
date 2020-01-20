"""

Created by Nicolas Raymond on 2019-05-31.

This python file provides all classes of object used during the loading process
(Crate, Stack, Warehouse, Trailer, CratesManager)

Last update : 2019-10-31
By : Nicolas Raymond

"""

import numpy as np
import matplotlib.pyplot as plt
from matplotlib.path import Path
import matplotlib.patches as patches

class Crate:

    """
    Representation of a square-based prism (crate) containing one or many models

    """

    def __init__(self, m_n, l, w, h, s_l, oh, mandatory, ranking):

        """

        :param m_n: list containing names of the models inside the crate
        :param l: length of the crate
        :param w: width of the crate
        :param h: height of the crate
        :param s_l: maximal quantity of the same crate than can be piled one above the other
        :param oh: boolean that specifies if the crate is allowed to exceed trailer's length (overhang)
        :param mandatory: boolean that indicates if the crate is marked as "MANDATORY"
        :param ranking: integer representing the ranking of the crate
        """

        self.model_names = m_n
        self.length = l
        self.width = w
        self.height = h
        self.stack_limit = s_l
        self.overhang = oh
        self.mandatory = mandatory
        self.ranking = ranking

    def __repr__(self):
        return self.model_names


class Stack:

    """
    Representation of a pile of crates

    """

    def __init__(self, l, w, h, models, oh, nb_of_mandatory, average_ranking):

        """

        :param l: length of the crate at the base
        :param w: width of the crate at the base
        :param h: height of the stack
        :param models: list containing names of the models inside the crate
        :param oh: boolean that specifies if the stack is allowed to exceed trailer's length (overhang)
        :param nb_of_mandatory: number of models marked as "MANDATORY" in the stack
        :param average_ranking: average ranking of the boxes inside the stacks
        """
        self.length = l
        self.width = w
        self.height = h
        self.volume = l*w*h
        self.models = models
        self.overhang = oh
        self.nb_of_mandatory = nb_of_mandatory
        self.average_ranking = average_ranking

    def nbr_of_models(self):

        """
        Computes the number of models in the stack (int)
        """
        return len(self.models)

    def better_rotated(self, trailer):

        """
        Checks if it is better to rotate the stack if we are about to put it in the trailer pass as parameter

        :param trailer: Object of class trailer
        :return: boolean specifying if it is better (True) or not (False)
        """

        # We calculate the area of the space that is going to be wasted if the stack is rotated
        lost_space_if_flipped = (trailer.width - self.length) * self.width

        # If the area is positive and smaller than the actual wasted space without rotation
        if (lost_space_if_flipped >= 0) and (lost_space_if_flipped < ((trailer.width - self.width) * self.length)):

            return True

        return False


class Trailer:

    """
    Representation of a square-based prism (trailer) that will contain stacks to ship

    """
    def __init__(self, cat, l, w, h, p, oh):

        """

        :param cat: name of the trailer category (Ex. Flatbed, Drybox)
        :param l: length of the trailer
        :param w: width of the trailer
        :param h: height of the trailer
        :param p: priority level (int) (higher = more important)
        :param oh: overhanging measure allowed
        """
        self.category = cat
        self.length = l
        self.length_used = 0                    # Length used on the trailer
        self.width = w
        self.height = h
        self.load = []                          # List that will contain the stack objects
        self.priority = p
        self.oh = oh
        self.score = 0
        self.crate_type = None
        self.packer = None

    def __repr__(self):
        return self.category

    def plot_load(self):
        """
        Plots the loading configuration of the trailer
        :return: plot
        """
        fig, ax = plt.subplots()
        codes = [Path.MOVETO, Path.LINETO, Path.LINETO, Path.LINETO, Path.CLOSEPOLY]
        rect_list = [self.packer[0][i] for i in range(len(self.packer[0]))]
        print(rect_list)

        for rect in rect_list:
            vertices = [
                (rect.left, rect.bottom),  # Left, bottom
                (rect.left, rect.top),  # Left, top
                (rect.right, rect.top),  # Right, top
                (rect.right, rect.bottom),  # Right, bottom
                (rect.left, rect.bottom),  # Ignored
            ]

            path = Path(vertices, codes)

            if self.crate_type == 'W':
                patch = patches.PathPatch(path, facecolor="brown", lw=2)

            else:  # crate_type == 'M'
                patch = patches.PathPatch(path, facecolor="grey", lw=2)

            ax.add_patch(patch)

        plt.axis('scaled')
        ax.set_xlim(0, self.width)
        ax.set_ylim(0, self.length + self.oh)

        if self.oh != 0:
            line = plt.axhline(self.length, color='black', ls='--')

        plt.show()
        plt.close()

    def area(self):

        """
        Computes the area of the trailer basis
        """
        return self.length*self.width

    def add_load(self, list_of_stacks):

        """
        Adds stacks to the trailer

        :param list_of_stacks: list of stack objects
        """
        self.load += list_of_stacks

    def add_stack(self, stack):

        """
        Adds a single stack to the trailer

        :param stack: stack object
        """
        self.load.append(stack)

    def load_summary(self):

        """
        Returns the quantities of every models in the trailer

        :return: Counter object
        """

        # Initialization of an empty list that will contain all model names in the trailer
        models = []

        for stack in self.load:

            for model in stack.models:
                models.append(model)

        return models

    def unload_trailer(self, unused_crates_list):

        """
        Removes all stacks from the trailer and save the names of the models unused in a list

        :param unused_crates_list: list that will store unused model names
        """

        for stack in self.load:
            unused_crates_list += stack.models

        self.load.clear()

    def nbr_of_units(self):

        """
        Compute the number of individual models in the trailer (int)
        """

        # Initialization of a units counter
        units = 0

        for stack in self.load:
            units += stack.nbr_of_models()

        return units

    def fit(self, stack, rotated=False):

        """
        Verifies if the stack can fit in the trailer. Does not care about the percentage of
        stack that will be on trailer's surface. A more strict verification is done during the loading.
        The current verification is just for a quick check.

        :param stack: stack object
        :param rotated boolean indicating if the stack should be considered rotated
        :return: boolean indicating if it fits or not
        """
        if not rotated:
            return stack.width <= self.width and stack.length <= self.length + int(stack.overhang) * self.oh
        else:
            return stack.length <= self.width and stack.width <= self.length + int(stack.overhang) * self.oh


class Warehouse:

    """
    Representation of a warehouse containing all stacks that we have to ship with trailers available

    """
    def __init__(self):
        self.stacks_to_ship = []  # List containing stacks to ship

    def __getitem__(self, key):
        return self.stacks_to_ship[key]

    def __len__(self):
        return len(self.stacks_to_ship)

    def not_empty(self):

        """
        Checks if the warehouse is empty

        :return : boolean (True if yes, False if no)
        """
        return len(self.stacks_to_ship) != 0

    def clear(self):

        """
        Removes all stacks from the warehouse
        """
        while self.not_empty():
            self.stacks_to_ship.pop()

    def add_stack(self, stack):

        """
        Adds a single stack to the warehouse

        :param stack: stack object
        """
        self.stacks_to_ship.append(stack)

    def remove_stacks(self, indexes):

        """
        Removes all stacks at the positions indicated in the list of indexes

        :param indexes: list of indexes
        """

        # We sort indexes in decreasing order to avoid conflict with stacks positions during the process
        indexes.sort(reverse=True)

        for i in indexes:
            self.stacks_to_ship.pop(i)

    def sort_by_volume(self):

        """
        Sorts stacks to ship by their volumes
        """
        self.stacks_to_ship.sort(key=lambda s: s.volume, reverse=True)

    def sort_by_ranking_and_volume(self):
        """
        Sorts stacks to ship by their average ranking, and their volume if their avg ranking is the same
        """
        self.stacks_to_ship.sort(key=lambda s: (s.average_ranking, -1*s.volume))

    def save_unused_crates(self, unused_crates_list):

        """
        Saves unused model names in the list indicated as parameter

        :param unused_crates_list: list of model names
        """
        for stack in self.stacks_to_ship:
            unused_crates_list += stack.models

        self.stacks_to_ship.clear()

    def merge_for_trailer(self, trailer):

        """
        Places and pre-rotates stacks in the best way possible for the loading, according to the trailer dimensions
        and the priority of the stacks

        :param trailer: trailer object
        :returns : list of booleans indicating if the stack should be seen as rotated during the loading process
                   and number of stacks that cannot fit in the trailer.

        """

        # We save the length of the longest element in the warehouse
        unique_tuples = set((stack.width, stack.length) for stack in self.stacks_to_ship)
        longest_item_length = max(max(dimensions) for dimensions in unique_tuples)

        # Initialization of a new warehouse
        new = Warehouse()

        # Initialization of a list of boolean indicating if the stacks are rotated or not
        config = []

        # Initialization of list that will contain index of stack that cannot fit in the trailer
        leftover = []

        # Initialization of an index that will go through the warehouse and a load_length variable
        # that approximate the length of the load during the loading process
        i, load_length = 0, 0

        # While the approximate difference between simulated load_length and trailer's length
        # is shorter than the longest item in the warehouse
        while i < len(self) and trailer.length - load_length > longest_item_length:

            # If the current stack can fit in the trailer
            if trailer.fit(self[i]) or trailer.fit(self[i], rotated=True):

                # We add the stack at index i in the new version of warehouse
                new.add_stack(self[i])

                # We initialize an index j and a boolean specifying if the merge operation was successful or not
                j, merged = i+1, False

                while j < len(self) and not merged:

                    # If the width of the next stack is smaller than the difference between trailer's width
                    # and current stack's width
                    if self[j].width <= trailer.width - self[i].width:

                        load_length += max(self[i].length, self[j].length)  # We update load_length
                        new.add_stack(self.stacks_to_ship.pop(j))  # We add the stack in the new warehouse
                        config += [False, False]  # We specify that both of them were not rotated
                        merged = True

                    j += 1

                # If no stack could fit side by side with the current stack
                if not merged:

                    # We compute if the current stack should be rotated or not
                    if self[i].better_rotated(trailer):
                        config += [True]
                        load_length += self[i].width
                    else:
                        config += [False]
                        load_length += self[i].length
            else:
                leftover.append(i)

            i += 1

        # We add the rest of stacks (that don't need any positioning verification) in the new warehouse
        while i < len(self):
            new.add_stack(self[i])
            i += 1

        # We add the stack that could not fit in the trailer at the end of the warehouse
        if len(leftover) > 0:
            for j in leftover:
                new.add_stack(self[j])

        # We clear the actual warehouse by precaution
        self.clear()

        # We introduce back our stacks in the new order
        self.stacks_to_ship = new.stacks_to_ship

        # We return the configuration list with booleans
        return config, len(leftover)


class CratesManager:

    """
    Representation of another kind of smaller warehouse where we manage leftover individual crates
    that were not stacked and introduced in the main warehouse at first.
    """

    def __init__(self, crates_type):

        self.crates = []                # List of individual crates (used in create_stacks function)
        self.stand_by_crates = []       # List of individual crates (used in create_incomplete_stacks function)
        self.crates_type = crates_type  # 'W' for wood, "M" for metal

    def add_crate(self, crate):

        """
        Add single crate to the list of crates

        :param crate: crate object
        """
        self.crates.append(crate)

    def remove_crates(self, indexes, option=1):

        """
        Remove all crates at locations specified in indexes list.
        NOTE : There's no need of sort indexes in decreasing order before the process because all indexes are
               one after the other at the beginning of the list specified with option parameter

        :param indexes: list of indexes
        :param option: indicates on which list to remove crates
        """

        if option == 1:
            for i in indexes:
                self.crates.pop(0)

        elif option == 2:
            for i in indexes:
                self.stand_by_crates.pop(0)

    def create_stacks(self, warehouse):

        """
        Build stacks with individual leftover crates

        :param warehouse: warehouse object (the one where the stack built will be stored)
        """

        # We first sort crates in decreasing order by width and then in decreasing order by length
        self.crates.sort(key=lambda s: (s.width, s.length), reverse=True)

        # Initialization of the number of crates needed to build a complete stack
        crates_needed = self.crates[0].stack_limit

        # While there is enough crates left to build a complete stack with the crate at the beginning of the list
        while len(self.crates) >= crates_needed:

            # We initialize a boolean indicating if it's possible to build a stack
            stacking_available = True

            # For all next crates needed
            for crate in [self.crates[i] for i in range(1, crates_needed)]:

                # We apply stacking rules of the crates type considered
                if crate.width != self.crates[0].width:
                    stacking_available = False

                    # We stop the process for this stack to avoid waste of time
                    break

                if self.crates_type == 'M' and crate.length != self.crates[0].length:
                    stacking_available = False

                    # We stop the process for this stack to avoid waste of time
                    break

            if stacking_available:

                # We save indexes of stacks that we will use
                index_list = range(0, crates_needed)

                # We save the crates that we will use
                crates_list = [self.crates[i] for i in index_list]

                # We build the stack and send it to the warehouse
                stack_height = np.sum([crate.height for crate in crates_list])
                stack_models = sum([crate.model_names for crate in crates_list], [])
                nb_of_mandatory = sum([crate.mandatory for crate in crates_list])
                stack_avg_rank = np.mean([crate.ranking for crate in crates_list])
                warehouse.add_stack(Stack(self.crates[0].length, self.crates[0].width, stack_height,
                                          stack_models, not self.crates[0].overhang, nb_of_mandatory, stack_avg_rank))

                # We remove crates from the list
                self.remove_crates(index_list)

            else:

                # We put the crate in stand by for a second phase of stack building
                self.stand_by_crates.append(self.crates[0])

                # We remove the stack from the list
                self.crates.pop(0)

            # Finally we reset the number of crates needed to build the next complete stack
            # if there is still crates available
            if len(self.crates) > 0:
                crates_needed = self.crates[0].stack_limit

        # We place leftover crates from the first phase in stand by for the second phase
        while len(self.crates) > 0:

            self.stand_by_crates.append(self.crates[0])
            self.crates.pop(0)

    def create_incomplete_stacks(self, warehouse):

        """
        Build stacks that will not be as high as a complete stack could be
        (they will not be to their full potential)

        :param warehouse: warehouse object were the stacks will be stored
        """

        # We first sort crates in decreasing order by width and then in decreasing order by length
        self.stand_by_crates.sort(key=lambda s: (s.width, s.length), reverse=True)

        while len(self.stand_by_crates) > 0:

            # We save the number of crates that would be optimal to build a complete stack
            crates_wanted = self.stand_by_crates[0].stack_limit

            # We initialize a list with the models that will be use to build the stack
            model_list = self.stand_by_crates[0].model_names

            # We initialize the height of the stack that we are building
            height = self.stand_by_crates[0].height

            # We initialize the nb of mandatory crates of the stack that we are building and the sum of rank
            nb_of_mandatory = int(self.stand_by_crates[0].mandatory)
            ranking_list = [self.stand_by_crates[0].ranking]

            # We initialize a counter of crates used
            crates_used = 1

            # We search in a range limited by the number of crates wanted and the number of crates available
            for crate in [self.stand_by_crates[i] for i in range(1, min(len(self.stand_by_crates), crates_wanted))]:

                if crate.width == self.stand_by_crates[0].width:
                    if self.crates_type == 'W' or (self.crates_type == 'M' and
                                                   crate.length == self.stand_by_crates[0].length):

                        model_list += crate.model_names
                        height += crate.height
                        nb_of_mandatory += crate.mandatory
                        ranking_list += [crate.ranking]
                        crates_used += 1

                # We stop the process if the next crate can't be on top of the actual one
                else:
                    break

            # We build the stack and send it to the specified warehouse
            avg_ranking = np.mean(ranking_list)
            warehouse.add_stack(Stack(self.stand_by_crates[0].length, self.stand_by_crates[0].width, height,
                                      model_list, self.stand_by_crates[0].overhang, nb_of_mandatory, avg_ranking))

            # We remove the crates from the stand by crates list
            self.remove_crates(range(crates_used), option=2)
