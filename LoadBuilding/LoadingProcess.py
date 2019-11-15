"""

Created by Nicolas Raymond on 2019-09-03.

This file contains all functions linked to the loading procedures.
The loading is based of a modified version skyline 2D packing algorithm from rectpack library.
The modified version allows overhang.

Last update : 2019-11-01
By : Nicolas Raymond

"""

import matplotlib.pyplot as plt
from matplotlib.path import Path
import matplotlib.patches as patches
from math import floor
import numpy as np
from packer import newPacker


def prepare_warehouse(warehouse, remaining_crates):

    """
    Finishes stacking procedure with crates that weren't stacked at first

    :param warehouse: Object of class warehouse
    :param remaining_crates: Object of class Crates_Manager
    """

    if len(remaining_crates.crates) > 0:
        remaining_crates.create_stacks(warehouse)

        if len(remaining_crates.stand_by_crates) > 0:
            remaining_crates.create_incomplete_stacks(warehouse)


def trailer_packing(warehouse, trailers, unused_list, plc_lb, plot_enabled=False):

    """

    Using a modified version of Skyline 2D bin packing algorithms provided by the rectpack library,
    this function performs trailer loading by considering the bottom surface of every stack
     as a rectangle that needs to be placed in a bin.

    Check https://github.com/secnot/rectpack for more informations on source code and
    http://citeseerx.ist.psu.edu/viewdoc/download;jsessionid=3A00D5E102A95EF7C941408817666342?doi=10.1.1.695.2918&rep=rep1&type=pdf
    for more information on algorithms implemented themselves.

    :param warehouse: Object of class warehouse
    :param trailers: List that contains objects of class Trailer
    :param unused_list: List of unused crates during loading process
    :param plc_lb: Lower bound of percentage of length covered that must be satisfied by all loading
    :param plot_enabled: Bool indicating if plotting is enable to visualize every load

    """

    # We sort trailer by the area of their surface
    trailers.sort(key=lambda s: (s.priority, s.area()), reverse=True)

    for t in trailers:

        # We initialize a list that will contain stacks stored in the trailer, and a list of different
        # "packer" loading strategies that were tried.
        stacks_used, packers = [], []

        # We compute all possible configurations of loading (efficiently) if there's still stacks available
        if len(warehouse) != 0:
            warehouse.sort_by_volume()
            all_configs = create_all_configs(warehouse, t)

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
                        packer.add_rect(warehouse[i].length, warehouse[i].width, rid=i, overhang=warehouse[i].overhang)

                    else:
                        packer.add_rect(warehouse[i].width, warehouse[i].length, rid=i, overhang=warehouse[i].overhang)

                # We two other dummy bins to store rectangles that do not enter in our trailer (first bin)
                for i in range(2):
                    packer.add_bin(t.width, t.length, bid=None, overhang=t.oh)

                # We execute the packing
                packer.pack()

                # We complete the packing (look if some unconsidered rectangles could enter at the end)
                complete_packing(warehouse, t, packer, len(config))

                # We save the loading configuration (the packer)
                packers.append(packer)

            # We save the index of the best loading configuration that respected the constraint of plc_lb
            best_packer_index = select_best_packer(warehouse, packers, plc_lb)

            # If an index is found (at least one load satisfies the constraint)
            if best_packer_index is not None:

                # We save the specified packer
                best_packer = packers[best_packer_index]

                # For every stack concerned by this loading configuration of the trailer
                for stack in best_packer[0]:

                    # We concretely assign the stack to the trailer and note his location (index) in the warehouse
                    t.add_stack(warehouse[stack.rid])
                    stacks_used.append(stack.rid)

                # We update the length_used of the trailer (using the top of the rectangle that is the most at the edge)
                t.length_used = max([rect.top for rect in best_packer[0]])

                # We remove stacks used from the warehouse
                warehouse.remove_stacks(stacks_used)

                # We print the loading configuration of the trailer to visualize the result
                if plot_enabled:
                    print_load(best_packer[0])

    # We remove trailer that we're not used during the loading process
    remove_leftover_trailers(trailers)

    # We save unused stacks
    warehouse.save_unused_crates(unused_list)


def select_best_packer(warehouse, packers_list, plc_lb):

    """
    Pick the best loading configuration done (the best packer) among the list according to the number of units
    placed in the trailer.

    :param warehouse: Object of class warehouse
    :param packers_list: List containing packers object
    :param plc_lb: (float) Constraint on the lower bound of percentage of trailer's length that must be covered
    :return: Index of the location of the best packer
    """

    i = 0
    best_packer_index = None
    best_nb_items_used = 0

    for packer in packers_list:

        # We check if packing respect plc lower bound and how many items it contains
        qualified, items_used = validate_packing(warehouse, packer, plc_lb)

        # If the packing respect constraints and has more items than the best one yet,
        # we change our best packer for this one.
        if qualified and items_used > best_nb_items_used:

            best_nb_items_used = items_used
            best_packer_index = i

        i += 1
    
    return best_packer_index


def validate_packing(warehouse, packer, plc_lb):

    """
    Verifies if the packing satisfies plc_lb constraint (Lower bound of percentage of length that must be covered)

    :param warehouse: Object of class warehouse
    :param packer: Packer object
    :param plc_lb: Float
    :returns: Boolean indicating if the loading satisfies constraint and number of units in the load
    """

    items_used = 0
    qualified = True
    trailer = packer[0]

    if max([rect.top for rect in trailer])/trailer.height < plc_lb:
        qualified = False
    else:
        items_used += sum([warehouse[rect.rid].nbr_of_models() for rect in trailer])

    return qualified, items_used


def max_rect_upperbound(warehouse, trailer, last_upper_bound):

    """
    Recursive function that approximates a maximum number of rectangle that can fit in the trailer,
    according to rectangles available that are going to enter in the trailer.

    :param warehouse: Object of class warehouse
    :param trailer: Object of class Trailer
    :param last_upper_bound: Last upper bound found (int)
    :return: Approximation of the maximal number (int)
    """

    # We build a set containing all pairs of objects' width and length in a certain range in the warehouse
    unique_tuples = set((warehouse[i].width, warehouse[i].length) for i in range(min(last_upper_bound, len(warehouse))))

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
            length_if_rotated = item_width/(floor(trailer.width/item_length))
            fit = True

        # If the item fits with the original positioning
        if item_length <= trailer.length and item_width <= trailer.width:

            # We save is length divided by the number of times it fits side by side
            length = item_length/(floor(trailer.width/item_width))
            fit = True

        # We update min_shortest_length value
        if fit:
            lengths_list = [l for l in [length_if_rotated, length] if l is not None]
            shortest_length = min([shortest_length] + lengths_list)

    # We compute the upper bound
    new_upper_bound = floor((trailer.length + trailer.oh)/shortest_length)

    # If the upper bound found equals the upper bound found in the last iteration we stop the process
    if new_upper_bound == last_upper_bound:
        return new_upper_bound

    else:
        return max_rect_upperbound(warehouse, trailer, new_upper_bound)


def create_all_configs(warehouse, trailer):

    """
    Creates all configurations of loading that can be done for the trailer. To avoid considering
    a large number of bad configurations and enhance the efficiency of the algorithm, we will pre-set
    wisely the positions of rectangles for a certain range of the trailer and THEN consider
    all possible configurations for the end of the loading of the trailer.

    :param warehouse: Object of class warehouse
    :param trailer: Object of class Trailer
    :return: List of list of boolean indicating permission of rotation
    """

    # We initialize a numpy array of configurations (only one configuration for now)
    configs, nb_oversize = warehouse.merge_for_trailer(trailer)
    configs = np.array([configs])
   
    # We compute an upper bound for the maximal number of rectangles that can fit in our trailer
    ub = max_rect_upperbound(warehouse, trailer, len(warehouse) - nb_oversize)

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
            true_vec = [[True]]*len(configs)
            new_configs = np.append(np.copy(configs), true_vec, axis=1)
            fit = True

        # If the i-th item fit not rotated in this trailer
        if trailer.fit(warehouse[i]):

            # We add the rotated rectangle indicator (False) to all configs found until now
            false_vec = [[False]]*len(configs)
            configs = np.append(configs, false_vec, axis=1)
            fit = True

        if not fit:

            # We add the stack index in the leftover list
            leftover.append(i)

            # (If possible) We extend our research space cause we didn't enter the current stack in our configs
            end_index = min(len(warehouse), end_index+1)

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


def complete_packing(warehouse, trailer, packer, start_index):

    """
    Verify if one (or multiple) item unconsidered in the first phase of packing fits at the end of the trailer

    :param warehouse: Object of class warehouse
    :param trailer: Object of class Trailer
    :param packer: Packer object
    :param start_index: If i is the index of the last item we considered in first phase, then start_index = i + 1
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
        for j in range(len(warehouse)-start_index+1):
            new_packer.add_bin(trailer.width, trailer.length, overhang=packer[0].overhang_measure)

        # We open the first bin
        new_packer._open_bins.append(packer[0]) 
       
        # We allow unlock rotation in this first bin.
        new_packer[0].rot = True

        # We pack the new packer
        new_packer.pack(reset_opened_bins=False)

        # We send a message if the second phase of packing was effective
        #if last_res < len(new_packer[0]):
            #print('COMPLETION EFFECTIVE')


def remove_leftover_trailers(trailers):

    """
    Remove trailers that contain no units after the loading

    :param trailers: List of trailers
    """

    # We initialize an index at the end of the list containing trailers
    i = len(trailers) - 1

    # We remove trailer that we're not used during the loading process
    while i >= 0:
        if trailers[i].nbr_of_units() == 0:
            trailers.pop(i)
        i -= 1


def print_load(trailer):

    """
    Plot a the loading configuration of the trailer

    :param trailer: Object of class Trailer
    """

    fig, ax = plt.subplots()
    codes = [Path.MOVETO, Path.LINETO, Path.LINETO, Path.LINETO, Path.CLOSEPOLY]
    rect_list = [trailer[i] for i in range(len(trailer))]

    for rect in rect_list:
        
        vertices = [
            (rect.left, rect.bottom),   # Left, bottom
            (rect.left, rect.top),      # Left, top
            (rect.right, rect.top),     # Right, top
            (rect.right, rect.bottom),  # Right, bottom
            (rect.left, rect.bottom),   # Ignored
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


def select_top_n(trailers, unused, n):

    """
    Selects the n best trailers in terms of units in the load

    :param trailers: List that contains objects of class Trailer
    :param unused: List of string that specifies names of models unused
    :param n: Max number of trailers to be considered

    """

    # If n is smaller then the number of trailers that we loaded
    if n < len(trailers):

        # We sort trailer in decreasing order with their number of units
        trailers.sort(key=lambda t: t.nbr_of_units(), reverse=True)

        # We initialize an index at the end of the list containing trailers
        i = len(trailers)-1

        # On unload toutes les trailers qui ne font pas parti de notre top n et on les retire
        while i >= n:
            trailers[i].unload_trailer(unused)
            trailers.pop(i)
            i -= 1


"""

       *** LOADING MAIN PROGRAM ***
        
"""


def solve(warehouse, trailers, remaining_crates, unused, plc_lb):

    """
    This is the core of the LoadingProcess.py file.
    It contains the two principal steps of the loading process.

    :param warehouse: Object of class Warehouse
    :param trailers: List of trailers
    :param remaining_crates: List of crates
    :param unused: List with the names of unused models
    :param plc_lb: Lower bound of percentage of length covered that must be satisfied by all loading

    """

    # We finish stacking process with leftover crates
    prepare_warehouse(warehouse, remaining_crates)

    # We execute the loading of the trailers
    trailer_packing(warehouse, trailers, unused, plc_lb)





