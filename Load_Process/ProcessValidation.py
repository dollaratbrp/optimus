"""
This file manages activities related to process tests and results validation

Author : Nicolas Raymond
"""

from tkinter import *
from openpyxl import load_workbook
from FastLoads import build_dataframe
from tkinter import messagebox
from P2PFunctions import *


class ApprovedItem:

    """Contains all datas related to a line of "APPROVED" worksheet of the P2P_Summary file"""

    def __init__(self, point_from, shipping_point, material_number, qty, days_to):

        """
        :param point_from: point where the item comes from (Ex : 4100)
        :param shipping_point: point where the item is going to be shipped (Ex : 4015)
        :param material_number: material number (SKU) of the item
        :param qty: number of times this SKU is shipped
        :param days_to: number of days until the item will be shipped
        """

        self.point_from = point_from
        self.shipping_point = shipping_point
        self.material_number = material_number
        self.qty = qty
        self.days_to = days_to

    def find_associated_wish(self, wishlist):

        """
        Look through the wishlist to see if the approved item was really in it!
        If the wish is found, then the wish is deleted from the wishlist

        :param wishlist: List of WishListObj
        :return : boolean indicating if the wish was found
        """

        for index, wish in enumerate(wishlist):
            if self.point_from == wish.POINT_FROM and self.shipping_point == wish.SHIPPING_POINT and \
                    self.material_number == wish.MATERIAL_NUMBER:
                wishlist.pop(index)
                return True

        return False

    def find_associated_inventory(self, inventory_list):

        """
        Look through list of INVobj to approve if the inventory was available for the item
        :param inventory_list: list of INVobj objects
        :return: boolean indicating if the inventory was available
        """

        for index, inventory_line in enumerate(inventory_list):

            if EquivalentPlantFrom(inventory_line.POINT, self.point_from) and \
                    self.material_number == inventory_line.MATERIAL_NUMBER and inventory_line.QUANTITY >= self.qty and\
                    (not inventory_line.Future or (inventory_line.Future and self.days_to > 0)):

                inventory_line.QUANTITY -= self.qty
                print('POINT FROM :', inventory_line.POINT)
                print('INVENTORY TAKEN :', self.qty)
                print('INVENTORY LEFT : ', inventory_line.QUANTITY, '\n')

                if inventory_line.QUANTITY == 0:
                    inventory_list.pop(index)

                return True

        return False


def warning(option):

    """
    Shows a warning message according to the options

    :param option: 'wish','inv' or 'p2p'
    """

    # We initialize a tkinter root and withdraw it
    root = Tk()
    root.withdraw()

    if option == 'inv':
        response = messagebox.askokcancel(title='Warning', message='No inventory available for the approved items.'
                                                                   ' Would you like to continue anyway?')

    elif option == 'p2p':
        response = messagebox.askokcancel(title='Warning', message='No plant to plant in the parameter box is matching'
                                                                   ' with this approved item.'
                                                                   ' Would you like to continue anyway? '
                                                                   ' If you click ok the item will not be considered'
                                                                   ' in the validation process')
    else:
        response = messagebox.askokcancel(title='Warning', message='No wish found for the approved items.'
                                                                   ' Would you like to continue anyway?')
    root.destroy()
    return response


def compare_maximum_sum(modified_parameters, original_parameters, residuals_counter):
    """
    Compares sum of maximums between commons POINT TO between original parameters and modified parameters

    :param modified_parameters: List of Parameters objects that were used in the process
    :param original_parameters: original list of Parameters from parameter box
    :param residuals_counter: residuals_counter used in the process
    :return: bool indicating if both sums are equal
    """

    # For both set of parameters (p2p) compute sums of maximum for each POINT TO
    original_sums = sum_maximums(original_parameters)
    modified_sums = sum_maximums(modified_parameters)

    # We add residuals to the modified sums
    for key in modified_sums.keys():
        modified_sums[key] += residuals_counter[key]

    print('ORIGINAL MAX SUMS : ', original_sums)
    print('PROCESS MAX SUMS : ', modified_sums)

    for key in modified_sums.keys():
        if original_sums[key] != modified_sums[key]:
            response = messagebox.askokcancel(title='Warning', message='Sum of maximum loads for POINT TO : '+str(key)+
                                                                       'is not matching between process and original'
                                                                       ' grid')
            if not response:
                print('\n', 'VALIDATION PROCESS STOPPED', '\n')
                sys.exit()

    print('SUMS ALL MATCHING!', '\n')


def sum_maximums(parameters_list):
    """
    Sum maximum for each POINT TO in the parameters list

    :param parameters_list: List of Parameters objects
    :return: dictionary with sums of each POINT To
    """
    # We initialize the dict that will contain the results
    results = {}

    # We sum the maximum
    for p2p in parameters_list:
        if p2p.POINT_TO in results:
            results[p2p.POINT_TO] += p2p.LOADMAX
        else:
            results[p2p.POINT_TO] = p2p.LOADMAX

    return results


def read_approved_items(workbook_path, parameters):
    """
    For each line saved in the "APPROVED" tab of the workbook at the path indicated, we'll build an
    ApprovedItem object

    :param workbook_path: path where the workbook was saved
    :param parameters: list with all the parameters (p2p line) from the ParameterBox
    :return: list with all the ApprovedItem
    """
    # We initialize an empty list that will contain all approved item
    approved_items = []

    # We recuperate all items approved in the "APPROVED" worksheet and save them in a pandas dataframe
    wb = load_workbook(workbook_path)
    ws = wb['APPROVED']
    data = build_dataframe(ws)

    # For every item approved
    for i in data.index:

        # We save the point from and shipping point
        point_from = data['POINT_FROM'][i]
        shipping_point = data['SHIPPING_POINT'][i]

        # We get the "days_to" associate with the item approved
        found = False
        for p2p in parameters:
            if p2p.POINT_FROM == point_from and p2p.POINT_TO == shipping_point:
                days_to = p2p.days_to
                found = True

        # We build the ApprovedItem object if the days_to was found
        if found:
            approved_items.append(ApprovedItem(point_from, shipping_point,
                                               data['MATERIAL_NUMBER'][i], data['QUANTITY'][i], days_to))
        else:
            if not warning('p2p'):
                raise Exception('Mismatching points between item approved and P2P available in the ParameterBox')

    return approved_items


def validate_process(workbook_path, modified_parameters, residuals):
    """
    Validates process results

    :param workbook_path: path of the workbook where the output was stored
    :param modified_parameters: List of Parameters objects that were used in the process
    :param residuals: residuals_counter used in the process
    """

    print('\n\n', 'PROCESS VALIDATION STARTED', '\n')

    # We import some SQL data for our validation
    downloaded = False
    nbr_of_try = 0
    while not downloaded and nbr_of_try < 3:  # 3 trials, SQL Queries sometime crash for no reason
        nbr_of_try += 1
        try:
            downloaded = True

            # We recuperate wishlist original data
            wishlist = get_wish_list()

            # We recuperate inventory original data
            inventory = get_inventory_and_qa()

            # We recuperate the original parameters grid (ParameterBox)
            parameters, connection = get_parameter_grid()

        except pyodbc.Error as err:
            downloaded = False
            sql_state = err.args[1]
            print('SQL Query failed :', sql_state)

    if downloaded:

        compare_maximum_sum(modified_parameters, parameters, residuals_counter)

        # We recuperate all items approved in the "APPROVED" worksheet
        approved_items = read_approved_items(workbook_path, parameters)

        # We initialize a list of problematic items
        conflict_items = []

        # We sort approved_items by days to to make sure that future items aren't link first to today's inventory
        approved_items.sort(key=lambda i: i.days_to)

        # We validate if every items were in the wishlist and available in inventory
        for index, item in enumerate(approved_items):
            print(index, '.', item.material_number)
            if not item.find_associated_wish(wishlist):
                if not warning('wish'):
                    return
                else:
                    conflict_items.append(item)

            if not item.find_associated_inventory(inventory):
                if not warning('inv'):
                    return
                else:
                    conflict_items.append(item)

        print('NUMBER OF CONFLICTS FOUND IN INVENTORY:', len(conflict_items))
