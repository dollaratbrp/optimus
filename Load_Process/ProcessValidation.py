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


def validate_process(workbook_path):

    print('\n\n', 'PROCESS VALIDATION STARTED', '\n')

    # We recuperate wishlist original data
    wishlist = get_wish_list()

    # We recuperate inventory original data
    inventory = get_inventory_and_qa()

    # We recuperate the original parameters grid (ParameterBox)
    parameters, connection = get_parameter_grid()

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

    print('NUMBER OF CONFLICTS FOUND :', len(conflict_items))
