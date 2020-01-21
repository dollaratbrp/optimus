"""
This file manages activities related to process tests and results validation

Author : Nicolas Raymond
"""

from tkinter import *
from openpyxl import load_workbook
from FastLoads import build_dataframe
from tkinter import messagebox
from P2PFunctions import *


class ApprovedItems:

    """Contains all datas related to a line of "APPROVED" worksheet of the P2P_Summary file"""

    def __init__(self, point_from, shipping_point, material_number, qty):

        """
        :param point_from: point where the item comes from (Ex : 4100)
        :param shipping_point: point where the item is going to be shipped (Ex : 4015)
        :param material_number: material number (SKU) of the item
        :param qty: number of times this SKU is shipped
        """

        self.point_from = point_from
        self.shipping_point = shipping_point
        self.material_number = material_number
        self.qty = qty

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
                    self.material_number == inventory_line.MATERIAL_NUMBER and inventory_line.QUANTITY >= self.qty:

                inventory_line.QUANTITY -= self.qty
                print('POINT FROM :', inventory_line.POINT)
                print('INVENTORY TAKEN :', self.qty)
                print('INVENTORY LEFT : ', inventory_line.QUANTITY, '\n')

                if inventory_line.QUANTITY == 0:
                    inventory_list.pop(index)

                return True

        return False


def warning(options):

    """
    Shows a warning message

    :param options: 'wish' or 'inv'
    """

    # We initialize a tkinter root and withdraw it
    root = Tk()
    root.withdraw()

    if options == 'inv':
        response = messagebox.askokcancel(title='Warning', message='No inventory available for the approved items.'
                                                                   ' Would you like to continue anyway?')
    else:
        response = messagebox.askokcancel(title='Warning', message='No wish found for the approved items.'
                                                                   ' Would you like to continue anyway?')
    root.destroy()
    return response


def validate_process(workbook_path):

    print('\n\n', 'PROCESS VALIDATION STARTED', '\n')

    # We recuperate wishlist data
    wishlist = get_wish_list()

    # We recuperate inventory data
    inventory = get_inventory_and_qa()

    # We recuperate all items approved in the "APPROVED" worksheet
    wb = load_workbook(workbook_path)
    ws = wb['APPROVED']
    data = build_dataframe(ws)
    items_approved = []
    for line in data.index:
        items_approved.append(ApprovedItems(data['POINT_FROM'][line], data['SHIPPING_POINT'][line],
                                            data['MATERIAL_NUMBER'][line], data['QUANTITY'][line]))

    # We initialize a list of problematic items
    conflict_items = []

    # We validate if every items were in the wishlist and available in inventory
    for index, item in enumerate(items_approved):
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
