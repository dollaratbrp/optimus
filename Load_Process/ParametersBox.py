"""

Author : Olivier Lefebre

This file manage everything about parameter box

Last update : 2020-01-06
By : Nicolas Raymond

"""

from InputOutput import *
import tkinter as tk
from tkinter import *
from tkinter import (messagebox)
from P2PFunctions import get_parameter_grid
import pandas as pd


largeurColonne = 12
Project = ['']

# If the user wants to execute the main program or not,
# becomes True if user click on 'Execute' button
ToExecute = [False]


# Define a frame
def frame(root, side): 
    w = Frame(root)
    w.pack(side=side, expand=YES, fill=BOTH)
    return w


# Look if user's inputs are integers
def IsInt(value):
    try:
        return isinstance(int(value), int)
    except:
        return False


# Each line inside the box
class ligne:
    pointFrom = []
    pointTo = []
    loadMin = []
    loadMax = []
    priority = []
    transit = []
    skip = []
    delete = []
    drybox = []
    flatbed = []
    parent = ''
    lineToSKip = ''
    root = ''
    index = 0
    columnLength = 0
    side = ''

    def __init__(self, parent, root, side, PF='', PT='', LMIN=0, LMAX=0, DRYBOX=None, FLATBED=None,
                 PTY=0, TRANS=0, SKIP=0, days=-1, largeurColonne=12):
        if SKIP == 1:
            IsToSkip = True
        else:
            IsToSkip = False
        if LMAX is None:
            LMAX = ''
        if DRYBOX is None:
            DRYBOX = ''
        if FLATBED is None:
            FLATBED = ''

        self.lineToSKip = BooleanVar()  # IntVar()
        self.lineToSKip.set(IsToSkip)
        self.root = root
        self.parent = parent
        self.index = len(parent.lignes)
        self.columnLength = largeurColonne
        self.side = side

        self.pointFrom = self.entryObj(PF)
        self.pointTo = self.entryObj(PT)
        self.loadMin = self.entryObj(LMIN)
        self.loadMax = self.entryObj(LMAX)
        self.drybox = self.entryObj(DRYBOX)
        self.flatbed = self.entryObj(FLATBED)
        self.priority = self.entryObj(PTY)
        self.transit = self.entryObj(TRANS)
        self.days = self.entryObj(days)

        self.skip = Button(root, text='X', width=largeurColonne, command=self.skipAction)
        # Checkbutton(root, text="", variable=self.lineToSKip)
        self.skip.pack(side=side)  # , expand=YES, fill=BOTH)

        self.delete = Button(root, text='Delete', width=largeurColonne, command=self.deleteAll)
        self.delete.pack(side=side)  # , expand=YES, fill=BOTH)

        self.changeColor()

    def entryObj(self, text):
        """Sets values in each entry"""
        tempo = tk.Entry(self.root, width=self.columnLength, borderwidth=2, relief="groove")
        tempo.insert(END, text)
        tempo.pack(side=self.side)
        return tempo

    def deleteAll(self):
        """Delete a line in the box"""
        self.changeColor('yellow')
        confirm = messagebox.askokcancel('Delete ?', "Are you sure you want to remove this line from the list?")
        if confirm:
            self.pointFrom.forget()
            self.pointTo.forget()
            self.loadMin.forget()
            self.loadMax.forget()
            self.drybox.forget()
            self.flatbed.forget()
            self.priority.forget()
            self.transit.forget()
            self.skip.forget()
            self.delete.forget()
            self.root.pack_forget()
            self.days.pack_forget()

            self.parent.ForgetToDelete(self.index)
        else:
            self.changeColor()

    def skipAction(self):
        """If the skip button is clicked"""
        self.lineToSKip.set(not self.lineToSKip.get())
        self.changeColor()

    def changeColor(self, color=''):
        """Changes the background color, red if skip is true, else white."""
        if color == '':
            if self.lineToSKip.get():
                color = 'red'
            else:
                color = 'white'
        self.pointFrom.config(bg=color)
        self.pointFrom.config(bg=color)
        self.pointTo.config(bg=color)
        self.loadMin.config(bg=color)
        self.loadMax.config(bg=color)
        self.drybox.config(bg=color)
        self.flatbed.config(bg=color)
        self.priority.config(bg=color)
        self.transit.config(bg=color)
        self.skip.config(bg=color)
        self.days.config(bg=color)


class VerticalScrolledFrame(tk.Frame):

    def __init__(self, parent, *args, **kw):
        """
        Creates a tkinter frame on which we can scroll vertically
        :param parent: tkinter root
        """
        tk.Frame.__init__(self, parent, *args, **kw)

        # Creation of a canvas object and a vertical scrollbar to scroll in it
        self.vscrollbar = tk.Scrollbar(self, orient=tk.VERTICAL)
        self.vscrollbar.pack(fill=tk.Y, side=tk.RIGHT, expand=tk.FALSE)
        self.canvas = tk.Canvas(self, bd=0, highlightthickness=0, height=650, yscrollcommand=self.vscrollbar.set)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=tk.TRUE)
        self.vscrollbar.config(command=self.canvas.yview)

        # Reset the view
        self.canvas.xview_moveto(0)
        self.canvas.yview_moveto(0)

        # Create a frame inside the canvas on which we'll be able to scroll
        self.interior = tk.Frame(self.canvas)
        self.interior_id = self.canvas.create_window(0, 0, window=self.interior, anchor=tk.NW)

        # Track changes of the canvas and frame width and sync them, also update the scrollbar
        self.interior.bind('<Configure>', self._configure_interior)
        self.canvas.bind('<Configure>', self._configure_canvas)

    def _configure_canvas(self, event):
        """
        Updates the inner frame's width to fill the canvas
        :param event: information on the activities in the GUI (This parameter must be there!)
        :return:
        """
        if self.interior.winfo_reqwidth() != self.canvas.winfo_width():

            # update the inner frame's width to fill the canvas
            self.canvas.itemconfigure(self.interior_id, width=self.canvas.winfo_width())

    def _configure_interior(self, event):
        """
        Updates the scrollbar to match the size of the inner frame
        :param event: information on the activities in the GUI (This parameter must be there!)
        """
        # update the scrollbars to match the size of the inner frame
        size = (self.interior.winfo_reqwidth(), self.interior.winfo_reqheight())
        self.canvas.config(scrollregion="0 0 %s %s" % size)

        if self.interior.winfo_reqwidth() != self.canvas.winfo_width():

            # update the canvas's width to fit the inner frame
            self.canvas.config(width=self.interior.winfo_reqwidth())


# The main box
class Box(Frame):

    lignes = []

    def __init__(self, largeurColonne):
        Frame.__init__(self)
        global Project

        headers = ['Point from', 'Point to', 'Load min', 'Load max', 'DRYBOX', 'FLATBED', 'Priority Order', 'Transit',
                   'DAYS_TO', 'Skip', '']  # Only for the box, not sql

        if Project[0] == 'P2P':
            ValuesParams, self.SQL = get_parameter_grid(parameter_box_output=True)
        else:
            ValuesParams, self.SQL = get_parameter_grid(forecast=True, parameter_box_output=True)
        print(ValuesParams)

        self.option_add('*Font', 'Verdana 12 bold')
        self.pack(expand=YES, fill=BOTH)
        self.master.title('Parameters Box')
        # self.master.geometry("400x450")

        # Section pour les headers
        option = ["raised", "sunken",  "solid"]  # choices for header's layout
        keyF = frame(self, TOP)
        for i, titre in enumerate(headers[0:-1]):
            tk.Label(keyF, text=titre, width=largeurColonne, bg='#ABA3A3', anchor='w',
                     borderwidth=2, relief=option[0], pady=10).pack(side=LEFT)
            # , expand=YES, fill=BOTH)#Label(keyF, text=titre, width=15).pack(side=LEFT)
        tk.Label(keyF, text=headers[-1], width=largeurColonne).pack(side=LEFT)

        # Define scrollbar
        keyF = frame(self, TOP)
        scframe = VerticalScrolledFrame(keyF)
        scframe.pack()

        self.verticalBar = scframe.interior

        for values in ValuesParams:
            keyF = frame(self.verticalBar, TOP)
            self.lignes.append(ligne(self, keyF, LEFT, *values, largeurColonne))

        keyF = frame(self, BOTTOM)
        tk.Button(keyF, text='Add New', command=lambda largeurColonne=largeurColonne: self.AddNew(largeurColonne),
                  borderwidth=2, bg='#ABA3A3', relief=option[0], pady=10).pack(side=LEFT, expand=YES, fill=BOTH)
        tk.Button(keyF, text='Modify Email list', command=change_emails_list, borderwidth=2, bg='#ABA3A3', relief=option[0],
                  pady=10).pack(side=LEFT, expand=YES, fill=BOTH)
        tk.Button(keyF, text='Execute', command=self.quit, borderwidth=2, bg='#ABA3A3', relief=option[0],
                  pady=10).pack(side=LEFT, expand=YES, fill=BOTH)

    def ForgetToDelete(self, index):
        """To delete a line of values"""
        self.lignes[index] = ''

    def quit(self):
        """Delete all from SQL and send new data if there is no mistakes"""
        WarningColor = 'yellow'
        errors = False
        DATA_TO_SEND = []
        priorityOrder =[]
        for ligne in self.lignes:

            if ligne != '':
                ligne.changeColor()  # reset color to default color
                if ligne.priority.get() in priorityOrder:
                    ligne.priority.config(bg=WarningColor)
                    errors = True
                else:
                    priorityOrder.append(ligne.priority.get())

                if not IsInt(ligne.pointFrom.get()):
                    ligne.pointFrom.config(bg=WarningColor)
                    errors = True

                if not IsInt(ligne.pointTo.get()):
                    ligne.pointTo.config(bg=WarningColor)
                    errors = True

                if not IsInt(ligne.loadMin.get()):
                    ligne.loadMin.config(bg=WarningColor)
                    errors = True

                if not (IsInt(ligne.loadMax.get()) or ligne.loadMax.get() == ''):
                    ligne.loadMax.config(bg=WarningColor)
                    errors = True

                if not (IsInt(ligne.drybox.get()) or ligne.drybox.get() == ''):
                    ligne.drybox.config(bg=WarningColor)
                    errors = True

                if not (IsInt(ligne.flatbed.get()) or ligne.flatbed.get() == ''):
                    ligne.flatbed.config(bg=WarningColor)
                    errors = True

                if IsInt(ligne.loadMin.get()) and IsInt(ligne.loadMax.get()):
                    if int(ligne.loadMin.get()) > int(ligne.loadMax.get()):
                        ligne.loadMin.config(bg=WarningColor)
                        ligne.loadMax.config(bg=WarningColor)
                        errors = True

                if not IsInt(ligne.priority.get()):
                    ligne.priority.config(bg=WarningColor)
                    errors = True

                if not IsInt(ligne.transit.get()):
                    ligne.transit.config(bg=WarningColor)
                    errors = True

                DATA_TO_SEND.append(
                        [ligne.pointFrom.get(), ligne.pointTo.get(), ligne.loadMin.get(), ligne.loadMax.get(),
                         ligne.drybox.get(), ligne.flatbed.get(), ligne.priority.get(), ligne.transit.get(),
                         ligne.days.get(), ligne.lineToSKip.get()])

        if not errors:
            timeOfExport = pd.datetime.now()
            # delete in SQL, only keep the last 3 modifications in history
            if Project[0] == 'P2P':
                self.SQL.deleteFromSQL("IMPORT_DATE not in (SELECT  DISTINCT top(3) IMPORT_DATE FROM"
                                       " [Business_Planning].[dbo].[OTD_1_P2P_F_PARAMETERS] ORDER BY IMPORT_DATE desc)")
            else:
                self.SQL.deleteFromSQL(
                    "IMPORT_DATE not in (SELECT  DISTINCT top(3) IMPORT_DATE"
                    " FROM [Business_Planning].[dbo].[OTD_1_P2P_F_FORECAST_PARAMETERS] ORDER BY IMPORT_DATE desc) ")

            # send data to SQL
            for DATALine in DATA_TO_SEND:
                for obj in range(2, len(DATALine)-1):
                    if DATALine[obj] == '':
                        DATALine[obj] = None
                    else:
                        DATALine[obj] = int(DATALine[obj])

                if DATALine[-1] is True:
                    DATALine[-1] = 1
                else:
                    DATALine[-1] = 0

                DATALine.append(timeOfExport)
                self.SQL.sendToSQL([(DATALine)])

            global ToExecute
            ToExecute = [True]

            self.master.destroy()
        else:
            messagebox.showerror('Invalid DATA', 'There are some invalid inputs in the table')

    def AddNew(self, largeurColonne):
        """To add a new line of values"""
        keyF = frame(self.verticalBar, TOP)
        global Project
        self.lignes.append(ligne(self, keyF, LEFT, '0000', '0000', 0, 0, '', '', 0, 0, False, 0, largeurColonne))


def change_emails_list():
    """
    Opens a new window to allow user to change emails list
    """
    email_box = EmailBox()


class EmailBox(VerticalScrolledFrame):

    def __init__(self):

        # We initialize the root
        self.root = Toplevel()
        self.root.title("EMAIL ADDRESS")
        self.root.configure(background="gray99")
        self.root.geometry("400x500+700+300")

        # We initialize our frame
        super().__init__(self.root)
        self.pack()

        # We recuperate the list of emails
        self.emails_list, self.connection = self.get_email_addresses(project_name=Project[0])

        # We initialize and pack a "ADD NEW" button
        self.add_new_button = Button(self.interior, height=1, width=40, relief=FLAT, bg="gray99", fg="purple3",
                                     font="Dosis", text='ADD NEW', command=self.add_new)
        self.add_new_button.pack()

        # We initialize and pack a label
        self.label = Label(self.interior, text="CLICK ON EMAIL TO DELETE", font=("Dosis", 12))
        self.label.pack(padx=10, pady=5, side=TOP)

        # We set the rest of the missing buttons associated to the email addresses
        self.email_buttons = []
        for i, email in enumerate(self.emails_list):
            self.email_buttons.append(Button(self.interior, height=1, width=40, relief=FLAT, bg="gray99",
                                             fg="purple3", font="Dosis", text=email,
                                             command=lambda i=i: self.warning(i)))
            self.email_buttons[i].pack(padx=10, pady=5, side=TOP)

    @staticmethod
    def get_email_addresses(project_name):

        sql_connection = SQLConnection('CAVLSQLPD2\pbi2', 'Business_Planning','OTD_1_P2P_F_PARAMETERS_EMAIL_ADDRESS',
                                       'EMAIL_ADDRESS,PROJECT')

        email_query = """ SELECT DISTINCT [EMAIL_ADDRESS] FROM [Business_Planning].[dbo].[OTD_1_P2P_F_PARAMETERS_EMAIL_ADDRESS]
        WHERE PROJECT = '{0}'""".format(project_name)

        return [item for sublist in sql_connection.GetSQLData(email_query) for item in sublist], sql_connection

    def warning(self, i):
        """
        Deletes an email address
        """
        email_to_delete = self.emails_list[i]

        confirm = messagebox.askokcancel('Confirmation',
                                         "Are you sure you want to remove"
                                         " {0} from the list?".format(email_to_delete))
        if confirm:
            self.connection.deleteFromSQL("[EMAIL_ADDRESS] = ('{0}') and PROJECT = '{1}'".format(email_to_delete, Project[0]))
            self.email_buttons[i].forget()
            self.email_buttons.pop(i)
            self.emails_list.pop(i)

        self.root.deiconify()

    def add_new(self):
        """
        Opens a new window to allow the user to add a new email address
        """
        new_email_box = NewEmailBox(self)


class NewEmailBox:

    def __init__(self, parent_email_box):

        # We save the parent of the new box
        self.parent = parent_email_box

        # Master initialization (root)
        self.master = Toplevel()
        self.master.geometry("500x80+700+500")
        self.master.title("ADD NEW EMAIL ADDRESS")

        # Email label initialization and positioning
        self.email_label = Label(self.master, text="Email", font=("Dosis", 12))
        self.email_label.grid(row=0, column=0)

        # Email label initialization and positioning
        self.email_entry = Entry(self.master, width=50)
        self.email_entry.grid(row=0, column=1)

        # Cancel button initialization and positioning
        self.cancel_button = Button(self.master, text='Cancel', command=self.master.destroy, font=("Dosis", 12))
        self.cancel_button.grid(row=1, column=0, sticky=W, pady=4, columnspan=1)

        # Add button initialization and positioning
        self.add = Button(self.master, text='Add', command=self.add_new_email, font=("Dosis", 12))
        self.add.grid(row=1, column=1, sticky=W, pady=4)

    def add_new_email(self):
        """
        Adds the email address written
        """
        # We retrieve the new address
        new_address = str(self.email_entry.get().lower())

        # If the address is not already in the parent emails list
        if new_address not in self.parent.emails_list:

            # We add the new addresse in the email list
            self.parent.emails_list.append(new_address)
            index = len(self.parent.emails_list) - 1

            # We add a new button for this address
            self.parent.email_buttons.append(Button(self.parent.interior, height=1, width=40, relief=FLAT,
                                                    bg="gray99", fg="purple3", font="Dosis", text=new_address,
                                                    command=lambda: self.parent.warning(index)))
            self.parent.email_buttons[-1].pack(padx=10, pady=5, side=TOP)

            # We send the data to SQL
            self.parent.connection.sendToSQL([(new_address, Project[0])])

        self.email_entry.delete(0, END)
        self.master.destroy()


def MissingP2PBox(MissingP2P):
    ToExecute[0] = False
    errorMess = 'Missing these P2P (point_from - point_to) : \n'
    for p2p in MissingP2P:
        errorMess += str(p2p) + '\n'
    temp_root = Tk()
    temp_root.withdraw()
    response = messagebox.askokcancel("Warning", errorMess + '\n Continue execution? \n')
    temp_root.destroy()
    return response


def OpenParameters(project_name='P2P'):

    set_project_name(project_name)
    Box(largeurColonne).mainloop()
    return ToExecute[0]


def set_project_name(name):
    """
    Sets the value of the global variable Project
    :param name: name of the project
    """
    global Project
    Project = [name]


if __name__ == '__main__':
    OpenParameters()
    print(ToExecute[0])