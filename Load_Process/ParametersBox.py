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


column_width = 8
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
class Line:

    def __init__(self, parent, root, side, point_from='', point_to='', lmin=0, lmax=0, drybox=None, flatbed=None,
                 priority=0, transit=0, skip=0, days=-1, col_width=12):
        if skip == 1:
            to_skip = True
        else:
            to_skip = False
        if lmax is None:
            lmax = ''
        if drybox is None:
            drybox = ''
        if flatbed is None:
            flatbed = ''

        self.lineToSKip = BooleanVar()  # IntVar()
        self.lineToSKip.set(to_skip)
        self.root = root
        self.parent = parent
        self.index = len(parent.lignes)
        self.columnLength = col_width
        self.side = side

        # Entries
        self.pointFrom = self.entryObj(point_from)
        self.pointTo = self.entryObj(point_to)
        self.loadMin = self.entryObj(lmin)
        self.loadMax = self.entryObj(lmax)
        self.drybox = self.entryObj(drybox)
        self.flatbed = self.entryObj(flatbed)
        self.priority = self.entryObj(priority)
        self.transit = self.entryObj(transit)
        self.days = self.entryObj(days)

        self.skip = Button(root, text='X', width=col_width, command=self.skipAction)
        self.skip.pack(side=side)  # , expand=YES, fill=BOTH)

        self.delete = Button(root, text='Delete', width=col_width, command=self.deleteAll)
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

            self.parent.delete_line(self.index)
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

    def __init__(self, parent, height=650, *args, **kw):
        """
        Creates a tkinter frame on which we can scroll vertically
        :param parent: tkinter root
        """
        tk.Frame.__init__(self, parent, *args, **kw)

        # Creation of a canvas object and a vertical scrollbar to scroll in it
        self.vscrollbar = tk.Scrollbar(self, orient=tk.VERTICAL)
        self.vscrollbar.pack(fill=tk.Y, side=tk.RIGHT, expand=tk.FALSE)
        self.canvas = tk.Canvas(self, bd=0, highlightthickness=0, height=height, yscrollcommand=self.vscrollbar.set)
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

    def __init__(self, col_width):

        Frame.__init__(self)

        global Project

        # We memorize column titles
        headers = ['From', 'To', 'Min', 'Max', 'Drybox', 'Flatbed', 'Priority', 'Transit',
                   'Days_to', 'Skip', '']

        # We get parameter box data
        if Project[0] == 'P2P':
            parameters_values, self.SQL = get_parameter_grid(parameter_box_output=True)
        else:
            parameters_values, self.SQL = get_parameter_grid(forecast=True, parameter_box_output=True)

        # We set some GUI details
        self.option_add('*Font', 'Verdana 12 bold')
        self.pack(expand=True)
        self.master.title('Parameters Box')

        # We set header labels
        option = ["raised", "sunken",  "solid"]  # choices for header's layout
        self.header_frame = frame(self, TOP)
        self.header_label = []
        i = 0
        for titre in headers[0:-1]:
            self.header_label.append(tk.Label(self.header_frame, text=titre, width=col_width,
                                              bg='#ABA3A3', anchor='center', borderwidth=2, relief=option[0], pady=10))
            self.header_label[i].grid(row=0, column=i)
            i += 1

        self.header_label.append(tk.Label(self.header_frame, text=headers[-1], width=col_width))
        self.header_label[-1].grid(row=0, column=i)

        # We define our vertical scrolled frame that will contain parameters value
        self.parameters_frame = VerticalScrolledFrame(frame(self, TOP), height=400)
        self.parameters_frame.pack()
        self.verticalBar = self.parameters_frame.interior
        for values in parameters_values:
            temp_frame = frame(self.verticalBar, TOP)
            self.lignes.append(Line(self, temp_frame, LEFT, *values, col_width))

        # We set GUI buttons
        self.buttons_frame = frame(self, BOTTOM)
        tk.Button(self.buttons_frame, text='Add New', command=lambda col_w=col_width: self.add_new_line(col_w),
                  borderwidth=2, bg='#ABA3A3', relief=option[0], pady=10).pack(side=LEFT, expand=YES, fill=BOTH)

        tk.Button(self.buttons_frame, text='Modify Email list', command=change_emails_list,
                  borderwidth=2, bg='#ABA3A3', relief=option[0], pady=10).pack(side=LEFT, expand=YES, fill=BOTH)

        tk.Button(self.buttons_frame, text='Save Parameters', command=self.save_parameters,
                  borderwidth=2, bg='#ABA3A3', relief=option[0], pady=10).pack(side=LEFT, expand=YES, fill=BOTH)

        tk.Button(self.buttons_frame, text='Execute', command=self.quit,
                  borderwidth=2, bg='#ABA3A3', relief=option[0], pady=10).pack(side=LEFT, expand=YES, fill=BOTH)

    def delete_line(self, index):

        """To delete a line of values"""

        self.lignes[index] = ''

    def save_parameters(self):

        """Save all parameters in the GUI"""

        # We init some useful variables
        warning_color = 'yellow'
        errors = False
        data_to_send = []

        # We valid each line
        for ligne in self.lignes:

            if ligne != '':
                ligne.changeColor()  # reset color to default color

                if not IsInt(ligne.pointFrom.get()):
                    ligne.pointFrom.config(bg=warning_color)
                    errors = True

                if not IsInt(ligne.pointTo.get()):
                    ligne.pointTo.config(bg=warning_color)
                    errors = True

                if not IsInt(ligne.loadMin.get()):
                    ligne.loadMin.config(bg=warning_color)
                    errors = True

                if not (IsInt(ligne.loadMax.get()) or ligne.loadMax.get() == ''):
                    ligne.loadMax.config(bg=warning_color)
                    errors = True

                if not (IsInt(ligne.drybox.get()) or ligne.drybox.get() == ''):
                    ligne.drybox.config(bg=warning_color)
                    errors = True

                if not (IsInt(ligne.flatbed.get()) or ligne.flatbed.get() == ''):
                    ligne.flatbed.config(bg=warning_color)
                    errors = True

                if IsInt(ligne.loadMin.get()) and IsInt(ligne.loadMax.get()):
                    if int(ligne.loadMin.get()) > int(ligne.loadMax.get()):
                        ligne.loadMin.config(bg=warning_color)
                        ligne.loadMax.config(bg=warning_color)
                        errors = True

                if not IsInt(ligne.priority.get()):
                    ligne.priority.config(bg=warning_color)
                    errors = True

                if not IsInt(ligne.transit.get()):
                    ligne.transit.config(bg=warning_color)
                    errors = True

                if not IsInt(ligne.days.get()):
                    ligne.days.config(bg=warning_color)
                    errors = True

                data_to_send.append(
                    [ligne.pointFrom.get(), ligne.pointTo.get(), ligne.loadMin.get(), ligne.loadMax.get(),
                     ligne.drybox.get(), ligne.flatbed.get(), ligne.priority.get(), ligne.transit.get(),
                     ligne.days.get(), ligne.lineToSKip.get()])

        if not errors:

            export_time = pd.datetime.now()

            # delete in SQL, only keep the last 3 modifications in history
            if Project[0] == 'P2P':
                self.SQL.deleteFromSQL("IMPORT_DATE not in (SELECT  DISTINCT top(3) IMPORT_DATE FROM"
                                       " [Business_Planning].[dbo].[OTD_1_P2P_F_PARAMETERS] ORDER BY IMPORT_DATE desc)")
            else:
                self.SQL.deleteFromSQL(
                    "IMPORT_DATE not in (SELECT  DISTINCT top(3) IMPORT_DATE"
                    " FROM [Business_Planning].[dbo].[OTD_1_P2P_F_FORECAST_PARAMETERS] ORDER BY IMPORT_DATE desc) ")

            # send data to SQL
            for DATALine in data_to_send:
                for obj in range(2, len(DATALine) - 1):
                    if DATALine[obj] == '':
                        DATALine[obj] = None
                    else:
                        DATALine[obj] = int(DATALine[obj])

                if DATALine[-1] is True:
                    DATALine[-1] = 1
                else:
                    DATALine[-1] = 0

                DATALine.append(export_time)
                self.SQL.sendToSQL([DATALine])
        else:
            messagebox.showerror('Invalid DATA', 'There are some invalid inputs in the table')

        return errors

    def quit(self):

        """Delete all from SQL and send new data if there is no mistakes"""

        errors = self.save_parameters()

        if not errors:
            global ToExecute
            ToExecute = [True]
            self.master.destroy()

    def add_new_line(self, col_width):

        """To add a new line of values"""

        global Project
        self.lignes.append(Line(self, frame(self.verticalBar, TOP), LEFT, '0000', '0000', 0, 0, '', '', 0, 0,
                                False, 0, col_width))


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
    Box(column_width).mainloop()
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