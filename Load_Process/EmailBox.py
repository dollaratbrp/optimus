"""
This file manages all activities linked to the Email list GUI

Author : Nicolas Raymond
"""
from ParametersBox import VerticalScrolledFrame
from tkinter import *
from tkinter import messagebox
from InputOutput import SQLConnection


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
        global Project
        self.emails_list, self.connection = self.get_email_addresses(project_name=Project[0])

        # We initialize and pack a "ADD NEW" button
        self.add_new_button = Button(self.interior, height=1, width=40, relief=FLAT, bg="gray99", fg="purple3",
                                     font="Dosis", text='ADD NEW', command=self.add_new)
        self.add_new_button.pack()

        # We initialize and pack a label
        self.label = Label(self.interior, text="CLICK ON EMAIL TO DELETE", font="Dosis")
        self.label.pack()

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
        print(email_query)

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
        self.email_label = Label(self.master, text="EMAIL ADDRESS")
        self.email_label.grid(row=0, column=0)

        # Email label initialization and positioning
        self.email_entry = Entry(self.master, width=50)
        self.email_entry.grid(row=0, column=1)

        # Cancel button initialization and positioning
        self.cancel_button = Button(self.master, text='Cancel', command=self.master.destroy)
        self.cancel_button.grid(row=1, column=0, sticky=W, pady=4)

        # Add button initialization and positioning
        self.add = Button(self.master, text='ADD', command=self.add_new_email)
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


if __name__ == '__main__':

    test = Tk()
    test.withdraw()
    global Project
    Project = ['P2P']
    box = EmailBox()
    test.mainloop()
