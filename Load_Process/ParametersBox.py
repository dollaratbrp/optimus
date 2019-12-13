from Import_Functions import *
import tkinter as tk
from tkinter import *
from tkinter import (messagebox)
import pandas as pd


largeurColonne=12
Project = ['']
ToExecute=[False]#If the user wants to execute the main program or not, becomes True if user click on 'Execute' button


#Define a frame
def frame(root, side): 
    w = Frame(root)
    w.pack(side=side, expand=YES, fill=BOTH)
    return w

#Look if user's inputs are integers
def IsInt(value):
    try:
        return isinstance(int(value), int)
    except:
        return False


#Each line inside the box
class ligne:
    pointFrom = []
    pointTo = []
    loadMin=[]
    loadMax=[]
    priority=[]
    transit=[]
    skip=[]
    delete=[]
    drybox=[]
    flatbed=[]
    parent=''
    lineToSKip=''
    root=''
    index=0
    columnLength=0
    side = ''

    def __init__(self,parent,root, side,PF='',PT='',LMIN=0,LMAX=0,DRYBOX=None,FLATBED=None,PTY=0,TRANS=0,SKIP=0,days=-1,largeurColonne=12):
        if SKIP ==1:
            IsToSkip = True
        else:
            IsToSkip = False
        if LMAX is None:
            LMAX = ''
        if DRYBOX is None:
            DRYBOX=''
        if FLATBED is None:
            FLATBED = ''

        self.lineToSKip=BooleanVar()#IntVar()
        self.lineToSKip.set(IsToSkip)
        self.root = root
        self.parent = parent
        self.index=len(parent.lignes)
        self.columnLength= largeurColonne
        self.side = side

        self.pointFrom   = self.entryObj(PF)
        self.pointTo = self.entryObj(PT)
        self.loadMin =self.entryObj(LMIN)
        self.loadMax = self.entryObj(LMAX)
        self.drybox = self.entryObj(DRYBOX)
        self.flatbed =  self.entryObj(FLATBED)
        self.priority = self.entryObj(PTY)
        self.transit = self.entryObj(TRANS)

        if days != -1:
            self.days = self.entryObj(days)
            self.NbDays = True
        else:
            self.NbDays = False

        self.skip = Button(root,text='X',width = largeurColonne,command = self.skipAction)#Checkbutton(root, text="", variable=self.lineToSKip)
        self.skip.pack(side=side)#, expand=YES, fill=BOTH)

        self.delete = Button(root,text='Delete',width = largeurColonne,command = self.deleteAll)
        self.delete.pack(side=side)#, expand=YES, fill=BOTH)

        self.changeColor()

    def entryObj(self, text):
        "Sets values in each entry"
        tempo = tk.Entry(self.root, width=self.columnLength, borderwidth=2, relief="groove")
        tempo.insert(END, text)
        tempo.pack(side=self.side)
        return tempo


    def deleteAll(self):
        "Delete a line in the box"
        self.changeColor('yellow')
        confirm = messagebox.askokcancel('Delete ?',"Are you sure you want to remove this line from the list?")
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
            if self.NbDays:
                self.days.pack_forget()

            self.parent.ForgetToDelete(self.index)
        else:
            self.changeColor()

    def skipAction(self):
        "If the skip button is clicked"
        self.lineToSKip.set(not self.lineToSKip.get())
        self.changeColor()


    def changeColor(self,color = ''):
        "Changes the background color, red if skip is true, else white."
        if color =='':
            if self.lineToSKip.get():
                color='red'
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
        if self.NbDays:
            self.days.config(bg = color)

#To create a vertical scrollBar
class VerticalScrolledFrame(tk.Frame):

        def __init__(self, parent, *args, **kw):
            tk.Frame.__init__(self, parent, *args, **kw)

            # create a canvas object and a vertical scrollbar for scrolling it
            vscrollbar = tk.Scrollbar(self, orient=tk.VERTICAL)
            vscrollbar.pack(fill=tk.Y, side=tk.RIGHT, expand=tk.FALSE)
            canvas = tk.Canvas(self, bd=0, highlightthickness=0,height=650,
                            yscrollcommand=vscrollbar.set)
            canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=tk.TRUE)
            vscrollbar.config(command=canvas.yview)

            # reset the view
            canvas.xview_moveto(0)
            canvas.yview_moveto(0)

            # create a frame inside the canvas which will be scrolled with it
            self.interior = interior = tk.Frame(canvas)
            interior_id = canvas.create_window(0, 0, window=interior,
                                               anchor=tk.NW)

            # track changes to the canvas and frame width and sync them,
            # also updating the scrollbar
            def _configure_interior(event):
                # update the scrollbars to match the size of the inner frame
                size = (interior.winfo_reqwidth(), interior.winfo_reqheight())
                canvas.config(scrollregion="0 0 %s %s" % size)
                if interior.winfo_reqwidth() != canvas.winfo_width():
                    # update the canvas's width to fit the inner frame
                    canvas.config(width=interior.winfo_reqwidth())

            interior.bind('<Configure>', _configure_interior)

            def _configure_canvas(event):
                if interior.winfo_reqwidth() != canvas.winfo_width():
                    # update the inner frame's width to fill the canvas
                    canvas.itemconfigure(interior_id, width=canvas.winfo_width())
            canvas.bind('<Configure>', _configure_canvas)


#The main box
class Box(Frame):
    lignes=[]
    verticalBar=''
    headers = ''
    SQL = ''


    def __init__(self,largeurColonne):
        Frame.__init__(self)
        global Project
        if Project[0] == 'P2P':
            self.headers='point_FROM,point_TO,LOAD_MIN,LOAD_MAX,DRYBOX,FLATBED,PRIORITY_ORDER,TRANSIT,SKIP,IMPORT_DATE'
            headers= ['Point from','Point to','Load min','Load max','DRYBOX','FLATBED','Priority Order','Transit','Skip' ,''] #Only for the box, not sql
            self.SQL = SQLConnection('CAVLSQLPD2\pbi2', 'Business_Planning', 'OTD_1_P2P_F_PARAMETERS', self.headers)
            # To set values
            SQLquery = """SELECT distinct [POINT_FROM]
                ,[POINT_TO]
                ,[LOAD_MIN]
                ,[LOAD_MAX]
                ,[DRYBOX]
                ,[FLATBED]
                ,[PRIORITY_ORDER]
                ,[TRANSIT]
                ,[SKIP]
                , -1 as [DAYS_TO]
            FROM [Business_Planning].[dbo].[OTD_1_P2P_F_PARAMETERS]
            where IMPORT_DATE = (select max(IMPORT_DATE) from [Business_Planning].[dbo].[OTD_1_P2P_F_PARAMETERS])
            order by [POINT_FROM]
                ,[POINT_TO] """
        else:
            self.headers='point_FROM,point_TO,LOAD_MIN,LOAD_MAX,DRYBOX,FLATBED,PRIORITY_ORDER,TRANSIT,[DAYS_TO],SKIP,IMPORT_DATE'
            headers= ['Point from','Point to','Load min','Load max','DRYBOX','FLATBED','Priority Order','Transit','DAYS_TO','Skip' ,''] #Only for the box, not sql
            self.SQL = SQLConnection('CAVLSQLPD2\pbi2', 'Business_Planning', 'OTD_1_P2P_F_FORECAST_PARAMETERS', self.headers)
            # To set values
            SQLquery = """SELECT distinct [POINT_FROM]
                ,[POINT_TO]
                ,[LOAD_MIN]
                ,[LOAD_MAX]
                ,[DRYBOX]
                ,[FLATBED]
                ,[PRIORITY_ORDER]
                ,[TRANSIT]
                ,[SKIP]
                ,[DAYS_TO]

            FROM [Business_Planning].[dbo].[OTD_1_P2P_F_FORECAST_PARAMETERS]
            where IMPORT_DATE = (select max(IMPORT_DATE) from [Business_Planning].[dbo].[OTD_1_P2P_F_FORECAST_PARAMETERS])
            order by [POINT_FROM]
                ,[POINT_TO] """

        self.option_add('*Font', 'Verdana 12 bold')
        self.pack(expand=YES, fill=BOTH)
        self.master.title('Parameters Box')
        #self.master.geometry("400x450")

        ##Section pour les headers
        option=["raised", "sunken",  "solid"] # choices for header's layout
        keyF = frame(self, TOP)
        for i, titre in enumerate(headers[0:-1]):
            tk.Label(keyF, text=titre, width=largeurColonne,bg='#ABA3A3',anchor='w', borderwidth=2, relief=option[0],pady=10 ).pack(side=LEFT)#, expand=YES, fill=BOTH)#Label(keyF, text=titre, width=15).pack(side=LEFT)
        tk.Label(keyF, text=headers[-1], width=largeurColonne).pack(side=LEFT)





        ##Define scrollbar
        keyF = frame(self, TOP)
        scframe = VerticalScrolledFrame(keyF)
        scframe.pack()


        self.verticalBar=scframe.interior



        ValuesParams= [ sublist for sublist in self.SQL.GetSQLData(SQLquery) ]

        for values in ValuesParams:
            keyF = frame(self.verticalBar, TOP)
            self.lignes.append(ligne(self,keyF,LEFT,*values,largeurColonne))


        keyF = frame(self, BOTTOM)
        tk.Button(keyF, text='Add New', command = lambda largeurColonne=largeurColonne : self.AddNew(largeurColonne), borderwidth=2,bg='#ABA3A3', relief=option[0],pady=10).pack(side=LEFT, expand=YES, fill=BOTH )
        tk.Button(keyF, text='Modify Email list', command = changeEmail, borderwidth=2,bg='#ABA3A3', relief=option[0],pady=10).pack(side=LEFT, expand=YES, fill=BOTH )
        tk.Button(keyF, text='Execute', command = self.quit, borderwidth=2,bg='#ABA3A3', relief=option[0],pady=10).pack(side=LEFT, expand=YES, fill=BOTH)

    def ForgetToDelete(self,index):
        "To delete a line of values"
        self.lignes[index] = ''



    def quit(self):
        "Delete all from SQL and send new data if there is no mistakes"
        WarningColor = 'yellow'
        errors = False
        DATA_TO_SEND = []
        priorityOrder=[]
        for ligne in self.lignes:

            if ligne !='':
                ligne.changeColor()#reset color to default color
                if ligne.priority.get() in priorityOrder:
                    ligne.priority.config(bg=WarningColor)
                    errors=True
                else:
                    priorityOrder.append(ligne.priority.get())

                if not IsInt(ligne.pointFrom.get()):
                    ligne.pointFrom.config(bg=WarningColor)
                    errors=True

                if not IsInt(ligne.pointTo.get()) :
                    ligne.pointTo.config(bg=WarningColor)
                    errors=True

                if not IsInt(ligne.loadMin.get()) :
                    ligne.loadMin.config(bg=WarningColor)
                    errors=True

                if not ( IsInt(ligne.loadMax.get())  or ligne.loadMax.get()=='' ):
                    ligne.loadMax.config(bg=WarningColor)
                    errors=True

                if not ( IsInt(ligne.drybox.get())  or ligne.drybox.get()=='' ):
                    ligne.drybox.config(bg=WarningColor)
                    errors=True

                if not ( IsInt(ligne.flatbed.get())  or ligne.flatbed.get()=='' ):
                    ligne.flatbed.config(bg=WarningColor)
                    errors=True

                # if ligne.flatbed.get() != '' and ligne.drybox.get() != '' and ligne.loadMax.get()!='':
                #     if int(ligne.flatbed.get()) + int(ligne.drybox.get()) < int(ligne.loadMax.get()):
                #         errors = True
                #         ligne.loadMax.config(bg=WarningColor)
                #         ligne.drybox.config(bg=WarningColor)
                #         ligne.flatbed.config(bg=WarningColor)


                if  IsInt(ligne.loadMin.get()) and IsInt(ligne.loadMax.get()):
                    if int(ligne.loadMin.get()) > int(ligne.loadMax.get()):
                        ligne.loadMin.config(bg=WarningColor)
                        ligne.loadMax.config(bg=WarningColor)
                        errors=True

                if not IsInt(ligne.priority.get()) :
                    ligne.priority.config(bg=WarningColor)
                    errors=True

                if not IsInt(ligne.transit.get()):
                    ligne.transit.config(bg=WarningColor)
                    errors=True

                if Project[0] == 'P2P':
                    DATA_TO_SEND.append([ligne.pointFrom.get(),ligne.pointTo.get(),ligne.loadMin.get(),ligne.loadMax.get(),ligne.drybox.get(),ligne.flatbed.get(),ligne.priority.get(),ligne.transit.get(),ligne.lineToSKip.get()])
                else:
                    DATA_TO_SEND.append(
                        [ligne.pointFrom.get(), ligne.pointTo.get(), ligne.loadMin.get(), ligne.loadMax.get(),
                         ligne.drybox.get(), ligne.flatbed.get(), ligne.priority.get(), ligne.transit.get(), ligne.days.get(),
                         ligne.lineToSKip.get()])

        if not errors:
            timeOfExport = pd.datetime.now()
            #delete in SQL, only keep the last 3 modifications in history
            if Project[0] == 'P2P':
                self.SQL.deleteFromSQL( "IMPORT_DATE not in (SELECT  DISTINCT top(3) IMPORT_DATE FROM [Business_Planning].[dbo].[OTD_1_P2P_F_PARAMETERS] ORDER BY IMPORT_DATE desc) ")
            else:
                self.SQL.deleteFromSQL(
                    "IMPORT_DATE not in (SELECT  DISTINCT top(3) IMPORT_DATE FROM [Business_Planning].[dbo].[OTD_1_P2P_F_FORECAST_PARAMETERS] ORDER BY IMPORT_DATE desc) ")

            #envoit du nouveau data dans SQL
            for DATALine in DATA_TO_SEND:
                for obj in range(2,len(DATALine)-1):
                    if DATALine[obj] =='':
                        DATALine[obj]=None
                    else:
                        DATALine[obj] = int(DATALine[obj])

                if DATALine[-1] == True:
                    DATALine[-1]=1
                else:
                    DATALine[-1]=0

                DATALine.append(timeOfExport)
                self.SQL.sendToSQL( [(DATALine)])


            global ToExecute
            ToExecute=[True]

            self.master.destroy()
        else:
            messagebox.showerror('Invalid DATA','There are some invalid inputs in the table')


    def AddNew(self,largeurColonne):
        "To add a new line of values"
        keyF =frame(self.verticalBar, TOP)
        global Project
        if Project[0] == 'P2P':
            self.lignes.append(ligne(self,keyF,LEFT,'0000','0000',0,'','','',0,0,False,-1,largeurColonne))
        else:
            self.lignes.append(ligne(self,keyF,LEFT,'0000','0000',0,'','','',0, 0, False,0, largeurColonne))


def changeEmail():
    "To change email list"


    class VerticalScrolledFrame(tk.Frame):
        #SQL = ''
        #headers= ''
        def __init__(self, parent, *args, **kw):
            tk.Frame.__init__(self, parent, *args, **kw)
            #self.headers='EMAIL_ADDRESS'
            #self.SQL = SQLConnection('CAVLSQLPD2\pbi2', 'Business_Planning','OTD_1_P2P_F_PARAMETERS_EMAIL_ADDRESS',self.headers)

            # create a canvas object and a vertical scrollbar for scrolling it
            vscrollbar = tk.Scrollbar(self, orient=tk.VERTICAL)
            vscrollbar.pack(fill=tk.Y, side=tk.RIGHT, expand=tk.FALSE)
            canvas = tk.Canvas(self, bd=0, highlightthickness=0,
                            yscrollcommand=vscrollbar.set)
            canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=tk.TRUE)
            vscrollbar.config(command=canvas.yview)

            # reset the view
            canvas.xview_moveto(0)
            canvas.yview_moveto(0)

            # create a frame inside the canvas which will be scrolled with it
            self.interior = interior = tk.Frame(canvas)
            interior_id = canvas.create_window(0, 0, window=interior,
                                               anchor=tk.NW)

            # track changes to the canvas and frame width and sync them,
            # also updating the scrollbar
            def _configure_interior(event):
                # update the scrollbars to match the size of the inner frame
                size = (interior.winfo_reqwidth(), interior.winfo_reqheight())
                canvas.config(scrollregion="0 0 %s %s" % size)
                if interior.winfo_reqwidth() != canvas.winfo_width():
                    # update the canvas's width to fit the inner frame
                    canvas.config(width=interior.winfo_reqwidth())

            interior.bind('<Configure>', _configure_interior)

            def _configure_canvas(event):
                if interior.winfo_reqwidth() != canvas.winfo_width():
                    # update the inner frame's width to fill the canvas
                    canvas.itemconfigure(interior_id, width=canvas.winfo_width())
            canvas.bind('<Configure>', _configure_canvas)


    root = tk.Tk()
    root.title("EMAIL ADDRESS")
    root.configure(background="gray99")
    root.geometry("400x500+700+300")

    scframe = VerticalScrolledFrame(root)
    scframe.pack()

    tempo = [] #list of all buttons
    #Get all email address
    headers='EMAIL_ADDRESS,PROJECT'
    SQL = SQLConnection('CAVLSQLPD2\pbi2', 'Business_Planning','OTD_1_P2P_F_PARAMETERS_EMAIL_ADDRESS',headers)
    global Project
    SQLquery = """ SELECT DISTINCT [EMAIL_ADDRESS] FROM [Business_Planning].[dbo].[OTD_1_P2P_F_PARAMETERS_EMAIL_ADDRESS] WHERE PROJECT = '{0}'""".format(Project[0])
    lis= [ item for sublist in SQL.GetSQLData(SQLquery) for item in sublist]


    btn = tk.Button(scframe.interior, height=1, width=40, relief=tk.FLAT,
            bg="gray99", fg="purple3",
            font="Dosis", text='ADD NEW',
            command=lambda  lis=lis: addNew(lis))
    btn.pack(padx=10, pady=5, side=tk.TOP)

    tk.Label(scframe.interior, text="CLICK ON EMAIL TO DELETE").pack()


    for i, x in enumerate(lis):
        btn = tk.Button(scframe.interior, height=1, width=40, relief=tk.FLAT,
            bg="gray99", fg="purple3",
            font="Dosis", text=lis[i],
            command=lambda i=i,x=x: openlink(i))
        btn.pack(padx=10, pady=5, side=tk.TOP)
        tempo.append(btn)

    def openlink(i):
        "To delete an email address"
        confirm = messagebox.askokcancel(lis[i],"Are you sure you want to remove {0} from the list?".format(lis[i]))
        if confirm:
            SQL.deleteFromSQL( "[EMAIL_ADDRESS] = ('{0}') and PROJECT = '{1}'".format(lis[i],Project[0]))
            tempo[i].forget()
            lis.pop(i)
           #Delete from SQL
    def addNew(lis):

        def addNewEmail():
            "Add new email address"
            newEmail =  [e1.get().lower()]
            if newEmail[0] not in lis:
                lis.append(newEmail[0])
                longueur = len(lis)-1

                btn = tk.Button(scframe.interior, height=1, width=40, relief=tk.FLAT, bg="gray99", fg="purple3",font="Dosis", text= newEmail[0] ,command=lambda : openlink(longueur))
                btn.pack(padx=10, pady=5, side=tk.TOP)
                tempo.append(btn)
                SQL.sendToSQL( [(newEmail[0], Project[0])])

            e1.delete(0,END)

        master = tk.Tk()
        master.geometry("500x80+700+500")
        master.title("ADD NEW EMAIL ADDRESS")
        tk.Label(master,
                 text="EMAIL ADDRESS").grid(row=0)


        e1 = tk.Entry(master, width=50)
        e1.grid(row=0, column=1)

        tk.Button(master,text='Cancel',command=master.destroy).grid(row=3, column=0,sticky=tk.W, pady=4)
        tk.Button(master,text='ADD', command= addNewEmail).grid(row=3,column=1,sticky=tk.W, pady=4)

        master.mainloop()





    root.mainloop()


def MissingP2PBox(MissingP2P):
    ToExecute[0]=False
    errorMess = 'Missing these P2P (point_from - point_to) : \n'
    for p2p in MissingP2P:
        errorMess += str(p2p ) + '\n'
    #root = tk.Tk()
    #root.withdraw()
    return messagebox.askokcancel("Warning",errorMess + '\n Continue execution? \n')
    #messagebox.showwarning("Warning", errorMess)
    #root.destroy()

def OpenParameters(projectName = 'P2P'):
    global Project
    Project = [projectName]
    Box(largeurColonne).mainloop()
    return ToExecute[0]


if __name__ == '__main__':
    OpenParameters()
    print(ToExecute[0])