"""
This file is used to generate option box at the beginning of P2P program

Author : Nicolas Raymond

"""

from tkinter import *

import sys
import os

sys.path.append(os.path.dirname(sys.argv[0]))  # We add current directory (of Load_Process) to sys.path
parent_dir = os.path.dirname(os.path.dirname(sys.argv[0]))
sys.path.append(os.path.join(parent_dir, 'LoadBuilding'))  # We add directory of LoadBuilding to sys.path

from P2PFullProcess import p2p_full_process

class ModeBox:

    def __init__(self, master):
        """
        Generates a box with both loading options

        :param master: root of the tkinter window
        """

        self.master = master
        master.title('Optimus')

        # Label initialization
        self.title = Label(self.master, text='Optimus Mode Selection')

        # Label positioning
        self.title.grid(row=0, column=0, columnspan=2)

        # Buttons initialization
        self.first_option = Button(self.master, text='P2P Full Process', padx=50, pady=20,
                                   borderwidth=3, command=self.run_full_process)
        self.second_option = Button(self.master, text='Fast Loads', padx=50, pady=20,
                                    borderwidth=3, command=self.run_fast_loads)
        # Buttons positioning
        self.first_option.grid(row=1, column=0)
        self.second_option.grid(row=1, column=1)

    def run_full_process(self):

        # We close the window
        self.master.destroy()

        # We run the full process function
        p2p_full_process()

    def run_fast_loads(self):
        pass


def open_mode_box():
    root = Tk()
    mode_box = ModeBox(root)
    root.mainloop()


if __name__ == '__main__':
    open_mode_box()