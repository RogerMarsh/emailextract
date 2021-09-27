# help.py
# Copyright 2017 Roger Marsh
# Licence: See LICENCE (BSD licence)

"""Functions to create Help widgets for emailextract.

"""

import tkinter

import solentware_misc.gui.textreadonly
from solentware_misc.gui.help import help_widget

from .. import help


def help_about(master):
    """Display About document"""

    help_widget(master, help.ABOUT, help)


def help_guide(master):
    """Display Guide document"""

    help_widget(master, help.GUIDE, help)


def help_notes(master):
    """Display Notes document"""

    help_widget(master, help.NOTES, help)


if __name__=='__main__':
    #Display all help documents without running Emailextract application

    root = tkinter.Tk()
    help_about(root)
    help_guide(root)
    help_notes(root)
    root.mainloop()
