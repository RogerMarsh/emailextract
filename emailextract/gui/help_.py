# help_.py
# Copyright 2017 Roger Marsh
# Licence: See LICENCE (BSD licence)

"""Functions to create Help widgets for emailextract."""

import tkinter

from solentware_misc.gui.help_ import help_widget

from .. import help_


def help_about(master):
    """Display About document."""
    help_widget(master, help_.ABOUT, help_)


def help_guide(master):
    """Display Guide document."""
    help_widget(master, help_.GUIDE, help_)


def help_notes(master):
    """Display Notes document."""
    help_widget(master, help_.NOTES, help_)


if __name__ == "__main__":
    # Display all help documents without running Emailextract application

    root = tkinter.Tk()
    help_about(root)
    help_guide(root)
    help_notes(root)
    root.mainloop()
