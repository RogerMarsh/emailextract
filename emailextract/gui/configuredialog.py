# configuredialog.py
# Copyright 2013 Roger Marsh
# Licence: See LICENCE (BSD licence)

"""A simple configuration file text editor.

"""

import tkinter
import tkinter.messagebox

from solentware_misc.gui.exceptionhandler import (
    ExceptionHandler,
    BAD_WINDOW,
    FOCUS_ERROR,
)


class ConfigureDialog(ExceptionHandler):

    """Configuration file text editor in a dialogue.

    Update methods are defined but do not change database.  Subclasses must
    override as needed.

    """

    def __init__(
        self,
        parent,
        configuration="",
        dialog_title="Text editor dialogue",
        dialog_cancel_hint="Quit without applying changes",
        dialog_update_hint="Apply changes",
        cnf=dict(),
        **kargs
    ):
        """Create a onfiguration file text editor dialogue."""

        super().__init__()
        self._config_text = None
        self.dialog = tkinter.Toplevel(parent)
        self.restore_focus = self.dialog.focus_get()
        self.dialog.wm_title(dialog_title)
        self.configuration = tkinter.Text(master=self.dialog)
        self.configuration.insert(tkinter.END, configuration)
        buttons_frame = tkinter.Frame(master=self.dialog)
        buttons_frame.pack(side=tkinter.BOTTOM, fill=tkinter.X)

        buttonrow = buttons_frame.pack_info()["side"] in ("top", "bottom")
        for i, b in enumerate(
            (
                ("Cancel", dialog_cancel_hint, True, 0, self.on_cancel),
                ("Update", dialog_update_hint, True, 0, self.on_update),
            )
        ):
            button = tkinter.Button(
                master=buttons_frame,
                text=b[0],
                underline=b[3],
                command=self.try_command(b[4], buttons_frame),
            )
            if buttonrow:
                buttons_frame.grid_columnconfigure(i * 2, weight=1)
                button.grid_configure(column=i * 2 + 1, row=0)
            else:
                buttons_frame.grid_rowconfigure(i * 2, weight=1)
                button.grid_configure(row=i * 2 + 1, column=0)
        if buttonrow:
            buttons_frame.grid_columnconfigure(len(b * 2), weight=1)
        else:
            buttons_frame.grid_rowconfigure(len(b * 2), weight=1)

        self.configuration.pack(
            side=tkinter.TOP, fill=tkinter.BOTH, expand=tkinter.TRUE
        )

        self.dialog.wait_visibility()
        self.dialog.grab_set()
        self.dialog.wait_window()

    @property
    def config_text(self):
        """Return True if tkMessageBox.askyesno closed by clicking Yes"""
        return self._config_text

    def on_cancel(self, event=None):
        """Show dialogue to confirm cancellation of edit"""
        if tkinter.messagebox.askyesno(
            parent=self.dialog,
            message="Confirm cancellation of edit",
            title=self.dialog.wm_title(),
        ):
            self.dialog.destroy()
        else:
            self.dialog.tkraise()

    def on_update(self, event=None):
        """Extract text from dialog and destroy dialog"""
        if tkinter.messagebox.askyesno(
            parent=self.dialog,
            message="".join(
                (
                    "Confirm apply changes to configuration file and update "
                    "main widget."
                )
            ),
            title=self.dialog.wm_title(),
        ):
            self._config_text = self.configuration.get(
                "1.0", " ".join((tkinter.END, "-1 chars"))
            )
            self.dialog.destroy()
        else:
            self._config_text = None
            self.dialog.tkraise()

    def __del__(self):
        """Restore focus to widget with focus before modal interaction."""
        try:
            # restore focus on dismissing dialogue.
            self.restore_focus.focus_set()
        except tkinter._tkinter.TclError as error:
            # application destroyed while confirm dialogue exists.
            if str(error) != FOCUS_ERROR:
                if not str(error).startswith(BAD_WINDOW):
                    raise
