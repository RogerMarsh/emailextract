# select.py
# Copyright 2014 Roger Marsh
# Licence: See LICENCE (BSD licence)

"""Email text exraction filter User Interface.

"""

import os
import tkinter
import tkinter.messagebox
import tkinter.filedialog
from email.utils import parseaddr, parsedate_tz
from time import strftime

from solentware_misc.gui.exceptionhandler import ExceptionHandler
from solentware_misc.gui import textreadonly

from . import help
from .. import APPLICATION_NAME
from .configuredialog import ConfigureDialog
from ..core.emailextractor import (
    EmailExtractor,
    EmailExtractorError,
    IGNORE_EMAIL,
    COLLECTED,
    EXTRACTED,
    MEDIA_TYPES,
    )

startup_minimum_width = 340
startup_minimum_height = 400


class SelectError(Exception):
    pass


class Select(ExceptionHandler):
    
    """Define and use an email select and store configuration file.
    
    """

    def __init__(self,
                 folder=None,
                 use_toplevel=False,
                 emailextractor=None,
                 **kargs):
        """Create the database and GUI objects.

        **kargs - passed to tkinter Toplevel widget if use_toplevel True

        """
        if use_toplevel:
            self.root = tkinter.Toplevel(**kargs)
        else:
            self.root = tkinter.Tk()
        try:
            if emailextractor:
                self._emailextractor = emailextractor
            else:
                self._emailextractor = EmailExtractor
            if folder is not None:
                self.root.wm_title(' - '.join((APPLICATION_NAME, folder)))
            else:
                self.root.wm_title(APPLICATION_NAME)
            self.root.wm_minsize(
                width=startup_minimum_width, height=startup_minimum_height)
            
            self._configuration = None
            self._configuration_edited = False
            self._email_collector = None
            self._tag_names = set()

            menubar = tkinter.Menu(self.root)

            menufile = tkinter.Menu(menubar, name='file', tearoff=False)
            menubar.add_cascade(label='File', menu=menufile, underline=0)
            menufile.add_command(
                label='Open',
                underline=0,
                command=self.try_command(self.file_open, menufile))
            menufile.add_command(
                label='New',
                underline=0,
                command=self.try_command(self.file_new, menufile))
            menufile.add_separator()
            #menufile.add_command(
            #    label='Save',
            #    underline=0,
            #    command=self.try_command(self.file_save, menufile))
            menufile.add_command(
                label='Save Copy As...',
                underline=7,
                command=self.try_command(self.file_save_copy_as, menufile))
            menufile.add_separator()
            menufile.add_command(
                label='Close',
                underline=0,
                command=self.try_command(self.file_close, menufile))
            menufile.add_separator()
            menufile.add_command(
                label='Quit',
                underline=0,
                command=self.try_command(self.file_quit, menufile))

            menuactions = tkinter.Menu(menubar, name='actions', tearoff=False)
            menubar.add_cascade(label='Actions', menu=menuactions, underline=0)
            menuactions.add_command(
                label='Source emails',
                underline=0,
                command=self.try_command(self.show_email_source, menuactions))
            menuactions.add_command(
                label='Decoded text',
                underline=0,
                command=self.try_command(self.show_decoded_text, menuactions))
            menuactions.add_command(
                label='Extracted text',
                underline=0,
                command=self.try_command(self.show_extracted_text, menuactions))
            menuactions.add_command(
                label='Update',
                underline=0,
                command=self.try_command(
                    self.update_difference_files, menuactions))
            menuactions.add_command(
                label='Clear selection',
                underline=0,
                command=self.try_command(self.clear_selection, menuactions))
            menuactions.add_separator()
            menuactions.add_command(
                label='Option editor',
                underline=0,
                command=self.try_command(
                    self.configure_email_selection, menuactions))

            menuhelp = tkinter.Menu(menubar, name='help', tearoff=False)
            menubar.add_cascade(label='Help', menu=menuhelp, underline=0)
            menuhelp.add_command(
                label='Guide',
                underline=0,
                command=self.try_command(self.help_guide, menuhelp))
            menuhelp.add_command(
                label='Notes',
                underline=0,
                command=self.try_command(self.help_notes, menuhelp))
            menuhelp.add_command(
                label='About',
                underline=0,
                command=self.try_command(self.help_about, menuhelp))

            self.root.configure(menu=menubar)

            self.statusbar = Statusbar(self.root)
            frame = tkinter.PanedWindow(
                self.root,
                background='cyan2',
                opaqueresize=tkinter.FALSE,
                orient=tkinter.HORIZONTAL)
            frame.pack(fill=tkinter.BOTH, expand=tkinter.TRUE)

            toppane = tkinter.PanedWindow(
                master=frame,
                opaqueresize=tkinter.FALSE,
                orient=tkinter.HORIZONTAL)
            originalpane = tkinter.PanedWindow(
                master=toppane,
                opaqueresize=tkinter.FALSE,
                orient=tkinter.VERTICAL)
            emailpane = tkinter.PanedWindow(
                master=toppane,
                opaqueresize=tkinter.FALSE,
                orient=tkinter.VERTICAL)
            self.configctrl = textreadonly.make_text_readonly(
                master=originalpane, width=80)
            self.emaillistctrl = textreadonly.make_text_readonly(
                master=originalpane, width=80)
            self.emailtextctrl = textreadonly.make_text_readonly(
                master=emailpane)
            originalpane.add(self.configctrl)
            originalpane.add(self.emaillistctrl)
            emailpane.add(self.emailtextctrl)
            toppane.add(originalpane)
            toppane.add(emailpane)
            toppane.pack(
                side=tkinter.TOP, expand=True, fill=tkinter.BOTH)
            for widget, sequence, function in (
                (self.configctrl, '<ButtonPress-3>', self.conf_popup),
                (self.emaillistctrl, '<ButtonPress-3>', self.list_popup),
                (self.emailtextctrl, '<ButtonPress-3>', self.text_popup),
                ):
                widget.bind(sequence, self.try_event(function))
            self._folder = folder
            self._most_recent_action = None

        except:
            self.root.destroy()
            del self.root

    def __del__(self):
        """"""
        if self._configuration:
            self._configuration = None

    def help_about(self):
        """Display information about EmailExtract."""
        help.help_about(self.root)

    def help_guide(self):
        """Display brief User Guide for EmailExtract."""
        help.help_guide(self.root)

    def help_notes(self):
        """Display technical notes about EmailExtract."""
        help.help_notes(self.root)

    def get_toplevel(self):
        """Return the toplevel widget."""
        return self.root

    def file_new(self):
        """Create and open a new league secretary database."""
        if self._configuration is not None:
            tkinter.messagebox.showinfo(
                parent=self.get_toplevel(),
                title='New Extraction Rules',
                message='Close the current extraaction rules first.')
            return
        config_file = tkinter.filedialog.asksaveasfilename(
            parent=self.get_toplevel(),
            title='New Extraction Rules',
            defaultextension='.conf',
            filetypes=(
                ('Extraction Rules', '*.conf'),
                ),
            initialfile='',
            initialdir='~')
        if not config_file:
            return
        if tkinter.messagebox.askquestion(
            parent=self.get_toplevel(),
            title='Close',
            message=''.join(
                ("Do you want to specify a directory containing CSV files of ",
                 "Media Types (like 'text/csv').\n\nOffical files are ",
                 "available at 'https://www.iana.org'.",
                 ))) == tkinter.messagebox.YES:
            media_types_directory = tkinter.filedialog.askdirectory(
                parent=self.get_toplevel(),
                title='Media Types CSV files',
                initialdir='~')
        else:
            media_types_directory = ''
        self.configctrl.delete('1.0', tkinter.END)
        self.configctrl.insert(
            tkinter.END, ''.join(
                ('# ', os.path.basename(config_file),
                 ' extraction rules')) + os.linesep)
        if media_types_directory:
            self.configctrl.insert(
                tkinter.END,
                ' '.join((MEDIA_TYPES, media_types_directory)) + os.linesep)
        self.configctrl.insert(
            tkinter.END, ' '.join((COLLECTED, COLLECTED)) + os.linesep)
        self.configctrl.insert(
            tkinter.END, ' '.join((EXTRACTED, EXTRACTED)) + os.linesep)
        fn = open(config_file, 'w', encoding='utf8')
        try:
            fn.write(self.configctrl.get(
                '1.0', ' '.join((tkinter.END, '-1 chars'))))
        finally:
            fn.close()
        self._configuration = config_file
        self._folder = os.path.dirname(config_file)

    def file_open(self):
        """Open an existing extraction rules file."""
        if self._configuration is not None:
            tkinter.messagebox.showinfo(
                parent=self.get_toplevel(),
                title='Extraction Rules',
                message='Close the current extraction rules first.')
            return
        config_file = tkinter.filedialog.askopenfilename(
            parent=self.get_toplevel(),
            title='Open Extraction Rules',
            defaultextension='.conf',
            filetypes=(
                ('Extraction Rules', '*.conf'),
                ),
            initialfile='',
            initialdir='~')
        if not config_file:
            return
        fn = open(config_file, 'r', encoding='utf8')
        try:
            self.configctrl.delete('1.0', tkinter.END)
            self.configctrl.insert(tkinter.END, fn.read())
        finally:
            fn.close()
        self._configuration = config_file
        self._folder = os.path.dirname(config_file)

    def file_close(self):
        """Close the open extraction rules file."""
        if self._configuration is None:
            tkinter.messagebox.showinfo(
                parent=self.get_toplevel(),
                title='Extraction Rules',
                message='Cannot close.\n\nThere is no database open.')
            return
        closemsg = 'Confirm Close.\n\nChanges not already saved will be lost.'
        dlg = tkinter.messagebox.askquestion(
            parent=self.get_toplevel(),
            title='Close',
            message=closemsg)
        if dlg == tkinter.messagebox.YES:
            self._clear_email_tags()
            self.configctrl.delete('1.0', tkinter.END)
            self.emailtextctrl.delete('1.0', tkinter.END)
            self.emaillistctrl.delete('1.0', tkinter.END)
            self.statusbar.set_status_text()
            self._configuration = None
            self._configuration_edited = False
            self._email_collector = None

    def file_quit(self):
        """Quit the extraction application."""
        quitmsg = 'Confirm Quit.\n\nChanges not already saved will be lost.'
        dlg = tkinter.messagebox.askquestion(
            parent=self.get_toplevel(),
            title='Quit',
            message=quitmsg)
        if dlg == tkinter.messagebox.YES:
            self.root.destroy()

    # Probably not going to be used because 'Actions | Option editor' does it.
    def file_save(self):
        """Save the open extraction rules file."""
        if self._configuration is None:
            tkinter.messagebox.showinfo(
                parent=self.get_toplevel(),
                title='Save',
                message='Cannot save.\n\nExtraction rules file not open.')
            return
        if tkinter.messagebox.askquestion(
            parent=self.get_toplevel(),
            title='Save',
            message=''.join(
                ('Confirm save extraction rules to\n',
                 self._configuration,
                 )),
            ) != tkinter.messagebox.YES:
            return
        fn = open(self._configuration, 'w')
        try:
            fn.write(self.configctrl.get(
                '1.0', ' '.join((tkinter.END, '-1 chars'))))
            self._clear_email_tags()
            self.emailtextctrl.delete('1.0', tkinter.END)
            self.emaillistctrl.delete('1.0', tkinter.END)
            self._configuration_edited = False
            self._email_collector = None
        finally:
            fn.close()
        return True

    def file_save_copy_as(self):
        """Save copy of open extraction rules and keep current open."""
        if self._configuration is None:
            tkinter.messagebox.showinfo(
                parent=self.get_toplevel(),
                title='Save Copy As',
                message='Cannot save.\n\nExtraction rules file not open.')
            return
        config_file = tkinter.filedialog.asksaveasfilename(
            parent=self.get_toplevel(),
            title='Save Extraction rules As',
            defaultextension='.conf',
            filetypes=(
                ('Extraction Rules', '*.conf'),
                ),
            initialfile=os.path.basename(self._configuration),
            initialdir=os.path.dirname(self._configuration))
        if not config_file:
            return
        if config_file == self._configuration:
            tkinter.messagebox.showinfo(
                parent=self.get_toplevel(),
                title='Save Copy As',
                message=''.join(
                    ('Cannot use "Save Copy As" to overwite the open ',
                     'extraction rules file.')),
                )
            return
        fn = open(config_file, 'w')
        try:
            fn.write(self.configctrl.get(
                '1.0', ' '.join((tkinter.END, '-1 chars'))))
        finally:
            fn.close()

    def configure_email_selection(self):
        """Set parameters that control extraction from emails."""
        if self._configuration is None:
            tkinter.messagebox.showinfo(
                parent=self.get_toplevel(),
                title='Configure Extraction',
                message='Open an extraction rules file.')
            return
        config_text = ConfigureDialog(
            self.root,
            self.configctrl.get(
                '1.0', ' '.join((tkinter.END, '-1 chars')))).config_text
        if config_text is None:
            return
        self._configuration_edited = True
        self.configctrl.delete('1.0', tkinter.END)
        self.configctrl.insert(tkinter.END, config_text)
        fn = open(self._configuration, 'w', encoding='utf-8')
        try:
            fn.write(config_text)
            self._clear_email_tags()
            self.emailtextctrl.delete('1.0', tkinter.END)
            self.emaillistctrl.delete('1.0', tkinter.END)
            self.statusbar.set_status_text()
            self._configuration_edited = False
            self._email_collector = None
        finally:
            fn.close()
        if self._most_recent_action:
            self._most_recent_action()

    def show_email_source(self):
        """Do the text extraction but do not copy the emails."""
        if self._configuration is None:
            tkinter.messagebox.showinfo(
                parent=self.get_toplevel(),
                title='Show Extraction Source',
                message='Open a extraction rules file')
            return
        if self._configuration_edited:
            tkinter.messagebox.showinfo(
                parent=self.get_toplevel(),
                title='Show Extraction Source',
                message=''.join((
                    'The edited configuration file has not been saved. ',
                    'It must be saved before "Show" action can be done.')))
            return
        if self._email_collector is None:
            emc = self._emailextractor(
                self._folder,
                configuration=self.configctrl.get(
                    '1.0', ' '.join((tkinter.END, '-1 chars'))),
                parent=self.get_toplevel())
            if not emc.parse():
                return
            if not emc.selected_emails:
                tkinter.messagebox.showinfo(
                    parent=self.get_toplevel(),
                    title='Show Extraction Source',
                    message='No emails match the selection rules.')
                return
            self._email_collector = emc
        self._show_email_source()
        self._most_recent_action = self.show_email_source

    def _show_email_source(self):
        """Populate widgets with email source."""
        self._clear_email_tags()
        tw = self.emailtextctrl
        lw = self.emaillistctrl
        tw.delete('1.0', tkinter.END)
        lw.delete('1.0', tkinter.END)

        # Tag the text put in the widgets such that the source entry in
        # selected_emails_text can be recovered from the pointer position
        # over the widget.
        tags = self._tag_names
        content_type_headers = (
            'Content-Type', 'Content-Transfer-Encoding', 'Content-Disposition')
        for e, em in enumerate(self._email_collector.selected_emails):
            m = em.message
            textname = 'x'.join(('T', str(e)))
            tags.add(textname)
            entryname = 'x'.join(('M', str(e)))
            tags.add(entryname)
            fromname = 'x'.join(('F', str(e)))
            tags.add(fromname)
            start = tw.index(tkinter.INSERT)
            tw.insert(tkinter.END, m.as_string())
            tw.insert(tkinter.END, '\n')
            tw.tag_add(textname, start, tw.index(tkinter.INSERT))
            tw.insert(tkinter.END, '\n\n\n')
            tw.tag_add(entryname, start, tw.index(tkinter.INSERT))
            start = lw.index(tkinter.INSERT)
            fromstart = lw.index(tkinter.INSERT)
            lw.insert(tkinter.END, m.get('From', ''))
            lw.insert(tkinter.END, '\n')
            lw.insert(tkinter.END, m.get('Date', ''))
            lw.tag_add(fromname, fromstart, lw.index(tkinter.INSERT))
            lw.insert(tkinter.END, '\n')
            lw.insert(tkinter.END, m.get('Subject', ''))
            lw.insert(tkinter.END, '\n')
            for p in m.walk():
                v = p.get_content_type()
                if v is not None:
                    lw.insert(tkinter.END, v)
                    lw.insert(tkinter.END, '\n')
            lw.tag_add(textname, start, lw.index(tkinter.INSERT))
            lw.insert(tkinter.END, '\n\n')
            lw.tag_add(entryname, start, lw.index(tkinter.INSERT))

    def _show_decoded_text(self):
        """Populate wigdets with decoded text."""
        self._clear_email_tags()
        tw = self.emailtextctrl
        lw = self.emaillistctrl
        tw.delete('1.0', tkinter.END)
        lw.delete('1.0', tkinter.END)

        # Tag the text put in the widgets such that the source entry in
        # selected_emails_text can be recovered from the pointer position
        # over the widget.
        tags = self._tag_names
        content_type_headers = (
            'Content-Type', 'Content-Transfer-Encoding', 'Content-Disposition')
        for e, em in enumerate(self._email_collector.selected_emails):
            m = em.message
            met = em.encoded_text
            textname = 'x'.join(('T', str(e)))
            tags.add(textname)
            entryname = 'x'.join(('M', str(e)))
            tags.add(entryname)
            fromname = 'x'.join(('F', str(e)))
            tags.add(fromname)
            fromstart = tw.index(tkinter.INSERT)
            tw.insert(tkinter.END, m.get('From', ''))
            tw.insert(tkinter.END, '\n')
            tw.insert(tkinter.END, m.get('Date', ''))
            tw.insert(tkinter.END, '\n\n')
            tw.tag_add(fromname, fromstart, tw.index(tkinter.INSERT))
            start = tw.index(tkinter.INSERT)
            tw.insert(tkinter.END, b'\n\n'.join(met))
            tw.insert(tkinter.END, '\n')
            tw.tag_add(textname, start, tw.index(tkinter.INSERT))
            tw.insert(tkinter.END, '\n\n\n')
            tw.tag_add(entryname, fromstart, tw.index(tkinter.INSERT))
            start = lw.index(tkinter.INSERT)
            fromstart = lw.index(tkinter.INSERT)
            lw.insert(tkinter.END, m.get('From', ''))
            lw.insert(tkinter.END, '\n')
            lw.insert(tkinter.END, m.get('Date', ''))
            lw.tag_add(fromname, fromstart, lw.index(tkinter.INSERT))
            lw.insert(tkinter.END, '\n')
            lw.insert(tkinter.END, m.get('Subject', ''))
            lw.insert(tkinter.END, '\n')
            for p in m.walk():
                v = p.get_content_type()
                if v is not None:
                    lw.insert(tkinter.END, v)
                    lw.insert(tkinter.END, '\n')
            lw.tag_add(textname, start, lw.index(tkinter.INSERT))
            lw.insert(tkinter.END, '\n\n')
            lw.tag_add(entryname, start, lw.index(tkinter.INSERT))

    def show_decoded_text(self):
        """Do the email selection and show decoded text for non-human readable
        parts in message.

        Parts with Content-Transfer-Encoding set to base64 in other words.

        The possible encodings are base64, quoted-printable, 8bit, 7bit,
        binary, and x-token.
        """
        if self._configuration is None:
            tkinter.messagebox.showinfo(
                parent=self.get_toplevel(),
                title='Show Decoded Text',
                message='Open a text extraction rules file')
            return
        if self._configuration_edited:
            tkinter.messagebox.showinfo(
                parent=self.get_toplevel(),
                title='Show Decoded Text',
                message=''.join((
                    'The edited configuration file has not been saved. ',
                    'It must be saved before "Show" action can be done.')))
            return
        if self._email_collector is None:
            emc = self._emailextractor(
                self._folder,
                configuration=self.configctrl.get(
                    '1.0', ' '.join((tkinter.END, '-1 chars'))),
                parent=self.get_toplevel())
            if not emc.parse():
                return
            if not emc.selected_emails:
                tkinter.messagebox.showinfo(
                    parent=self.get_toplevel(),
                    title='Show Decoded Text',
                    message='No emails match the selection rules.')
                return
            self._email_collector = emc
        try:
            self._show_decoded_text()
        except EmailExtractorError:
            tkinter.messagebox.showinfo(
                parent=self.get_toplevel(),
                title='Show Decoded Text',
                message=''.join((
                    'KeyError exception occurred, probably due to missing ',
                    'or incorrect entry in configuration file.',
                    )))
            return
        self._most_recent_action = self.show_decoded_text
        return True

    def _show_extracted_text(self):
        """Populate widgets with text extracted from email."""
        self._clear_email_tags()
        tw = self.emailtextctrl
        lw = self.emaillistctrl
        tw.delete('1.0', tkinter.END)
        lw.delete('1.0', tkinter.END)

        # Tag the text put in the widgets such that the source entry in
        # selected_emails_text can be recovered from the pointer position
        # over the widget.
        tags = self._tag_names
        content_type_headers = (
            'Content-Type', 'Content-Transfer-Encoding', 'Content-Disposition')
        for e, em in enumerate(self._email_collector.selected_emails):
            m = em.message
            met = em.extracted_text
            textname = 'x'.join(('T', str(e)))
            tags.add(textname)
            entryname = 'x'.join(('M', str(e)))
            tags.add(entryname)
            fromname = 'x'.join(('F', str(e)))
            tags.add(fromname)
            fromstart = tw.index(tkinter.INSERT)
            tw.insert(tkinter.END, m.get('From', ''))
            tw.insert(tkinter.END, '\n')
            tw.insert(tkinter.END, m.get('Date', ''))
            tw.insert(tkinter.END, '\n\n')
            tw.tag_add(fromname, fromstart, tw.index(tkinter.INSERT))
            start = tw.index(tkinter.INSERT)
            tw.insert(tkinter.END, '\n\n'.join(met))
            tw.insert(tkinter.END, '\n')
            tw.tag_add(textname, start, tw.index(tkinter.INSERT))
            tw.insert(tkinter.END, '\n\n\n')
            tw.tag_add(entryname, fromstart, tw.index(tkinter.INSERT))
            start = lw.index(tkinter.INSERT)
            fromstart = lw.index(tkinter.INSERT)
            lw.insert(tkinter.END, m.get('From', ''))
            lw.insert(tkinter.END, '\n')
            lw.insert(tkinter.END, m.get('Date', ''))
            lw.tag_add(fromname, fromstart, lw.index(tkinter.INSERT))
            lw.insert(tkinter.END, '\n')
            lw.insert(tkinter.END, m.get('Subject', ''))
            lw.insert(tkinter.END, '\n')
            for p in m.walk():
                v = p.get_content_type()
                if v is not None:
                    lw.insert(tkinter.END, v)
                    lw.insert(tkinter.END, '\n')
            lw.tag_add(textname, start, lw.index(tkinter.INSERT))
            lw.insert(tkinter.END, '\n\n')
            lw.tag_add(entryname, start, lw.index(tkinter.INSERT))

    def show_extracted_text(self):
        """Do the email selection and extract text from email as directed by
        settings in configuration file (usually etracted.conf)."""
        if self._configuration is None:
            tkinter.messagebox.showinfo(
                parent=self.get_toplevel(),
                title='Show Extracted Text',
                message='Open a text extraction rules file')
            return
        if self._configuration_edited:
            tkinter.messagebox.showinfo(
                parent=self.get_toplevel(),
                title='Show Extracted Text',
                message=''.join((
                    'The edited configuration file has not been saved. ',
                    'It must be saved before "Show" action can be done.')))
            return
        if self._email_collector is None:
            emc = self._emailextractor(
                self._folder,
                configuration=self.configctrl.get(
                    '1.0', ' '.join((tkinter.END, '-1 chars'))),
                parent=self.get_toplevel())
            if not emc.parse():
                return
            if not emc.selected_emails:
                tkinter.messagebox.showinfo(
                    parent=self.get_toplevel(),
                    title='Show Extracted Text',
                    message='No emails match the selection rules.')
                return
            self._email_collector = emc
        try:
            self._show_extracted_text()
        except EmailExtractorError:
            tkinter.messagebox.showinfo(
                parent=self.get_toplevel(),
                title='Show Extracted Text',
                message=''.join((
                    'KeyError exception occurred, probably due to missing ',
                    'or incorrect entry in configuration file.',
                    )))
            return
        self._most_recent_action = self.show_extracted_text
        return True

    def update_difference_files(self):
        """Do the text extraction and save difference files."""
        if not self.show_extracted_text():
            tkinter.messagebox.showinfo(
                parent=self.get_toplevel(),
                title='Update Extracted Text',
                message='Unable to update extracted text.')
            return
        ce = self._email_collector.copy_emails()
        if ce is None:
            return
        difference_tags, additional = ce
        if difference_tags is not None:
            self.emailtextctrl.see(
                self.emailtextctrl.tag_ranges(difference_tags[-1])[0])
            self.emaillistctrl.see(
                self.emaillistctrl.tag_ranges(
                    ''.join(('F', difference_tags[-1][1:])))[0])
            w = ' emails ' if len(difference_tags) > 1 else ' email '
            tkinter.messagebox.showinfo(
                parent=self.get_toplevel(),
                title='Update Extracted Text',
                message=''.join(
                    ('Text extracted from ',
                     str(len(difference_tags)),
                     w,
                     'differs from version held in database.',
                     )))
            return
        elif len(additional):
            w = ' emails ' if len(additional) > 1 else ' email '
            tkinter.messagebox.showinfo(
                parent=self.get_toplevel(),
                title='Update Extracted Text',
                message=''.join(
                    ('Text from ',
                     str(len(additional)),
                     w,
                     ' added to version held in database.',
                     )))
            return
        else:
            tkinter.messagebox.showinfo(
                parent=self.get_toplevel(),
                title='Update Extracted Text',
                message=''.join((
                    'No additional emails.\n\n',
                    'No text added to version held in database.')))
            return

    def clear_selection(self):
        """Clear the lists of extracted text."""
        if tkinter.messagebox.askquestion(
            parent=self.get_toplevel(),
            title='Clear Extracted Text',
            message='Confirm request to clear the lists of extracted text.',
            ) != tkinter.messagebox.YES:
            return
        self._clear_email_tags()
        self.emailtextctrl.delete('1.0', tkinter.END)
        self.emaillistctrl.delete('1.0', tkinter.END)
        self.statusbar.set_status_text()
        self._email_collector = None
        self._most_recent_action = None

    def _clear_email_tags(self):
        """Clear the tags identifying data for each email."""
        for w in (self.emailtextctrl, self.emaillistctrl, self.configctrl):
            for t in self._tag_names:
                w.tag_delete(t)
        self._tag_names.clear()
        
    def conf_popup(self, event=None):
        """"""
        wconf = self.configctrl
        index = wconf.index(''.join(('@', str(event.x), ',', str(event.y))))
        start = wconf.index(' '.join((index, 'linestart')))
        end = wconf.index(' '.join((index, 'lineend', '+1 char')))
        text = wconf.get(start, end)
        if not text.startswith(''.join((IGNORE_EMAIL, ' '))):
            tkinter.messagebox.showinfo(
                parent=self.get_toplevel(),
                title='Cancel Ignore Email',
                message=''.join(
                    ('The text under the pointer does not refer to an ',
                     'email to be ignored.',
                     )))
            return
        if tkinter.messagebox.askquestion(
            parent=self.get_toplevel(),
            title='Cancel Ignore Email',
            message=''.join(
                ('Confirm request to cancel ignore \n\n',
                 text.split(' ', 1)[-1],
                 '\n\nemail.\n\nThe file is not copied to the output ',
                 'directory in this action; use "Update" later to do this.',
                 )),
            ) != tkinter.messagebox.YES:
            return
        wconf.delete(start, end)
        if self._email_collector is not None:
            self._email_collector.include_email(text.split(' ', 1)[-1].strip())
        self._configuration_edited = True
        fn = open(self._configuration, 'w', encoding='utf-8')
        try:
            fn.write(wconf.get('1.0', ' '.join((tkinter.END, '-1 chars'))))
            self._configuration_edited = False
        finally:
            fn.close()
        return
        
    def list_popup(self, event=None):
        """"""
        wtext = self.emailtextctrl
        wlist = self.emaillistctrl
        wconf = self.configctrl
        tags = wlist.tag_names(
            wlist.index(''.join(('@', str(event.x), ',', str(event.y)))))
        for t in tags:
            if t.startswith('F'):
                text = wlist.get(*wlist.tag_ranges(t))
                if tkinter.messagebox.askquestion(
                    parent=self.get_toplevel(),
                    title='Show Email in List',
                    message=''.join(
                        ('Confirm request to scroll text to \n\n',
                         text,
                         '\n\nemail.',
                         )),
                    ) != tkinter.messagebox.YES:
                    return
                wtext.see(wtext.tag_ranges(''.join(('T', t[1:])))[0])
                trconf = wconf.tag_ranges(t)
                if trconf:
                    wconf.see(trconf[-1])
                return
        
    def text_popup(self, event=None):
        """"""
        wtext = self.emailtextctrl
        wlist = self.emaillistctrl
        tags = wtext.tag_names(
            wtext.index(''.join(('@', str(event.x), ',', str(event.y)))))
        for t in tags:
            if t.startswith('T'):
                ftag = ''.join(('F', t[1:]))
                fm, dt = wlist.get(*wlist.tag_ranges(ftag)).split('\n')
                dt = parsedate_tz(dt)
                fm = parseaddr(fm)
                if not (fm and dt):
                    tkinter.messagebox.showinfo(
                        parent=self.get_toplevel(),
                        title='Remove Email from Selection',
                        message='Email from or date invalid.')
                    return
                date = strftime('%Y%m%d%H%M%S', dt[:-1])
                utc = ''.join((format(dt[-1] // 3600, '0=+3'), '00'))
                filename = ''.join((date, fm[-1], utc))
                emailname = ''.join((filename, '.mbs'))
                if filename in self._email_collector.excluded_emails:
                    tkinter.messagebox.showinfo(
                        parent=self.get_toplevel(),
                        title='Remove Email from Selection',
                        message=''.join(
                            (emailname,
                             '\n\n',
                             'is already one of the emails ignored from the ',
                             'selection.',
                             )))
                    return
                fp = os.path.join(
                    os.path.expanduser(self._email_collector.outputdirectory),
                    filename)
                if os.path.exists(fp):
                    if tkinter.messagebox.askquestion(
                        parent=self.get_toplevel(),
                        title='Remove Email from Selection',
                        message=''.join(
                            (filename,
                             '\n\nexists in the output directory.  You will ',
                             "have to use your system's file manager to ",
                             'delete the file.\n\nConfirm request to add \n\n',
                             wlist.get(*wlist.tag_ranges(ftag)),
                             '\n\nto ignored email list in selection rules.',                             
                             )),
                        ) != tkinter.messagebox.YES:
                        return
                elif tkinter.messagebox.askquestion(
                    parent=self.get_toplevel(),
                    title='Remove Email from Selection',
                    message=''.join(
                        ('Confirm request to add \n\n',
                         wlist.get(*wlist.tag_ranges(ftag)),
                         '\n\nto ignored email list in selection rules.',
                         )),
                    ) != tkinter.messagebox.YES:
                    return
                wconf = self.configctrl
                start = wconf.index(tkinter.END)
                wconf.insert(tkinter.END, '\n')
                wconf.insert(tkinter.END, ' '.join((IGNORE_EMAIL, emailname)))
                wconf.tag_add(
                    ftag, start, wconf.index(' '.join((start, 'lineend'))))
                wconf.tag_bind(ftag, '<ButtonPress-1>', self._file_exists)
                self._email_collector.ignore_email(emailname)
                self._configuration_edited = True
                fn = open(self._configuration, 'w', encoding='utf-8')
                try:
                    fn.write(wconf.get(
                        '1.0', ' '.join((tkinter.END, '-1 chars'))))
                    self._configuration_edited = False
                finally:
                    fn.close()
                return
        
    def _file_exists(self, event=None):
        """"""
        w = event.widget
        ti = w.index(''.join(('@', str(event.x), ',', str(event.y))))
        start = w.index(' '.join((ti, 'linestart')))
        end = w.index(' '.join((ti, 'lineend')))
        filename = w.get(start, end).split(' ', 1)[-1]
        od = os.path.expanduser(self._email_collector.outputdirectory)
        fp = os.path.join(od, filename)
        if os.path.exists(fp):
            self.statusbar.set_status_text(
                ' '.join((filename, 'exists in output directory', od)))
        else:
            self.statusbar.set_status_text(
                ' '.join((filename, 'does not exist in output directory', od)))


class Statusbar(object):
    
    """Status bar for EmailExtract application.
    
    """

    def __init__(self, root):
        """Create status bar widget."""
        self.status = tkinter.Text(
            root,
            height=0,
            width=0,
            background=root.cget('background'),
            relief=tkinter.FLAT,
            state=tkinter.DISABLED,
            wrap=tkinter.NONE)
        self.status.pack(side=tkinter.BOTTOM, fill=tkinter.X)

    def get_status_text(self):
        """Return text displayed in status bar."""
        return self.status.cget('text')

    def set_status_text(self, text=''):
        """Display text in status bar."""
        self.status.configure(state=tkinter.NORMAL)
        self.status.delete('1.0', tkinter.END)
        self.status.insert(tkinter.END, text)
        self.status.configure(state=tkinter.DISABLED)

