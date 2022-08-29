# emailextractor.py
# Copyright 2017 Roger Marsh
# Licence: See LICENCE (BSD licence)

"""Extract text from emails and save for application specific extraction.

These classes assume text from emails are held in files in a directory, with
no sub-directories, where each file contains a single email.

Each file may start with a 'From ' line, formatted as in mbox mailbox files,
but lines within the email which start 'From ' will only have been changed to
lines starting '>From ' if the email client which accepted delivery of the
email did so.  It depends on which mailbox format the email client uses.

"""
# A slightly modified results.core.emailextractor version 2.2, with all the
# stuff specific to the ChessResults application reomoved, is the initial
# version of this module.

import os
from datetime import date
import re
from email import message_from_binary_file
from email.utils import parseaddr, parsedate_tz
from email.message import EmailMessage
import email.header
from time import strftime
import subprocess
import io
import csv
import difflib
import shutil
import tkinter.messagebox
import zipfile
import xml.etree.ElementTree
import base64

try:
    import tnefparse
except ImportError:  # Not ModuleNotFoundError for Pythons earlier than 3.6
    tnefparse = None

try:
    import xlsx2csv
except ImportError:
    xlsx2csv = None
try:
    import pdfminer
    from pdfminer import pdfinterp, layout, converter
except ImportError:
    pdfminer = None

# Added when finding out how to use pdfminer3k to extract data from PDF files,
# and how to use xlsx2csv to extract data fron xlsx files.
import sys

from solentware_misc.core.utilities import AppSysDate

# Directory which holds emails one per file copied from email client mailboxes.
# Use imported COLLECTED attribute if available because emailextract expects to
# work with the emailstore package but can work with arbitrary collections of
# emails.
try:
    from emailstore.core.emailcollector import COLLECTED
except ImportError:
    COLLECTED = "collected"

# Python is installed in C: by default on Microsft Windows, so it is deemed
# acceptable to install pdftotext.exe in HOME even though this is done by
# direct copying rather than by an installer.  Per user installation is done by
# copying pdftotext.exe to HOMEDRIVE.

if sys.platform != "win32":
    _PDFTOTEXT = "pdftotext"
else:
    _PDFTOTEXT = "pdftotext.exe"

    # xpdf installation notes for Microsoft Windows say 'copy everything to an
    # installation directory, e.g. C:/Program Files/Xpdf'.
    # Try to choose the 32-bit or 64-bit executable as appropriate.
    if sys.maxsize > 2 ** 32:
        xpdf = os.path.join("Xpdf", "bin64")
    else:
        xpdf = os.path.join("Xpdf", "bin32")

    if os.path.isfile(
        os.path.join(os.environ["USERPROFILE"], xpdf, _PDFTOTEXT)
    ):
        _PDFTOTEXT = os.path.join(os.environ["USERPROFILE"], xpdf, _PDFTOTEXT)
    elif os.path.isfile(
        os.path.join(os.environ["HOMEDRIVE"], xpdf, _PDFTOTEXT)
    ):
        _PDFTOTEXT = os.path.join(os.environ["HOMEDRIVE"], xpdf, _PDFTOTEXT)
    else:
        _PDFTOTEXT = None
    del xpdf

# ssconvert, a tool supplied with Gnumeric, converts many spreadsheet formats
# to csv files.
# Python provides a module to handle csv files almost trivially.
# In September 2014 it was noticed that gnumeric.org withdrew all pre-built
# binaries of Gnumeric for Microsoft Windows in August 2014, citing crashes
# with profiles suggesting Gtk+ problems and lack of resources to do anything
# about it.
# I had downloaded the pre-built binary for Gnumeric-1-9-16 in March 2010.
# Later versions introduced support for putting each sheet in a separate csv
# file, but I never downloaded a Microsoft Windows binary.
# xls2csv and xlsx2csv were considered as alternatives to ssconvert to do the
# spreadsheet to csv file conversion.
# xls2csv works only on Python 2.  I do not know it's capabilities.
# xlsx2csv handles xlsx format only, and it's output is not compatible with the
# output from ssconvert as far as this application is concerned.
# For cross-platform consistency a date format has to be specified to xlsx2csv
# otherwise one may get the raw number representing a date output as text to
# the csv file.  ssconvert outputs a date according to the formatting given
# for the cell in the source spreadsheet.

# The workaround is attach csv files created using the spreadsheet application
# to the email.

_SSTOCSV = "ssconvert.exe"
if sys.platform != "win32":
    _SSTOCSV = "ssconvert"
else:

    # Version directories exist in ../Program Files/Gnumeric for each version
    # of Gnumeric installed.  Pick one of them, effectively at random, if any
    # exist.
    # If Gnumeric is not installed, look for xlsx2csv.py in site-packages.
    sstocsvdir = os.path.join(os.environ["PROGRAMFILES"], "Gnumeric")
    if os.path.isdir(sstocsvdir):
        sstocsv = set(os.listdir(sstocsvdir))
    else:
        sstocsv = None
    if sstocsv:
        sstocsv = os.path.join(sstocsvdir, sstocsv.pop(), "bin", _SSTOCSV)
        if os.path.isfile(sstocsv):
            _SSTOCSV = sstocsv
        else:
            _SSTOCSV = None
    else:
        _SSTOCSV = None
    del sstocsvdir, sstocsv

# _SSTOCSV is left alone so that ssconvert is used if it is available, despite
# the problems cited in August 2014.

_MAIL_STORE = "mailstore"
_EARLIEST_DATE = "earliestdate"
_MOST_RECENT_DATE = "mostrecentdate"
_EMAIL_SENDER = "emailsender"
IGNORE_EMAIL = "ignore"

# The name of the configuration file for extracting text from emails.
EXTRACTED_CONF = "extracted.conf"

# The extract configuration file entry naming the collect configuration file
# for collecting emails from email client mailboxes.
COLLECT_CONF = "collect_conf"

# MEDIA_TYPES directory holds csv files of media types in the format provided
# by https://www.iana.org/assignments/media-types/media-types.xhtml
MEDIA_TYPES = "media_types"

# Directory which holds difference files of text extracted from emails.
EXTRACTED = "extracted"

# Identify a pdf content-type to be included in the extracted data
PDF_CONTENT_TYPE = "pdf_content_type"

# Identify a text content-type to be included in the extracted data
TEXT_CONTENT_TYPE = "text_content_type"

# Identify a spreadsheet content-type to be included in the extracted data
_SS_CONTENT_TYPE = "ss_content_type"

# Identify a comma separated value (csv) content-type to be included in the
# extracted  data
CSV_CONTENT_TYPE = "csv_content_type"

# Identify a spreadsheet sheet name or csv file name to be included or
# excluded in the extracted data.
# Probably want to do ss version but maybe not csv, not done for pdf or text.
# Maybe this should be done in application specific configuration too, so
# there is choice about best place to do this.
INCLUDE_CSV_FILE = "include_csv_file"
EXCLUDE_CSV_FILE = "exclude_csv_file"
INCLUDE_SS_FILE_SHEET = "include_ss_file_sheet"
EXCLUDE_SS_FILE_SHEET = "exclude_ss_file_sheet"

# csv files should not contain '\x00' bytes but in practice some do.
# _NUL is used when decoding a csv email attachment to offer the opportunity of
# not processing such csv files.
_NUL = "\x00"

# Identify an Office Open XML word processing content-type to be included in
# the extracted data.  Also known as (Microsoft) docx format.
DOCX_CONTENT_TYPE = "docx_content_type"

# Identify an Office Open XML spreadsheet content-type to be included in the
# extracted data.  Also known as (Microsoft) xlsx format.
XLSX_CONTENT_TYPE = "xlsx_content_type"

# Identify an Open Office XML word processing content-type to be included in
# the extracted data.  Also known as Open Document Format (*.odt).
ODT_CONTENT_TYPE = "odt_content_type"

# Identify an Open Office XML spreadsheet content-type to be included in the
# extracted data.  Also known as Open Document Format (*.ods).
ODS_CONTENT_TYPE = "ods_content_type"

# Constants to handle TNEF attachments easily with non-TNEF attachments.
# Non-TNEF can use content-type for attachment directly but TNEF is
# content-type application/ms-tnef and we use the extension to decide.  The
# non-TNEF version sets the appropriate constant below.
# _SS is any spreadsheet, so do not set to an extension: cannot use TNEF here.
# TNEF means emails sent by Microsoft Outlook in some circumstances.
_PDF = ".pdf"
_SS = None
_XLSX = ".xlsx"
_ODS = ".ods"
_CSV = ".csv"
_TXT = ".txt"
_DOCX = ".docx"
_ODT = ".odt"


class EmailExtractorError(Exception):
    """Exception class for emailextractor module."""


# There are two distinct sets of configuration settings; email selection and
# parsing rules. EmailExtractor will end up a subclass of "Parse" which can be
# shared with EventSeason for text parsing rules.
class EmailExtractor(object):
    """Extract emails matching selection criteria from email client store.

    By default look for emails sent or received using the Opera email client
    in the most recent twelve months.

    """

    email_select_line = re.compile(
        "".join(
            (
                r"\A",
                "(?:",
                "(?:",  # whitespace line
                r"\s*",
                ")|",
                "(?:",  # comment line
                r"\s*#.*",
                ")|",
                "(?:",  # parameter line
                r"\s*(\S+?)\s+([^#]*).*",
                ")",
                ")",
                r"\Z",
            )
        )
    )
    replace_value_columns = re.compile(r"\+|\*")

    def __init__(
        self,
        folder,
        configuration=None,
        parser=None,
        extractemail=None,
        parent=None,
    ):
        """Define the email extraction rules from configuration.

        folder - the directory containing the event's data
        configuration - the rules for extracting emails

        """
        self.configuration = configuration
        self.criteria = None
        self.email_client = None
        self._folder = folder
        self.parent = parent
        if parser is None:
            self._parser = Parser
        else:
            self._parser = parser
        if extractemail is None:
            self._extractemail = ExtractEmail
        else:
            self._extractemail = extractemail

    def parse(self):
        """Set rules, from configuration file, for email text extraction."""
        self.criteria = None
        criteria = self._parser(parent=self.parent).parse(self.configuration)
        if criteria:
            self.criteria = criteria
            return True
        if criteria is False:
            return False
        return True

    def _select_emails(self):
        """Calculate and return emails matching requested from addressees."""
        if self.criteria is None:
            return
        self.email_client = self._extractemail(
            eventdirectory=self._folder, **self.criteria
        )
        return self.email_client.selected_emails

    @property
    def selected_emails(self):
        """Return emails selected by matching requested from addressees."""
        if self.email_client:
            return self.email_client.selected_emails
        return self._select_emails()

    @property
    def excluded_emails(self):
        """Return set of emails to ignore."""
        if not self.email_client:
            if not self._select_emails():
                return
        return self.email_client.excluded_emails

    @property
    def eventdirectory(self):
        """Return the path name of the document directory."""
        return self.email_client.eventdirectory

    @property
    def outputdirectory(self):
        """Return the path name of the extracted directory."""
        return self.email_client._extracts

    def copy_emails(self):
        """Copy text of emails selected from collected to extracted directory.

        Each email file in the collected directory will have a corresponding
        text difference file in the extracted directory.

        """
        difference_tags = []
        additional = []
        for e, em in enumerate(self.selected_emails):

            # The interaction between universal newlines and difflib can cause
            # problems.  In particular when \r is used as a field separator.
            # This way such text extracted from an email is readable because
            # \r shows up as a special glyph in the tkinter Text widget.
            # Later, when processing text, it shows up as a newline (\n).
            if tuple(
                (s.rstrip("\r\n"), s[-1] in "\r\n")
                for s in "\n".join(em.extracted_text).splitlines(True)
            ) != tuple(
                (s.rstrip("\r\n"), s[-1] in "\r\n")
                for s in list(difflib.restore(em.edit_differences, 1))
            ):
                difference_tags.append("x".join(("T", str(e))))

            if em.difference_file_exists is False:
                additional.append(em)
        if difference_tags:
            return difference_tags, None
        elif additional:
            w = " emails " if len(additional) > 1 else " email "
            if (
                tkinter.messagebox.askquestion(
                    parent=self.parent,
                    title="Update Extracted Text",
                    message="".join(
                        (
                            "Confirm that text from ",
                            str(len(additional)),
                            w,
                            " be added to version held in database.",
                        )
                    ),
                )
                != tkinter.messagebox.YES
            ):
                return
        else:
            return None, additional
        try:
            os.mkdir(os.path.dirname(additional[0].difference_file_path))
        except FileExistsError:
            pass
        for em in additional:
            try:
                em.write_additional_file()
            except FileNotFoundError as exc:
                excdir = os.path.basename(os.path.dirname(exc.filename))
                tkinter.messagebox.showinfo(
                    parent=self.parent,
                    title="Update Extracted Text",
                    message="".join(
                        (
                            "Write additional file to directory\n\n",
                            os.path.basename(os.path.dirname(exc.filename)),
                            "\n\nfailed.\n\nHopefully because the directory ",
                            "does not exist yet: it could have been deleted.",
                        )
                    ),
                )
                return
        return None, additional

    def ignore_email(self, filename):
        """Add email to list of ignored emails."""
        if self.email_client.ignore is None:
            self.email_client.ignore = set()
        self.email_client.ignore.add(filename)

    def include_email(self, filename):
        """Remove an email from list of ignored emails."""
        if self.email_client.ignore is None:
            self.email_client.ignore = set()
        self.email_client.ignore.remove(filename)


class Parser(object):
    """Parse configuration file."""

    def __init__(self, parent=None):
        """Set parent widget to own configuration rule error dialogues.

        The rules for handling configuration file keywoords are set.

        """
        self.parent = parent
        # Rules for processing conf file
        self.keyword_rules = {
            _MAIL_STORE: self.assign_value,
            _EARLIEST_DATE: self.assign_value,
            _MOST_RECENT_DATE: self.assign_value,
            _EMAIL_SENDER: self.add_value_to_set,
            IGNORE_EMAIL: self.add_value_to_set,
            COLLECT_CONF: self.assign_value,
            COLLECTED: self.assign_value,
            EXTRACTED: self.assign_value,
            MEDIA_TYPES: self.assign_value,
            PDF_CONTENT_TYPE: self.add_value_to_set,
            TEXT_CONTENT_TYPE: self.add_value_to_set,
            _SS_CONTENT_TYPE: self.add_value_to_set,
            DOCX_CONTENT_TYPE: self.add_value_to_set,
            ODT_CONTENT_TYPE: self.add_value_to_set,
            XLSX_CONTENT_TYPE: self.add_value_to_set,
            ODS_CONTENT_TYPE: self.add_value_to_set,
            CSV_CONTENT_TYPE: self.add_value_to_set,
            INCLUDE_CSV_FILE: self.add_value_to_set,
            EXCLUDE_CSV_FILE: self.add_value_to_set,
            INCLUDE_SS_FILE_SHEET: self.add_values_to_dict_of_sets,
            EXCLUDE_SS_FILE_SHEET: self.add_values_to_dict_of_sets,
        }

    def assign_value(self, v, args, args_key):
        """Set dict item args[args_key] to v from configuration file."""
        args[args_key] = v

    def add_value_to_set(self, v, args, args_key):
        """Add v, from configuration file, to set args[args_key]."""
        if args_key not in args:
            args[args_key] = set()
        args[args_key].add(v)

    def add_values_to_dict_of_sets(self, v, args, args_key):
        """Add v, from configuration file, to set at dict args[args_key]."""
        sep = v[0]
        v = v.split(sep=v[0], maxsplit=2)
        if len(v) < 3:
            args[args_key].setdefault(v[-1], set())
            return
        if args_key not in args:
            args[args_key] = dict()
        args[args_key].setdefault(v[1], set()).update(v[2].split(sep=sep))

    def _parse_error_dialogue(self, message):
        """Show dialogue for errors reading configuration file."""
        tkinter.messagebox.showinfo(
            parent=self.parent,
            title="Configuration File",
            message="".join(
                (
                    "Extraction rules are invalid.\n\nFailed rule is:\n\n",
                    message,
                )
            ),
        )

    def parse(self, configuration):
        """Return arguments from configuration file."""
        args = {}
        for line in configuration.split("\n"):
            g = EmailExtractor.email_select_line.match(line)
            if not g:
                self._parse_error_dialogue(line)
                return False
            key, value = g.groups()
            if key is None:
                continue
            if not value:
                self._parse_error_dialogue(line)
                return False
            args_type = self.keyword_rules.get(key.lower())
            if args_type is None:
                self._parse_error_dialogue(line)
                return False
            try:
                args_type(value, args, key.lower())
            except (re.error, ValueError):
                self._parse_error_dialogue(line)
                return False
        return args


class MessageFile(EmailMessage):
    """Extend EmailMessage class with a method to generate a filename.

    The From and Date headers are used.

    """

    def generate_filename(self):
        """Return a base filename or None when headers are no available."""
        t = parsedate_tz(self.get("Date"))
        f = parseaddr(self.get("From"))[-1]
        if t and f:
            ts = strftime("%Y%m%d%H%M%S", t[:-1])
            utc = "".join((format(t[-1] // 3600, "0=+3"), "00"))
            return "".join((ts, f, utc, ".mbs"))
        else:
            return False


class ExtractEmail(object):
    """Extract emails matching selection criteria from email store."""

    def __init__(
        self,
        extracttext=None,
        earliestdate=None,
        mostrecentdate=None,
        emailsender=None,
        eventdirectory=None,
        ignore=None,
        collect_conf=None,
        collected=None,
        extracted=None,
        media_types=None,
        pdf_content_type=None,
        text_content_type=None,
        ss_content_type=None,
        csv_content_type=None,
        docx_content_type=None,
        odt_content_type=None,
        xlsx_content_type=None,
        ods_content_type=None,
        include_csv_file=None,
        exclude_csv_file=None,
        include_ss_file_sheet=None,
        exclude_ss_file_sheet=None,
        parent=None,
        **soak
    ):
        """Define the email extraction rules from configuration.

        mailstore - the directory containing the email files
        earliestdate - emails before this date are ignored
        mostrecentdate - emails after this date are ignored
        emailsender - iterable of from addressees to select emails
        eventdirectory - directory to contain the event's data
        ignore - iterable of email filenames to be ignored
        schedule - difference file for event schedule
        reports - difference file for event result reports

        """
        self.parent = parent
        if extracttext is None:
            self._extracttext = ExtractText
        else:
            self._extracttext = extracttext
        if collect_conf:
            try:
                cc = open(
                    os.path.join(eventdirectory, collect_conf), encoding="utf8"
                )
                try:
                    for line in cc.readlines():
                        line = line.split(" ", 1)
                        if line[0] == COLLECTED:
                            if len(line) == 2:
                                from_conf = line[1].strip()
                    if from_conf:
                        collected = from_conf
                except Exception:
                    tkinter.messagebox.showinfo(
                        parent=self.parent,
                        title="Read Configuration File",
                        message="".join(
                            (
                                "Unable to determine collected directory.\n\n",
                                "Using name from extract configuration.",
                            )
                        ),
                    )
                finally:
                    cc.close()
            except Exception:
                pass
        if collected is None:
            ms = COLLECTED
        else:
            ms = os.path.join(eventdirectory, collected)
        self.mailstore = os.path.expanduser(os.path.expandvars(ms))
        if extracted is None:
            self._extracts = os.path.join(eventdirectory, EXTRACTED)
        else:
            self._extracts = os.path.join(eventdirectory, extracted)
        d = AppSysDate()
        if earliestdate is not None:
            if d.parse_date(earliestdate) == -1:
                tkinter.messagebox.showinfo(
                    parent=self.parent,
                    title="Read Configuration File",
                    message="".join(
                        (
                            "Format error in earliest date argument.\n\n",
                            "Please fix configuration file.",
                        )
                    ),
                )
                self.earliestdate = False
            else:
                self.earliestdate = d.iso_format_date()
        else:
            self.earliestdate = earliestdate
        if mostrecentdate is not None:
            if d.parse_date(mostrecentdate) == -1:
                tkinter.messagebox.showinfo(
                    parent=self.parent,
                    title="Read Configuration File",
                    message="".join(
                        (
                            "Format error in most recent date argument.\n\n",
                            "Please fix configuration file.",
                        )
                    ),
                )
                self.mostrecentdate = False
            else:
                self.mostrecentdate = d.iso_format_date()
        else:
            self.mostrecentdate = mostrecentdate
        self.emailsender = emailsender
        self.eventdirectory = eventdirectory
        self.ignore = ignore
        self._selected_emails = None
        self._selected_emails_text = None
        self._text_extracted_from_emails = None
        if pdf_content_type:
            self.pdf_content_type = pdf_content_type
        else:
            self.pdf_content_type = frozenset()
        if text_content_type:
            self.text_content_type = text_content_type
        else:
            self.text_content_type = frozenset()
        if ss_content_type:
            self.ss_content_type = ss_content_type
        else:
            self.ss_content_type = frozenset()
        if csv_content_type:
            self.csv_content_type = csv_content_type
        else:
            self.csv_content_type = frozenset()
        if docx_content_type:
            self.docx_content_type = docx_content_type
        else:
            self.docx_content_type = frozenset()
        if odt_content_type:
            self.odt_content_type = odt_content_type
        else:
            self.odt_content_type = frozenset()
        if xlsx_content_type:
            self.xlsx_content_type = xlsx_content_type
        else:
            self.xlsx_content_type = frozenset()
        if ods_content_type:
            self.ods_content_type = ods_content_type
        else:
            self.ods_content_type = frozenset()
        if include_csv_file is None:
            self.include_csv_file = []
        else:
            self.include_csv_file = include_csv_file
        if exclude_csv_file is None:
            self.exclude_csv_file = []
        else:
            self.exclude_csv_file = exclude_csv_file
        if include_ss_file_sheet is None:
            self.include_ss_file_sheet = []
        else:
            self.include_ss_file_sheet = include_ss_file_sheet
        if exclude_ss_file_sheet is None:
            self.exclude_ss_file_sheet = []
        else:
            self.exclude_ss_file_sheet = exclude_ss_file_sheet

    def get_emails(self):
        """Return email files in order stored in mail store.

        Each email is stored in a file named:
        <self.mailstore>/yyyymmddHHMMSS<sender><utc offset>.mbs

        """
        emails = []
        if self.earliestdate is not None:
            try:
                date(*tuple([int(d) for d in self.earliestdate.split("-")]))
            except Exception:
                tkinter.messagebox.showinfo(
                    parent=self.parent,
                    title="Get Emails",
                    message="".join(
                        (
                            "Earliest date format error.\n\n",
                            "Please fix configuration file.",
                        )
                    ),
                )
                return emails
        if self.mostrecentdate is not None:
            try:
                date(*tuple([int(d) for d in self.mostrecentdate.split("-")]))
            except Exception:
                tkinter.messagebox.showinfo(
                    parent=self.parent,
                    title="Get Emails",
                    message="".join(
                        (
                            "Most recent date format error.\n\n",
                            "Please fix configuration file.",
                        )
                    ),
                )
                return emails
        try:
            ms = self.mailstore
            ems = self.emailsender
            for a in os.listdir(ms):
                if self.ignore:
                    if a in self.ignore:
                        continue
                if ems:
                    for e in ems:
                        if e == a[8 : 8 + len(e)]:
                            break
                    else:
                        continue
                emd = "-".join((a[:4], a[4:6], a[6:8]))
                if self.earliestdate is not None:
                    if emd < self.earliestdate:
                        continue
                if self.mostrecentdate is not None:
                    if emd > self.mostrecentdate:
                        continue
                emails.append(self._extracttext(a, self))
        except FileNotFoundError:
            emails.clear()
            tkinter.messagebox.showinfo(
                parent=self.parent,
                title="Get Emails",
                message="".join(
                    (
                        "Attempt to get files in directory\n\n",
                        str(self.mailstore),
                        "\n\nfailed.",
                    )
                ),
            )
        emails.sort()
        return emails

    def _get_emails_for_from_addressees(self):
        """Return selected email files in order stored in mail store.

        Emails are selected by 'From Adressee' using the email addresses in
        the emailsender argument of ExtractEmail() call.

        """
        return [
            e
            for e in self.get_emails()
            if e.is_from_addressee_in_selection(self.emailsender)
        ]

    @property
    def selected_emails(self):
        """Return emails selected by matching requested from addressees."""
        if self._selected_emails is None:
            self._selected_emails = self._get_emails_for_from_addressees()
        return self._selected_emails

    @property
    def excluded_emails(self):
        """Return set of emails to ignore."""
        if not self.ignore:
            return set()
        return set(self.ignore)


class ExtractText(object):
    """Repreresent the stages in processing an email."""

    def __init__(self, filename, emailstore):
        """Set up to extract text from mailstore into filename."""
        self.filename = filename
        self._emailstore = emailstore
        self._message = None
        self._encoded_text = None
        self._extracted_text = None
        self._edit_differences = None
        self._difference_file_exists = None
        self._date = None
        self._delivery_date = None

    def __eq__(self, other):
        """Return True if self.filename == other.filename."""
        return self.filename == other.filename

    def __lt__(self, other):
        """Return True if self.filename < other.filename."""
        return self.filename < other.filename

    @property
    def email_path(self):
        """Return mailstore path name."""
        return os.path.join(self._emailstore.mailstore, self.filename)

    @property
    def difference_file_path(self):
        """Return difference file's path name without extension."""
        return os.path.join(
            self._emailstore._extracts, os.path.splitext(self.filename)[0]
        )

    @property
    def difference_file_exists(self):
        """Return True if difference file existed when edit_differences set.

        True means the email does not change the original text.
        False means the edited version of text is copied from the email text.
        None means answer is unknown because edit_differences has not been set.

        """
        return self._difference_file_exists

    def is_from_addressee_in_selection(self, selection):
        """Return filename if addressee is in selection.

        The file name is generated by self.message.generate_filename function.

        """
        if selection is None:
            return True
        from_ = parseaddr(self.message.get("From"))[-1]

        if not selection:
            return self.message.generate_filename()

        # Ignore emails not sent by someone in self.emailsender.
        # Account owners may be in that set, so emails sent from one
        # account owner to another can get selected.
        if from_ in selection:
            return self.message.generate_filename()
        else:
            return False

    @property
    def message(self):
        """Return object created by email.message_from_binary_file function."""
        if self._message is None:
            mf = open(self.email_path, "rb")
            try:
                self._message = message_from_binary_file(
                    mf, _class=MessageFile
                )
                self._date = [
                    parsedate_tz(d)[:-1]
                    for d in self.message.get_all("date", [])
                ]
                self._delivery_date = [
                    parsedate_tz(d)[:-1]
                    for d in self.message.get_all("delivery-date", [])
                ]
            finally:
                mf.close()
        return self._message

    @property
    def encoded_text(self):
        """Return encoded text extracted from emails."""
        if self._encoded_text is None:
            ems = self._emailstore
            text = []
            for p in self.message.walk():
                cte = p.get("Content-Transfer-Encoding")
                if cte:
                    if cte.lower() == "base64":
                        t = p.get_payload(decode=True)
                        if t:
                            text.append(t)

            # If no text at all is extracted return a single blank line.
            if not text:
                text.append(b"\n")

            # Ensure the extracted text ends with a newline so that editing of
            # the last line causes difflib processing to append the '?   ---\n'
            # or '?   +++\n' as a separate line in the difference file.
            # This may make the one character adjustment done by
            # _insert_entry() in the SourceEdit class redundant.
            # The reason to not do this all along was to avoid making any
            # change at all between the selected email payload and the original
            # version held in the difference file, including an extra newline.
            # Attempts to wrap difflib functions to cope seem not worth it, if
            # such prove possible at all.
            if not text[-1].endswith(b"\n"):
                text[-1] = b"".join((text[-1], b"\n"))

            self._encoded_text = text
        return self._encoded_text

    def _extract_text(
        self, content_type, filename, payload, text, charset=None
    ):
        """Bind text extracted from emails to _emailtore attribute."""
        ems = self._emailstore
        if content_type == _PDF:
            if _PDFTOTEXT:
                self._get_pdf_text_using_xpdf(filename, payload, text)
            elif pdfminer:
                self.get_pdf_text_using_pdfminer3k(None, payload, text)
        elif content_type == _SS:
            if _SSTOCSV:
                self._get_ss_text_using_gnumeric(filename, payload, text)
        elif content_type == _XLSX:
            if _SSTOCSV:
                self._get_ss_text_using_gnumeric(filename, payload, text)
            elif xlsx2csv:
                self._get_ss_text_using_xlsx2csv(filename, payload, text)
        elif content_type == _ODS:
            if _SSTOCSV:
                self._get_ss_text_using_gnumeric(filename, payload, text)
            else:
                self._get_ods_text_using_python_xml(filename, payload, text)
        elif content_type == _CSV:
            if ems.include_csv_file:
                if filename not in ems.include_csv_file:
                    return
            elif ems.exclude_csv_file:
                if filename in ems.exclude_csv_file:
                    return
            text.append(self.get_csv_text(payload, charset))
        elif content_type == _TXT:
            text.append(payload.decode(charset))
        elif content_type == _DOCX:
            text.append(self.get_docx_text(payload, ems.eventdirectory))
        elif content_type == _ODT:
            text.append(self.get_odt_text(payload, ems.eventdirectory))

    @property
    def extracted_text(self):
        """Return text extracted from emails."""
        if self._extracted_text is None:
            ems = self._emailstore
            text = []
            for p in self.message.walk():
                ct = p.get_content_type()
                if ct in ems.pdf_content_type:
                    self._extract_text(
                        _PDF,
                        p.get_filename(),
                        p.get_payload(decode=True),
                        text,
                    )
                elif ct in ems.ss_content_type:
                    self._extract_text(
                        _SS, p.get_filename(), p.get_payload(decode=True), text
                    )
                elif ct in ems.xlsx_content_type:
                    self._extract_text(
                        _XLSX,
                        p.get_filename(),
                        p.get_payload(decode=True),
                        text,
                    )
                elif ct in ems.ods_content_type:
                    self._extract_text(
                        _ODS,
                        p.get_filename(),
                        p.get_payload(decode=True),
                        text,
                    )
                elif ct in ems.csv_content_type:
                    self._extract_text(
                        _CSV,
                        p.get_filename(),
                        p.get_payload(decode=True),
                        text,
                        charset=p.get_param("charset", failobj="utf-8"),
                    )
                elif ct in ems.text_content_type:
                    self._extract_text(
                        _TXT,
                        p.get_filename(),
                        p.get_payload(decode=True),
                        text,
                        charset=p.get_param("charset", failobj="iso-8859-1"),
                    )
                elif ct in ems.docx_content_type:
                    self._extract_text(
                        _DOCX,
                        p.get_filename(),
                        p.get_payload(decode=True),
                        text,
                    )
                elif ct in ems.odt_content_type:
                    self._extract_text(
                        _ODT,
                        p.get_filename(),
                        p.get_payload(decode=True),
                        text,
                    )
                elif ct == "application/ms-tnef":
                    if not tnefparse:
                        text.append(
                            "".join(
                                (
                                    "Cannot process attachment: ",
                                    "tnefparse is not installed.",
                                )
                            )
                        )
                        continue

                    # I do not know if the wrapped attachment type names are
                    # encoded within the application/ms-tnef attachment, or the
                    # mapping if so.  Fall back on assumption from file name
                    # extension.
                    # Similar point applies for charset argument used when
                    # extracting csv or txt attachments.
                    # As far as I know, the only application/ms-tnef
                    # attachments ever received by me at time of writing
                    # contain just *.txt attachments. These are not processed
                    # by this route.
                    tnef = tnefparse.TNEF(base64.b64decode(p.get_payload()))
                    for attachment in tnef.attachments:
                        name = attachment.name
                        for e in _PDF, _XLSX, _ODS, _CSV, _TXT, _DOCX, _ODT:
                            if name.lower().endswith(e):
                                self._extract_text(
                                    e,
                                    attachment.name,
                                    attachment.data,
                                    text,
                                    charset="iso-8859-1",
                                )
                                break
                        else:
                            text.append(
                                "".join(
                                    (
                                        "Cannot process '",
                                        name,
                                        "' extracted from application/",
                                        "ms-tnef attachment.",
                                    )
                                )
                            )

            # If no text at all is extracted return a single blank line.
            if not text:
                text.append("\n")

            # Ensure the extracted text ends with a newline so that editing of
            # the last line causes difflib processing to append the '?   ---\n'
            # or '?   +++\n' as a separate line in the difference file.
            # This may make the one character adjustment done by
            # _insert_entry() in the SourceEdit class redundant.
            # The reason to not do this all along was to avoid making any
            # change at all between the selected email payload and the original
            # version held in the difference file, including an extra newline.
            # Attempts to wrap difflib functions to cope seem not worth it, if
            # such prove possible at all.
            if not text[-1].endswith("\n"):
                text[-1] = "".join((text[-1], "\n"))

            self._extracted_text = text
        return self._extracted_text

    def _is_attachment_to_be_extracted(self, attachment_filename):
        ems = self._emailstore
        if ems.include_ss_file_sheet:
            if attachment_filename not in ems.include_ss_file_sheet:
                return
        elif ems.exclude_ss_file_sheet:
            if attachment_filename in ems.exclude_ss_file_sheet:
                return
        if _decode_header(attachment_filename) is None:
            tkinter.messagebox.showinfo(
                parent=self._emailstore.parent,
                title="Extract Spreadsheet Data",
                message="Spreadsheet attachment does not have a filename.",
            )
            return
        return True

    def _create_temporary_attachment_file(self, filename, payload, dirbase):
        a = _decode_header(filename)
        try:
            os.mkdir(os.path.join(dirbase, "xls-attachments"))
        except FileExistsError:
            pass
        op = open(os.path.join(dirbase, "xls-attachments", a), "wb")
        try:
            op.write(payload)
        finally:
            op.close()
        return a

    def _get_ss_text_using_gnumeric(self, filename, payload, text):
        fn = filename
        if not self._is_attachment_to_be_extracted(fn):
            return
        ems = self._emailstore
        taf = self._create_temporary_attachment_file(
            filename, payload, ems.eventdirectory
        )
        process = subprocess.Popen(
            (_SSTOCSV, "--recalc", "-S", taf, "%s.csv"),
            cwd=os.path.join(ems.eventdirectory, "xls-attachments"),
        )
        process.wait()
        if process.returncode == 0:
            sstext = []
            for sheet, sheettext in self.get_spreadsheet_text(
                ems.eventdirectory
            ):
                if fn in ems.include_ss_file_sheet:
                    if ems.include_ss_file_sheet[fn]:
                        if sheet not in ems.include_ss_file_sheet[fn]:
                            continue
                elif fn in ems.exclude_ss_file_sheet:
                    if ems.exclude_ss_file_sheet[fn]:
                        if sheet in ems.exclude_ss_file_sheet[fn]:
                            continue
                    else:
                        continue
                sstext.append(sheettext)
            text.append("\n\n".join(sstext))
        shutil.rmtree(
            os.path.join(ems.eventdirectory, "xls-attachments"),
            ignore_errors=True,
        )

    def _get_ss_text_using_xlsx2csv(self, filename, payload, text):
        fn = filename
        if not self._is_attachment_to_be_extracted(fn):
            return
        ems = self._emailstore
        taf = os.path.join(
            ems.eventdirectory,
            "xls-attachments",
            self._create_temporary_attachment_file(
                filename, payload, ems.eventdirectory
            ),
        )

        # Extract all sheets and filter afterwards.
        # xlsx2csv can do the filtering and will be used to do so later.
        # outputencoding has to be given, even though it is default value, to
        # avoid a KeyError exception on options passed to Xlsx2csv in Python 3.
        # The defaults of other arguments are used as expected.
        xlsx2csv.Xlsx2csv(
            taf,
            skip_empty_lines=True,
            sheetid=0,
            dateformat="%Y-%m-%d",
            outputencoding="utf-8",
        ).convert(
            os.path.join(ems.eventdirectory, "xls-attachments"), sheetid=0
        )

        sstext = []
        for sheet, sheettext in self.get_spreadsheet_text(ems.eventdirectory):
            if fn in ems.include_ss_file_sheet:
                if ems.include_ss_file_sheet[fn]:
                    if sheet not in ems.include_ss_file_sheet[fn]:
                        continue
            elif fn in ems.exclude_ss_file_sheet:
                if ems.exclude_ss_file_sheet[fn]:
                    if sheet in ems.exclude_ss_file_sheet[fn]:
                        continue
                else:
                    continue
            sstext.append(sheettext)
        text.append("\n\n".join(sstext))
        shutil.rmtree(
            os.path.join(ems.eventdirectory, "xls-attachments"),
            ignore_errors=True,
        )

    def _get_ods_text_using_python_xml(self, filename, payload, text):
        nstable = "{urn:oasis:names:tc:opendocument:xmlns:table:1.0}"
        nstext = "{urn:oasis:names:tc:opendocument:xmlns:text:1.0}"
        nsoffice = "{urn:oasis:names:tc:opendocument:xmlns:office:1.0}"

        def get_rows(table):
            rows = []
            for row in table.iter(nstable + "table-row"):
                rows.append({})
                cells = rows[-1]
                for cell in row.iter(nstable + "table-cell"):
                    repeat = int(
                        cell.attrib.get(
                            nstable + "number-columns-repeated", "1"
                        )
                    )
                    if cell.attrib.get(nsoffice + "value-type") == "date":
                        paragraphs = [cell.attrib[nsoffice + "date-value"]]
                    else:
                        paragraphs = []
                        for element in cell.iter(nstext + "p"):
                            if element.text is not None:
                                paragraphs.append(element.text)
                    text = "\n\n".join(paragraphs)
                    for r in range(repeat):
                        cells[len(cells)] = text if paragraphs else None

            # Discard leading and trailing empty columns, and empty rows.
            # No need to convert remaining Nones to ''s because csv module
            # DictWriter.writerows() method does it.
            trailing = -1
            leading = max(len(r) for r in rows)
            for r in rows:
                notnone = sorted(k for k, v in r.items() if v is not None)
                if notnone:
                    trailing = max(notnone[-1], trailing)
                    leading = min(notnone[0], leading)
            for r in rows:
                columns = list(r.keys())
                for c in columns:
                    if c > trailing or c < leading:
                        del r[c]
            return [
                r
                for r in rows
                if len([v for v in r.values() if v is not None])
            ]

        ems = self._emailstore
        fn = filename
        if ems.include_ss_file_sheet:
            if fn not in ems.include_ss_file_sheet:
                return
        elif ems.exclude_ss_file_sheet:
            if fn in ems.exclude_ss_file_sheet:
                return
        xmlzip = zipfile.ZipFile(io.BytesIO(payload))
        archive = {}
        for n in xmlzip.namelist():
            with xmlzip.open(n) as f:
                archive[n] = f.read()
        for k, v in archive.items():
            if os.path.basename(k) == "content.xml":
                tree = xml.etree.ElementTree.XML(v)
                sstext = []
                for spreadsheet in tree.iter(nsoffice + "spreadsheet"):
                    for table in spreadsheet.iter(nstable + "table"):
                        sheet = table.attrib[nstable + "name"]
                        if fn in ems.include_ss_file_sheet:
                            if ems.include_ss_file_sheet[fn]:
                                if sheet not in ems.include_ss_file_sheet[fn]:
                                    continue
                        elif fn in ems.exclude_ss_file_sheet:
                            if ems.exclude_ss_file_sheet[fn]:
                                if sheet in ems.exclude_ss_file_sheet[fn]:
                                    continue
                            else:
                                continue
                        rows = get_rows(table)
                        if not rows:
                            continue
                        fieldnames = [c for c in sorted(rows[0].keys())]
                        csvfile = io.StringIO()
                        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
                        writer.writerows(rows)
                        try:
                            sstext.append(
                                self.extract_text_from_csv(
                                    csvfile, sheet=sheet
                                )
                            )
                        except KeyError:
                            raise EmailExtractorError
                text.append("\n\n".join(sstext))

    def get_docx_text(self, payload, dirbase):
        """Return text from payload, an Office Open XML email attachment.

        This is Microsoft's *.docx format, not to be confused with *.odt format
        which is Open Office XML (or Open Document Format).

        """
        # Thanks to
        # http://etienned.github.io/posts/extract-text-from-word-docs-simply/
        # for which pieces to take from *.docx file.
        nsb = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

        xmlzip = zipfile.ZipFile(io.BytesIO(payload))
        archive = {}
        for n in xmlzip.namelist():
            with xmlzip.open(n) as f:
                archive[n] = f.read()
        text = []
        for k, v in archive.items():
            if os.path.basename((os.path.splitext(k)[0])) == "document":
                paragraphs = []
                tree = xml.etree.ElementTree.XML(v)
                for p in tree.iter(nsb + "p"):
                    t = [n.text for n in p.iter(nsb + "t") if n.text]
                    if t:
                        paragraphs.append("".join(t))
                text.append("\n\n".join(paragraphs))
        return "\n".join(text)

    def get_odt_text(self, payload, dirbase):
        """Return text from payload, an Open Office XML email attachment.

        This is *.odt format (or Open Document Format), not to be confused
        with Microsoft's *.docx format (or Office Open XML).

        """
        nsb = "{urn:oasis:names:tc:opendocument:xmlns:text:1.0}"
        topelems = {nsb + "p", nsb + "h"}

        def get_text(element):
            # Thanks to https://github.com/deanmalmgren/textract
            # for which pieces to take from *.odt file.
            text = []
            if element.text is not None:
                text.append(element.text)
            for child in element:
                if child.tag == nsb + "tab":
                    text.append("\t")
                    if child.tail is not None:
                        text.append(child.tail)
                elif child.tag == nsb + "s":
                    text.append(" ")
                    if child.get(nsb + "c") is not None:
                        text.append(" " * (int(child.get(nsb + "c")) - 1))
                    if child.tail is not None:
                        text.append(child.tail)
                else:
                    text.append(get_text(child))
            if element.tail is not None:
                text.append(element.tail)
            return "".join(text)

        xmlzip = zipfile.ZipFile(io.BytesIO(payload))
        archive = {}
        for n in xmlzip.namelist():
            with xmlzip.open(n) as f:
                archive[n] = f.read()
        text = []
        for k, v in archive.items():
            if os.path.basename((os.path.splitext(k)[0])) == "content":
                for child in xml.etree.ElementTree.fromstring(v).iter():
                    if child.tag in topelems:
                        text.append(get_text(child))
        return "\n".join(text)

    def _get_pdf_text_using_xpdf(self, filename, payload, text):
        """Use pdf2text utility (part of xpdf) to extract text."""
        a = _decode_header(filename)
        aout = a + ".txt"
        if a is None:
            tkinter.messagebox.showinfo(
                parent=self._emailstore.parent,
                title="Extract PDF Data",
                message="PDF attachment does not have a filename.",
            )
            return ""
        dirbase = self._emailstore.eventdirectory
        try:
            os.mkdir(os.path.join(dirbase, "pdf-attachments"))
        except FileExistsError:
            pass
        op = open(os.path.join(dirbase, "pdf-attachments", a), "wb")
        try:
            op.write(payload)
        finally:
            op.close()
        process = subprocess.Popen(
            (
                _PDFTOTEXT,
                "-nopgbrk",  # no way of saying this in pdfminer3k.
                "-layout",
                a,
                aout,
            ),
            cwd=os.path.join(dirbase, "pdf-attachments"),
        )
        process.wait()
        if process.returncode == 0:
            if os.path.exists(os.path.join(dirbase, "pdf-attachments", aout)):
                op = open(
                    os.path.join(dirbase, "pdf-attachments", aout),
                    "r",
                    encoding="iso-8859-1",
                )
                try:
                    text.append(op.read())
                finally:
                    op.close()
        shutil.rmtree(
            os.path.join(dirbase, "pdf-attachments"), ignore_errors=True
        )

    def get_pdf_text_using_pdfminer3k(
        self, filename, payload, text, char_margin=150, word_margin=1, **k
    ):
        """Use pdfminer3k functions to extract text from pdf by line (row).

        The char_margin and word_margin defaults give a reasonable fit to
        pdftotext (part of xpdf) behaviour with the '-layout' option.

        char_margin seems to have no upper limit as far as fitting with the
        '-layout' option is concerned, but a word_margin value of 1.5 caused
        words to be concatenated from the PDF document tried.  However the
        word_margin default (0.1) caused 'W's at the start of a word to be
        treated as a separate word: 'Winchester' becomes 'W inchester'.  There
        must be something else going on because 'Winchester' remained as
        'Winchester' in another, less tabular, PDF document.

        The PDF document has a tabular layout (read each row) which seems to
        get treated as a column layout (read each column) with the LAParams
        defaults set out below.

        **k captures other arguments which override defaults for pdfminer3k's
        LAParams class.

        At pdfminer3k-1.3.1 the arguments and their defaults are:
        line_overlap=0.5
        char_margin=2.0
        line_margin=0.5
        word_margin=0.1
        boxes_flow=0.5
        detect_vertical=False
        all_texts=False
        paragraph_indent=None
        heuristic_word_margin=False

        """
        # Adapted from pdf2txt.py script included in pdfminer3k-1.3.1.
        # On some *.pdf inputs the script raises UnicodeEncodeError:
        # 'ascii' codec can't encode character ...
        # which does not happen with the adaption below.
        # A sample ... is '\u2019 in position 0: ordinal not in range(128)'.
        # Changing 'outfp = io.open(...)' to 'outfp = open(...)' was sufficient
        # but here it is most convenient to say 'outfp = io.StringIO()'.
        caching = True
        rsrcmgr = pdfinterp.PDFResourceManager(caching=caching)
        laparams = layout.LAParams(
            char_margin=char_margin, word_margin=word_margin
        )
        for a in laparams.__dict__:
            if a in k:
                laparams.__dict__[a] = k[a]
        outfp = io.StringIO()
        device = converter.TextConverter(rsrcmgr, outfp, laparams=laparams)
        try:
            fp = io.BytesIO(payload)
            try:
                pdfinterp.process_pdf(
                    rsrcmgr,
                    device,
                    fp,
                    pagenos=set(),
                    maxpages=0,
                    password="",
                    caching=caching,
                    check_extractable=True,
                )
            finally:
                fp.close()
        finally:
            device.close()
        text.append(outfp.getvalue())
        outfp.close()

    def get_spreadsheet_text(self, dirbase):
        """Return (sheetname, text) from spreadsheet attachment part.

        dirbase is the event directory where a temporary directory is created
        to hold temporary files for the attachment extracts.
        """
        text = []
        for fn in os.listdir(os.path.join(dirbase, "xls-attachments")):
            sheetname, e = os.path.splitext(fn)
            if e.lower() != ".csv":
                continue
            sheetname = sheetname.lower()
            csvp = os.path.join(dirbase, "xls-attachments", fn)
            if not os.path.exists(csvp):
                continue
            try:
                text.append(
                    (
                        sheetname,
                        self.extract_text_from_csv(
                            self._read_file(csvp), sheet=sheetname
                        ),
                    )
                )
            except KeyError:
                raise EmailExtractorError
            except csv.Error as exc:
                tkinter.messagebox.showinfo(
                    parent=self._emailstore.parent,
                    title="Extract Text from CSV",
                    message="".join(
                        (
                            str(exc),
                            "\n\nreported by csv module for sheet\n\n",
                            os.path.splitext(fn)[0],
                            "\n\nwhich is not included in extracted text.",
                        )
                    ),
                )
        return text

    def extract_text_from_csv(self, text, sheet=None, filename=None):
        """Return text if it looks like CSV format, otherwise ''.

        A csv.Sniffer determines the csv dialect and text is accepted as csv
        format if the delimeiter seems to be in ',/t;:'.
        """
        text = text.getvalue()
        dialect = csv.Sniffer().sniff(text)
        if dialect.delimiter not in ",/t;:":
            return ""

        # All the translation in code taken from results.core.emailextractor
        # at results-2.2 is removed because it is specific to application.
        # Maybe this method should return the list of rows in case caller will
        # do more filtering? (and the other extract_* methods)
        return text

    def get_csv_text(self, payload, charset):
        """Return text from part, a csv attachment to an email."""
        try:
            return self.extract_text_from_csv(
                io.StringIO(self._decode_payload(payload, charset))
            )
        except KeyError:
            raise EmailExtractorError

    def _decode_payload(self, payload, charset):
        """Return decoded payload; try 'utf-8' then 'iso-8859-1'.

        iso-8859-1 should not fail but if it does fall back to ascii with
        replacement of bytes that do not decode.

        The current locale is not used because the decode must be the same
        every time it is done.

        """
        for c in charset, "iso-8859-1":
            try:
                return self._accept_csv_file_with_nul_characters(
                    payload.decode(encoding=c)
                )
                return t
            except UnicodeDecodeError:
                pass
        else:
            return self._accept_csv_file_with_nul_characters(
                payload.decode(encoding="ascii", errors="replace")
            )

    def _accept_csv_file_with_nul_characters(self, csvstring):
        """Dialogue asking what to do with csv file with NULs."""
        nulcount = csvstring.count(_NUL)
        if nulcount:
            if (
                tkinter.messagebox.askquestion(
                    parent=self._emailstore.parent,
                    title="Update Extracted Text",
                    message="".join(
                        (
                            "A csv file attachment to an email contains NUL ",
                            "characters.\n\nDo you wish to include the ",
                            "significant characters from this file?\n\n",
                            str(len(csvstring)),
                            " characters in file.\n",
                            str(nulcount),
                            " NULs in file.\n",
                            str(len(csvstring) - nulcount),
                            " significant characters in file.",
                        )
                    ),
                )
                != tkinter.messagebox.YES
            ):
                return ""
            csvstring = csvstring.replace(_NUL, "")
        return csvstring

    def _read_file(self, csvpath):
        """Return StringIO object containing decoded payload.

        Try 'utf-8' then 'iso-8859-1' and finally 'ascii' with errors replaced.

        """
        for c in "utf-8", "iso-8859-1":
            csvfile = open(csvpath, encoding=c)
            try:
                return io.StringIO(csvfile.read())
            except UnicodeDecodeError:
                pass
            finally:
                csvfile.close()
        else:
            csvfile = open(csvpath, encoding="ascii", errors="replace")
            try:
                return io.StringIO(csvfile.read())
            except UnicodeDecodeError:
                pass
            finally:
                csvfile.close()

    @property
    def edit_differences(self):
        """Return list(difflin.ndiff()) of original and edited email text."""
        if self._edit_differences is None:
            try:
                text = self._read_file(self.difference_file_path).readlines()
                self._difference_file_exists = True
            except FileNotFoundError:
                lines = "\n".join(self.extracted_text).splitlines(1)
                text = list(difflib.ndiff(lines, lines))
                self._difference_file_exists = False
            self._edit_differences = text
        return self._edit_differences

    def write_additional_file(self):
        """Write difference file, utf-8 encoding, if file does not exist."""
        if self._difference_file_exists is False:
            f = open(
                self.difference_file_path,
                mode="w",
                encoding="utf8",
            )
            try:
                f.writelines(self.edit_differences)
                self._difference_file_exists = True
            finally:
                f.close()

    @property
    def dates(self):
        """Return tuple(date, delivery_dates)."""
        return self._date, self._delivery_date


def _decode_header(value):
    """Decode value according to RFC2231 and return the decoded string."""
    b, c = email.header.decode_header(value)[0]
    return b if c is None else b.decode(c)
