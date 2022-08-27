# __init__.py
# Copyright 2017 Roger Marsh
# Licence: See LICENCE (BSD licence)

"""Extract text from emails and their attachments stored in a directory.

Each email is in a separate file (the emailstore project splits mailbox style
mailstores into files with one emial in each).

Text can be extracted from the email body and attachments containing text which
usually means '*.txt', '*.csv', '*.pdf', and various spreadsheets.

"""
APPLICATION_NAME = "EmailExtract"
ERROR_LOG = "ErrorLog"
