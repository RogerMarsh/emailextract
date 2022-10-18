# __init__.py
# Copyright 2014 Roger Marsh
# Licence: See LICENCE (BSD licence)

"""Help files for emailextract package."""

import os

ABOUT = "About"
GUIDE = "Guide"
NOTES = "Text Extraction"

_textfile = {
    ABOUT: ("about",),
    GUIDE: ("guide",),
    NOTES: ("emailextract",),
}

folder = os.path.dirname(__file__)

for k in list(_textfile.keys()):
    _textfile[k] = tuple(
        [os.path.join(folder, ".".join((n, "txt"))) for n in _textfile[k]]
    )

del folder, k, os
