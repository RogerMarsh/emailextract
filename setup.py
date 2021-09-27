# setup.py
# Copyright 2017 Roger Marsh
# Licence: See LICENCE (BSD licence)
"""emailextract setup file."""

from setuptools import setup

if __name__ == "__main__":

    long_description = open("README").read()

    setup(
        name="emailextract",
        version="0.7.3.dev0",
        description="Extract text from emails",
        author="Roger Marsh",
        author_email="roger.marsh@solentware.co.uk",
        url="http://www.solentware.co.uk",
        packages=[
            "emailextract",
            "emailextract.core",
            "emailextract.gui",
            "emailextract.help",
        ],
        package_data={
            "emailextract.help": ["*.txt"],
        },
        long_description=long_description,
        license="BSD",
        classifiers=[
            "License :: OSI Approved :: BSD License",
            "Programming Language :: Python :: 3.6",
            "Programming Language :: Python :: 3.7",
            "Programming Language :: Python :: 3.8",
            "Programming Language :: Python :: 3.9",
            "Operating System :: OS Independent",
            "Topic :: Other/Nonlisted Topic",
            "Intended Audience :: End Users/Desktop",
            "Intended Audience :: Developers",
            "Development Status :: 3 - Alpha",
        ],
        install_requires=["solentware-misc==1.3.1.dev0"],
        dependency_links=[
            "http://solentware.co.uk/files/solentware-misc-1.3.1.dev0.tar.gz",
        ],
    )
