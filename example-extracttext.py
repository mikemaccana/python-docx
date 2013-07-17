#!/usr/bin/env python
"""
This file opens a docx (Office 2007) file and dumps the text.

If you need to extract text from documents, use this file as a basis for your
work.

Part of Python's docx module - http://github.com/mikemaccana/python-docx
See LICENSE for licensing information.
"""

import sys

from docx import opendocx, getdocumenttext

if __name__ == '__main__':
    try:
        document = opendocx(sys.argv[1])
        newfile = open(sys.argv[2], 'w')
    except:
        print(
            "Please supply an input and output file. For example:\n"
            "  example-extracttext.py 'My Office 2007 document.docx' 'outp"
            "utfile.txt'"
        )
        exit()

    # Fetch all the text out of the document we just created
    paratextlist = getdocumenttext(document)

    # Make explicit unicode version
    newparatextlist = []
    for paratext in paratextlist:
        newparatextlist.append(paratext.encode("utf-8"))

    # Print out text of document with two newlines under each paragraph
    newfile.write('\n\n'.join(newparatextlist))
