#!/usr/bin/env python2.6
'''
This file opens a docx (Office 2007) file and dumps the text.

If you need to extract text from documents, use this file as a basis for your work.

Part of Python's docx module - http://github.com/mikemaccana/python-docx
See LICENSE for licensing information.
'''
from docx import *
import sys
if __name__ == '__main__':        
    try:
        document = opendocx(sys.argv[1])
    except:
        print('Please supply a filename. For example:')    
        print('''  example-extracttext.py 'My Office 2007 document.docx' ''')    
        exit()
    ## Fetch all the text out of the document we just created        
    paratextlist = getdocumenttext(document)    

    # Note that if using shell redirection &>, 1> 2> etc) Python tries to 
    # change the unicode into ASCII and fails - even with a UTF-8 $LANG
    # As a workaround, create our own ASCII copy of the list.
    asciiparatextlist = []
    for paratext in paratextlist:
        asciiparatextlist.append(paratext.encode("ascii", "backslashreplace"))
    
    ## Print our documnts test with two newlines under each paragraph
    print '\n\n'.join(asciiparatextlist)
    
        