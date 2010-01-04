#!/usr/bin/env python2.6
'''
This file makes an docx (Office 2007) file from scratch, showing off most of python-docx's features.

If you need to make documents from scratch, use this file as a basis for your work.

Part of Python's docx module - http://github.com/mikemaccana/python-docx
See LICENSE for licensing information.
'''
from docx import *

if __name__ == '__main__':        
    # Make a new document tree - this is the main part of a Word document
    document = newdocument()
    
    # This xpath location is where most interesting content lives 
    docbody = document.xpath('/w:document/w:body', namespaces=nsprefixes)[0]
    
    # Append two headings and a paragraph
    docbody.append(heading('''Welcome to Python's docx module''',1)  )   
    docbody.append(heading('Make and edit docx in 200 lines of pure Python',2))
    docbody.append(paragraph('The module was created when I was looking for a Python support for MS Word .doc files on PyPI and Stackoverflow. Unfortunately, the only solutions I could find used:'))

    # Add a numbered list
    for point in ['''COM automation''','''.net or Java''','''Automating OpenOffice or MS Office''']:
        docbody.append(paragraph(point,style='ListNumber'))
    docbody.append(paragraph('''For those of us who prefer something simpler, I made docx.''')) 
    
    docbody.append(heading('Making documents',2))
    docbody.append(paragraph('''The docx module has the following features:'''))

    # Add some bullets
    for point in ['Paragraphs','Bullets','Numbered lists','Multiple levels of headings','Tables']:
        docbody.append(paragraph(point,style='ListBullet'))

    docbody.append(paragraph('Tables are just lists of lists, like this:'))
    # Append a table
    docbody.append(table([['A1','A2','A3'],['B1','B2','B3'],['C1','C2','C3']]))

    docbody.append(heading('Editing documents',2))
    docbody.append(paragraph('Thanks to the awesomeness of the lxml module, we can:'))
    for point in ['Search and replace','Extract plain text of document','Add and delete items anywhere within the document']:
        docbody.append(paragraph(point,style='ListBullet'))
        
    # Add an image (beta)
    docbody.append(picture())
 
    # Search and replace 
    document = replace(document,'the','the goshdarned')

    # Add a pagebreak
    docbody.append(pagebreak(type='page', orient='portrait'))

    docbody.append(heading('Ideas? Questions? Want to contribute?',2))
    docbody.append(paragraph('''Email <python.docx@librelist.com>'''))
    
    properties = docproperties('Python docx demo','A practical example of making docx from Python','Mike MacCana',['python','Office Open XML','Word'])
    contenttypes = contenttypes()
    websettings = websettings()
    
    # Save our document
    savedocx(document,properties,contenttypes,websettings,'Welcome to the Python docx module.docx')