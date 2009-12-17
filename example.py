#!/usr/bin/env python2.6
'''
Open and modify Microsoft Word 2007 docx files (called 'OpenXML' and 'Office OpenXML' by Microsoft)
'''
from docx import *

if __name__ == '__main__':        
    #document = opendocx('Hello world.docx')
    document = newdocument()
    
    # This location is where most document content lives 
    docbody = document.xpath('/w:document/w:body', namespaces=wordnamespaces)[0]
    
    # Append two headings
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
 
    # Search and replace 
    document = replace(document,'the','the goshdarned')

    docbody.append(heading('Ideas? Questions? Want to chat?',2))
    docbody.append(paragraph('''Email <python.docx@librelist.com>'''))
    
    ## Fetch all the text out of the document we just created        
    #print getdocumenttext(document)
    #print etree.tostring(document, pretty_print=True)

    # Save our document
    savedocx(document,'Welcome to the Python docx module.docx')