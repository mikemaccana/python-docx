#!/usr/bin/env python

"""
This file makes a .docx (Word 2007) file from scratch, showing off most of the
features of python-docx.

If you need to make documents from scratch, you can use this file as a basis
for your work.

Part of Python's docx module - http://github.com/mikemaccana/python-docx
See LICENSE for licensing information.
"""

from docx import *

if __name__ == '__main__':
    # Default set of relationshipships - the minimum components of a document
    relationships = relationshiplist()

    # Make a new document tree - this is the main part of a Word document
    document = newdocument()

    # This xpath location is where most interesting content lives
    body = document.xpath('/w:document/w:body', namespaces=nsprefixes)[0]

    # Append two headings and a paragraph
    body.append(heading("Welcome to Python's docx module", 1))
    body.append(heading('Make and edit docx in 200 lines of pure Python', 2))
    body.append(paragraph('The module was created when I was looking for a '
        'Python support for MS Word .doc files on PyPI and Stackoverflow. '
        'Unfortunately, the only solutions I could find used:'))

    # Add a numbered list
    points = [ 'COM automation'
             , '.net or Java'
             , 'Automating OpenOffice or MS Office'
             ]
    for point in points:
        body.append(paragraph(point, style='ListNumber'))
    body.append(paragraph([('For those of us who prefer something simpler, I '
                          'made docx.', 'i')]))    
    body.append(heading('Making documents', 2))
    body.append(paragraph('The docx module has the following features:'))

    # Add some bullets
    points = ['Paragraphs', 'Bullets', 'Numbered lists',
              'Multiple levels of headings', 'Tables', 'Document Properties']
    for point in points:
        body.append(paragraph(point, style='ListBullet'))

    body.append(paragraph('Tables are just lists of lists, like this:'))
    # Append a table
    tbl_rows = [ ['A1', 'A2', 'A3']
               , ['B1', 'B2', 'B3']
               , ['C1', 'C2', 'C3']
               ]
    body.append(table(tbl_rows))

    body.append(heading('Editing documents', 2))
    body.append(paragraph('Thanks to the awesomeness of the lxml module, '
                          'we can:'))
    points = [ 'Search and replace'
             , 'Extract plain text of document'
             , 'Add and delete items anywhere within the document'
             ]
    for point in points:
        body.append(paragraph(point, style='ListBullet'))

    # Add an image
    relationships, picpara = picture(relationships, 'image1.png',
                                     'This is a test description')
    body.append(picpara)

    # Search and replace
    print 'Searching for something in a paragraph ...',
    if search(body, 'the awesomeness'):
        print 'found it!'
    else:
        print 'nope.'

    print 'Searching for something in a heading ...',
    if search(body, '200 lines'):
        print 'found it!'
    else:
        print 'nope.'

    print 'Replacing ...',
    body = replace(body, 'the awesomeness', 'the goshdarned awesomeness')
    print 'done.'

    # Add a pagebreak
    body.append(pagebreak(type='page', orient='portrait'))

    body.append(heading('Ideas? Questions? Want to contribute?', 2))
    body.append(paragraph('Email <python.docx@librelist.com>'))

    # Create our properties, contenttypes, and other support files
    title    = 'Python docx demo'
    subject  = 'A practical example of making docx from Python'
    creator  = 'Mike MacCana'
    keywords = ['python', 'Office Open XML', 'Word']

    coreprops = coreproperties(title=title, subject=subject, creator=creator,
                               keywords=keywords)
    appprops = appproperties()
    contenttypes = contenttypes()
    websettings = websettings()
    wordrelationships = wordrelationships(relationships)

    # Save our document
    savedocx(document, coreprops, appprops, contenttypes, websettings,
             wordrelationships, 'Welcome to the Python docx module.docx')

