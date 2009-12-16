#!/usr/bin/env python26
'''
Open and modify Microsoft Word 2007 docx files (called 'OpenXML' and 'Office OpenXML' by Microsoft)

TODO:
- Return all text in document
- return document properties dict
- Read word XML reference 
- Functions to recieve dict and put into table
key is left column, data is right column
- Package for easy_install
- rest converter??
'''

from lxml import etree
import zipfile
import sys
#import ipdb

wordnamespaces = {
    'mv':'urn:schemas-microsoft-com:mac:vml',
    'mo':'http://schemas.microsoft.com/office/mac/office/2008/main',
    've':'http://schemas.openxmlformats.org/markup-compatibility/2006',
    'o':'urn:schemas-microsoft-com:office:office',
    'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'm':'http://schemas.openxmlformats.org/officeDocument/2006/math',
    'v':'urn:schemas-microsoft-com:vml',
    'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'w10':'urn:schemas-microsoft-com:office:word',
    'wne':'http://schemas.microsoft.com/office/word/2006/wordml',
    'wp':'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    }

def opendocx(file):
    '''Open a docx file, return a document XML tree'''
    mydoc = zipfile.ZipFile(file)
    xmlcontent = mydoc.read('word/document.xml')
    document = etree.fromstring(xmlcontent)    
    return document

def makeelement(tagname,tagattributes=None,tagtext=None):
    '''Create an element & return it'''
    namespace = '{'+wordnamespaces['w']+'}'    
    newelement = etree.Element(namespace+tagname)
    # Add attributes with namespaces
    if tagattributes:
        for tagattribute in tagattributes:
            newelement.set(namespace+tagattribute, tagattributes[tagattribute])
    if tagtext:
        newelement.text = tagtext    
    return newelement
    

def addparagraph(paratext):
    '''Make a new paragraph element, containing a run, and some text. Return the paragraph element.'''
    # Make our elements
    paragraph = makeelement('p')
    run = makeelement('r')
    text = makeelement('t',tagtext=paratext)
    # Add the text the run, and the run to the paragraph
    run.append(text)
    paragraph.append(run)    
    # Return the combined paragraph
    return paragraph

def addheading(headingtext,headinglevel):
    '''Make a new heading, return the heading element'''
    # Make our elements
    paragraph = makeelement('p')
    pr = makeelement('pPr')
    pStyle = makeelement('pStyle',tagattributes={'val':'Heading'+str(headinglevel)})    
    run = makeelement('r')
    text = makeelement('t',tagtext=headingtext)
    # Add the text the run, and the run to the paragraph
    pr.append(pStyle)
    run.append(text)
    paragraph.append(pr)   
    paragraph.append(run)    
    # Return the combined paragraph
    return paragraph   


def addtable(contents):
    '''Get a list of lists, return a table'''
    table = makeelement('tbl')
    columns = len(contents[0][0])
    
    # Table properties
    tableprops = makeelement('tblPr')
    tablestyle = makeelement('tblStyle',tagattributes={'val':'ColorfulGrid-Accent1'})
    tablewidth = makeelement('tblW',tagattributes={'w':'0','type':'auto'})
    tablelook = makeelement('tblLook',tagattributes={'val':'0400'})
    for tableproperty in [tablestyle,tablewidth,tablelook]:
        tableprops.append(tableproperty)
    table.append(tableprops)    
    
    # Table Grid    
    tablegrid = makeelement('tblGrid')
    for _ in range(columns):
        tablegrid.append(makeelement('gridCol',tagattributes={'gridCol':'2390'}))
    table.append(tablegrid)     

    # Heading Row    
    row = makeelement('tr')
    rowprops = makeelement('trPr')
    cnfStyle = makeelement('cnfStyle',tagattributes={'val':'000000100000'})
    rowprops.append(cnfStyle)
    row.append(rowprops)
    for heading in contents[0]:
        cell = makeelement('tc')  
        # Cell properties  
        cellprops = makeelement('tcPr')
        cellwidth = makeelement('tcW',tagattributes={'w':'2390','type':'dxa'})
        cellstyle = makeelement('shd',tagattributes={'val':'clear','color':'auto','fill':'548DD4','themeFill':'text2','themeFillTint':'99'})
        cellprops.append(cellwidth)
        cellprops.append(cellstyle)
        cell.append(cellprops)
        
        # Paragraph (Content)
        cell.append(addparagraph(heading))
        row.append(cell)
    table.append(row)    
        
    # Contents Rows   
    for contentrow in contents[1:]:
        row = makeelement('tr')     
        for content in contentrow:   
            cell = makeelement('tc')
            # Properties
            cellprops = makeelement('tcPr')
            cellwidth = makeelement('tcW',tagattributes={'type':'dxa'})
            cellprops.append(cellwidth)
            cell.append(cellprops)

            # Paragraph (Content)
            cell.append(addparagraph(content))
            row.append(cell)    
        table.append(row)   
    return table                 
                        

def search(phrase):
    '''Recieve a search, return the results'''
    results = False
    return results

def replace(search,replace):
    '''Replace all occurences of string with a different string'''
    results = False
    return results
    
def savedocx(document,newfilename):
    '''Save a modified document'''
    documentstring = etree.tostring(document, pretty_print=True)
    newfile = zipfile.ZipFile(newfilename,mode='w')
    newfile.writestr('word/document.xml',documentstring)
    # Add support files
    for xmlfile in [ 
    '[Content_Types].xml',
    '_rels/.rels',
    'docProps/core.xml',
    'docProps/thumbnail.jpeg',
    'docProps/app.xml',
    'word/webSettings.xml',
    'word/_rels/document.xml.rels',
    'word/styles.xml',
    'word/theme/',
    'word/theme/theme1.xml',
    'word/settings.xml',
    'word/fontTable.xml']:
        newfile.write('template/'+xmlfile,xmlfile)
    print 'Saved new file to: '+newfilename
    return    

    
if __name__ == '__main__':        
    #document = opendocx('Hello world.docx')
    document = etree.parse('template/word/document.xml')
    #ipdb.set_trace()
    
    # This location is where most document content lives 
    docbody = document.xpath('/w:document/w:body', namespaces=wordnamespaces)[0]
    
    # Append two headings
    docbody.append(addheading('All your base are belong to us',1)  )   
    docbody.append(addheading('You have no chance to survive. ',2))

    # Append a table
    docbody.append(addtable([['A1','A2','A3'],['B1','B2','B3'],['C1','C2','C3']]))

    # Append a paragraph element 
    newpara = addparagraph(paratext='Make your time. Hahaha')    
    docbody.append(newpara)
    
    # Save our document
    savedocx(document,sys.argv[1])
