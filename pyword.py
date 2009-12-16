#!/usr/bin/env python26
'''
Open and modify Microsoft Word 2007 docx files (called 'Office OpenXML' by Microsoft)
'''

from lxml import etree
import zipfile
#import ipdb

def opendocx(file):
    '''Open a docx file, return a document XML tree'''
    mydoc = zipfile.ZipFile(file)
    xmlcontent = mydoc.read('word/document.xml')
    document = etree.fromstring(xmlcontent)    
    return document

def makeelement(tagname,tagattributes=None,tagtext=None,**kwargs):
    '''Create an element & return it'''
    newelement = etree.Element(tagname)
    if tagattributes:
        for tagattribute in tagattributes:
            newelement.set(tagattribute, tagattributes[tagattribute])
    if tagtext:
        newelement.text = tagtext    
    return newelement
    
def appendelement(addlocation,tagname,tagattributes=None,tagtext=None,**kwargs):
    '''Make and append an element at a path location'''
    global document
    newelement = makeelement(tagname,tagattributes,tagtext)
    location = document.xpath(addlocation, namespaces={'mv':'urn:schemas-microsoft-com:mac:vml',
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
        })
    location[0].insert(1, newelement)
    return

def appendparagraph(document,text):
    '''Make a new text paragraph, return the modified document tree'''
    #newelement
    return
    
def search(phrase):
    '''Recieve a search, return the results'''
    results = False
    return results
    
def savedocx(oldfilename,document,newfilename=None):
    return    
    
if __name__ == '__main__':        
    #document = opendocx('Hello world.docx')
    document = etree.parse('sample.xml')
    testelement = makeelement('success')
    #ipdb.set_trace()
    appendelement('/w:document/w:body','p',tagattributes={'rsidR':'008D6863','rsidRDefault':'00590D07'})
    appendelement('/w:document/w:body/p','r')
    appendelement('/w:document/w:body/p/r','t',tagtext='success')

    resultsfile = open('results','w')
    newxml = etree.tostring(document, pretty_print=True)
    print newxml
    resultsfile.write(newxml)    