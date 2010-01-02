#!/usr/bin/env python2.6
'''
Open and modify Microsoft Word 2007 docx files (called 'OpenXML' and 'Office OpenXML' by Microsoft)

Part of Python's docx module - http://github.com/mikemaccana/python-docx
See LICENSE for licensing information.
'''

from lxml import etree
import zipfile
import re
import time

# All Word prefixes / namespace matches used in document.xml & core.xml
# LXML doesn't actually use prefixes (just the real namespace) , but these
# make it easier to copy Word output more easily. 
nsprefixes = {
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
    'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
    'pic':'http://schemas.openxmlformats.org/drawingml/2006/picture',
    'cp':"http://schemas.openxmlformats.org/package/2006/metadata/core-properties", 
    'dc':"http://purl.org/dc/elements/1.1/", 
    'dcterms':"http://purl.org/dc/terms/",
    'dcmitype':"http://purl.org/dc/dcmitype/",
    'xsi':"http://www.w3.org/2001/XMLSchema-instance",
    }

def opendocx(file):
    '''Open a docx file, return a document XML tree'''
    mydoc = zipfile.ZipFile(file)
    xmlcontent = mydoc.read('word/document.xml')
    document = etree.fromstring(xmlcontent)    
    return document

def newdocument():
    document = makeelement('document')
    document.append(makeelement('body'))
    return document

def makeelement(tagname,tagtext=None,nsprefix='w',attributes=None,attributenamespace=None):
    '''Create an element & return it''' 
    namespace = '{'+nsprefixes[nsprefix]+'}'
    newelement = etree.Element(namespace+tagname)
    # Add attributes with namespaces
    if attributes:
        # If they haven't bothered setting attribute namespace, use the same one as the tag
        if not attributenamespace:
            attributenamespace = namespace    
        for tagattribute in attributes:
            newelement.set(attributenamespace+tagattribute, attributes[tagattribute])
    if tagtext:
        newelement.text = tagtext    
    return newelement

def pagebreak(type='page', orient='portrait'):
    '''Insert a break, default 'page'.
    See http://openxmldeveloper.org/forums/thread/4075.aspx
    Return our page break element.'''
    # Need to enumerate different types of page breaks.
    if type not in ['page', 'section']:
            raiseError('Page break style "%s" not implemented.' % type)

    pagebreak = makeelement('p')
    if type == 'page':
        run = makeelement('r')
        br = makeelement('br',attributes={'type':type})

        run.append(br)
        pagebreak.append(run)

    elif type == 'section':
        pPr = makeelement('pPr')
        sectPr = makeelement('sectPr')
        if orient == 'portrait':
            pgSz = makeelement('pgSz',attributes={'w':'12240','h':'15840'})
        elif orient == 'landscape':
            pgSz = makeelement('pgSz',attributes={'h':'12240','w':'15840', 'orient':'landscape'})

        sectPr.append(pgSz)
        pPr.append(sectPr)
        pagebreak.append(pPr)

    return pagebreak    

def paragraph(paratext,style='BodyText',breakbefore=False):
    '''Make a new paragraph element, containing a run, and some text. 
    Return the paragraph element.'''
    # Make our elements
    paragraph = makeelement('p')
    run = makeelement('r')    

    # Insert lastRenderedPageBreak for assistive technologies like
    # document narrators to know when a page break occurred.
    if breakbefore:
        lastRenderedPageBreak = makeelement('lastRenderedPageBreak')
        run.append(lastRenderedPageBreak)
    
    text = makeelement('t',tagtext=paratext)
    pPr = makeelement('pPr')
    pStyle = makeelement('pStyle',attributes={'val':style})
    pPr.append(pStyle)

                
    # Add the text the run, and the run to the paragraph
    run.append(text)    
    paragraph.append(pPr)    
    paragraph.append(run)    
    # Return the combined paragraph
    return paragraph


def heading(headingtext,headinglevel):
    '''Make a new heading, return the heading element'''
    # Make our elements
    paragraph = makeelement('p')
    pr = makeelement('pPr')
    pStyle = makeelement('pStyle',attributes={'val':'Heading'+str(headinglevel)})    
    run = makeelement('r')
    text = makeelement('t',tagtext=headingtext)
    # Add the text the run, and the run to the paragraph
    pr.append(pStyle)
    run.append(text)
    paragraph.append(pr)   
    paragraph.append(run)    
    # Return the combined paragraph
    return paragraph   


def table(contents):
    '''Get a list of lists, return a table'''
    table = makeelement('tbl')
    columns = len(contents[0][0])    
    # Table properties
    tableprops = makeelement('tblPr')
    tablestyle = makeelement('tblStyle',attributes={'val':'ColorfulGrid-Accent1'})
    tablewidth = makeelement('tblW',attributes={'w':'0','type':'auto'})
    tablelook = makeelement('tblLook',attributes={'val':'0400'})
    for tableproperty in [tablestyle,tablewidth,tablelook]:
        tableprops.append(tableproperty)
    table.append(tableprops)    
    # Table Grid    
    tablegrid = makeelement('tblGrid')
    for _ in range(columns):
        tablegrid.append(makeelement('gridCol',attributes={'gridCol':'2390'}))
    table.append(tablegrid)     
    # Heading Row    
    row = makeelement('tr')
    rowprops = makeelement('trPr')
    cnfStyle = makeelement('cnfStyle',attributes={'val':'000000100000'})
    rowprops.append(cnfStyle)
    row.append(rowprops)
    for heading in contents[0]:
        cell = makeelement('tc')  
        # Cell properties  
        cellprops = makeelement('tcPr')
        cellwidth = makeelement('tcW',attributes={'w':'2390','type':'dxa'})
        cellstyle = makeelement('shd',attributes={'val':'clear','color':'auto','fill':'548DD4','themeFill':'text2','themeFillTint':'99'})
        cellprops.append(cellwidth)
        cellprops.append(cellstyle)
        cell.append(cellprops)        
        # Paragraph (Content)
        cell.append(paragraph(heading))
        row.append(cell)
    table.append(row)            
    # Contents Rows   
    for contentrow in contents[1:]:
        row = makeelement('tr')     
        for content in contentrow:   
            cell = makeelement('tc')
            # Properties
            cellprops = makeelement('tcPr')
            cellwidth = makeelement('tcW',attributes={'type':'dxa'})
            cellprops.append(cellwidth)
            cell.append(cellprops)
            # Paragraph (Content)
            cell.append(paragraph(content))
            row.append(cell)    
        table.append(row)   
    return table                 

def picture():
    '''Create a pragraph containing an image - FIXME - not implemented yet'''
    # Word uses paragraphs to contain images
    # http://openxmldeveloper.org/articles/462.aspx
    #resourceid = rId5
    #newrelationship = makeelement('Relationship',attributes={'Id':resourceid,'Type':'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'},Target=filename)
    
    # Now make drawing element
    #newpara = makeelement('deleteme',style='BodyText')

    
    blipfill = makeelement('blipFill',nsprefix='a')
    blipfill.append(makeelement('blip',nsprefix='a',attributes={'embed':'rId5'}))
    stretch = makeelement('stretch',nsprefix='a')
    stretch.append(makeelement('fillRect',nsprefix='a'))
    blipfill.append(stretch)

    sppr = makeelement('spPr',nsprefix='pic')
    xfrm = makeelement('xfrm',nsprefix='a')
    xfrm.append(makeelement('off',nsprefix='a',attributes={'x':'0','y':'0'}))
    xfrm.append(makeelement('ext',nsprefix='a',attributes={'cx':'5486400','cy':'3429000'}))
    prstgeom = makeelement('prstGeom',nsprefix='a',attributes={'prst':'rect'})
    prstgeom.append(makeelement('avLst',nsprefix='a'))
    sppr.append(xfrm)
    sppr.append(prstgeom)

    nvpicpr = makeelement('nvPicPr',nsprefix='a')
    cnvpr = makeelement('cNvPr',nsprefix='a',attributes={'id':'0','name':'aero_glow_v2_1920x1200.png'})
    cnvpicpr = makeelement('cNvPicPr')
    nvpicpr.append(cnvpicpr)
    nvpicpr.append(cnvpr)
    
    pic = makeelement('pic',nsprefix='pic')
    pic.append(blipfill)
    pic.append(sppr)
    pic.append(nvpicpr)


    graphicdata = makeelement('graphicData',nsprefix='a',attributes={'uri':'http://schemas.openxmlformats.org/drawingml/2006/picture'})
    graphicdata.append(pic)

    graphic = makeelement('graphic',nsprefix='a')
    graphic.append(graphicdata)

    framepr = makeelement('cNvGraphicFramePr',nsprefix='a')
    framelocks = makeelement('graphicFrameLocks',nsprefix='a',attributes={'noChangeAspect':'1'})
    framepr.append(framelocks)

    makeelement('drawing')
    inline = makeelement('inline',attributes={'distT':"0",'distB':"0",'distL':"0",'distR':"0"},nsprefix='wp')
    extent = makeelement('extent',nsprefix='a',attributes={'cx':'5486400','cy':'3429000'})
    effectextent = makeelement('effectExtent',nsprefix='a',attributes={'l':'25400','t':'0','r':'0','b':'0'})
    docpr = makeelement('docPr',nsprefix='a',attributes={'id':'1','name':'Picture 0','descr':'aero_glow_v2_1920x1200.png'})
    inline.append(extent)
    inline.append(effectextent)
    inline.append(docpr)
    inline.append(framepr)
    inline.append(graphic)
    drawing = makeelement('drawing')
    drawing.append(inline)
    return drawing
    '''
    <w:drawing>
    	<wp:inline distT="0" distB="0" distL="0" distR="0">
    		<wp:extent cx="5486400" cy="3429000"/>
    		<wp:effectExtent l="25400" t="0" r="0" b="0"/>
    		<wp:docPr id="1" name="Picture 0" descr="aero_glow_v2_1920x1200.png"/>
    		<wp:cNvGraphicFramePr>
    			<a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/>
    		<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
    			<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
    				<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
    					<pic:nvPicPr>
    						<pic:cNvPr id="0" name="aero_glow_v2_1920x1200.png"/>
    						<pic:cNvPicPr/>
    					<pic:blipFill>
    						<a:blip r:embed="rId5"/>
    						<a:stretch>
    							<a:fillRect/>
    					<pic:spPr>
    						<a:xfrm>
    							<a:off x="0" y="0"/>
    							<a:ext cx="5486400" cy="3429000"/>
    						<a:prstGeom prst="rect">
    							<a:avLst/>
    '''


def search(document,search):
    '''Search a document for a regex, return '''
    results = False
    searchre = re.compile(search)
    for element in document.iter():
        if element.tag == 'w'+'t':
            if element.text:
                if searchre.match(element.text):
                    results = True
    return results

def replace(document,search,replace):
    '''Replace all occurences of string with a different string, return updated document'''
    newdocument = document
    searchre = re.compile(search)
    for element in newdocument.iter():
        if element.tag == 'w'+'t':
            if element.text:
                if searchre.search(element.text):
                    element.text = re.sub(search,replace,element.text)
    return newdocument


def getdocumenttext(document):
    '''Return the raw text of a document, as a list of paragraphs.'''
    paratextlist=[]   
        
    # Compile a list of all paragraph (p) elements
    paralist = []
    for element in document.iter():
        # Find p (paragraph) elements
        if element.tag == 'w'+'p':
            paralist.append(element)
    
    # Since a single sentence might be spread over multiple text elements, iterate through each 
    # paragraph, appending all text (t) children to that paragraphs text.     
    for para in paralist:      
        paratext=u''  
        # Loop through each paragraph
        for element in para.iter():
            # Find t (text) elements
            if element.tag == 'w'+'t':
                if element.text:
                    paratext = paratext+element.text

        # Add our completed paragraph text to the list of paragraph text    
        if not len(paratext) == 0:
            paratextlist.append(paratext)                    
    return paratextlist        

def docproperties(title,subject,creator,keywords,lastmodifiedby=None):
    '''Makes document properties. '''
    # OpenXML uses the term 'core' to refer to the 'Dublin Core' specification used to make the properties.  
    docprops = makeelement('coreProperties',nsprefix='cp')    
    docprops.append(makeelement('title',tagtext=title,nsprefix='dc'))
    docprops.append(makeelement('subject',tagtext=subject,nsprefix='dc'))
    docprops.append(makeelement('creator',tagtext=creator,nsprefix='dc'))
    docprops.append(makeelement('keywords',tagtext=','.join(keywords),nsprefix='cp'))    
    if not lastmodifiedby:
        lastmodifiedby = creator
    docprops.append(makeelement('lastModifiedBy',tagtext=lastmodifiedby,nsprefix='cp'))
    docprops.append(makeelement('revision',tagtext='1',nsprefix='cp'))
    docprops.append(makeelement('category',tagtext='Examples',nsprefix='cp'))
    docprops.append(makeelement('description',tagtext='Examples',nsprefix='dc'))
    currenttime = time.strftime('%Y-%m-%dT-%H:%M:%SZ')
    # FIXME - creating these items manually fails - but we can live without them for now.
    '''	What we're going for:
    <dcterms:created xsi:type="dcterms:W3CDTF">2010-01-01T21:07:00Z</dcterms:created>
    <dcterms:modified xsi:type="dcterms:W3CDTF">2010-01-01T21:20:00Z</dcterms:modified>
    currenttime'''
    #docprops.append(makeelement('created',nsprefix='dcterms',
    #attributes={'type':'dcterms:W3CDTF'},tagtext='2010-01-01T21:07:00Z',attributenamespace='xsi'))
    #docprops.append(makeelement('modified',nsprefix='dcterms',
    #attributes={'type':'dcterms:W3CDTF'},tagtext='2010-01-01T21:07:00Z',attributenamespace='xsi'))
    return docprops



def savedocx(document,properties,newfilename):
    '''Save a modified document'''
    newfile = zipfile.ZipFile(newfilename,mode='w')
    # Write our generated document
    documentstring = etree.tostring(document, pretty_print=True)
    newfile.writestr('word/document.xml',documentstring)
    # And it's properties
    propertiesstring = etree.tostring(properties, pretty_print=True)
    newfile.writestr('docProps/core.xml',propertiesstring)
    # Add support files
    for xmlfile in [ 
    '[Content_Types].xml',
    '_rels/.rels',
    'docProps/thumbnail.jpeg',
    'docProps/app.xml',
    'word/webSettings.xml',
    'word/_rels/document.xml.rels',
    'word/styles.xml',
    'word/numbering.xml',
    'word/theme/',
    'word/theme/theme1.xml',
    #'word/media/image1.png',
    'word/settings.xml',
    'word/fontTable.xml']:
        newfile.write('template/'+xmlfile,xmlfile)
    print 'Saved new file to: '+newfilename
    return
    

