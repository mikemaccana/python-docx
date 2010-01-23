Python docx
===========

## Introduction

The docx module creates, reads and writes Microsoft Office Word 2007 docx files.

These are referred to as 'WordML', 'Office Open XML' and 'Open XML' by Microsoft.

These documents can be opened in Microsoft Office 2007 / 2010, Microsoft Mac Office 2008, Google Docs, OpenOffice.org 3, and Apple iWork 08.

They also [validate as well formed XML](http://validator.w3.org/check).

The module was created when I was looking for a Python support for MS Word .doc files, but could only find various hacks involving COM automation, calling .net or Java, or automating OpenOffice or MS Office.

The docx module has the following features:

### Making documents

Features for making documents include:

- Paragraphs
- Bullets
- Numbered lists
- Document properties (author, company, etc)
- Multiple levels of headings
- Tables
- Section and page breaks
- Images

### Editing documents

Thanks to the awesomeness of the lxml module, we can:

- Search and replace
- Extract plain text of document
- Add and delete items anywhere within the document
- Change document properties
- Run xpath queries against particular locations in the document - useful for retrieving data from user-completed templates.

# Getting started

## Making a Document

- Just [download python docx](http://github.com/mikemaccana/python-docx/tarball/master).
- Use **easy_install** to fetch the **lxml** and **PIL** modules. 
- Then run: 

<pre>example-makedocument.py</pre>

Congratulations, you just made and then modified a Word document!

## Extracting Text from a Document

If you just want to extract the text from a Word file, run: 

    example-extracttext.py 'Some word file.docx' 'new file.txt' 

### Ideas & To Do List

- Further improvements to image handling
- Document health checks
- Egg
- Markdown conversion support

### We love forks & changes!

Check out the [HACKING](HACKING.markdown) to add your own changes!

### Want to talk? Need help?

Email <python.docx@librelist.com>.

### License

Licensed under the [MIT license](http://www.opensource.org/licenses/mit-license.php)
Short version: this code is copyrighted to me (Mike MacCana), I give you permission to do what you want with it except remove my name from the credits. See the LICENSE file for specific terms.
