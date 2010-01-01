Python docx
===========

## Introduction

The docx module reads and writes Microsoft Office Word 2007 docx files.

These are referred to as 'WordML', 'Office Open XML' and 'Open XML' by Microsoft.

They can be opened in Microsoft Office 2007, Microsoft Mac Office 2008, Google Docx, OpenOffice.org 2.2, and Apple iWork 08.

They also validate as well formed XML at http://validator.w3.org/check

The module was created when I was looking for a Python support for MS Word .doc files, but could only find various hacks involving COM automation, calling .net or Java, or automating OpenOffice or MS Office.

The docx module has the following features:

### Making documents

Features for making documents include:

- Paragraphs
- Bullets
- Numbered lists
- Set document properties (author, company, etc)
- Multiple levels of headings
- Tables
- Section and page breaks

### Editing documents

Thanks to the awesomeness of the lxml module, we can:

- Search and replace
- Extract plain text of document
- Add and delete items anywhere within the document
- Change document properties
- Run xpath queries against particular locations in the document - useful for retrieving data from user-completed templates.

### Getting started

- Use **easy_install** to fetch the **lxml** module
- Download the files above
- Open **example.py** which creates and modifies a sample docx document. 

### Ideas & To Do List

- Images
- Document health checks
- Egg
- Markdown conversion support

### Authors & Contact

If you have idea, or would like to add functionality contact the Python docx mailing list at <python.docx@librelist.com>

- Mike MacCana - main developer
- Marcin Wielgoszewski - support for breaks & document narrators in paragraphs

### License

Licensed under the [MIT license](http://www.opensource.org/licenses/mit-license.php)
Short version: this code is copyrighted to me (Mike MacCana), I give you permission to do what you want with it except remove my name from the credits. See the LICENSE file for specific terms.
