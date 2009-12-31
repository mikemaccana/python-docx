Python docx
===========

## Introduction

The docx module reads and writes Microsoft Office Word 2007 docx files.

These are referred to as 'WordML', 'Office Open XML' and 'Open XML' by Microsoft.

They can be opened in Microsoft Office 2007, Microsoft Mac Office 2008, OpenOffice.org 2.2, and Apple iWork 08.

The module was created when I was looking for a Python support for MS Word .doc files, but could only find various hacks involving COM automation, calling .net or Java, or automating OpenOffice or MS Office.

The docx module has the following features:

### Making documents

The docx module has the following features:
- Paragraphs
- Bullets
- Numbered lists
- Set document properties (author, company, etc)
- Multiple levels of headings
- Tables

### Editing documents

Thanks to the awesomeness of the lxml module, we can:
- Search and replace
- Extract plain text of document
- Add and delete items anywhere within the document
- Change document properties
- Run xpath queries against particular locations in the document - useful for retrieving data from user-completed templates.

### Ideas & To Do List

- Images
- Document health checks
- Egg
- Markdown conversion support

Licensed under the MIT license: http://www.opensource.org/licenses/mit-license.php

Ideas? Questions? Want to chat?
Email <python.docx@librelist.com>
