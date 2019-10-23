---
title: "Release Notes for Office Insiders"
ms.author: andrewmo
author: v-almuzz
manager: andrewmo
ms.date: 5/17/2019
ms.audience: Win32 Fast
ms.topic: reference
ms.service: o365-proplus-
localization_priority: Critical
ms.collection: RelNotes_ProPlus
description: "Provides Insiders Fast audience with the latest list of key new features, fixes or known issues"
---

# Release Notes for Office Insiders

This article contains release notes for Insider builds of Word, Excel, PowerPoint, Outlook, Access, and Project for Windows desktop. Every week, we’ll highlight interesting new features, important fixes, and any significant issues we want you to know about. Note that we often roll out features (and sometimes even fixes) to Insiders over a period of time. This allows us to ensure that things are working smoothly before releasing the feature to a wider audience. So, if you don’t see something described below, don't worry you'll get it eventually.  

> [!NOTE]
> - Release notes are posted weekly and may be a compilation of multiple builds
> - The release notes publication date may not match the actual build release date

[//]: # (DO NOT REMOVE)

## Version 1910: October 23
*Version 1910 (Build 12130.20210)*


[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates
### Excel

- **Create more accessible PDFs:** Create a PDF and the accessibility checker will point out accessibility issues to fix before you save. [Learn more](https://support.office.com/en-us/article/064625e0-56ea-4e16-ad71-3aa33bb4b7ed)

- **Convert files to improve accessibility:** Upgrade your files to the modern format to make them more accessible for everyone.

- **Data visualizer add-in:** Quickly create Visio flowcharts from Excel. [Learn more](https://support.office.com/en-us/article/bee3b5aa-aaaf-4401-acc6-276b711c763c)

### Outlook

- **TBD:** TBD

### PowerPoint

- **Ink-stant replay:** Animate an ink drawing so that it replays forward or backward during your slide show. [Learn more](https://support.office.com/en-us/article/fa4f044f-810b-43fe-b774-da04a0b37496)

- **Create more accessible PDFs:** Create a PDF and the accessibility checker will point out accessibility issues to fix before you save.

- **Convert files to improve accessibility:** Upgrade your files to the modern format to make them more accessible for everyone.

- **Optimize your presentation for all:** Accessibility Checker helps you arrange objects on your slides with screen readers in mind.

### Visio

- **Make polished Visio diagrams in Excel:** Quickly and easily visualize your data into polished Visio diagrams within Excel. [Learn more](https://support.office.com/en-us/article/bee3b5aa-aaaf-4401-acc6-276b711c763c)

### Word

- **Create more accessible PDFs:** Create a PDF and the accessibility checker will point out accessibility issues to fix before you save. [Learn more](https://support.office.com/en-us/article/064625e0-56ea-4e16-ad71-3aa33bb4b7ed)

- **Convert files to improve accessibility:** Upgrade your files to the modern format to make them more accessible for everyone.

- **Others see your changes quickly:** Co-authoring improvements mean your collaborators can see your changes faster than ever before.

- **Coauthoring improvements:** Improved the coauthoring experience by making it more likely that content changes will be received by others in real time.


[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

<br/>

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues
### Access

- <div>We fixed an issue where users could receive an &quot;inconsistent state&quot; error when using a shared database.</div>


### Excel

- <div>We resolved an performance issue with Asynchronous User-Defined Functions that was causing them to be run Synchronously.</div>


- <div>Resolved an issue where workbooks created in earlier versions of Office could cause Excel to hang when opened in current versions of Office</div>


- <div><span>We fixed an issue which could prevent a user from pasting hyperlinks in some protected sheets</span></div>


- <div><span>We fixed an issue which could cause the font name in the ribbon to be different from the font being used</span></div>


- <div>Resolved an issue where the hyperlink format could not be properly applied to some content</div>


- <div>Resolved an issue where formulas containing structured absolute references may be adjusted incorrectly</div>


- <div>Resolved an issue where content copied from Excel could appear incorrect when pasted into other Office applications</div>


- <h5>User Scenario</h5><br><h5>Defect</h5><br><h5>Fix</h5><br><h5>Dev/Test Notes</h5><br><h5>Code Review</h5><br>


- <div><span>We fixed an issue which could have caused scatter line charts from rendering properly when changing the series collection</span></div>


- <div><span>We fixed an issue where Excel could sometimes hang at launch</span></div>


- <div>Resolved an issue in inserting files as object from OneDrive</div>


- <div><span>We fixed an issue which could have prevented pivot tables from being refreshed during a co-authoring session</span></div>


- <div><span>We fixed an issue where the print area in print preview was not correct</span></div>


- <div>Fix to correct colors used in previews when inserting charts using chart templates</div>


### OneNote

- <div><span>We fixed an issue where a link to a copied page could navigate to the original page instead of the newly copied one</span></div>


### OPC

- <div><span></span></div><span style="font-size:11.0pt;font-family:&quot;Calibri&quot;,sans-serif;">We fixed an issue
that could result in inappropriate resource consumption by Outlook when
Protected Mode is disabled for Restricted Sites in Internet Explorer</span>


### Outlook

- <div><span style="font-size:11.0pt;font-family:&quot;Calibri&quot;,sans-serif;">We fixed an issue
that can prevent mailbox sync for customers with multiple mailboxes in Outlook
when migrating to modern authentication in Office 365.</span><br></div>


- <div><span>We fixed an issue where some characters in signature labels would not display in the dropdown menu</span></div>


- <div><span>We significantly improved the performance of room selection when there are a large number of rooms available</span></div>


- <div><span>We fixed an issue which could have prevented users from adding attachments to calendars</span></div>


- <div><span>We fixed an issue which caused error messages to display when replying to a digitally signed message</span></div>


- <div><span>We fixed an issue which prevented Expanded Find Flight results from appearing in search results</span></div>


- <div><span>We fixed an issue where the UI could get stuck in a compact view</span></div>


- <div><span>We fixed an issue where some users would incorrectly appear as Offline in a Group Schedule view</span></div>


- <div><span>We fixed an issue which could sometimes cause Unicode characters to appear when pasting text from an ANSI source</span></div>


- <div><span>We fixed an issue which could have caused duplication of mail folders</span></div>


- <div><span>We fixed an issue which could have caused an incorrect error message when attempting to send s/MIME encrypted e-mail</span></div>


- <div><span>We fixed an issue which could sometimes result in a crash when a user receives a 'Missed Conversation' message from Skype</span></div>


- <div><span>We fixed an issue which could have resulted in a memory leak</span></div>


- <div><span>We fixed an issue which could have reported permission errors when interacting with shared calendar folders</span></div>


- <div><span>We fixed an issue which could have caused the senders name to change when saved as a draft</span></div>


### PowerPoint

- <div><span>We fixed an issue where a user could experience an error upon printing to PDF</span></div>


- <div>Identified an issue where ARC Devices could lose connection when running under AirSpace WinComp</div>


- <div><span></span></div>We fixed an issue which prevented hyperlink from being created when pasting text with hyperlink


- <div>We fixed an issue which would prevent statistics from working correctly</div>


- <div><span></span></div>We fixed an issue which could cause TextRanges to become broken after pasting text into the header/footer/slide number placeholders on slide master and slide layout


### Project

- <div><span>We fixed an issue which could cause a crash when replacing an enterprise resource with a local resource</span></div>


- <div>Fixed an issue w<span><span style="display:inline !important;background-color:rgba(255, 255, 255, 1);">hen projects are saved to Project Web App via a master project that the subprojects local
resources, rates and other properties may not be saved.</span><br></span></div>


- <div><span>We fixed an issue where <span style="display:inline !important;background-color:rgb(255, 255, 255);font-size:13.33px;">changes to a work value on a summary task could cause a crash if protected work is enabled</span></span></div>


- <div><span>Fixed an issue where if you have a master project and copy/paste local custom field values from the master project to an inserted project, some information on the task is not correctly saved.</span></div>


- <font size=2 style="background-color:rgb(255, 255, 255);">We fixed an issue which could inhibit the ability to sync events with enterprise calendars</font>


### Word

- <div><span>We fixed an issue which could break the ctrl+v keyboard shortcut</span></div>


- <div><span>We fixed an issue which could prevent synchronous scrolling from working properly in draft view</span></div>


- <div><span>We fixed an issue which could have caused scaling issues when printing to Deskjet printers</span></div>


- <div>We fixed an issue which could prevent Tool Tips from displaying properly after saving a document for the first time</div>


- <div><span>We fixed an issue where table formatting could be lost</span></div>


- <div>Resolved an issue in inserting files as object from OneDrive</div>


- <div><span>We fixed an issue where font colors were not being committed</span></div>


- <div>Improved our recovery steps to f<span>ix an issue that caused graphical content to get deleted from email threads.&nbsp;</span></div>


### Office Suite

- <div><span>We fixed an issue where a user could be given an incorrect error message when closing a file with a pending upload</span></div>


- <div><span>We fixed an issue which could cause a repeated warning to discard local versions of a file</span></div>


- <div>We fixed an issue which could offer version history when that feature was disabled</div>


- <div>Login name is now being pre-populated <span style="background-color:rgb(255, 255, 255);display:inline !important;">for basic auth prompts in Outlook</span></div>


- <div><span>We fixed an issue where medium bold text could be incorrectly styled</span></div>



[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## May 24, 2019
Version 1906 (build 11715.20002)

## What's New:

#### User Experience Updates

Updates that have been in Coming Soon are now here, featuring the Simplified Ribbon, and a visual refresh of the folder pane, message list, and reading pane.

## Notable Fixes:

### All

- We fixed an issue where the Chat Pane would not display

### Word 
- We fixed an issue where in some cases Word could incorrectly highlight text for grammatical errors

### Excel
- We fixed an issue where an incorrect icon was used in for Chart Elements
- We fixed an issue where the incorrect workbook could be activated in a VBA script when the same book was already open

### PowerPoint
- Various performance and stability fixes

### Outlook
- Various performance and stability fixes

### Access
- Various performance and stability fixes

### Project
- We fixed an issue where Project could crash after switching to the taskbar

</BR></BR>
