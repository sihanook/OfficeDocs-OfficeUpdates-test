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

## Version 1911: November 14
*Version 1911 (Build 12228.20120)*


[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues
### Access

- <div><span>The record count could be incorrect</span></div>


### Excel

- <div>Resolved an issue where deleting sheets containing sparklines referencing data on another sheet could cause the file to be identified as corrupted when re-opened.</div>


- <div>C<span>hanges to a chart size could not be saved</span></div>


- <div><span>Checkboxes could not render correctly</span></div>


- <div><span>Select Data Source dialogs were not case sensitive for some fields</span></div>


- <div><span>Some VBA functions would return an error on new chart types</span></div>


- <div>U<span>sers could be prevented from saving in Office 365 Excel Workbook format</span></div>


- <div><span>Check box controls could shrink when using autofit to adjust row height</span></div>


- <div>Incorrect results&nbsp;<span style="color:rgba(0, 0, 0, 0.9);display:inline !important;">when converting report filters along with the rest of the PivotTable for queries to SQL tabular servers.</span></div>


- <div><span>Excel may have issues when editing a protected file from an untrusted network share</span></div>


- <div><span>Using Narrator and Magnifier at the same time may result in a crash</span></div>


- <div><span>Resolved an issue where selecting a cell after scrolling could result in the wrong cell being selected</span></div>


- <div>S<span>ignificant improvements to the performance of deleting columns with merged cells</span></div>


### OneNote

- <div><span>Identified an issue which could affect syncing from a local resource to a cloud resource</span></div>


### Outlook

- <div><span>Room Finder tool may be displaying &quot;None&quot; for available rooms</span></div>


- <div>Identified an issue where the search box could disappear when the ribbon is set to hide automatically</div>


- <div><span>Identified an issue which could cause digital signatures to become broken when signing an e-mail with a digitally signed attachment</span></div>


- <div>A forwarded e-mail may be missing embedded images</div>


- <span style="display:inline !important;background-color:rgba(255, 255, 255, 1);font-size:13.33px;">Users may not be able to create Outlook profiles with strict tenant restriction</span>


- <div><span>Identified an issue where long filenames were truncated after draging and droping to the message body</span></div>


### PowerPoint

- <div>C<span>hanges to a chart size could not be saved</span></div>


- <div><span>Using Narrator and Magnifier at the same time may result in a crash</span></div>


- <div><span></span></div><span style="font-size:11.0pt;font-family:&quot;Calibri&quot;,sans-serif;">Identified an issue
where aspect ratio for the slide preview was not being properly locked/unlocked</span>


### Project

- <div>Identified an issue where notes might not persist if entered while doing update tasks<br></div>


- <div style="box-sizing:border-box;"><span style="box-sizing:border-box;">An assignment with zero work on a task, the task may be unable to be marked complete and will always show up at 99%.</span></div>


- <div style="box-sizing:border-box;"><span style="box-sizing:border-box;">Level All command doesn't properly resolve a resource overallocation.</span></div><div></div>


- <div><span>Identified an issue where users could get several messages when opening a read-only project</span></div>


- <div>Identified an issue where a file could be locked by a user, but no username would be displayed in the error message</div>


### Publisher

- <div><span>Shapes could appear outside of the graphics border</span></div>


### Word

- <div><span>Proofing suggestions are not displaying in contextual menus.</span></div>


- <div>C<span>hanges to a chart size could not be saved</span></div>


- <div><span>Shapes could appear outside of the graphics border</span></div>


- <div><span>Identified an issue when viewing comments while using a screen reader</span></div>


- <div><span>Identified an issue where some critiques were misidentified as being spelling or grammar critiques</span></div>


- <div style="box-sizing:border-box;font-family:&quot;Segoe UI&quot;, &quot;Helvetica Neue&quot;, &quot;Apple Color Emoji&quot;, &quot;Segoe UI Emoji&quot;, Helvetica, Arial, sans-serif;text-align:start;">The links of cid: images from Outlook messages&nbsp;can now be successfully broken when requested.</div>


- <div><span>Incorrect characters may appear when using Korean/English autocorrect</span></div>


- <div><span>Searching from the Navigation pane may fail</span></div>


- <div><span>Using Narrator and Magnifier at the same time may result in a crash</span></div>


- <div><span>Content policies are being incorrectly applied to comments</span></div>


- <div>Lower policy labels may be applied when a higher policy label should have taken priority</div>


- <div><font size=2><font style="background-color:rgba(255, 255, 255, 1);">Opening legacy<span style="display:inline !important;"> documents and then going to the Info tab can cause a crash</span></font></font><br></div>


- <div><span>Identified an issue where a new comment dialog could sometimes not obtain focus</span></div>


- <div>A<span> contact card could be prevented from opening after apply formatting to an @ mention</span></div>


- <div><span>Highlighting text could be challenging</span></div>


- <div><span><p style="margin:0in 0in 0pt;font-family:&quot;Calibri&quot;,sans-serif;font-size:11pt;"><span style="font-size:10.5pt;">Resolved
an issue where users may be unable to save Word, Excel, and PowerPoint
documents.&nbsp; This issue affects users that create a new file and bring up
the &quot;Save as Model Dialog&quot; option after clicking on the Save icon or
pressing Ctrl + S.</span><br></p></span></div>


- <div><span>Legacy comments written with dark text is not visible in Dark Mode</span></div>


- <div>A<span> user could be prevented from navigating to an individual item in the editor</span></div>


- <div>G<span>rammar or spelling errors might not be highlighted</span></div>


### Office Suite

- <div>Some d<span>rawings may not display in preview or slide shows</span></div>


- <div><span>We fixed an issue where an upgrade could be prevented by a incorrect error message of &quot;Another install in progress&quot;</span></div>


- <div><span>Some katakana characters may display incorrectly in a vertical text box</span></div>


- <div>Attempting to save a file to a disconnected network share may result in issues</div>


- <div><span>Performance issue when using Shapes on Windows 7</span></div>



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
