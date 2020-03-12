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

## Version 2003: March 12
*Version 2003 (Build 12624.20176)*


[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates
### Security

- **Sensitivity labels:** You can now apply a sensitivity label that your organization has configured to prompt you for custom permissions.

### Office Suite

- **Tabbed Panes:** Now you can switch between multiple panes using a tab UI on the right hand side of the app. The UI will only be visible when you have 2+ panes open.


[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

<br/>

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues
### Excel

- <div>Fixed a cosmetic issue where the 'OK' button on the File \ Options dialog displayed as being grayed out but functionality was not impacted.</div>


- <div><span>Fixed an issue that users may have experienced when renaming pivot table measures.&nbsp;</span></div>


- <div>Fixed a performance issue that users may have experienced when using a VBA macro to clear the contents of a range.&nbsp;</div>


- <p style="margin:0in 0in 0pt;font-family:&quot;Calibri&quot;,sans-serif;font-size:11pt;">Fixed an issue that caused the UI to flash when users executed a macro that interacted with the ribbon.&nbsp;</p><div></div>


- <div>Fixed an issue where CSV files were loaded incorrectly when the first word in the file was TABLE.</div>


- <div>Fixed an issue where users may have experienced crashes when switching between two workbooks that had different zoom levels.&nbsp;</div>


- <div>Fixed an issue where text in a slicer isn't scaled properly in Print Preview.</div>


- <div>Fixed an issue where CUBEVALUE functions would sometimes return an incorrect result.&nbsp;</div>


- <div><span><span style="display:inline !important;color:black;font-family:Calibri,Helvetica,sans-serif;font-size:10pt;font-size-adjust:none;">This change addresses a run-time error in the object model and potential crash of the App (Excel, Word) when Add-ins ask for Host Items on documents/worksheets that contain shapes with noSelect locks.&nbsp;</span><br></span></div>


### excel.exe

- <div>Addresses an issue that caused Outlook users to experience a crash when synchronizing settings.</div>


### Outlook

- <div>Fixed an issue where creating a rule with Outlook Web Access did not persist to the Exchange server and resulted in a conflict.</div>


- <div>Addresses an issue that caused Outlook users to experience a crash when synchronizing settings.</div>


- <div>Fixed an issue with Outlook in dark mode would not display the drop down list in the 'From:' field.</div>


- <div>Addresses an issue that caused Outlook to unexpectedly generate logging output in some scenarios, even when logging was turned off.</div>


- <div>Addresses an issue that caused users to be unable to open public folder messages when Outlook was left running overnight.</div>


- <div>Fixed a race condition where the 'Allow' and 'Deny' buttons on the permissions page are disabled during&nbsp; the authentication workflow of adding a Gmail account.</div>


- <div><span style="">Fixed an issue that may result in the alert: &quot;Sorry we're having trouble opening this item&quot; when opening some recurring meeting instances sent from a different time zone.</span><br></div>


- <div>Addressed an issue that could cause users to be unable to reopen a .msg file after dragging and dropping an attachment from that message.</div>


- <div><span style="">Fixed an issue where after uploading a file attachment from Outlook to OneDrive could result in the file name being changed if the attachment's name contained parenthesis.</span><br></div>


- <div>Addressed an issue that caused users to lose access to the &quot;Free Busy Options&quot; calendar permissions dialog.</div>


- <div>Addresses an issue that caused users to be unable to attach a file to their mail message via the file explorer when that file was open in another application.</div>


### outlook.exe

- <div>Addresses an issue that caused Outlook users to experience a crash when synchronizing settings.</div>


### PowerPoint

- <div>Fixed an issue where the recommended thumbnails flash when hovering your mouse over the thumbnails. In some cases this could cause PowerPoint to crash.</div>


- <div>Fixed a cosmetic issue where the 'OK' button on the File \ Options dialog displayed as being grayed out but functionality was not impacted.</div>


- <div>Fixed an issue that could result in a failure to save a document in PowerPoint or Word containing an Excel chart.</div>


- <div>Improved a copy-paste scenario:&nbsp;<span style="font-size:13.3333px;display:inline !important;">Copying the Shape in powerpoint slide and paste it in other slide in a loop might fail with exception.&nbsp;</span></div>


### Project

- <div>Fixed an issue where the ProjectBeforeTaskChange event does not detect when a task has been inactivated/activated via the Inactivate button.</div>


- <div><span style="display:inline !important;">Fixed an issue where the ProjectBeforeTaskChange event does not detect when a task has been inactivated/activated via the Inactivate button.</span><br></div>


- <div>Fixed an issue where task percent complete was incorrectly changing to a value less than 100% complete after it was marked complete.</div>


- <div><span style="display:inline !important;">Fixed an issue where the OnUndoOrRedo event doesn't fire&nbsp;</span><span style="box-sizing:border-box;font-size:13.3333px;display:inline !important;">without first running the OpenUndoTransaction method.</span><br></div>


- <div><span style="display:inline !important;">Fixed an issue where summary task dates weren't always getting calculated correctly.</span><br></div>


- <div><span style="display:inline !important;">Fixed an issue where the ProjectBeforeTaskChange event does not detect when a task has been inactivated/activated via the Inactivate button.</span><br></div>


- <div><span style="display:inline !important;">Fixed an issue where the ProjectBeforeTaskChange event does not detect when a task has been inactivated/activated via the Inactivate button.</span><br></div>


### Security

- <div>Fixes an issue when multiple documents are open in Word/Excel/PowerPoint from the same SharePoint library, only the first document opened will be scanned for Policy compliance.</div>


### Visio

- <div><span><div style="box-sizing:border-box;font-family:&quot;Segoe UI&quot;, system-ui, &quot;Apple Color Emoji&quot;, &quot;Segoe UI Emoji&quot;, sans-serif;text-align:start;">Shape info pane was showing inconsistent details under Shape Data section, with respect to the file when opened in Visio Desktop. It has now been fixed<font face="&quot;Segoe UI VSS (Regular)&quot;,&quot;Segoe UI&quot;,&quot;-apple-system&quot;,BlinkMacSystemFont,Roboto,&quot;Helvetica Neue&quot;,Helvetica,Ubuntu,Arial,sans-serif,&quot;Apple Color Emoji&quot;,&quot;Segoe UI Emoji&quot;,&quot;Segoe UI Symbol&quot;">.</font><br></div></span></div>


- <div>Bitmaps imported in versions before 2016 were not being rendered due to some security checks. We have fixed this issue in Visio Subscription &nbsp;</div>


### Word

- <div>Fixed an issue where comment cards don't always get highlighted when a mouse pointer hovers over the comment card.</div>


- <div>Fixed an issue that when tabbing through a comment card, the focus on the comment edit box would not be visible.</div>


- <div>Fixed a cosmetic issue where the 'OK' button on the File \ Options dialog displayed as being grayed out but functionality was not impacted.</div>


- <div><span style="color:rgb(34, 34, 34);font-family:&quot;Segoe UI&quot;, &quot;Helvetica Neue&quot;, Helvetica, Arial, Verdana;display:inline !important;">During an active document co-authoring session, adding an image directly in to a comment card may result in the addition of a tag. This issue has been fixed.</span></div><div><br></div>


- <div><span><span>Inserting a control (such as a Text Content Control) in an equation then saving and opening the file results in an un-readable content error.</span><br></span></div>


- <div>Fixed an issue where saving a previously password protected file to a cloud storage would not work.</div>


- <div>Fixed an issue with Compare feature for <span style="display:inline !important;background-color:rgba(255, 255, 255, 1);font-size:13.33px;">documents that were protected for editing.</span></div>


- <div><span>Fixed an issue where pictures in documents lose transparency when exported to PDF.<br></span></div>


### Office Suite

- <div><p style="margin:0in 0in 8pt;font-size:11pt;font-family:Calibri, sans-serif;"><span style="font-size:10.0pt;font-family:&quot;Segoe UI&quot;,sans-serif;color:black;background:white;">When using Multichoice/Lookup/Managed-metadata
properties with Word/Excel/PowerPoint documents and saving to a SharePoint
Document Library, these properties were previously limited to 255 characters.
When these properties exceeded 255 characters, such documents could not be
saved. With this change, this limit has been increased to 2048 characters.</span></p></div>


- <div>Fixed an issue Word/Excel/PowerPoint where the User Principal Name (UPN) is no longer case sensitive resulting in less failures when working with files on SharePoint.</div>



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
