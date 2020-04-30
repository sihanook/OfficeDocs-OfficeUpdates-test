
title: "Release Notes Monthly Channel (Targeted"
ms.author: anankani
author: v-lislo
manager: andrewmo
ms.audience: Win32 Fast
ms.topic: reference
ms.service: o365-proplus-
localization_priority: Critical
ms.collection: RelNotes_ProPlus
description: "Provides Insiders Slow audience with the latest list of key new features, fixes or known issues"

# Release Notes for Office Monthly Channel (Targeted)

This article contains release notes for Insider builds of Word, Excel, PowerPoint, Outlook, Access, and Project for Windows desktop. Every week, we’ll highlight interesting new features, important fixes, and any significant issues we want you to know about. Note that we often roll out features (and sometimes even fixes) to Insiders over a period of time. This allows us to ensure that things are working smoothly before releasing the feature to a wider audience. So, if you don’t see something described below, don't worry you'll get it eventually.  

> [!NOTE]
> - Release notes are posted weekly and may be a compilation of multiple builds
> - The release notes publication date may not match the actual build release date
> - Microsoft Teams on existing installations of Office 365 ProPlus - Beginning in late June, Microsoft Teams will be included in existing installations of Office 365 ProPlus (and Office 365 Business) upon updates of these installations. The date when Teams will be added depends on which update channel you're using. Please refer to [Deploy Microsoft Teams with Office 365 ProPlus](https://docs.microsoft.com/en-us/deployoffice/teams-install) for additional information.

[//]: # (DO NOT REMOVE)


## Version 2002: April 30
*Version 2002 (Build 12527.20470)*


[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues
### Excel

- <div>Addresses a performance issue when creating charts from templates</div>


- <div>Addressed&nbsp;an&nbsp;issue&nbsp;where&nbsp;external&nbsp;links&nbsp;don't&nbsp;update&nbsp;on&nbsp;fill (fill down, fill across, etc.) if&nbsp;the&nbsp;source&nbsp;book&nbsp;is&nbsp;closed.<br></div>


- <div>Using VBA's Application.Evaluate was not working for User-defined functions in some cases.</div>


- <div>When saving as a CSV file, Excel could merge all columns into a single column in some cases.</div>


- <div>Using Range.ClearContents on a range in a protected sheet could take longer than expected.</div>


- <div>Fixed a problem with scaling of text in form controls when displayed in Print Preview.</div>


- <div>VBA macros that interact with the ribbon may run with ScreenUpdating set to True unexpectedly.</div>


- <div>Excel would crash in certain cases when re-opening a workbook embedded in Word or PowerPoint.</div>


- <div style="box-sizing:border-box;">Fixed an issue where CUBEVALUE functions would sometimes return an incorrect result.&nbsp;</div><div><span style="display:inline !important;"></span><br></div>


### Outlook

- <div>Addresses an issue that caused the Save to Cloud button to be missing from Attachment Tools.</div>


- <div><span style="display:inline !important;">This change restores the ability to view multi-line subjects in the message header.</span><br></div>


- <div><div style="box-sizing:border-box;font-family:&quot;Segoe UI&quot;, system-ui, &quot;Apple Color Emoji&quot;, &quot;Segoe UI Emoji&quot;, sans-serif;">Addresses an issue that caused users to experience a crash when selecting certain search results.</div></div>


- <div>Addresses an issue that caused users to occasionally experience a crash when using the X button on the mouse.&nbsp;</div>


- <div>Addresses an issue that caused third party applications to be unable to send email.</div>


- <div>Addresses an issue that caused Outlook to unexpectedly sync all mail even when the sync slider is set to a smaller setting.&nbsp;</div>


- <div>Addresses an issue that caused users with Black Theme to see the &quot;From&quot; dropdown show white text on a white background.</div>


- <div>Addresses an issue that caused the option to disable flagged item highlighting to fail to be respected in some scenarios.</div>


- <div>Addresses an issue that could result in a crash when viewing the same item in multiple windows.</div>


- <div>Addresses an issue that caused the Group header to expand unexpectedly in some scenarios.</div>


- <div>Addresses an issue that caused commas in the location field of a meeting to turn into semicolons.</div>


### PowerPoint

- <div>Improved a copy-paste scenario:&nbsp;<span style="font-size:13.3333px;display:inline !important;">Copying the Shape in powerpoint slide and paste it in other slide in a loop might fail with exception.&nbsp;</span></div>


### Project

- <div><span style="display:inline !important;">Fixed an issue where summary task dates weren't always getting calculated correctly.</span><br></div>


- <div><span style="display:inline !important;">Fixed an issue where the OnUndoOrRedo event doesn't fire&nbsp;</span><span style="box-sizing:border-box;font-size:13.3333px;display:inline !important;">without first running the OpenUndoTransaction method.</span><br></div>


- <div>Fixed an issue where a task that is marked 100% complete is wrongly changing to be less than 100% complete.</div>


- <div>Fixed an issue where the ProjectBeforeTaskChange event does not detect when a task has been inactivated/activated via the Inactivate button.</div>


- <div><span style="display:inline !important;">Fixed an issue where the ProjectBeforeTaskChange event does not detect when a task has been inactivated/activated via the Inactivate button.</span><br></div>


- <div><span style="display:inline !important;">Fixed an issue where Project may crash when saving projects created with older versions of Project.</span><br></div>


- <div><span style="display:inline !important;">Fixed an issue where Project may crash when saving projects created with older versions of Project.</span><br></div>


### Security

- <div>Fixes an issue when multiple documents are open in Word/Excel/PowerPoint from the same SharePoint library, only the first document opened will be scanned for Policy compliance.</div>


### Word

- <div>We fixed an issue where inserting horizontal lines are not shorter and centered</div>


- <div>We fixed an issue with fit text in a table</div>


- <div>Addresses an issue that caused users to occasionally experience a crash when using the X button on the mouse.&nbsp;</div>



[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

