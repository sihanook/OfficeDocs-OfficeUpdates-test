---
title: "Release notes for Monthly Enterprise Channel releases in 2020"
ms.author: andrewmo
author: anankani
manager: andrewmo
ms.audience: ITPro
ms.topic: reference
ms.service: o365-proplus-itpro
localization_priority: Critical
ms.collection: RelNotes_ProPlus
description: "Provides IT Pros with release notes for Monthly Enterprise Channel releases for Microsoft 365 Apps in 2020"
---

# Release notes for Monthly Enterprise Channel releases in 2020

These release notes provide information about new features and non-security updates that are included in Monthly Enterprise Channel updates in 2020 for Microsoft 365 Apps for enterprise, Microsoft 365 Apps for business, and the subscription versions of the desktop apps for Project and Visio.


[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)


[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

<br/>

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)


[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)




> [!NOTE]
> If you need help with an issue with using Office, we recommend that you post your question on [Microsoft's Answers forum](https://answers.microsoft.com/) or [Tech Community](https://techcommunity.microsoft.com/), or you can contact [support](https://support.microsoft.com/contactus).

[//]: # (DO NOT REMOVE)


## Version 2002: May 08
*Version 2002 (Build 12527.20470)*


[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates
### PowerPoint

- **Better diagrams:** With better connectors and a smoother ink conversion process you can ink your ideas more confidently. [Learn more](https://support.office.com/article/0740dec3-6291-4c1f-8baa-011d18449919)


[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

<br/>

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

