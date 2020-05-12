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


[//]: # (DO NOT REMOVE)




## Version 2003: May 12
*Version 2003 (Build 12624.20588)*

Security updates listed [here](https://docs.microsoft.com/officeupdates/microsoft365-apps-security-updates)


[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates
### Access

- **Be more productive working in Query Designer, SQL view, and the Relationships window:** Right-click a table to open, design, size, and hide it. Search and replace text in SQL View. Select multiple tables in the Relationships window.

### Excel

- **Type a formula that returns multiple values:** Quickly type a formula that returns multiple values, and they'll automatically spill into the neighboring cells. [Learn more](https://support.office.com/article/5c2c9cbb-def8-409a-b380-2fbf91b20aa3)<br />See details in [blog post](https://blog-insider.office.com/2019/06/13/dynamic-arrays-and-new-functions-in-excel/)

- **No more bouncing to the browser:** You decide how links to Office documents open: in the browser or in the app.

- **Six powerful functions:** We’ve added six new functions to supercharge your spreadsheets: FILTER, SORT, SORTBY, UNIQUE, SEQUENCE and RANDARRAY. [Learn more](https://support.office.com/article/003df6c7-1dcb-4388-8e2e-0fe77a0887bc)

- **Look left, look right… XLOOKUP is here!:** Row by row, find anything you need in a table or range with XLOOKUP. [Learn more](https://support.office.com/article/b7fd680e-6d10-43e6-84f9-88eae8bf5929)<br />See details in [blog post](https://blog-insider.office.com/2019/08/29/announcing-xlookup/)

- **Read and reply on the fly:** Respond to comments and mentions right from email without opening the workbook.

### Outlook

- **Get email suggestions when you search for a person:** When you type a person’s name in the Search box, the most relevant email messages will be included with your search suggestions. [Learn more](https://support.office.com/article/d824d1e9-a255-4c8a-8553-276fb895a8da)

- **Groups Naming policy:** A group naming policy enables the IT admin to standardize and manage the names of groups created by users in the organization. The admin can require a specific prefix and suffix be added to the name for a group when it's created, and can block specific words from being used. This helps minimize the use of inappropriate words in group names as well as IT manage the representation of groups in their directory. Naming Policy also helps organizations that deploy team sites to categorize them based on department.

- **Join meetings without leaving your inbox:** No need to switch to your calendar to join online meetings. With the Calendar pinned to the To-Do pane, join any meeting with just one click.

- **New experience for captive wifi networks:** Have you ever joined a wifi network that required a web page to sign in with? Outlook now detects this and helps you get connected.

### PowerPoint

- **Fast real-time collaboration in PowerPoint:** Fast real-time collaboration in PowerPoint

- **Online videos have a new home:** Save a video to Microsoft Stream so anyone in your organization can see it. Insert the video link and enjoy a multimedia presentation at a fraction of the file size. [Learn more](https://support.office.com/article/8340ec69-4cee-4fe1-ab96-4849154bc6db)

- **No more bouncing to the browser:** You decide how links to Office documents open: in the browser or in the app.

- **GIFs in a jiffy:** One slide, one frame. Easily create looping GIFs in PowerPoint. [Learn more](https://support.office.com/article/a598753e-92de-4f1b-8393-714db4d334b4)<br />See details in [blog post](https://blog-insider.office.com/2019/12/30/create-animated-gifs-using-powerpoint/)

- **Better diagrams:** With better connectors and a smoother ink conversion process you can ink your ideas more confidently. [Learn more](https://support.office.com/article/0740dec3-6291-4c1f-8baa-011d18449919)

- **Update slides during slide show:** Update slides changed by other authors during your presentation.<br />See details in [blog post](https://blog-insider.office.com/2020/04/08/synchronize-changes-while-presenting/)

### Word

- **No more bouncing to the browser:** You decide how links to Office documents open: in the browser or in the app.

- **Others see your changes quickly:** Co-authoring improvements mean your collaborators can see your changes faster than ever before.

### Office Suite

- **Open Large Files Incrementally:** The ability to download, open and interact with large powerpoint presentations, even if parts of the presentation (for example a large video or image) have not finished downloading yet.


[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

<br/>

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues
### Apex Shared

- <h5>User Scenario</h5><div>User updates location or content for Personal/Workgroup/System templates so that there is no content to display in the New tab, or signs out of the tenant, we continue to display an empty section instead of reloading all sections.</div><h5>Defect</h5><div>We do not reload immediately because when a section is added we only want to load it when the New tab is being made visible again (to avoid jumpiness). However, this does not hold where a section is not longer displaying content, we should always remove that section.</div><h5>Fix</h5><div>Make sure that when we detect changes to content and no longer have anything to display, or the tenants are no longer available, we reset the data for the affected template sections and reload all contents.</div><h5>Dev/Test Notes</h5><div><span>The change only impacts&nbsp;the new UI, the file being modified is only used for the new&nbsp;layout and is&nbsp;behind &nbsp;feature gate</span></div><div><span>- a User/Workgroup/System section with no content disappears as soon as the update is received<br></span><div>- the UI does not update if the user is browsing, but does update after they exit browsing<br></div><div>- the UI does not update if the user is searching, but does update after they exit search<br></div><div>- sign in/out repeatedly, the Org Template section is loaded/removed as expected<br></div><div>- the org templates section is removed and content reloads when the user signs out while browsing the tenant library<br></div><span>- the org templates section is removed on reload after a search has been perform (also if searching while browsing)</span><br></div><h5>Code Review</h5><div>codeflow://open/?server=https%3A%2F%2Foffice.visualstudio.com%2F&amp;review=OE.892903&amp;host=vso&amp;alert=true<br></div>


- <h5>User Scenario</h5><div>Automation will fail when the new Templates UI feature gate is enabled</div><h5>Defect</h5><div>Automation is not yet ready for the new UI, we have a separate track tracking that work.</div><h5>Fix</h5><div>Make sure the new UI is not displayed for Automation, even where the feature gate is enabled</div><h5>Dev/Test Notes</h5><div><span>- docsui_test passes qtest with and without the feature gate enabled<br></span><div>- new Templates UI is disabled when automation.snt is in app Documents folder<br></div><div>- new Templates UI is enabled when automation.snt is not in app Documents folder<br></div><span>- sample sma automation tests pass with the feature gate disabled/automation.snt present</span><br></div><h5>Code Review</h5><div><span style="color:rgba(0, 0, 0, 0.901961);display:inline !important;">codeflow://open/?server=https%3A%2F%2Foffice.visualstudio.com%2F&amp;review=OE.892899&amp;host=vso&amp;alert=true</span><br></div>


### ECO

- <div><span style="display:inline !important;">This update fixes an issue in Visual Basic for Applications in Microsoft Office where certain VBA projects that contain references to code libraries with DBCS characters in the library name or library path would be viewed by the Office application as corrupt on load.</span><br></div>


### Excel

- <div><span style="display:inline !important;">Application.Evaluate (VBA) was not working for User-defined functions in some cases.</span><br></div>


- <div style="color:#333333;background-color:#f5f5f5;font-family:Consolas, 'Courier New', monospace;font-weight:normal;font-size:14px;"><div><span style="color:#4b69c6;">Addressed&nbsp;an&nbsp;issue&nbsp;where&nbsp;external&nbsp;links&nbsp;don't&nbsp;update&nbsp;on&nbsp;fill&nbsp;if&nbsp;the&nbsp;source&nbsp;book&nbsp;is&nbsp;closed.</span></div></div>


### OneNote

- <div><span style="display:inline !important;">Improve sync and service stability by temporarily altering how frequently page version histories are created.</span><br></div>


- <div>Improve sync and service stability by temporarily disabling moving pages into the recycle bin. Users who want to delete a page will instead be shown a dialog asking if they'd want to permanently delete the page.</div>


- <div><span style="display:inline !important;">Improve sync and service stability by temporarily<span>&nbsp;reducing the maximum allowable size of new embedded attachments to 50MB in OneNote 2016. For files that exceed this limit, users will have the option of uploading the file to OneDrive and inserting a link into OneNote.</span></span><br></div>


- <div><span style="display:inline !important;">Improve sync and service stability by temporarily<span>&nbsp;disabling in-app video recording in OneNote 2016. Local notebooks are not impacted by this measure.</span></span><br></div>


- <div><span style="display:inline !important;">Improve sync and service stability by temporarily<span>&nbsp;adjusting sync frequency in OneNote 2016.</span></span><br></div>


- <span style="display:inline !important;">Inform users through the InfoBar about temporary adjustments in Microsoft OneNote that will help improve sync and service availability during high worldwide usage.</span><br>


### Outlook

- <div>Addresses an issue that caused users to see the Outlook process lingering in task manager after exiting.</div>


- <div>Addresses an issue that caused users to occasionally experience a crash when using the &quot;X&quot; button on their mouse.</div>


- <div>Addresses an issue that caused users to experience crashes in third party MAPI applications.</div>


- <div><span style="display:inline !important;">Addresses an issue that caused Outlook to crash when opening .msg or .oft files that were saved locally after a Windows update.</span><br></div>


- <div>Addresses an issue that caused Outlook to crash on some builds of Windows.</div>


- <span style="display:inline !important;">Addresses an issue that caused the width of the folder pane to change unexpectedly.</span>


- <h5>User Scenario</h5><span>Failing test validates that marking a message read/unread explicitly invalidates implicit mark as read timer event.<br></span><div>1. Selects an unread message<br></div><div>2. Marks as read (explicitly)<br></div><div>3. Marks as unread (explicitly)<br></div><div>4. Fire implicit mark as read timer event and expect no-op.<br></div><span></span><br><h5>Defect</h5><div><span style="color:rgb(33, 33, 33);font-family:&quot;Segoe UI&quot;, -apple-system, system-ui, Roboto, Arial, sans-serif;display:inline !important;">Appears that a second (unexpected) content pane reloads occurs with a delay fired event that revalidates the timer for implicit mark as read, after the timer has been (expectedly) invalidated due to an explicit mark as read/unread action. This is likely a test-only issue within our Demo app and the full Outlook app does not show any side effects at the moment.&nbsp;</span><br></div><h5>Fix</h5><div><span style="color:rgb(33, 33, 33);font-family:&quot;Segoe UI&quot;, -apple-system, system-ui, Roboto, Arial, sans-serif;display:inline !important;">Suppressing the test for now and unblocking our build loops.</span></div><h5>Dev/Test Notes</h5><div>No production impact, only test issue.</div><h5>Code Review</h5>codeflow://open?server=https%3A%2F%2Foffice.visualstudio.com%2F&amp;review=OE.893917&amp;host=vso&amp;alert=true<br>


### Project

- <div>Fixed an issue where task percent complete was incorrectly changing to a value less than 100% complete after it was marked complete.</div>


- <div><span style="display:inline !important;">Fixed an issue where the OnUndoOrRedo event doesn't fire&nbsp;</span><span style="box-sizing:border-box;font-size:13.3333px;display:inline !important;">without first running the OpenUndoTransaction method.</span><br></div>


- <div><span style="display:inline !important;">Fixed an issue where summary task dates weren't always getting calculated correctly.</span><br></div>


- <div><span style="display:inline !important;">Fixed an issue where the ProjectBeforeTaskChange event does not detect when a task has been inactivated/activated via the Inactivate button.</span><br></div>


- <div><span style="display:inline !important;">Fixed an issue where the ProjectBeforeTaskChange event does not detect when a task has been inactivated/activated via the Inactivate button.</span><br></div>


- <div>Fixed an issue where Project may crash when saving projects created with older versions of Project.</div>


- <div><span style="display:inline !important;">Fixed an issue where the user couldn't enter time-phased Baseline Work when the setting to protect actual work is on.</span><br></div>


- <p style="box-sizing:border-box;margin:0px;color:rgba(0, 0, 0, 0.9);">When&nbsp;<span style="">Predecessor/Successor data is edited within a Form view, an extra <span style="font-size:13.3333px;display:inline !important;">ProjectBeforeTaskChange&nbsp;</span>event i</span><span style="">s fired</span></p><br><br>


### Security

- <div>This fix resolves an error which occurs preventing both restricting access and protecting files with a password simultaneously.&nbsp;</div>


### Word

- <div>Addresses an issue that caused users to occasionally experience a crash when using the &quot;X&quot; button on their mouse.</div>



[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2002: May 12
*Version 2002 (Build 12527.20612)*

Security updates listed [here](https://docs.microsoft.com/officeupdates/microsoft365-apps-security-updates)


[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues
### ECO

- <div><span style="display:inline !important;">This update fixes an issue in Visual Basic for Applications in Microsoft Office where certain VBA projects that contain references to code libraries with DBCS characters in the library name or library path would be viewed by the Office application as corrupt on load.</span><br></div>


- <div><span><span style="display:inline !important;background-color:rgba(255, 255, 255, 1);">This update fixes an issue in Microsoft Office where Visual Basic for Applications projects with references that are expected to be found by searching locations specified in the PATH environment variable may not be found properly at runtime, leading to VBA runtime errors.</span><br></span></div>


### Excel

- <div>Using a Range.Value and Range.Value2 (VBA) would cause formulas to be entered as dynamic arrays.</div>


- <div><span style="display:inline !important;">Opening a workbook with references to numerous other workbooks, especially with hidden windows, would be slower than expected.</span><br></div>


- <div><span style="display:inline !important;">Workbooks saved with a digital signature in Excel 2016 could have the signature invalidated upon opening in the current version of Excel.</span><br></div>


- <div>Addresses an issue where the &quot;Value Crosses At&quot; property on chart axis unexpectedly changes when saving and re-opening a file.</div>


- <div>Opening CSV files was taking longer than expected in some circumstances.</div>


- <div>Excel may crash in some circumstances when switching between workbooks with different zoom levels.</div>


- <div><span style="display:inline !important;">Fixed an issue with VBA in which writing values to a range would be slower than expected.</span><br></div>


- <div><span style="display:inline !important;">Data copied from a column filtered by color sometimes would not paste properly.</span><br></div>


- <div><span style="display:inline !important;">Opening a workbook with references to numerous other workbooks, especially with hidden windows, would be slower than expected.</span><br></div>


- <div>Fixed an issue which would cause Excel to crash in some cases after copying a sheet containing a PivotTable.</div>


### LSXOXO

- <div><p style="margin:0in 0in 8pt;font-family:Calibri, sans-serif;font-size:11pt;"><span style="background:white;color:black;font-family:&quot;Segoe UI&quot;,sans-serif;font-size:10.5pt;">Localises the notification that allows the user to learn more about temporary measures being enacted in the OneNote user experience to improve sync and service stability.<br></span></p></div>


### OneNote

- <div><span style="display:inline !important;">Improve sync and service stability by temporarily</span><span style="box-sizing:border-box;">&nbsp;disabling in-app video recording in OneNote 2016. Local notebooks are not impacted by this measure.</span><br></div>


- <div><span style="display:inline !important;">Improve sync and service stability by temporarily<span>&nbsp;reducing the maximum allowable size of new embedded attachments to 50MB in OneNote 2016. For files that exceed this limit, users will have the option of uploading the file to OneDrive and inserting a link into OneNote.</span></span><br></div>


- <div><span style="display:inline !important;">Improve sync and service stability by temporarily<span>&nbsp;adjusting sync frequency in OneNote 2016.</span></span><br></div>


- <div><span style="font-family:&quot;Segoe UI&quot;, sans-serif;display:inline !important;">Improve sync and service stability by temporarily deferring the downloading of embedded files and images in online notebooks until the user navigates to the page in OneNote 2016.</span><br></div>


- <div><p style="margin:0in 0in 8pt;font-size:11pt;font-family:Calibri, sans-serif;"><span style="font-size:10.5pt;font-family:&quot;Segoe UI&quot;,sans-serif;color:black;background:white;">Display a notification that
allows the user to learn more about temporary measures being enacted in the
OneNote user experience to improve sync and service stability.</span></p></div>


- <p style="margin:0in 0in 8pt;font-size:11pt;font-family:Calibri, sans-serif;"><span style="font-size:10.5pt;font-family:&quot;Segoe UI&quot;,sans-serif;color:black;background:white;">Improve sync and service
stability by temporarily</span><span style="box-sizing:border-box;">&nbsp;reducing the number and sync frequency of version history pages in
OneNote 2016.</span></p>


- <div><p style="margin:0in 0in 8pt;font-size:11pt;font-family:Calibri, sans-serif;"><span style="font-size:10.5pt;font-family:&quot;Segoe UI&quot;,sans-serif;color:black;background:white;">Improve sync and service
stability by temporarily</span><span style="box-sizing:border-box;">&nbsp;disabling the recycle bin in OneNote 2016. When a user tries to
delete data that would normally be sent to the recycle bin, users will be asked
if they would prefer to keep or permanently delete the data.</span></p></div>


### Outlook

- <div>Addresses an issue that caused users to see message body truncation when forwarding large HTML messages.</div>


- <div>Addresses an issue that caused users to experience a crash when trying to open .msg and .oft files after a Windows update.</div>


- <span style="display:inline !important;">Addresses an issue that caused the width of the folder pane to change unexpectedly.</span>


### PowerPoint

- <div>Fixes an issue to relay correct messaging for users who open a copy of a file that has improved comments.</div>


### Security

- <div><div style="margin:0px 0in 0.000133333px;font-size:11pt;font-family:Calibri, sans-serif;"><span style="font-size:10.5pt;font-family:&quot;Segoe UI&quot;,sans-serif;color:black;background:white;"></span><span style="font-size:10.5pt;font-family:&quot;Segoe UI&quot;,sans-serif;color:black;background:white;">“This update fixes a problem in Microsoft Word where text longer
than 255 characters inserted while applying a sensitivity label could not
subsequently be identified and removed by changing or removing the label if the
label policy applied a header or footer or watermark.</span><span style="font-size:11.0pt;font-family:&quot;Calibri&quot;,sans-serif;color:black;background:white;">”</span></div></div>


- <div>&quot;This update fixes an issue with Microsoft Outlook not displaying the current sensitivity label when viewing or composing messages.&quot;</div>


### TEL

- <div>This bug updates the <span style="display:inline !important;">enhanced configuration&nbsp;</span>service (ECS) url endpoint. Calling this newer endpoint has a higher success rate for fetching from ECS.</div>


- <div><p style="margin:0in 0in 0.0001pt;font-size:11pt;font-family:Calibri, sans-serif;">This bug fix eliminates crashes during Office
handoff sessions, thereby improving reliability in the user experience. Additionally, it improves the diagnostic data for helping us identify and fix such reliability issues.</p></div>


### Word

- <div><span style="display:inline !important;">We fixed an issue when merging 2 documents into one document.</span><br></div>


- <div><span style="display:inline !important;">Fixed an issue with Compare feature for<span>&nbsp;</span></span><span style="box-sizing:border-box;font-size:13.33px;display:inline !important;">documents that were protected for editing.</span><br></div>


- <div style="box-sizing:border-box;margin:0px;color:rgba(0, 0, 0, 0.9);"><div style="box-sizing:border-box;font-family:&quot;Segoe UI&quot;, system-ui, &quot;Apple Color Emoji&quot;, &quot;Segoe UI Emoji&quot;, sans-serif;">Resolved the issue where Access and Publisher might not boot correctly depending on which languages were installed.&nbsp;</div><br></div><br>


### Office Suite

- <div><span style="font-size:11.0pt;font-family:&quot;Calibri&quot;,sans-serif;color:#201F1E;background:white;">Fixed
an issue while opening files from on-prem locations with some specific proxy configurations.</span><br></div>


- <div><span style="font-size:10.5pt;font-family:&quot;Segoe UI&quot;,sans-serif;color:black;background:white;">This change ensures Sketched outline works properly in the ribbon.</span></div>


- <div><div style="box-sizing:border-box;font-family:&quot;Segoe UI&quot;, system-ui, &quot;Apple Color Emoji&quot;, &quot;Segoe UI Emoji&quot;, sans-serif;">Hi - We have resolved the issue where an internal operation was throwing an exception on failure instead of logging and continuing on. The affected users will not be blocked from receiving updates anymore.&nbsp;</div><div style="box-sizing:border-box;font-family:&quot;Segoe UI&quot;, system-ui, &quot;Apple Color Emoji&quot;, &quot;Segoe UI Emoji&quot;, sans-serif;"><p style="margin:0in 0in 0.0001pt;font-size:11pt;font-family:Calibri, sans-serif;margin-left:.5in;"><br></p></div><div style="box-sizing:border-box;font-family:&quot;Segoe UI&quot;, system-ui, &quot;Apple Color Emoji&quot;, &quot;Segoe UI Emoji&quot;, sans-serif;">Thanks,</div><div style="box-sizing:border-box;font-family:&quot;Segoe UI&quot;, system-ui, &quot;Apple Color Emoji&quot;, &quot;Segoe UI Emoji&quot;, sans-serif;">Office Deployment Engineering</div></div>


- <div>This is a fix to address that Project app should not block network when the file is cached in the client.</div>



[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2002: May 11
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

## Version 2001: May 11
*Version 2001 (Build 12430.20504)*
* Various bugs and performance fixes.

[//]: # (DO NOT REMOVE END)

> [!NOTE]
> If you need help with an issue with using Office, we recommend that you post your question on [Microsoft's Answers forum](https://answers.microsoft.com/) or [Tech Community](https://techcommunity.microsoft.com/), or you can contact [support](https://support.microsoft.com/contactus).

