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

> [!IMPORTANT]
> We’re making some changes to the update channels for Microsoft 365 Apps, including adding a new update channel (Monthly Enterprise Channel) and changing the names of the existing update channels. To learn more, [read this article](https://go.microsoft.com/fwlink/p/?linkid=2127441).

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

### Excel

- Application.Evaluate (VBA) was not working for User-defined functions in some cases.

- Addressed an issue where external links don't update on fill if the source book is closed.


### OneNote

- Improved sync and service stability by temporarily altering how frequently page version histories are created.


- Improved sync and service stability by temporarily disabling moving pages into the recycle bin. Users who want to delete a page will instead be shown a dialog asking if they'd want to permanently delete the page.


- Improved sync and service stability by temporarily reducing the maximum allowable size of new embedded attachments to 50MB in OneNote 2016. For files that exceed this limit, users will have the option of uploading the file to OneDrive and inserting a link into OneNote.


- Improved sync and service stability by temporarily disabling in-app video recording in OneNote 2016. Local notebooks are not impacted by this measure.


- Improved sync and service stability by temporarily adjusting sync frequency in OneNote 2016.


- Inform users through the InfoBar about temporary adjustments in Microsoft OneNote that will help improve sync and service availability during high worldwide usage.


### Outlook

- Addressed an issue that caused users to see the Outlook process lingering in task manager after exiting.


- Addressed an issue that caused users to occasionally experience a crash when using the button on their mouse.


- Addressed an issue that caused users to experience crashes in third party MAPI applications.


- Addressed an issue that caused Outlook to crash when opening .msg or .oft files that were saved locally after a Windows update.


- Addressed an issue that caused Outlook to crash on some builds of Windows.


- Addressed an issue that caused the width of the folder pane to change unexpectedly.


### Project

- Fixed an issue where task percent complete was incorrectly changing to a value less than 100% complete after it was marked complete.


- Fixed an issue where the OnUndoOrRedo event doesn't fire without first running the OpenUndoTransaction method.


- Fixed an issue where summary task dates weren't always getting calculated correctly.


- Fixed an issue where the ProjectBeforeTaskChange event does not detect when a task has been inactivated/activated via the Inactivate button.


- Fixed an issue where the ProjectBeforeTaskChange event does not detect when a task has been inactivated/activated via the Inactivate button.


- Fixed an issue where Project may crash when saving projects created with older versions of Project.


- Fixed an issue where the user couldn't enter time-phased Baseline Work when the setting to protect actual work is on.


- When Predecessor/Successor data is edited within a Form view, an extra ProjectBeforeTaskChange event is fired.


### Word

- Addresses an issue that caused users to occasionally experience a crash when using the "X" button on their mouse.

### Office Suite

- This fix resolves an error which occurs preventing both restricting access and protecting files with a password simultaneously.

- This update fixes an issue in Visual Basic for Applications in Microsoft Office where certain VBA projects that contain references to code libraries with DBCS characters in the library name or library path would be viewed by the Office application as corrupt on load.


[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)


[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

[//]: # (DO NOT REMOVE END)

> [!NOTE]
> If you need help with an issue with using Office, we recommend that you post your question on [Microsoft's Answers forum](https://answers.microsoft.com/) or [Tech Community](https://techcommunity.microsoft.com/), or you can contact [support](https://support.microsoft.com/contactus).