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


## Version 2005: May 04
*Version 2005 (Build 12827.20030)*


[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates
### Unknown ****** Please Review ******

- **:** description

- **:** 

- **:** 


[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

<br/>

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues
### Excel

- <div>Fixed an issue where chart data table could render values in a date axis incorrectly.</div>


- <div>Fixed an issue where&nbsp;page breaks could not be disabled after going into Page Layout or Page Break Preview.</div>


- <div>Inserting a column in a filtered list would take longer than expected.</div>


- <div>Fixed an issue where chart line styles could be lost after hiding and unhiding columns with series data.</div>


- <div>A crash could occur when trying to list changes on a new sheet for a workbook using legacy &quot;Shared Workbook&quot; mode.</div>


- <div>Fixed an issue where custom formatting for a single data point in a Pivot chart was not saved if the &quot;Invert if negative&quot; option was selected.</div>


- <div>Fixed the issue where custom formatting in Pivot charts may not be saved when the &quot;Invert if negative&quot; option was enabled.</div>


- <div>This change fixes an issue where the '@' character uploaded in a CSV file, would result in the string after the '@' character to be converted to a formula.</div>


- <div>Fixed an issue where decimal values in the SEQUENCE function were not rounded correctly.</div>


### Outlook

- <div>Addresses an issue that caused very long safelinks that users clicked on in the Outlook Desktop client to fail to load due to truncation.</div>


- <div>Fixed an issue where Outlook folders with names containing DBCS (Double Byte Character Set) characters would intermittently disappear when synchronizing with the server. For this to happen, Outlook had to be configured with an IMAP account and running on a system with the locale set to Japanese.</div>


### PowerPoint

- <div>Fixed an issue where if a user created a comment without posting it and closed the Comments pane, then opened a new window, navigated through multiple slides and, closed the window, and finally re-opened the Comments pane in the original presentation, the draft comments would not be available.</div>


### Project

- <div><span style="display:inline !important;">Fixed an issue where if Project is connected to Project Web App and the decimal separator is a comma, TaskDependencies Add method fails when Lag is added.</span><br></div>


### Word

- <div>Fixed an issue where inserting comments on a document in collaboration mode would not always work.</div>


- <div>This change fixes an issue where the People card would flash if the @ mention was clicked.</div>


- <div>Fixed the issue where closing a document with draft comments would prompt the user if they wanted to close the document without saving the draft comments. Cancelling the prompt would close the document rather than leaving it open.</div>


- <div>Fixed an issue where translating a posted comment would result in the error &quot;Inserting translated text failed'.</div>


- <span style="display:inline !important;">In Web View/Immersive reader, clicking on a hint would scroll to the top even though it was already in view. This has been fixed.</span><br>


- <div>We fixed an issue that, when attempting to save a file containing a macro under a new name, &nbsp;would cause it to be saved with .docx extension and the filename WRO0004.docx, regardless of what the user entered, rendering the document unusable.</div>



[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2005: May 04
*Version 2005 (Build 12827.20030)*


[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates
### Unknown ****** Please Review ******

- **:** description

- **:** 

- **:** 


[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

<br/>

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues
### Excel

- <div>Fixed an issue where chart data table could render values in a date axis incorrectly.</div>


- <div>Fixed an issue where&nbsp;page breaks could not be disabled after going into Page Layout or Page Break Preview.</div>


- <div>Inserting a column in a filtered list would take longer than expected.</div>


- <div>Fixed an issue where chart line styles could be lost after hiding and unhiding columns with series data.</div>


- <div>A crash could occur when trying to list changes on a new sheet for a workbook using legacy &quot;Shared Workbook&quot; mode.</div>


- <div>Fixed an issue where custom formatting for a single data point in a Pivot chart was not saved if the &quot;Invert if negative&quot; option was selected.</div>


- <div>Fixed the issue where custom formatting in Pivot charts may not be saved when the &quot;Invert if negative&quot; option was enabled.</div>


- <div>This change fixes an issue where the '@' character uploaded in a CSV file, would result in the string after the '@' character to be converted to a formula.</div>


- <div>Fixed an issue where decimal values in the SEQUENCE function were not rounded correctly.</div>


### Outlook

- <div>Addresses an issue that caused very long safelinks that users clicked on in the Outlook Desktop client to fail to load due to truncation.</div>


- <div>Fixed an issue where Outlook folders with names containing DBCS (Double Byte Character Set) characters would intermittently disappear when synchronizing with the server. For this to happen, Outlook had to be configured with an IMAP account and running on a system with the locale set to Japanese.</div>


### PowerPoint

- <div>Fixed an issue where if a user created a comment without posting it and closed the Comments pane, then opened a new window, navigated through multiple slides and, closed the window, and finally re-opened the Comments pane in the original presentation, the draft comments would not be available.</div>


### Project

- <div><span style="display:inline !important;">Fixed an issue where if Project is connected to Project Web App and the decimal separator is a comma, TaskDependencies Add method fails when Lag is added.</span><br></div>


### Word

- <div>Fixed an issue where inserting comments on a document in collaboration mode would not always work.</div>


- <div>This change fixes an issue where the People card would flash if the @ mention was clicked.</div>


- <div>Fixed the issue where closing a document with draft comments would prompt the user if they wanted to close the document without saving the draft comments. Cancelling the prompt would close the document rather than leaving it open.</div>


- <div>Fixed an issue where translating a posted comment would result in the error &quot;Inserting translated text failed'.</div>


- <span style="display:inline !important;">In Web View/Immersive reader, clicking on a hint would scroll to the top even though it was already in view. This has been fixed.</span><br>


- <div>We fixed an issue that, when attempting to save a file containing a macro under a new name, &nbsp;would cause it to be saved with .docx extension and the filename WRO0004.docx, regardless of what the user entered, rendering the document unusable.</div>



[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2005: April 24
*Version 2005 (Build 12816.20006)*


[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates
### Outlook

- **Improved links in email:** When you include a link to a file, the file name replaces the URL. You can change permissions so all recipients have access.

- **File name and permission errors for links in the body of emails:** Send online attachments with ease
and modify it's permissions right within your message.

- **Friendlier link names:** When you insert a document link, only the document name will be displayed. [Learn more](https://support.office.com/article/02040f47-bd56-4806-8311-fc913fed54c0)


[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

<br/>

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues
### Excel

- <div>This change fixes an issue where the chart trendline R-Squared value (in the forced y-intercept case) was incorrect even though the LINEST function returns the correct value.</div>


- <div>This change fixes an issue where customized chart trendline formatting was not always being saved.</div>


### Outlook

- <div>Fixed an issue where the categorize button for group calendars in the Office Ribbon was disabled.</div>


- <div>Fixed an issue where enterprise customers with group folders not implemented or not working, would result in Outlook displaying a &quot;not responding&quot; message.</div>


### PowerPoint

- <div>Fixed an issue where hovering over the asterisk (*) symbol did not display the user name and date of the last person to update the document.<br></div>


### Word

- <div>Enabling the option &quot;Show bookmarks&quot; would not display bookmarks. This has been fixed.</div>


- <div>This change fixes an issue where text with hyperlinks may not display if the option: &quot;Show field codes instead of their values&quot; was enabled.</div>



[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2005: April 17
*Version 2005 (Build 12810.20002)*


[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues
### Excel

- <div>Increased the size of the cell reference edit controls on the Custom Error Bars dialog used with charts.</div>


- <div>Workbooks saved with a digital signature in Excel 2016 could have the signature invalidated upon opening in the current version of Excel.</div>


- <div><span style="display:inline !important;">Fixed a problem with the scaling of checkboxes in form controls when printed.</span><br></div>


- <div><span style="display:inline !important;">Application.Evaluate (VBA) was not working for User-defined functions in some cases.</span><br></div>


- <div>This change fixes an issue where conditional formatting (CF) information was not being saved to XLSB files correctly.</div>


### OneNote

- <div>Fixed an issue where line breaks were being stored as vertical tabs.</div>


### Outlook

- <div>Addresses an issue that caused users to be unable to add a Personal Contact Group as a Meeting attendee.</div>


- <div>Addresses an issue that caused meetings that are more than 2 months away to fail to display a meeting subject in the Scheduling Assistant.</div>


- <div><span style="display:inline !important;">Addresses an issue that caused users to see message body truncation when forwarding large HTML messages.</span><br></div>


- <div>Added the ability to enforce S/MIME default signing configuration via group policy</div>


- <div>Addresses an issue that caused delete rules created for mailboxes other than the user's primary mailbox to become invalid.</div>


- <div>Addresses an issue that caused attachments to get dropped when forwarding an encrypted message.</div>


### Project

- <div><span style="color:rgba(0, 0, 0, 0.9);display:inline !important;">When&nbsp;</span><span style="box-sizing:border-box;color:rgba(0, 0, 0, 0.9);">Predecessor/Successor data is edited within a Form view, an extra<span>&nbsp;</span><span style="box-sizing:border-box;font-size:13.3333px;display:inline !important;">ProjectBeforeTaskChange&nbsp;</span>event i</span><span style="box-sizing:border-box;color:rgba(0, 0, 0, 0.9);">s fired</span><br></div>


- <div>Fixed an issue where Project may crash when changing the board status field on a project that is connected to a SharePoint task list.</div>


- <div><span style="display:inline !important;">Fixed an issue where Project may crash when saving projects created with older versions of Project.</span><br></div>



[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2004: April 10
*Version 2004 (Build 12730.20024)*


[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates
### Access

- **Bypass the Show Table dialog box, go directly to a task pane, and streamline adding tables to relationships and queries.:** Add tables and queries by clicking four tabs, multi-selecting names, searching by text, and dragging from a task pane that stays open while you work.

### Excel

- **M365 Premium Content Picker:** Bring your documents to life! Explore 1000’s of royalty free stock images, icons and stickers [Learn more](https://support.office.com/article/c7b78cdf-2503-4993-8664-851085c30fce)

### Outlook

- **M365 Premium Content Picker:** Bring your documents to life! Explore 1000’s of royalty free stock images, icons and stickers [Learn more](https://support.office.com/article/c7b78cdf-2503-4993-8664-851085c30fce)

- **Keep your pictures high fidelity when sending them as part of an email:** A new Outlook setting is available to limit picture compression when you send pictures as part of the email contents

### PowerPoint

- **M365 Premium Content Picker:** Bring your documents to life! Explore 1000’s of royalty free stock images, icons and stickers [Learn more](https://support.office.com/article/c7b78cdf-2503-4993-8664-851085c30fce)

- **Synchronize changes while you are presenting:** Synchronize changes whenever they are made even when the presentation is in slide show mode.

### Word

- **M365 Premium Content Picker:** Bring your documents to life! Explore 1000’s of royalty free stock images, icons and stickers [Learn more](https://support.office.com/article/c7b78cdf-2503-4993-8664-851085c30fce)

- **Annotate your private copy:** Create hand written notes for your eyes by making a private copy of a shared document. Go to View > Create a Private Copy to get started. [Learn more](https://support.office.com/article/275ea4ab-0125-48b4-b528-a20a1ae392d5)


[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

<br/>

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues
### Access

- <div>Fixed issues with resizing and refreshing tables in the task pane.</div>


### Excel

- <div>Fixed an issue where inserting a user defined chart template as default would result in saving it as a column chart.</div>


- <div>Fixed an issue where an Excel sheet with R1C1 cell referencing enabled and is being co-authored / shared, hovering over the user presence icon does not display the active cell reference in R1C1 mode.</div>


### Outlook

- <div>Addresses an issue that caused delegates to see different folder hierarchies on different machines for shared mailboxes.</div>


- <div>Addresses an issue that caused categories to occasionally disappear from email messages.</div>


### PowerPoint

- <div>This change fixes an issue where the&nbsp;<span style="display:inline !important;">rendering of a legacy Excel chart embedded as an OLE object in PowerPoint or Word may not always display the chart title.</span></div>


- <div><span style="font-size:12.0pt;font-family:&quot;Calibri&quot;,serif;color:black;">This change fixes an issue where finding special characters using 'find whole words only' did not always work as expected.</span><br></div>


### Project

- <div>Fixed an issue where the user couldn't enter time-phased Baseline Work when the setting to protect actual work is on.</div>


### Word

- <div>This change fixes an issue where the&nbsp;<span style="display:inline !important;">rendering of a legacy Excel chart embedded as an OLE object in PowerPoint or Word may not always display the chart title.</span></div>


- <div>This change fixes an issue in two page view, when creating a comment, the comment anchor did not always come into view.</div>


- <div><span style="color:rgb(0, 0, 0);">Fixed an issue where if a&nbsp;</span><font color="rgba(0, 0, 0, 0.901960784313726)"><span style="display:inline !important;color:rgb(0, 0, 0);">paragraph whose style is an ancestor of a style linked to a list, then the numbering of that list could be lost.</span></font></div>



[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2004: April 03
*Version 2004 (Build 12725.20006)*


[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates
### Access

- **"Add Tables" Task Pane:** Access's new "Add Tables" Task Pane is finally here! This feature allows you to easily select/multi-select which tables they'd like to add/remove into a query window, without navigating to the "Show Tables" dialog for queries and for relationship view. This also includes a new "links" tab to display linked tables, a search box to filter the current list, "drag and drop" behavior, and more!

### Excel

- **M365 Premium Content Picker:** Bring your documents to life! Explore 1000’s of royalty free stock images, icons and stickers

### Outlook

- **M365 Premium Content Picker:** Bring your documents to life! Explore 1000’s of royalty free stock images, icons and stickers

### PowerPoint

- **M365 Premium Content Picker:** Bring your documents to life! Explore 1000’s of royalty free stock images, icons and stickers

- **Keep slides updated in slide show:** Automatically receive content updates during slide show

### Word

- **M365 Premium Content Picker:** Bring your documents to life! Explore 1000’s of royalty free stock images, icons and stickers


[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

<br/>

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues
### Excel

- <div>Fixed an issue where selecting a range of cells on a sheet would result in the selection of a single cell.</div>


- <div>Fixed an issue which could cause Excel to stop responding when reducing the size of a chart with some x-axis ranges.</div>


- <div>Fixed an issue where Data labels on charts would display as blank when the underlying data cells did not have a caption.</div>


### Outlook

- <div>Addresses an issue that caused users to experience a crash when attempting to view the properties of an Organizational Form.</div>


- <div>Addresses an issue that caused some reminders to fail to fire when changing the timezone on a machine.</div>


### PowerPoint

- <div>We have fixed an issue when copying&nbsp;text from Excel to PowerPoint might change its formatting.&nbsp;</div>


### Word

- <div><span style="display:inline !important;">This change fixes an issue where hovering a cursor over a hint would not highlight its card.</span></div>


- <div>This change fixes an issue that would cause the text in grouped shapes to disappear temporarily when using the Lasso selection tool.</div>



[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2004: March 27
*Version 2004 (Build 12718.20010)*


[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates
### Outlook

- **New option to disable @ mention suggestions when composing mail:** Do you find the @ mention picker more annoying than useful? Now you can turn it off if you prefer.


[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

## Version 2004: March 20
*Version 2004 (Build 12711.20000)*


[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates
### Outlook

- **Calendar visual refresh:** Last year, we brought you a refreshed mail experience, and, this year, it is the calendar’s turn to get a facelift! The updates are fresh but familiar so, as a seasoned Outlook user, you can jump in and be more productive right away.

- **Help protect data in your group:** The Sensitivity label you choose when creating a group is applied to group email, documents, and team sites

### PowerPoint

- **Update slides during slide show:** Update slides changed by other authors during your presentation.


[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

<br/>

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues
### Excel

- <div><span>This change addresses delays when processing images with malformed or invalid protocol information.</span></div>


### Outlook

- <div><span>This change addresses delays when processing images with malformed or invalid protocol information.</span></div>


- <div>This change fixes an issue where the latest changes to draft emails were not being updated.</div>


- <div>Fixed an issue where right-mouse clicking on a file and using 'Send to' would not work.</div>


- <div>Fixed an issue where if a user had a customized the search path for the Address book<span style="display:inline !important;">, Outlook's name resolution scope would be limited to the customized path rather than including the Global Address List (GAL).</span></div><div><span style="display:inline !important;"><br></span></div><div><span style="display:inline !important;"><br></span></div>


- <div>Fixed an issue where w<span style="display:inline !important;">ithin a set of returned search results, sorting the results by Categories would not display the Category colors.</span></div>


### Project

- <div><span style="font-size:11.0pt;font-family:&quot;Calibri&quot;,sans-serif;color:#1F497D;">Fixed an issue where the 'ProjectBeforeTaskChange' Visual Basic Applications (VBA) event did not fire when a user clicked the “Inactivate” button found on the Tasks Ribbon within the Scheduling
grouping.</span><br></div>


- <div><span style="font-size:11.0pt;font-family:&quot;Calibri&quot;,sans-serif;color:#1F497D;">If
you set predecessor or successor details from within a Form type view, the
ProjectBeforeTaskChange Visual Basic Applications (VBA) event didn't always capture
the changes. For example, if you deleted a dependency and clicked OK on the form,
the event did not fire. This behavior has been fixed.</span><br></div>


- <div>Fixed an issue where the latest values for the&nbsp;<span style="color:rgb(30, 30, 30);font-family:&quot;Segoe UI&quot;, &quot;Segoe UI Web&quot;, wf_segoe-ui_normal, &quot;Helvetica Neue&quot;, &quot;BBAlpha Sans&quot;, &quot;S60 Sans&quot;, Arial, sans-serif;font-size:16px;display:inline !important;">Actual Cost of Work Performed (ACWP) would not be displayed after making a change, such as a date change.</span></div>


- <div>Fixed an issue where opening a project using the Most Recently Used (MRU) menu opened the project file with Read/Write access.</div>


- <div>This change fixes an issue where if you created a manual task with a start date and a time (but no duration), it would be displayed with an incorrect time on the timeline.</div>


- <div>Fixed an issue where printing a timeline using a&nbsp;<span style="display:inline !important;">Hijri calendar would result in a month being skipped or duplicated in the print view.</span></div>


- <div>This change addresses an issue where working in Team Planner with GDI objects, could result in the over allocation of GDI objects and create low memory conditions.</div>


### Word

- <div>Fixed an issue where the functionality to post comments was disabled.</div>


- <div><span>This change addresses delays when processing images with malformed or invalid protocol information.</span></div>


- <div>This change addresses an issue where the account manager would not dispatch messages resulting in a hang with third party applications.</div>


- <div>This change fixes an issue where the&nbsp;<span style="display:inline !important;">Table of Contents would get updated with heading styles which were not present in the document.</span></div>


- <div>Fixed an issue where digital signatures saved in Word documents would be removed when mailing the documents.</div>



[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2004: March 12
*Version 2004 (Build 12703.20010)*


[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates
### Security

- **Sensitivity labels:** You can now apply a sensitivity label that your organization has configured to prompt you for custom permissions.


[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

## Version 2003: March 06
*Version 2003 (Build 12624.20086)*


[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues
### Excel

- <div>Fixed a cosmetic issue where the 'OK' button on the File \ Options dialog displayed as being grayed out but functionality was not impacted.</div>


### Outlook

- <div>Fixed an issue where creating a rule with Outlook Web Access did not persist to the Exchange server and resulted in a conflict.</div>


- <div>Fixed an issue with Outlook in dark mode would not display the drop down list in the 'From:' field.</div>


- <div>Addresses an issue that caused users to be unable to attach a file to their mail message via the file explorer when that file was open in another application.</div>


### PowerPoint

- <div>Fixed an issue where the recommended thumbnails flash when hovering your mouse over the thumbnails. In some cases this could cause PowerPoint to crash.</div>


- <div>Fixed a cosmetic issue where the 'OK' button on the File \ Options dialog displayed as being grayed out but functionality was not impacted.</div>


### Word

- <div>Fixed a cosmetic issue where the 'OK' button on the File \ Options dialog displayed as being grayed out but functionality was not impacted.</div>


- <div>Fixed an issue with Compare feature for <span style="display:inline !important;background-color:rgba(255, 255, 255, 1);font-size:13.33px;">documents that were protected for editing.</span></div>


### Office Suite

- <div>Fixed an issue Word/Excel/PowerPoint where the User Principal Name (UPN) is no longer case sensitive resulting in less failures when working with files on SharePoint.</div>



[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2003: February 19
*Version 2003 (Build 12615.20000)*


[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates
### Excel

- **Pick the perfect color:** Use hex color codes to choose exactly the color you want for your font, text highlight, and more.

### Outlook

- **Pick the perfect color:** Use hex color codes to choose exactly the color you want for your font, text highlight, and more.

### PowerPoint

- **Pick the perfect color:** Use hex color codes to choose exactly the color you want for your font, text highlight, and more.

### Word

- **Pick the perfect color:** Use hex color codes to choose exactly the color you want for your font, text highlight, and more.


[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

<br/>

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues
### Excel

- <div><u style="box-sizing:border-box;font-size:16px;color:rgb(0, 128, 0);"><p style="margin:0in 0in 0.0001pt;font-size:9pt;font-family:&quot;Helvetica Neue&quot;;font-weight:normal;"><br></p></u></div>


- <p style="margin:0in 0in 0pt;font-family:&quot;Calibri&quot;,sans-serif;font-size:11pt;"><br></p><div></div>


- <div>We fixed a bug where CSV files were loaded incorrectly when the first word in the file was TABLE.</div>


- Please contact the PM(Theresa Estrada ) to get the Accept message that was sent to the customer.


### Outlook

- <div>Addresses an issue that caused Outlook to unexpectedly generate logging output in some scenarios, even when logging was turned off.</div>


- <div>Addresses an issue that caused users to be unable to open public folder messages when Outlook was left running overnight.</div>


### Project Online

- <div>Fixed an issue where you couldn't edit a resource after deleting a resource custom field.</div>


### Office Suite

- <div><p style="margin:0in 0in 8pt;font-size:11pt;font-family:Calibri, sans-serif;"><span style="font-size:10.0pt;font-family:&quot;Segoe UI&quot;,sans-serif;color:black;background:white;">When using Multichoice/Lookup/Managed-metadata
properties with Word/Excel/PowerPoint documents and saving to a SharePoint
Document Library, these properties were previously limited to 255 characters.
When these properties exceeded 255 characters, such documents could not be
saved. With this change, this limit has been increased to 2048 characters.</span></p></div>



[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2003: February 14
*Version 2003 (Build 12607.20000)*

## Version 2002: February 07
*Version 2002 (Build 12527.20040)*


[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates
### Access

- **Be more productive working in Query Designer, SQL view, and the Relationships window:** Right-click a table to open, design, size, and hide it. Search and replace text in SQL View. Select multiple tables in the Relationships window.

### Word

- **:** 


[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

<br/>

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues
### Access

- <div><span style="display:inline !important;">This update fixes an issue where using an ADODB. Recorder object in VB code may incorrectly report an error.&nbsp;</span><br></div>


- <div><span><span style="display:inline !important;background-color:rgba(255, 255, 255, 1);">This update fixes an issue that can cause Microsoft Access to fail to identify an Identity Column in a linked SQL Server table, which can cause rows to be reported as deleted incorrectly</span><br></span></div>


### Excel

- <div>Fixed an issue where Excel would crash when using Text To Columns with dynamic arrays.</div>


### Outlook

- <div>Fixed an issue where scrolling in calendar with month view, fails to show previous calendar events.<br></div>


- <div>Addresses an issue that caused users to experience a crash when viewing more than 30 calendars in a Citrix environment.</div>


### PowerPoint

- <div>Fixed an issue where&nbsp;<span style="font-size:13.3333px;display:inline !important;">After closing a file, PowerPoint does not immediately remove it from the Presentations collection if there are any event handlers running. Hence the number of open presentations reported by the object model is incorrect, and shutdown of PowerPoint is prevented.</span>&nbsp;</div>


- <div>Fixed an issue with highlighter : White texts with dark highlighter colors are printed as black in Grayscale.&nbsp;</div>


### Word

- <div>Updating and scrolling through a table of contents may sometimes display a gray area over the document.</div>


- <div>Fixed an issue where if a document is being coauthored, the draft version of a root comment may not be preserved.</div>


- <div>Fixed an issue where going back and forth between comment cards would sometimes display the initially selected comment with a selection highlight.</div>


- <div>Fixed an issue where using 'Browse' to save a file did not work if a comment was written but not posted and the user tried to save the file.</div>


- <div>With SlideTrack enabled and the comments pane closed,&nbsp;Ctrl+Alt+M may not open the comments pane.</div>


- <div>Fixed an issue when adding @mention in a table could generate the error message: 'A table in this document has become corrupted'.</div>


### Office Suite

- <div>Resolves an issue that may have caused Norway Nynorsk (nn-no) proofing tools package to be installed incorrectly.</div>



[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2002: January 31
*Version 2002 (Build 12513.20010)*


[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates
### Excel

- **Read and reply on the fly:** Respond to comments and mentions right from email without opening the workbook.

- **Data Profiling in Query Editor:** Get at-a-glance analysis of the data in your columns, identify error and empty values, see distribution histograms and more.

### Word

- **:** 


[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

<br/>

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues
### Outlook

- <div><span>Fixed an issue where emails expiring based on a retention policy would display two labels. One showing that the mail will expire in one day and another displaying that it will expire in two days.<br></span></div>


### Word

- <div><span>Fixed an issue where </span><font style="background-color:rgba(255, 255, 255, 1);">c<span style="display:inline !important;">omment hint was not visible in read mode with &quot;Inverse&quot; page color.</span></font></div>


- <div><span>Fixed an issue where italics formatting is lost after editing a comment, italicizing the text and then posting it.<br></span></div>


- <div><span>Fixed an issue where comment commands (Edit comment, Reply to comment, Delete comment, Resolve comment) in the comments context menu were not being displayed.</span></div>



[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2002: January 17
*Version 2002 (Build 12508.20000)*


[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates
### Word

- **Mention & comment notification emails now contain previews:** Email notifications for mentions & comments will now include previews of the comment text and context for what it is referring to.


[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

<br/>

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues
### Excel

- <h5>In some cases, the Accessibility checker would not show the recommendations for shapes.</h5>


### Outlook

- <div><span>Folders saved in 'Favorites' in the left navigation pane may intermittently disappear.</span></div>


### PowerPoint

- <div><span><p style="margin:0cm 0cm 0.0001pt;font-size:11pt;font-family:&quot;Calibri&quot;, sans-serif;"></p><p style="margin:0cm 0cm 0.0001pt;font-size:11pt;font-family:&quot;Calibri&quot;, sans-serif;"><span><span style="font-size:11.0pt;font-family:&quot;Calibri&quot;,sans-serif;">Fixed an issue where Ink
may not render completely or get skipped when used in a PowerPoint ink
animations.</span></span></p><span></span><p></p></span></div>



[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2001: January 03
*Version 2001 (Build 12425.20000)*


[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues
### Access

- <div><span style="font-size:11.0pt;font-family:&quot;Calibri&quot;,sans-serif;">RelNotesNotNeeded</span><br></div>


### Excel

- <div>Some border lines were not printing as expected on A4 size paper.</div>


- <div><p style="box-sizing:border-box;margin-top:0px;margin-bottom:0px;font-family:Calibri, Arial, Helvetica, sans-serif;color:rgb(32, 31, 30);font-size:15px;background-color:white;"><b style="box-sizing:border-box;"><span style="box-sizing:border-box;margin:0px;font-size:16pt;">Excel has decided to<span style="box-sizing:border-box;margin:0px;">&nbsp;</span><span style="box-sizing:border-box;margin:0px;">accept</span><span style="box-sizing:border-box;margin:0px;">&nbsp;</span>this QFE:&nbsp;&nbsp;</span><span style="box-sizing:border-box;margin:0px;"><span style="box-sizing:border-box;margin:0px;font-size:16pt;color:rgb(149, 79, 114);text-decoration:underline;"><span style="box-sizing:border-box;margin:0px;"><a href="https://office.visualstudio.com/OC/_workitems/edit/3771884/" rel="noopener noreferrer" target=_blank title="https://office.visualstudio.com/OC/_workitems/edit/3771884/" style="box-sizing:border-box;cursor:pointer;margin:0px;">3771884</a></span></span></span></b></p><p style="box-sizing:border-box;margin-top:0px;margin-bottom:0px;font-family:Calibri, Arial, Helvetica, sans-serif;color:rgb(32, 31, 30);font-size:15px;background-color:white;"><span style="box-sizing:border-box;margin:0px;">&nbsp;</span></p><p style="box-sizing:border-box;margin-top:0px;margin-bottom:0px;font-family:Calibri, Arial, Helvetica, sans-serif;color:rgb(32, 31, 30);font-size:15px;background-color:white;"><span style="box-sizing:border-box;margin:0px;">Thank you for reporting this problem to our team.&nbsp; We have reviewed the issue you sent us and have decided to<span style="box-sizing:border-box;margin:0px;">&nbsp;</span><span style="box-sizing:border-box;margin:0px;">accept</span><span style="box-sizing:border-box;margin:0px;">&nbsp;</span>the hotfix you have requested&nbsp;for QFE&nbsp;<span style="box-sizing:border-box;margin:0px;color:rgb(149, 79, 114);text-decoration:underline;"><span style="box-sizing:border-box;margin:0px;"><a href="https://office.visualstudio.com/OC/_workitems/edit/3771884/" rel="noopener noreferrer" target=_blank title="https://office.visualstudio.com/OC/_workitems/edit/3771884/" style="box-sizing:border-box;cursor:pointer;margin:0px;">3771884</a></span></span>.</span></p><p style="box-sizing:border-box;margin-top:0px;margin-bottom:0px;font-family:Calibri, Arial, Helvetica, sans-serif;color:rgb(32, 31, 30);font-size:15px;background-color:white;"><b style="box-sizing:border-box;"><span style="box-sizing:border-box;margin:0px;">&nbsp;</span></b></p><p style="box-sizing:border-box;margin-top:0px;margin-bottom:0px;font-family:Calibri, Arial, Helvetica, sans-serif;color:rgb(32, 31, 30);font-size:15px;background-color:white;"><b style="box-sizing:border-box;"><span style="box-sizing:border-box;margin:0px;">Reason for<span style="box-sizing:border-box;margin:0px;">&nbsp;</span><span style="box-sizing:border-box;margin:0px;">accept</span>ance:</span></b></p><p style="box-sizing:border-box;margin-top:0px;margin-bottom:0px;font-family:Calibri, Arial, Helvetica, sans-serif;color:rgb(32, 31, 30);font-size:15px;background-color:white;"><span style="box-sizing:border-box;margin:0px;">The scenario caused a large enough problem for customers that it was deemed necessary to fix.&nbsp; The issue in question, having to manually add in a header image to a chart using a macro because the macro would fail, happened often enough to make it necessary.</span></p><p style="box-sizing:border-box;margin-top:0px;margin-bottom:0px;font-family:Calibri, Arial, Helvetica, sans-serif;color:rgb(32, 31, 30);font-size:15px;background-color:white;"><span style="box-sizing:border-box;margin:0px;">&nbsp;</span></p><p style="box-sizing:border-box;margin-top:0px;margin-bottom:0px;font-family:Calibri, Arial, Helvetica, sans-serif;color:rgb(32, 31, 30);font-size:15px;background-color:white;"><b style="box-sizing:border-box;"><span style="box-sizing:border-box;margin:0px;">Timeline for the fix:</span></b></p><p style="box-sizing:border-box;margin-top:0px;margin-bottom:0px;font-family:Calibri, Arial, Helvetica, sans-serif;color:rgb(32, 31, 30);font-size:15px;background-color:white;"><span style="box-sizing:border-box;margin:0px;">Pending successful validation of the fix.&nbsp; The fix will be released on the following schedule:</span></p><ul style="box-sizing:border-box;padding:0px 0px 0px 40px;font-family:Calibri, Arial, Helvetica, sans-serif;color:rgb(32, 31, 30);font-size:15px;margin-top:0px;margin-bottom:0px;"><li style="box-sizing:border-box;font-size:11pt;font-family:Calibri, sans-serif;margin:0px;"><b style="box-sizing:border-box;"><u style="box-sizing:border-box;">Insiders Fast</u></b><b style="box-sizing:border-box;">:</b>&nbsp;&nbsp;Initial fix will be checked into our&nbsp;codebase in the next two weeks.&nbsp;&nbsp;If no issues are found that fix will reach our insiders fast audience in the next 4-5 weeks.&nbsp;&nbsp;If you would like to validate that the fix will work for your&nbsp;scenario,&nbsp;please install an insider's fast build on a test box within your organization.&nbsp;&nbsp;Instructions for how to update to the&nbsp;insider's&nbsp;fast release of Office can be found here:&nbsp;&nbsp;<a href="https://nam06.safelinks.protection.outlook.com/?url=https://insider.office.com/en-us/&amp;data=04%7c01%7cbrslinin%40microsoft.com%7c00c933a0bfa940d0d37a08d77856972c%7c72f988bf86f141af91ab2d7cd011db47%7c1%7c0%7c637110185608667687%7cUnknown%7cTWFpbGZsb3d8eyJWIjoiMC4wLjAwMDAiLCJQIjoiV2luMzIiLCJBTiI6Ik1haWwiLCJXVCI6Mn0%3D%7c-1&amp;sdata=/AfF0nhEbu0EqiEuNQJLSWSZbjVS6cxK9TtY4llvfCs%3D&amp;reserved=0" rel="noopener noreferrer" target=_blank title="Original URL: https://insider.office.com/en-us/. Click or tap if you trust this link." style="box-sizing:border-box;text-decoration:underline;cursor:pointer;margin:0px;"><span style="box-sizing:border-box;margin:0px;color:rgb(149, 79, 114);">https://insider.office.com/en-us/</span></a>.&nbsp; Release notes for these builds will have information detailing what build your bug was fixed in:&nbsp;<a href="https://nam06.safelinks.protection.outlook.com/?url=https://docs.microsoft.com/en-us/officeupdates/release-notes-office-insider&amp;data=04%7c01%7cbrslinin%40microsoft.com%7c00c933a0bfa940d0d37a08d77856972c%7c72f988bf86f141af91ab2d7cd011db47%7c1%7c0%7c637110185608677686%7cUnknown%7cTWFpbGZsb3d8eyJWIjoiMC4wLjAwMDAiLCJQIjoiV2luMzIiLCJBTiI6Ik1haWwiLCJXVCI6Mn0%3D%7c-1&amp;sdata=iSL0qytoTlu3RusX239l5tr5NequjohK3Zn1I3LQoRw%3D&amp;reserved=0" rel="noopener noreferrer" target=_blank title="Original URL: https://docs.microsoft.com/en-us/officeupdates/release-notes-office-insider. Click or tap if you trust this link." style="box-sizing:border-box;text-decoration:underline;cursor:pointer;margin:0px;"><span style="box-sizing:border-box;margin:0px;color:rgb(149, 79, 114);">https://docs.microsoft.com/en-us/officeupdates/release-notes-office-insider</span></a>.</li><li style="box-sizing:border-box;font-size:11pt;font-family:Calibri, sans-serif;margin:0px;"><b style="box-sizing:border-box;"><u style="box-sizing:border-box;">Monthly Channel</u></b><b style="box-sizing:border-box;">:&nbsp;</b>The fix will be available in our current channel for the Office 365 Monthly update with the December 2019 (1912) fork. We expect the fix to be available for production customers around the second week of<span style="box-sizing:border-box;margin:0px;color:black;background-color:white;">&nbsp;<b style="box-sizing:border-box;"><u style="box-sizing:border-box;">February 2020</u></b></span>.</li><li style="box-sizing:border-box;font-size:11pt;font-family:Calibri, sans-serif;margin:0px 0px 12pt;"><b style="box-sizing:border-box;"><u style="box-sizing:border-box;"><span style="box-sizing:border-box;margin:0px;color:black;background-color:white;">Semi Annual Targeted Channel</span></u></b><b style="box-sizing:border-box;"><span style="box-sizing:border-box;margin:0px;color:black;background-color:white;">:&nbsp;</span></b><span style="box-sizing:border-box;margin:0px;color:black;background-color:white;">The fix will be available in our&nbsp;Semi-Annual Targeted channel update for Office 365 subscription customers because the January 2020 (2001) fork will be the new SACT release</span></li><li style="box-sizing:border-box;font-size:11pt;font-family:Calibri, sans-serif;margin:0px;"><b style="box-sizing:border-box;"><u style="box-sizing:border-box;"><span style="box-sizing:border-box;margin:0px;color:black;">Semi Annual Channel</span></u></b><b style="box-sizing:border-box;"><span style="box-sizing:border-box;margin:0px;color:black;">:</span></b><span style="box-sizing:border-box;margin:0px;color:black;">&nbsp;The fix will be available in our&nbsp;<span style="box-sizing:border-box;margin:0px;background-color:white;">Semi-Annual channel update for Office 365 subscription customers in the July 2019 (1907) fork</span>. &nbsp;We expect to release for production customers in the patch Tuesday release for the month of&nbsp;<b style="box-sizing:border-box;"><u style="box-sizing:border-box;">March 2020</u></b>.</span></li><li style="box-sizing:border-box;font-size:11pt;font-family:Calibri, sans-serif;margin:0px;"><b style="box-sizing:border-box;"><u style="box-sizing:border-box;"><span style="box-sizing:border-box;margin:0px;color:black;">Office 2019</span></u></b><b style="box-sizing:border-box;"><span style="box-sizing:border-box;margin:0px;color:black;">:</span></b><span style="box-sizing:border-box;margin:0px;color:black;">&nbsp;</span><span style="box-sizing:border-box;margin:0px;">The fix will&nbsp;</span><b style="box-sizing:border-box;"><i style="box-sizing:border-box;">not</i></b><span style="box-sizing:border-box;margin:0px;">&nbsp;be available for Office 2016 Customers.&nbsp; The issue is not found in Excel 2019 LTSB (it started in 1809 and LTSB is 1807), so there is nothing there to fix.</span></li><li style="box-sizing:border-box;font-size:11pt;font-family:Calibri, sans-serif;margin:0px;"><b style="box-sizing:border-box;"><u style="box-sizing:border-box;">Office 2016</u></b><b style="box-sizing:border-box;">:&nbsp;</b>The fix will&nbsp;<b style="box-sizing:border-box;"><i style="box-sizing:border-box;">not</i></b>&nbsp;be available for Office 2016 Customers.&nbsp; The issue is not found in Excel 2016, so there is nothing there to fix.</li></ul><p style="box-sizing:border-box;margin-top:0px;margin-bottom:0px;font-family:Calibri, Arial, Helvetica, sans-serif;color:rgb(32, 31, 30);font-size:15px;background-color:white;"><span style="box-sizing:border-box;margin:0px;">&nbsp;</span></p><p style="box-sizing:border-box;margin-top:0px;margin-bottom:0px;font-family:Calibri, Arial, Helvetica, sans-serif;color:rgb(32, 31, 30);font-size:15px;background-color:white;"><i style="box-sizing:border-box;"><span style="box-sizing:border-box;margin:0px;">*If this plan changes due to issue found during validation we will notify you and provide an updated plan for delivery of the fix.</span></i></p><br></div>


- <div><span>When formatting a chart axis, the interval between labels is limited to 255.</span></div>


- <div>Fixed an issue where an error would occur trying to refresh an XML query in which the URL to the datasource was being truncated.</div>


- <div><span style="display:inline !important;background-color:rgba(255, 255, 255, 1);">Workbook Statistics reports a macro count
from all open workbooks, including the personal macro workbook.</span><br></div>


### Outlook

- <div>Switching folders may result in a brief white 'flash' in the mail list / mail preview. This behavior is more pronounced in dark mode.</div>


### PowerPoint

- <div><span style="color:rgba(0, 0, 0, 0.9);display:inline !important;">Fixed an OM issue where calling Shape.Paste method results in the pasted shape receiving focus.&nbsp;</span><br></div>


- <div>Improved a copy-paste scenario:&nbsp;<span style="font-size:13.3333px;display:inline !important;">Copying the Shape in powerpoint slide and paste it in other slide in a loop might fail with exception.&nbsp;</span></div>


- <div><span>Animation in the section headers of slides may not render properly after <span style="display:inline !important;background-color:rgba(255, 255, 255, 1);">collapsing and expanding section headers.</span></span></div>


### Project

- <div>Fixed an issue where if a resource has more than one cost rate, cost value on assignments may not be correct.</div>


- <div>Fixed an issue where text wrapping wasn't working in the task and resource usage views.</div>


### User Lifecycle

- <div>Removed showing erroneous expiry date of the valid license when trying to change license with only one license</div>


### Word

- <div><span><span>Inserting a control (such as a Text Content Control) in an equation then saving and opening the file results in an un-readable content error.</span><br></span></div>


- <div><span>When co-authoring, adding a comment using Word online may not appear in Word desktop.</span></div>



[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2001: December 13
*Version 2001 (Build 12410.20000)*


[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates
### Outlook

- **Drag email to a group you own:** Move and copy messages and conversations by dragging them from your inbox. Messages you drag will be shared with all group members.


[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

<br/>

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues
### Access

- <div><span><span style="display:inline !important;background-color:rgba(255, 255, 255, 1);font-size:13.33px;">Executing a union query that references linked ODBC table(s) and contains an Order By clause crashes 64 bit Access.</span><br></span></div>


- <div><span>Summing of data from union queries in Access (O365) may result in truncation of decimal data.<br></span></div>


- <div><span>COM Interfaces for ACE are not exposed for use outside of Office applications.</span></div>


### Excel

- <div><span>Inserting a 3D model (animated or static) and trying to 'Save as Picture' doesn't work.<br></span></div>


- <div><span><span style="display:inline !important;background-color:rgba(255, 255, 255, 1);">The shortkey (Alt+Ctrl + 7/8)  to launch feedback from backstage is in conflict with AZERTY keyboards (Alt-Gr + 7/8). Impact: users may not be able to use some characters such as: '\'</span><br></span></div>


### Outlook

- <div><span style="display:inline !important;">Addresses an issue that caused email to be sent to the wrong address for the recipient.</span><br></div>


- <div>Addresses an issue that caused email to be sent to the wrong address for the recipient.</div>


- <div>This addresses an issue that caused Outlook to incorrectly allow users with &quot;read&quot; access to a mailbox to change the read/unread state of a message.</div>


- <div>The revocation of the security certificate on the web site is not re-producible by Product Support. Logging needs to be added to help root cause the issue.</div>


- <div>Addresses an issue that caused users to see synchronization failures.</div>


### PowerPoint

- <div><span>Inserting a 3D model (animated or static) and trying to 'Save as Picture' doesn't work.<br></span></div>


### Project

- <div><span>Task work is not calculated in Summary roll-up for manually scheduled child tasks.<br></span></div>


- <div><span>Project VBA Code invoked from a Ribbon button may not work when trying to save server-based Projects.<br></span></div>


- <div><span>Unless Project is already running, opening Project files from a SharePoint document library displays an error and the file will not open.</span></div>


### Word

- <div><span>Saving existing files may not work.</span></div>


- <div><span>Using arrow keys in the Spelling and Grammar editor window may result in intermittent flickering.<br></span></div>


- <div><span>When resolving a follow-up, associated comments may not convert to point comments.</span></div>


- <div><span>Inserting a 3D model (animated or static) and trying to 'Save as Picture' doesn't work.<br></span></div>


### Office Suite

- <p style="margin:0in 0in 0.0001pt;font-size:11pt;font-family:Calibri, sans-serif;"><span style="font-size:11pt;">Fixed an issue where Office update messages appear in a different language than expected.&nbsp;</span><span style="font-size:11pt;">Going forward, Office update messages will correctly match the Windows display language.</span><br></p>



[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 1912: December 06
*Version 1912 (Build 12325.20012)*


[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates
### Outlook

- **Advanced group email settings:** This feature helps groups users to customize which emails or events to receive/follow in their inbox.

- **Groups Naming policy:** A group naming policy enables the IT admin to standardize and manage the names of groups created by users in the organization. The admin can require a specific prefix and suffix be added to the name for a group when it's created, and can block specific words from being used. This helps minimize the use of inappropriate words in group names as well as IT manage the representation of groups in their directory. Naming Policy also helps organizations that deploy team sites to categorize them based on department.

### Office Suite

- **Tabbed Panes:** Now you can switch between multiple panes using a tab UI on the right hand side of the app. The UI will only be visible when you have 2+ panes open.


[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

<br/>

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues
### Excel

- <div><span>Users may encounter an error when saving changes while using some non-English character sets</span></div>


- <div><span>Users may encounter an error when accessing a hidden named range</span></div>


- <div><span>Disabling hardware graphics acceleration with 4K resolution may result in delayed rendering of cells when scrolling around.</span></div>


- <div>Clearing a long formula that overlaps a cell boundary may still display across the cell boundary.</div>


- <div>Resolved an issue with ribbon customization not loading when opening embedded workbook.</div>


- <div><span>Margin dropdown menu may not render correctly</span></div>


### Modern Project

- <div>Setting effort on tasks that have no assignments are rounded to 1 day.</div>


### OneNote

- <div><span><span style="display:inline !important;background-color:rgba(255, 255, 255, 1);color:rgba(0, 0, 0, 0.9);">OneNote may not open via the 'Meeting Notes' Outlook add-in.</span><br></span></div>


### Outlook

- <div><font size=2><font style="background-color:rgba(255, 255, 255, 1);">R<span style="display:inline !important;">etention policy labels may display the retention time period in parenthesis.</span></font></font><br></div>


- <div><span><p style="box-sizing:border-box;color:rgba(0, 0, 0, 0.9);margin-bottom:0px;margin-top:0px;">Blank spaces may appear in Contact cards with Japanese language pack.</p><br></span></div>


- <div><span>Images inserted inline to Outlook e-mail messages can sometimes get resized</span></div>


### PowerPoint

- <div><span><span style="font-size:12.0pt;font-family:&quot;Calibri&quot;,sans-serif;color:black;">If a user has two (or more) different videos on a slide in a cloud file, the video images are rendered correctly, but
when the user clicks on each one to play, the video content is the same.</span><br></span></div>


- <div><span>In some cases, scrolling with touch devices will not work</span></div>


- <div><span>Margin dropdown menu may not render correctly</span></div>


- <div>Safelinks from one Office application to another may not launch the linked application</div>


### Project

- <div>Project may crash when you use the Compare Projects feature.</div>


- <div>If you are in Dark mode, when you go to the task inspector panel on a task with an overallocated resource, you are unable to read the table.</div>


### Word

- <div>Saving a file after doing a mail merge may not work under certain conditions.</div>


- <div><span>Building blocks organizer may display an invalid alert: &quot;You have modified styles, building blocks&quot;.</span></div>


- <div><span>Comment pane sometimes gets reloaded when using copy/paste</span></div>


- <div>Comments are sometimes not pasted in the correct order</div>


- <div><span>Applying a template consisting of a multi-level list with custom styles to existing documents may not preserve the style under certain conditions.</span></div>


- <div><span>Resizing a split screen border may introduce an additional split screen.</span></div>


- <div><span>Margin dropdown menu may not render correctly</span></div>


- <div><span>At-mentioning a user in a comment card may show <span style="display:inline !important;font-family:Calibri,Helvetica,sans-serif;font-size:16px;">JSON</span></span></div>


- <div>Safelinks from one Office application to another may not launch the linked application</div>


### Office Suite

- <div><span>For Japanese based products, account user first name, last name may appear in incorrect order.</span></div>


- <div>Hovering a mouse pointer over comments may display a textbox outline around the comment.</div>



[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 1912: November 15
*Version 1912 (Build 12307.20000)*


[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates
### Insights Services

- **Natural Language Queries in Ideas in Excel:** Excel Ideas's new capability to ask a natural language question about your data


[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

<br/>

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues
### Excel

- <div><span>Text to Column functionality may fail for some localizations</span></div>


- <h5>Editing dynamic array formulas within a cell may result in text being aligned outside of the boundary of the cell<br></h5>


### Outlook

- <div>Added the ability to enforce S/MIME configuration via group policy</div>


- <div><span>Embedded images may appear smaller than expected</span></div>


### PowerPoint

- <div><span>Cursor may disappear after moving focus from text</span></div>


### Project

- <div><span>Users may experience an error regarding licensing</span></div>


- <div><span>The Today button in the date picker may sometimes set the incorrect date</span></div>


### Word

- <div><span>Right-clicking can sometimes not result in selecting the whole word</span></div>


- <div>The cursor may remain active within an object after converting to a suggested format</div>


- <div>Images in messages may be incorrectly scaled in some scenarios</div>


- <div>Some themes may make it difficult to determine which comment is selected</div>


- <div><span>Selecting a comment hint should now show the modern comments pane when hidden in pane switcher</span></div>


### Office Suite

- <div><span>Replying to a comment may cause the textbox to expand vertically beyond the edge of the pane</span></div>



[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 1911: October 23
*Version 1911 (Build 12215.20006)*


[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates
### Excel

- **See Your Pen Options When You Pick Up Your Surface Pen:** When you first pick up your Surface Pen in Word, Excel, or PowerPoint, the Draw tab will be activated to make selecting your pen colors easy.

### Outlook

- **.:** .

- **NA:** NA

### PowerPoint

- **See Your Pen Options When You Pick Up Your Surface Pen:** When you first pick up your Surface Pen in Word, Excel, or PowerPoint, the Draw tab will be activated to make selecting your pen colors easy.

### Visio

- **Make polished Visio diagrams in Excel:** Quickly and easily visualize your data into polished Visio diagrams within Excel. [Learn more](https://support.office.com/en-us/article/bee3b5aa-aaaf-4401-acc6-276b711c763c)

### Word

- **See Your Pen Options When You Pick Up Your Surface Pen:** When you first pick up your Surface Pen in Word, Excel, or PowerPoint, the Draw tab will be activated to make selecting your pen colors easy.

- **Coauthoring improvements:** Improved the coauthoring experience by making it more likely that content changes will be received by others in real time.


[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

<br/>

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues
### Access

- <div><span>The record count could be incorrect</span></div>


### Excel

- <div><span>We significantly improved the performance of deleting columns with merged cells</span></div>


- <div>U<span>sers could be prevented from saving in Office 365 Excel Workbook format</span></div>


- <div><span>Checkboxes could not render correctly</span></div>


- <div>C<span>hanges to a chart size could not be saved</span></div>


- <div><span>Select Data Source dialogs were not case sensitive for some fields</span></div>


- <div><span>Some VBA functions would return an error on new chart types</span></div>


### PowerPoint

- <div>C<span>hanges to a chart size could not be saved</span></div>


### Publisher

- <div><span>Shapes could appear outside of the graphics border</span></div>


### Word

- <div><span>Shapes could appear outside of the graphics border</span></div>


- <div><span>Highlighting text could be challenging</span></div>


- <div>A<span> user could be prevented from navigating to an individual item in the editor</span></div>


- <div>G<span>rammar or spelling errors might not be highlighted</span></div>


- <div>C<span>hanges to a chart size could not be saved</span></div>


- <div>A<span> contact card could be prevented from opening after apply formatting to an @ mention</span></div>


- <div><span><p style="margin:0in 0in 0pt;font-family:&quot;Calibri&quot;,sans-serif;font-size:11pt;"><span style="font-size:10.5pt;">Resolved
an issue where users may be unable to save Word, Excel, and PowerPoint
documents.&nbsp; This issue affects users that create a new file and bring up
the &quot;Save as Model Dialog&quot; option after clicking on the Save icon or
pressing Ctrl + S.</span><br></p></span></div>


### Office Suite

- <div><span>Performance issue when using Shapes on Windows 7</span></div>



[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 1911: October 18
*Version 1911 (Build 12209.20010)*


[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates
### Outlook

- **.:** .

- **NA:** NA

- **Send accessible mail to those who need it most:** Outlook will display a mail tip to help you ensure that your content is accessible when sending to a user who prefers accessible content

### PowerPoint

- **Optimize your presentation for all:** Accessibility Checker helps you arrange objects on your slides with screen readers in mind.

### Visio

- **Visio Data Visualizer Excel Add-in:** Quickly and easily visualize your data into polished Visio diagrams within Excel. [Learn more](https://support.office.com/en-us/article/bee3b5aa-aaaf-4401-acc6-276b711c763c)

### Office Suite

- **The Upload Center is being replaced by the Files Needing Attention experience:** The Upload Center is being replaced by the Files Needing Attention experience that will show up inside the Office applications under File > Open. This new experience is more modern, integrated, and less intrusive compared to the Upload Center.


[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

<br/>

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Non-security updates
### Excel

- <div><span>Resolved an issue where check box controls could shrink when using autofit to adjust row height</span></div>


- <div><span>Resolved an issue where selecting a cell after scrolling could result in the wrong cell being selected</span></div>


### OneNote

- <div><span>Identified an issue which could affect syncing from a local resource to a cloud resource</span></div>


### Outlook

- <div><span>Identified an issue which could cause digital signatures to become broken when signing an e-mail with a digitally signed attachment</span></div>


- <div><span>Identified an issue where long filenames were truncated after draging and droping to the message body</span></div>


- <div>Identified an issue where the search box could disappear when the ribbon is set to hide automatically</div>


### PowerPoint

- <div><span></span></div><span style="font-size:11.0pt;font-family:&quot;Calibri&quot;,sans-serif;">Identified an issue
where aspect ratio for the slide preview was not being properly locked/unlocked</span>


### Project

- <div>Identified an issue where notes might not persist if entered while doing update tasks<br></div>


- <div>Identified an issue where a file could be locked by a user, but no username would be displayed in the error message</div>


- <div><span>Identified an issue where users could get several messages when opening a read-only project</span></div>


### Security

- <div><span>Identified an issue where a welcome message contained an invalid link</span></div>


### Word

- <div><span>Identified an issue when viewing comments while using a screen reader</span></div>


- <div><span>Identified an issue where some critiques were misidentified as being spelling or grammar critiques</span></div>


- <div><span>Identified an issue where a new comment dialog could sometimes not obtain focus</span></div>


### Office Suite

- <div><span>We fixed an issue where an upgrade could be prevented by a incorrect error message of &quot;Another install in progress&quot;</span></div>



[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 1910: October 11
*Version 1910 (Build 12130.20112)*


[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates
### Outlook

- **.:** .

- **NA:** NA

- **Send accessible mail to those who need it most:** Outlook will display a mail tip to help you ensure that your content is accessible when sending to a user who prefers accessible content

### Visio

- **Visio Data Visualizer Excel Add-in:** Quickly and easily visualize your data into polished Visio diagrams within Excel. [Learn more](https://support.office.com/en-us/article/bee3b5aa-aaaf-4401-acc6-276b711c763c)


[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

<br/>

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Non-security updates
### Excel

- <div>Resolved an issue in inserting files as object from OneDrive</div>


- <div>Resolved an issue where the hyperlink format could not be properly applied to some content</div>


- <div>Resolved an issue where formulas containing structured absolute references may be adjusted incorrectly</div>


- <div>Resolved an issue where workbooks created in earlier versions of Office could cause Excel to hang when opened in current versions of Office</div>


- <div>Resolved an issue where content copied from Excel could appear incorrect when pasted into other Office applications</div>


- <div>Fix to correct colors used in previews when inserting charts using chart templates</div>


### PowerPoint

- <div>Identified an issue where ARC Devices could lose connection when running under AirSpace WinComp</div>


### Word

- <div>Resolved an issue in inserting files as object from OneDrive</div>


- <div>Improved our recovery steps to f<span>ix an issue that caused graphical content to get deleted from email threads.&nbsp;</span></div>



[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 1910: October 04
*Version 1910 (Build 12126.20000)*


[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates
### Excel

- **Data visualizer add-in:** Quickly create Visio flowcharts from Excel. [Learn more](https://support.office.com/en-us/article/bee3b5aa-aaaf-4401-acc6-276b711c763c)

### Outlook

- **.:** .

### Visio

- **Visio Data Visualizer Excel Add-in:** Quickly and easily visualize your data into polished Visio diagrams within Excel. [Learn more](https://support.office.com/en-us/article/bee3b5aa-aaaf-4401-acc6-276b711c763c)


[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

<br/>

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Non-security updates
### Excel

- <div><span>We fixed an issue where the print area in print preview was not correct</span></div>


- <div><span>We fixed an issue which could have prevented pivot tables from being refreshed during a co-authoring session</span></div>


### OPX

- <div>We fixed an issue where users could receive an &quot;inconsistent state&quot; error when using a shared database.</div>


### Outlook

- <div><span>We fixed an issue which could sometimes result in a crash when a user receives a 'Missed Conversation' message from Skype</span></div>


- <div><span>We fixed an issue which could have caused duplication of mail folders</span></div>


- <div><span>We fixed an issue which could have caused an incorrect error message when attempting to send s/MIME encrypted e-mail</span></div>


- <div><span>We fixed an issue which could have resulted in a memory leak</span></div>


- <div><span>We fixed an issue which could have caused the senders name to change when saved as a draft</span></div>


### PowerPoint

- <div><span></span></div>We fixed an issue which could cause TextRanges to become broken after pasting text into the header/footer/slide number placeholders on slide master and slide layout


- <div><span></span></div>We fixed an issue which prevented hyperlink from being created when pasting text with hyperlink


- <div>We fixed an issue which would prevent statistics from working correctly</div>


### Word

- <div><span>We fixed an issue where font colors were not being committed</span></div>


### Office Suite

- <div>We fixed an issue which could offer version history when that feature was disabled</div>



[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 1910: September 27
*Version 1910 (Build 12119.20000)*


[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates
### Outlook

- **.:** .


[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

<br/>

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Non-security updates
### Excel

- <div><span>We fixed an issue which could have caused scatter line charts from rendering properly when changing the series collection</span></div>


### Outlook

- <div><span>We fixed an issue which could have reported permission errors when interacting with shared calendar folders</span></div>


- <div><span>We fixed an issue which could have prevented users from adding attachments to calendars</span></div>


- <div><span>We fixed an issue which caused error messages to display when replying to a digitally signed message</span></div>


### Word

- <div><span>We fixed an issue which could have caused scaling issues when printing to Deskjet printers</span></div>


### Office Suite

- <div><span>We fixed an issue where medium bold text could be incorrectly styled</span></div>


- <div><span>We fixed an issue where a user could be given an incorrect error message when closing a file with a pending upload</span></div>



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

## May 17, 2019
Version 1906 (build 11708.20006)

## What's New:

### Outlook

#### User Experience Updates

Updates that have been in Coming Soon are now here, featuring the Simplified Ribbon, and a visual refresh of the folder pane, message list, and reading pane.

##### Getting Started:

These change will be part of the new default UI; it has been available behind the Coming Soon switch since mid Dec for 100% prod

#### Customizable Simplified Ribbon

Easily customizable to switch between classic and Simplified views and pin/unpin commands.

##### Getting Started:

Users can get to the simplified ribbon by turning on Coming Soon (initially) and clicking the chevron in the ribbon to toggle between the classic multi-line ribbon and the new simplified single-line ribbon.

##### Scenarios to Try:

Switch from Classic ribbon to Simplified ribbon

#### Pick your favorite action

Don't use Flag and Delete? How about Archive or Mark as Read? Customize the quick action menu with the commands you use most.

##### Getting Started:

To select your Quick Actions, right click on an email in the message list to bring up the Context Menu. Then click "Set Quick Actions..."

##### Scenarios to Try:

Change the defaults from flag and delete to either archive, move,  mark as read, or none for a cleaner message list

#### Relaxed or tighter layout? You choose

Use Tighter Spacing lets you decide if you want more space between items, or a tighter layout to see more.

##### Getting Started:

View tab, use tighter spacing checkbox - in Messages group for classic ribbon, Current View settings for simplified ribbon

##### Scenarios to Try:

Use Outlook to triage and write email with and without the setting enabled. With Use tighter spacing on, more messages fit per page, and controls on the compose forms are more streamlined.

#### Dedupe MRU entries when using the Onedrive sync client

Enable better integration with onedrive sync client with cloud attachments by deduping the mru entries and to enable faster attach as copy behavior for synchronized data

##### Getting Started:

If you use the OneDrive sync client, you will no longer see file duplicates in the Attach File MRU.

##### Scenarios to Try:

Enable the OneDrive sync client and use the Attach File menu in Outlook Desktop

#### Improved shared folder synchronization for mailboxes with many folders

For years Outlook has been limited to a maximum of 500 folders when synchronizing shared mailboxes. With this change Outlook has been improved to sync in a way that will no longer encounter this 500 folder limit.

##### Getting Started:

Create 1000 folders in a mailbox, give someone else access to the mailbox, create an Outlook profile for the "someone else" and verify that sync works.

### Word

#### Erase just a little bit

##### Getting Started:

Go to the Draw Tab. 
Select the Eraser dropdown. 
Choose Small Eraser or Medium Eraser.

##### Scenarios to Try:

Go to the Draw Tab. 
Select a pen. 
Draw an ink stroke. 
Select the Eraser dropdown. 
Choose Small Eraser or Medium Eraser. 
Erase just bits of the ink stroke.

## Notable Fixes:

### All 
- We fixed an issue which could prevent some users from saving as PDF
- We fixed an issue which could impact users saving large files on a 32-bit system

### Word 
- We significantly improved the responsiveness of the dictation feature

### Excel
- We fixed an issue where double-click events could fail on touch screen devices
- We fixed an issue which could prevent some users from being able to edit VBA macros
- We fixed an issue which could impact performance when using slicers

### PowerPoint
- Various performance and stability fixes

### Outlook
- We fixed an issue where the wrong template could be displayed from what was selected

### Access
- We fixed an issue where using the zoom builder to display long rich text, could be hard to read

### Project
- Various performance and stability fixes

</BR></BR>

## May 10, 2019
Version 1906 (build 11702.20000)

## Notable Fixes:

### All
- We fixed an issue where the Save As dialog could display the incorrect path

### Word 
- We fixed an issue where some selections from Tell Me would not get inserted

### Excel
- Various performance and stability fixes

### PowerPoint
- Various performance and stability fixes

### Outlook
- Various performance and stability fixes

### Access
- Various performance and stability fixes

### Project
- We fixed an issue where Task ID's could require highlighting to see

</BR></BR>

## May 3, 2019
Version 1906 (build 11629.20008)

## Notable Fixes:

### All
- We fixed an issue where some users would experience problems syncing with OneDrive for Business

### Word 
- We fixed an issue where in some cases Word would take a long time to start

### Excel
- We fixed an issue where external links were sometimes removed from workbooks after upgrading to a newer version of Excel
- We fixed an issue where some users could experience difficulty selecting cells in a new workbook

### PowerPoint
- We fixed an issue where font sizes were not consistant when converting drawings to text

### Outlook
- We fixed an issue where saving a contact from a .VCF file could result in empty fields
- We fixed an issue where a message could get stuck in the outbox folder even though it had been sent
- We fixed an issue where Outlook could crash when viewing a DRM message

### Access
- Various performance and stability fixes

### Project
- We fixed an issue where the editor would switch from Chinese to English
- We fixed an issue where unpublished tasks could appear in the published copy of a master project

</BR></BR>

## April 26, 2019
Version 1905 (build 11617.20002)

## Notable Fixes:

### Word 
- Various performance and stability fixes

### Excel
- We fixed an issue where Solver macros would fail to run
- We fixed an issue which could prevent Excel files from being imported into SharePoint

### PowerPoint
- Various performance and stability fixes

### Outlook
- Various performance and stability fixes

### Access
- Various performance and stability fixes

### Project
- Various performance and stability fixes

</BR></BR>

## April 19, 2019
Version 1905 (build 11609.20002)

## What's New:

### Excel

#### Improved Filled Maps experience using Data Types

This feature is an improvement for users who plot Filled Map Charts using Excel's Geographic Data Types. The benefit to the end users will be richer integration between the features and better accuracy of the region the end user wants to map. Additional benefits include - ability to map city polygons.

##### Getting Started:

- This feature is an improvement to the existing features within Excel. To use the improvement - convert locations into Rich Entities and plot with Filled Maps. 

##### Scenarios to Try:

- Users can try mapping cities, states, counties, countries and zip codes. 


## Notable Fixes:

### All Applications
- We fixed an issue where the First Run dialog would display whenever an application was launched
- We fixed an issue where a SharePoint link in the "save as" dialog could be missing.
- We fixed an issue where users would incorrectly see a "Repair Now" dialog

### Word 
- We fixed an issue where some users could receive an error for insufficient memory or disk space when requesting a font
- We fixed an issue where a window could lose focus when switching from the comments pane

### Excel
- Various performance and stability fixes

### PowerPoint
- We fixed an issue preventing the resizing of branded shapes
- We fixed an issue where PowerPoint could crash when opening a file in protected view mode

### Outlook
- We fixed an issue which prevented some users from selecting Chinese words
- We fixed an issue where expiry dates were not calculated correctly

### Access
- We fixed an issue which prevented some users from using the Macro Builder
- We fixed an issue where printing a report would only print the first page
- We fixed an issue where the application could crash when hovering over a hyperlink
- We fixed an issue which caused some items to appear off screen when using relationships view

### Project
- Various performance and stability fixes

</BR></BR>

## April 12, 2019
Version 1905 (build 11601.20042)

## What's New:

### Access

#### New in Access - Data Connector to Microsoft Graph

Link to or import form Microsoft Graph services to build applications that can leverage the smart contextual data stored in the Graph

#### Getting Started:

External Data tab in the ribbon, click on New Data Sources and find the new Graph connector in the Online Services menu

#### Scenarios to Try:

Import from or link to various Graph services including People, Groups and OneDrive Items.

## Notable Fixes:

### All Applications
 - We fixed an issue which prevented some users from saving files to cloud locations
 - We fixed an issue where the wrong pane could open from the ribbon

### Word 
- Various performance and stability fixes

### Excel
- We fixed an issue where users would see an error message for linked data types when the workbook did not contain linked data types
- We fixed an issue where URL links within a Word document could change when viewed locally vs. online

### PowerPoint
- We fixed an issue where the application could crash after undoing changes from the animations tab

### Outlook
- We fixed an issue which prevented some users from modifying the Notes field for contacts in a Public Folder
- We fixed an issue where a conflict could occur between expiration dates and deletion dates

### Access
- Various performance and stability fixes

### Project
- Various performance and stability fixes

</BR></BR>

## April 5, 2019
Version 1904 (build 11527.20014)

## What's New:

### Outlook

#### Outlook for Windows:  set and share your Focused Inbox settings

Your Focused Inbox preferences are stored in the cloud so you can enjoy the same consistent experience when using Outlook for Windows and Outlook on the web on any computer.

#### Getting Started:

Under File > Options > General tab, there is a new preference for 'Store my Outlook settings in the cloud'. Users will need to check the box to enable their Focused Inbox setting to roam to other Desktop Outlook installations and OWA.

#### Scenarios to Try:

Change Focused Inbox on the machine that has cloud settings preference turned on. Navigate to OWA and see the preference applied there as well. Change Focused Inbox in OWA and start Desktop Outlook to see the preference reflected.

### Word

#### Learning Tools mode has additional support for more page colors

Learning Tools in Word adds support for more page theme colors, which allows the changing of the background color of the page.  Many people have challenges reading with an all-white or all-black background, so we’ve expanded the choice of colors in Word on PC and Mac.

#### Getting Started:

To try this out, go to the View tab and choose Learning Tools, and then Page Color.

#### Scenarios to Try:

To try this out, go to the View tab and choose Learning Tools, and then Page Color.

### Excel

#### Elevate Creativity with Animated 3D Models

Office now supports animated models, which will playback in the editor so you can bring your sheets to life!

#### Getting Started:

1. Open Excel
2. Insert an animated 3D Model (coming to Remix soon, but for now, access animated models here: \\osan\ogx\Public\TestFiles\3D Models\Animated3D\C3Art)
3. The animated model will play in the editor! Check Slideshow mode - it will play there too!
4. In the 3D Format Ribbon, explore more animation scenes in the model

#### Scenarios to Try:

1. Insert an animated model and watch it play in the editor
2. Explore the animation scenes available in the animated model via the Scenes Gallery, available in the 3D Format Ribbon
3. Easily play/pause the animation via the ribbon, floatie or space bar

## Notable Fixes:

### All Applications
- We fixed an issue where the incorrect app icon could appear for Excel in contextual menus
- We fixed an issue where the File Menu button could disappear after installing an update
- We fixed an issue which could change your user license

### Word 
- We fixed an issue where text would not render correctly at certain zoom levels

### Excel
- We fixed an issue where users would not be prompted to save a workbook after making edits
- We fixed an issue where a BeforeSave event would not be triggered if the user shared the workbook.
- We fixed an issue where resizing a column to fewer than 6 pixels could throw an incorrect error message.
- We fixed an issue where Excel would ignore the Application.Visible flag
- We fixed an issue where trace arrows would remain on non-active frozen panes
- We fixed an issue where cell formatting of dates an currency could change when opening a workbook
- We fixed an issue where tooltips would move unexpectedly
- We fixed localization issues for the Power Query editor
- We fixed an issue where a workbook would be removed as an attachment when sending via e-mail

### PowerPoint
- We fixed an issue where copying shapes would take longer than expected

### Outlook
- We fixed an issue where Outlook could crash while using the drawing tool
- We fixed a localization issue when composing html e-mails
- We fixed an issue where the user would have difficulty in selecting the lower pane

### Access
- Various performance and stability fixes

### Project
- Various performance and stability fixes

</BR></BR>

## March 22, 2019
Version 1904 (build 11514.20004)

## Notable Fixes:

### Word 
- We fixed an issue where the UI would constantly display "Checking for Changes"

### Excel
- We fixed an issue where the application could crash after moving a worksheet
- We fixed an issue where the application could crash after saving as a PDF
- We fixed an issue where the save dialog would not accept some Korean characters

### PowerPoint
- Various performance and stability fixes

### Outlook
- Various performance and stability fixes

### Access
- We fixed the error message in Access where an extra shortcut to Access was created
- We fixed an issue where data from a linked SharePoint would display incorrectly

### Project
- We fixed an issue where the language settings would switch from Chinese to English
- We fixed an issue where the application could crash when synching to SharePoint

</BR></BR>

## March 15, 2019
Version 1904 (build 11504.20000)

## What's New:

### Word

#### Focus Mode

Switch to Focus on the View menu to remove distractions and concentrate on your work. For Office 365 subscribers only.

#### Getting Started:

View tab "Focus" Button in the Ribbon
Status Bar "Focus" Button

#### Scenarios to Try:

Enter Focus Mode and experience the Focused Experience

## Notable Fixes:

### Word 
- We fixed an issue where images in a document saved as a PDF would have the incorrect DPI

### Excel
- Various performance and stability fixes

### PowerPoint
- We fixed an issue where the comments pane would not open or close properly
- We fixed an issue where the application could crash when deleting a video
- We fixed an issue where in some instances the application would fail to launch

### Outlook
- We fixed an issue where read receipts were incorrect when viewed in Japanese

### Access
- Various performance and stability fixes

### Project
- Various performance and stability fixes

</BR></BR>

## March 8, 2019 
Version 1903 (build 11425.20036)

## What's New:

### Word

### Find What You're Looking For with Microsoft Search

With Microsoft Search, you can find all the files, actions, people, and help you need to get work done.

#### Getting Started:

- The feature is prominently displayed on top of the UI in the header.

#### Scenarios to Try:

- Search for a college, a recent document or search for the ribbon commands you use most often
- Look up a topic or subject to get more information on it
- 
#### CoAuthoring

Tired of being locked out of your document with macros? Now your docm files on OneDrive for Business allow simultaneous editing by multiple authors.

#### Getting Started:

The user doesn't need to press any buttons in the UI to access this feature. It is enabled by default on OneDrive for Business docm files.
So, the user should save a docm file to OneDrive for Business to try it out.

#### Scenarios to Try:

Create a docm file on OneDrive for Business, share it with your colleagues, and collaborate!

## Notable Fixes:

### Word 
- We fixed a crashing issue that occurred when pressing ‘ESC’ while in Options
- We fixed a crashing issue that occurred when replying to comments
- We fixed an issue with copy & paste from Word to PowerPoint Online

### Excel
- We fixed an issue where copying a cell in Excel caused high CPU usage when protected document and editable document were opened

### PowerPoint
- Various performance and stability fixes

### Outlook
- We fixed an issue where Outlook Search was not honoring the selected chronological sorting
- We fixed an issue where the "Open this task" workflow ribbon button was unresponsive for certain emails
- We fixed an issue where Outlook did not clear on premise rooms after users selected an available room in Room Finder

### Access
- We fixed the saved import/export dialog that had white text on white background in Dark Theme
- We fixed an issue where users could not set the DisplayControl property for a Yes/No field to Textbox in table design

### Project
- Various performance and stability fixes


## March 1, 2019 
Version 1903 (build 11414.20014)

## What's New

### Word

#### Colors for Track Changes, Comments and Real-Time Collaboration in Sync

Fixes in our product now ensure that the comments, track changes and the cursor for a collaborator show up in the same color.

#### Getting Started:

Open a SharePoint or OneDrive document that others have open. Verify that track changes and comments color for a user matches the color of that user's cursor.

#### Scenarios to Try:

Open a SharePoint or OneDrive document that others have open. Verify that track changes and comments color for a user matches the color of that user's cursor.

## Notable Fixes:

### Word 
- We fixed a crashing issue that occurred when pressing ‘ESC’ while in Options
- We fixed an issue with copy & paste from Word to PowerPoint Online

### Excel
- We fixed an issue where copying a cell in Excel caused high CPU usage when protected document and editable document were opened

### PowerPoint
- We fixed an issue with slide image size when using @Mentions in PowerPoint

### Outlook
- We fixed an issue where Outlook Search was not honoring the selected chronological sorting
- We fixed an issue where the "Open this task" workflow ribbon button was unresponsive for certain emails
- We fixed an issue where Outlook did not clear on premise rooms after users selected an available room in Room Finder

### Access
- We updated the prompt text that showed when confirming the relinking tables with a datasource
- We fixed the saved import/export dialog that had white text on white background in Dark Theme
- We fixed an issue where users could not set the Display Control property for a Yes/No field to Textbox in table design

### Project
- Various performance and stability fixes

</BR></BR>



## February 15, 2019 
Version 1903 (build 11310.20016)

## What's New:

### PowerPoint


### Morph Transition Enhancements - Morph by Name

Specify the shapes you want to morph

#### Getting Started:

- To get Morph to treat two objects as the same object, the user can rename the shapes using the Selection Pane.
- The name must be prefaced with “!!” (two exclamation points) for Morph to use it to override our default matching behavior, e.g. “!!Name”
- Users can continue to rename shapes with any name that doesn’t start with “!!” without worrying that it will change the way Morph works

#### Scenarios to Try:

- Insert a Shape in a slide, let's say Rectangle
- Create a new slide
- Insert a different shape in the 2nd slide, let's say Triangle
- Open SelectionPane, rename the Rectangle in slide 1 to "!!shape", and rename the Triangle in slide 2 to "!!shape"
- Apply Morph on the 2nd slide

</BR>

### Morph Transition Enhancements - SmartArt

SmartArt morph with smoother transitions

#### Getting Started:

Use Morph the same way you would with SmartArt

#### Scenarios to Try:

- Insert a SmartArt in a slide
- Duplicate the Slide
- Resize/Change/Move the SmartArt on the duplicated slide
- Apply Morph on the duplicated slide

</BR>

### Morph Transition Enhancements - Tables

Tables morph with smoother transitions

#### Getting Started:
Use Morph the same way you would with tables

#### Scenarios to Try:

- Insert a Table in a slide
- Duplicate the slide
- Resize/Change/Move the Table on the duplicated slide
- Apply Morph on the duplicated slide

</BR>

### Word, Excel, PowerPoint, OneNote, Access, Project, Publisher & Visio

### Seamlessly Switch Between Accounts

The new account manager shows all of your work and personal accounts in one place, and puts you in control of switching between them. This updated experience makes it clear how you're logged in, and now you can toggle between work and personal accounts without having to sign out first or deal with complex dialogs.


![MeMock.png](Images/MeMock.png)

#### Scenarios to Try:
- Switch between accounts
- Add a new account [Note: you may want to first go to File | Account | Connected Services and remove any personal services connected to work accounts or vice versa]
- Sign out from an account
</BR>

## Notable Fixes:

### Word 
- We fixed an issue with context preview for tables & images

### Excel
- We fixed an issue where text in autofilter Search field is white in Black theme
- We fixed a consent UI issue with New Office Add-in

### PowerPoint
- We fixed an issue with automatically extending display when presenting SlideShows on laptops or tablets.

### Outlook
- We fixed an issue with the Send to OneNote button display

### Access
- Various performance and stability fixes

### Project
- Various performance and stability fixes


</BR></BR>
## February 11, 2019
Version 1903 (build 11330.20014)


## Notable Fixes:

### Word 
- We fixed an issue where some customized styles could not be applied to word online
- We fixed Context Preview issues with rich objects in Word
- We fixed an issue where pasting lists  would result in Word crashing

### Excel
- We fixed an issue where appended spaces after number formats are no longer showing when there is no currency symbol
- We fixed an issue with auto detect for stocks

### PowerPoint
- Various performance and stability fixes

### Outlook
- Various performance and stability fixes

### Access
- Various performance and stability fixes

### Project
- Various performance and stability fixes

</BR></BR>


## February 1, 2019 
Version 1902 (build 11326.20000)


## Notable Fixes

### Word 
- We fixed an issue with resizing cells in an embedded Excel table
- We fixed an issue with copy/paste of shapes in a Drawing Canvas

### Excel
- We fixed an issue with opening files from the Excel Web app
- We fixed an issue where saving a CSV file as .XLSX was failing due to file name size
- We fixed the context menu to display the context menu options

### PowerPoint
- We fixed an issued where users were unable to use the keyboard shortcut ctrl+alt+7/ctrl+alt+8 to enter square brackets
- We fixed an issue where inserting a local video into the PPT would reduce the ‘C’ drive disk space
- We fixed the Publish to Microsoft Stream button which was not displaying to some users

### Outlook
- We fixed an issue where the task view in calendar was  not correctly showing the task subject

### Access
- We fixed a scaling issue with charts

### Project
- Various performance and stability fixes
