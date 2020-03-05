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

## Version 2002: March 04
*Version 2002 (Build 12527.20242)*


[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues
### Access

- <div><span><div style="box-sizing:border-box;"><span style="box-sizing:border-box;"><span style="box-sizing:border-box;">This update fixes an issue in Microsoft Visual Basic for Applications (VBA) that can allow modification of certain parts of a VBA project without invalidating the digital signature on the project. These modifications can be constructed in such a way to allow an attacker to entice an end user who opens the file and runs the embedded VBA project to execute code different than that which was signed.<br style="box-sizing:border-box;"></span></span></div><div style="box-sizing:border-box;"><span style="box-sizing:border-box;"><div style="box-sizing:border-box;">​</div><span style="box-sizing:border-box;">This fix mitigates the issue in two ways. First, it ensures that the data read during digital signing is the same as the data read during project load. Second, it prevents VBA from loading external project references (i.e. type libraries) by path from untrusted internet locations if they are not found registered on the local machine. It also prevents VBA from loading external references by path from untrusted intranet locations unless administrators explicitly enable the functionality via group policy. This can be done by setting the HKCU\Software\Policies\Microsoft\VBA\Security\AllowIntranetReferences DWORD registry value to 1. This setting should only be enabled if necessary to allow for solution migration for critical solutions adversely impacted by this update, after which it should be disabled again for added security. The recommended alternative pattern for introducing external component dependencies into VBA projects is to ensure the references are located in secured locations, and to register only those that are required on the individual users' machines under HKCR. Alternatively, if a well secured location for such external references is established and registering the components on users' machines is not possible, this secure location can be added to the internet trusted sites list.</span></span></div><br></span></div>


### Excel

- <div><span><div style="box-sizing:border-box;"><span style="box-sizing:border-box;"><span style="box-sizing:border-box;">This update fixes an issue in Microsoft Visual Basic for Applications (VBA) that can allow modification of certain parts of a VBA project without invalidating the digital signature on the project. These modifications can be constructed in such a way to allow an attacker to entice an end user who opens the file and runs the embedded VBA project to execute code different than that which was signed.<br style="box-sizing:border-box;"></span></span></div><div style="box-sizing:border-box;"><span style="box-sizing:border-box;"><div style="box-sizing:border-box;">​</div><span style="box-sizing:border-box;">This fix mitigates the issue in two ways. First, it ensures that the data read during digital signing is the same as the data read during project load. Second, it prevents VBA from loading external project references (i.e. type libraries) by path from untrusted internet locations if they are not found registered on the local machine. It also prevents VBA from loading external references by path from untrusted intranet locations unless administrators explicitly enable the functionality via group policy. This can be done by setting the HKCU\Software\Policies\Microsoft\VBA\Security\AllowIntranetReferences DWORD registry value to 1. This setting should only be enabled if necessary to allow for solution migration for critical solutions adversely impacted by this update, after which it should be disabled again for added security. The recommended alternative pattern for introducing external component dependencies into VBA projects is to ensure the references are located in secured locations, and to register only those that are required on the individual users' machines under HKCR. Alternatively, if a well secured location for such external references is established and registering the components on users' machines is not possible, this secure location can be added to the internet trusted sites list.</span></span></div><br></span></div>


### OneNote

- <div><span><div style="box-sizing:border-box;"><span style="box-sizing:border-box;"><span style="box-sizing:border-box;">This update fixes an issue in Microsoft Visual Basic for Applications (VBA) that can allow modification of certain parts of a VBA project without invalidating the digital signature on the project. These modifications can be constructed in such a way to allow an attacker to entice an end user who opens the file and runs the embedded VBA project to execute code different than that which was signed.<br style="box-sizing:border-box;"></span></span></div><div style="box-sizing:border-box;"><span style="box-sizing:border-box;"><div style="box-sizing:border-box;">​</div><span style="box-sizing:border-box;">This fix mitigates the issue in two ways. First, it ensures that the data read during digital signing is the same as the data read during project load. Second, it prevents VBA from loading external project references (i.e. type libraries) by path from untrusted internet locations if they are not found registered on the local machine. It also prevents VBA from loading external references by path from untrusted intranet locations unless administrators explicitly enable the functionality via group policy. This can be done by setting the HKCU\Software\Policies\Microsoft\VBA\Security\AllowIntranetReferences DWORD registry value to 1. This setting should only be enabled if necessary to allow for solution migration for critical solutions adversely impacted by this update, after which it should be disabled again for added security. The recommended alternative pattern for introducing external component dependencies into VBA projects is to ensure the references are located in secured locations, and to register only those that are required on the individual users' machines under HKCR. Alternatively, if a well secured location for such external references is established and registering the components on users' machines is not possible, this secure location can be added to the internet trusted sites list.</span></span></div><br></span></div>


### Outlook

- <div><span><div style="box-sizing:border-box;"><span style="box-sizing:border-box;"><span style="box-sizing:border-box;">This update fixes an issue in Microsoft Visual Basic for Applications (VBA) that can allow modification of certain parts of a VBA project without invalidating the digital signature on the project. These modifications can be constructed in such a way to allow an attacker to entice an end user who opens the file and runs the embedded VBA project to execute code different than that which was signed.<br style="box-sizing:border-box;"></span></span></div><div style="box-sizing:border-box;"><span style="box-sizing:border-box;"><div style="box-sizing:border-box;">​</div><span style="box-sizing:border-box;">This fix mitigates the issue in two ways. First, it ensures that the data read during digital signing is the same as the data read during project load. Second, it prevents VBA from loading external project references (i.e. type libraries) by path from untrusted internet locations if they are not found registered on the local machine. It also prevents VBA from loading external references by path from untrusted intranet locations unless administrators explicitly enable the functionality via group policy. This can be done by setting the HKCU\Software\Policies\Microsoft\VBA\Security\AllowIntranetReferences DWORD registry value to 1. This setting should only be enabled if necessary to allow for solution migration for critical solutions adversely impacted by this update, after which it should be disabled again for added security. The recommended alternative pattern for introducing external component dependencies into VBA projects is to ensure the references are located in secured locations, and to register only those that are required on the individual users' machines under HKCR. Alternatively, if a well secured location for such external references is established and registering the components on users' machines is not possible, this secure location can be added to the internet trusted sites list.</span></span></div><br></span></div>


- <div>Addresses an issue that caused third party applications to be unable to send email.</div>


### PowerPoint

- <div><span><div style="box-sizing:border-box;"><span style="box-sizing:border-box;"><span style="box-sizing:border-box;">This update fixes an issue in Microsoft Visual Basic for Applications (VBA) that can allow modification of certain parts of a VBA project without invalidating the digital signature on the project. These modifications can be constructed in such a way to allow an attacker to entice an end user who opens the file and runs the embedded VBA project to execute code different than that which was signed.<br style="box-sizing:border-box;"></span></span></div><div style="box-sizing:border-box;"><span style="box-sizing:border-box;"><div style="box-sizing:border-box;">​</div><span style="box-sizing:border-box;">This fix mitigates the issue in two ways. First, it ensures that the data read during digital signing is the same as the data read during project load. Second, it prevents VBA from loading external project references (i.e. type libraries) by path from untrusted internet locations if they are not found registered on the local machine. It also prevents VBA from loading external references by path from untrusted intranet locations unless administrators explicitly enable the functionality via group policy. This can be done by setting the HKCU\Software\Policies\Microsoft\VBA\Security\AllowIntranetReferences DWORD registry value to 1. This setting should only be enabled if necessary to allow for solution migration for critical solutions adversely impacted by this update, after which it should be disabled again for added security. The recommended alternative pattern for introducing external component dependencies into VBA projects is to ensure the references are located in secured locations, and to register only those that are required on the individual users' machines under HKCR. Alternatively, if a well secured location for such external references is established and registering the components on users' machines is not possible, this secure location can be added to the internet trusted sites list.</span></span></div><br></span></div>


### Project

- <div><span><div style="box-sizing:border-box;"><span style="box-sizing:border-box;"><span style="box-sizing:border-box;">This update fixes an issue in Microsoft Visual Basic for Applications (VBA) that can allow modification of certain parts of a VBA project without invalidating the digital signature on the project. These modifications can be constructed in such a way to allow an attacker to entice an end user who opens the file and runs the embedded VBA project to execute code different than that which was signed.<br style="box-sizing:border-box;"></span></span></div><div style="box-sizing:border-box;"><span style="box-sizing:border-box;"><div style="box-sizing:border-box;">​</div><span style="box-sizing:border-box;">This fix mitigates the issue in two ways. First, it ensures that the data read during digital signing is the same as the data read during project load. Second, it prevents VBA from loading external project references (i.e. type libraries) by path from untrusted internet locations if they are not found registered on the local machine. It also prevents VBA from loading external references by path from untrusted intranet locations unless administrators explicitly enable the functionality via group policy. This can be done by setting the HKCU\Software\Policies\Microsoft\VBA\Security\AllowIntranetReferences DWORD registry value to 1. This setting should only be enabled if necessary to allow for solution migration for critical solutions adversely impacted by this update, after which it should be disabled again for added security. The recommended alternative pattern for introducing external component dependencies into VBA projects is to ensure the references are located in secured locations, and to register only those that are required on the individual users' machines under HKCR. Alternatively, if a well secured location for such external references is established and registering the components on users' machines is not possible, this secure location can be added to the internet trusted sites list.</span></span></div><br></span></div>


- <div><span style="display:inline !important;">Fixed an issue where summary task dates weren't always getting calculated correctly.</span><br></div>


### Publisher

- <div><span><div style="box-sizing:border-box;"><span style="box-sizing:border-box;"><span style="box-sizing:border-box;">This update fixes an issue in Microsoft Visual Basic for Applications (VBA) that can allow modification of certain parts of a VBA project without invalidating the digital signature on the project. These modifications can be constructed in such a way to allow an attacker to entice an end user who opens the file and runs the embedded VBA project to execute code different than that which was signed.<br style="box-sizing:border-box;"></span></span></div><div style="box-sizing:border-box;"><span style="box-sizing:border-box;"><div style="box-sizing:border-box;">​</div><span style="box-sizing:border-box;">This fix mitigates the issue in two ways. First, it ensures that the data read during digital signing is the same as the data read during project load. Second, it prevents VBA from loading external project references (i.e. type libraries) by path from untrusted internet locations if they are not found registered on the local machine. It also prevents VBA from loading external references by path from untrusted intranet locations unless administrators explicitly enable the functionality via group policy. This can be done by setting the HKCU\Software\Policies\Microsoft\VBA\Security\AllowIntranetReferences DWORD registry value to 1. This setting should only be enabled if necessary to allow for solution migration for critical solutions adversely impacted by this update, after which it should be disabled again for added security. The recommended alternative pattern for introducing external component dependencies into VBA projects is to ensure the references are located in secured locations, and to register only those that are required on the individual users' machines under HKCR. Alternatively, if a well secured location for such external references is established and registering the components on users' machines is not possible, this secure location can be added to the internet trusted sites list.</span></span></div><br></span></div>


### Security

- <div>Fixes an issue when multiple documents are open in Word/Excel/PowerPoint from the same SharePoint library, only the first document opened will be scanned for Policy compliance.</div>


### Visio

- <div><span><div style="box-sizing:border-box;"><span style="box-sizing:border-box;"><span style="box-sizing:border-box;">This update fixes an issue in Microsoft Visual Basic for Applications (VBA) that can allow modification of certain parts of a VBA project without invalidating the digital signature on the project. These modifications can be constructed in such a way to allow an attacker to entice an end user who opens the file and runs the embedded VBA project to execute code different than that which was signed.<br style="box-sizing:border-box;"></span></span></div><div style="box-sizing:border-box;"><span style="box-sizing:border-box;"><div style="box-sizing:border-box;">​</div><span style="box-sizing:border-box;">This fix mitigates the issue in two ways. First, it ensures that the data read during digital signing is the same as the data read during project load. Second, it prevents VBA from loading external project references (i.e. type libraries) by path from untrusted internet locations if they are not found registered on the local machine. It also prevents VBA from loading external references by path from untrusted intranet locations unless administrators explicitly enable the functionality via group policy. This can be done by setting the HKCU\Software\Policies\Microsoft\VBA\Security\AllowIntranetReferences DWORD registry value to 1. This setting should only be enabled if necessary to allow for solution migration for critical solutions adversely impacted by this update, after which it should be disabled again for added security. The recommended alternative pattern for introducing external component dependencies into VBA projects is to ensure the references are located in secured locations, and to register only those that are required on the individual users' machines under HKCR. Alternatively, if a well secured location for such external references is established and registering the components on users' machines is not possible, this secure location can be added to the internet trusted sites list.</span></span></div><br></span></div>


### Word

- <div><span><div style="box-sizing:border-box;"><span style="box-sizing:border-box;"><span style="box-sizing:border-box;">This update fixes an issue in Microsoft Visual Basic for Applications (VBA) that can allow modification of certain parts of a VBA project without invalidating the digital signature on the project. These modifications can be constructed in such a way to allow an attacker to entice an end user who opens the file and runs the embedded VBA project to execute code different than that which was signed.<br style="box-sizing:border-box;"></span></span></div><div style="box-sizing:border-box;"><span style="box-sizing:border-box;"><div style="box-sizing:border-box;">​</div><span style="box-sizing:border-box;">This fix mitigates the issue in two ways. First, it ensures that the data read during digital signing is the same as the data read during project load. Second, it prevents VBA from loading external project references (i.e. type libraries) by path from untrusted internet locations if they are not found registered on the local machine. It also prevents VBA from loading external references by path from untrusted intranet locations unless administrators explicitly enable the functionality via group policy. This can be done by setting the HKCU\Software\Policies\Microsoft\VBA\Security\AllowIntranetReferences DWORD registry value to 1. This setting should only be enabled if necessary to allow for solution migration for critical solutions adversely impacted by this update, after which it should be disabled again for added security. The recommended alternative pattern for introducing external component dependencies into VBA projects is to ensure the references are located in secured locations, and to register only those that are required on the individual users' machines under HKCR. Alternatively, if a well secured location for such external references is established and registering the components on users' machines is not possible, this secure location can be added to the internet trusted sites list.</span></span></div><br></span></div>



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
