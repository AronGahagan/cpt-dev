![header](https://github.com/AronGahagan/cpt-dev/blob/develop/images/header.png?raw=true)

![cpt](https://github.com/AronGahagan/cpt-dev/blob/develop/images/cpt.png?raw=true)

<sub>_License: This software is provided gratis (free of charge), AS IS and without Warranty. It is free to use, and it is free to distribute **with prior written consent from the contributors/copyright holders** provided **no modifications are made**. Contributors retain their patents. All other rights reserved. Copyright 2019, contributors and ClearPlan Consulting, LLC._</sub>

---
[![Average time to resolve an issue](http://isitmaintained.com/badge/resolution/AronGahagan/cpt-dev.svg)](http://isitmaintained.com/project/AronGahagan/cpt-dev "Average time to resolve an issue")
[![Percentage of issues still open](http://isitmaintained.com/badge/open/AronGahagan/cpt-dev.svg)](http://isitmaintained.com/project/AronGahagan/cpt-dev "Percentage of issues still open")

---
## Purpose
* The purpose of this project is to provide schedulers with a time-saving and error-free tool _to be used in support of industry best practics and **solid processes**_. (Read: don't be like the TSA conveyor belt.)
* It is **not** the purpose of this project to absolve schedulers of the duty to build, analyze, and maintain good schedules. No tool, and no automation, can replace the mind, experience, judgment, and instincts of a living, breathing human being. Therefore:
* Use only as directed. "When in doubt, read the directions."

---
## Table of Contents
1. [Prerequisites](#prerequisites)
1. [Installation](#installation)
1. [Use](#use)
    1. [View and Count Tools](#view-and-count-tools)
    1. [Text Tools](#text-tools)
    1. [Trace Tools](#trace-tools)
    1. [Status Tools](#status-tools)
    1. [Resource Demand](#resource-demand)
    1. [Integration Tool](#integration-tools)
1. [Help](#help)
1. [Credits](#credits)

[[top]](#table-of-contents)

---
## Prerequisites
1. Microsoft Project Standard or Professional 2013+
1. Microsoft Office (Outlook, Excel, Word, PowerPoint)
1. Internet Connection preferred, but not required

[[top]](#table-of-contents)

---
## Installation
_Note: this tool is not currently designed (nor has it been tested) for use with Microsoft Project Server._

1. Enable Macros
    1. Open Microsoft Project, go to File > Options > Trust Center > Trust Center Settings...
    1. In the **Macro Settings** pane, click **Enable all macros**
    ![installation-01](https://github.com/AronGahagan/cpt-dev/blob/develop/images/installation-01.PNG?raw=true)
    1. In the **Legacy Formats** pane, click **Allow loading files with legacy or non-default formats** _(this will allow import/export of settings required for various features)_ 
    ![installation-02](https://github.com/AronGahagan/cpt-dev/blob/develop/images/installation-02.PNG?raw=true)
    1. Click **OK** a couple of times to close the dialogs
    1. Completely exit, and then restart, Microsoft Project (this makes the settings above 'stick')
1. Download and open [cpt.mpp](https://github.com/AronGahagan/cpt-dev/releases/download/1.3.1/cpt_v1.3.1.mpp)
1. Open the **Organizer** and select the **Modules** tab:
    1. **If you have an internet connection:** copy the `cptSetup_bas` module into your Global.MPT
    1. **If you do not have an internet connection:** copy **all** modules prefixed with `cpt` into your Global.MPT
1. On the Ribbon, click **View** > **Visual Basic** > **Macro**s > and run the macro `cptSetup()`
    1. This macro installs necessary core modules if they are not already installed
    1. Changes will be made to your ThisProject module, but if you have exising code it will not be overwritten. cpt-related code will be inserted at the very top of the procedures **Project_Activate** and **Project_Open** and each line is appended with `'</cpt>` for reference.
1. The ClearPlan Toolbar will be added. Click **ClearPlan** > **Help** > **Help** > **Check for Upgrades** to download the latest hotfixes. If you do not have an internet connection, please contact [cpt@ClearPlanConsulting.com](mailto:cpt@ClearPlanConsulting.com) for the latest hotfixes.
    
[[top]](#table-of-contents)

---
## Use
---
### View and Count Tools
![view](https://github.com/AronGahagan/cpt-dev/blob/develop/images/view.png?raw=true)

1. _**Reset All**_ - removes all groups and filters, expands all tasks, and reorders by Task.ID
1. _**WrapItUp**_ - collapses the currently visible outline (or group) of your project starting from the lowest level up.
1. _**Count Tasks**_ - Just what it sounds like - optionally count all tasks, visible (filtered) tasks, or selected tasks. User is prompted with a tally of subtasks vs. summary tasks, and includes inactive tasks separately. External and 'Nothing' Tasks are skipped.

[[up]](#use) | [[toc]](#table-of-contents)

---
### Text Tools
![text](https://github.com/AronGahagan/cpt-dev/blob/develop/images/text.png?raw=true)

1. _**Advanced Text Tools**_ - a single dialog aggregating the text tools below. Presents a preview of changes, and can be applied to any selected tasks. Undo works.
1. The following utilities are also available standalone in the top portion of the splitButton menu:
    1. _**Bulk Prepend**_ - does just what it says. Select some tasks, click the button, enter a prefix, done.
    1. _**Bulk Append**_ - same as above but adds a suffix instead of a prefix.
    1. _**Enumerate**_ - select some tasks, click _enumerate_. Enter how many characters your sequence should be (i.e., '3' = 000), and what number you wish to start by.
    1. _**MyReplace**_ - a better (and safer) find/replace tool that:
        1. Limits found/replaced texts to the selected tasks
        1. Includes all Text (and Outline Code) fields in the selection also.
1. Other tools:
    1. _**Trim Task Names**_ - removes all the pesky leading and trailing spaces when folks copy/paste from other tools.
    1. _**Replicate a Process**_ - WIP (future release)
    1. _**Find Duplicate Task Names**_ - finds duplicate task names and exports them into an Excel workbook.
    1. _**Reset Row Height**_ = for the type-A-er in you; when task rows auto-adjust and it gives you chest pains, click this button to restore your mental sanity and enter a limited state of ~ nirvana ~.
1. _**Dynamic Filter**_ - This 'find-as-you-type' tool is arguably the greatest addition ever made to Microsoft Project (but we are biased). It allows you to optionally 'pin' a given task while searching for others. Optionally highlight the results instead of filtering. Include Summaries, or related summaries, in your results. Press CTRL+BKSP to clear the filter and return to your originally selected task, or click 'Clear All' to reset the file. This tool is particularly helpful when working with a team to link up a schedule.

[[up]](#use) | [[toc]](#table-of-contents)

---
### Trace Tools
![trace](https://github.com/AronGahagan/cpt-dev/blob/develop/images/trace.png?raw=true)

1. _**Driving Path**_ - Select any target task, and then click this button. The Primary, Secondary, and Tertiary Driving Path will be displayed. _Note: as of v1.3 "`New Orleans`" this does not yet work with master/subprojects. We are targeting v1.5 "`Coronado`" for upgrade._
![drivingpath](https://github.com/AronGahagan/cpt-dev/blob/develop/images/driving-path.png?raw=true)

1. _**Export to PowerPoint**_ - _With the target task still selected_, click this button to export this `VBA wizardry` into a PowerPoint slide deck. Each Driving Path (Primary, Secondary, and Tertiary) will start a new set of slides. If the task count is too high, the network will be split over multiple slides to keep it readable (anti-eye-charting). If multiple slides are requried, the last task on (e.g.,) slide 1 will be repeated as the first task on the following slide to retain context and presentation flow.

[[up]](#use) | [[toc]](#table-of-contents)

---
### Status Tools
![status](https://github.com/AronGahagan/cpt-dev/blob/develop/images/status.png?raw=true)

1. _**Status Sheet**_ - Set a Project Status Date, then click this to create a simple Status Sheet turnaround document. Optionally add custom fields between UID and Task Name, which settings are saved for future runs. Options to create multiple worksheets or workbooks (per user-selected field) and/or send via email, are slated for a future release to be named.
1. _**Smart Duration**_ - "That task will be done on November 25th." "Ok..uh...let's starg guessing how many days of remainingi duration I need to hit that date..." Ever done that? NO MORE! Select a task, click this button, enter your finish date, and it will Figure Out The Things for you automagically. Please be sure to confirm remaining work after the magic subsides.

[[up]](#use) | [[toc]](#table-of-contents)

---
### Resource Demand
![resource](https://github.com/AronGahagan/cpt-dev/blob/develop/images/resource-demand.png?raw=true)

1. _**Resource Demand**_ - Export timephased remaining work to Excel (similar to the Task Usage view), with automatic PivotTable and PivotCharts. Add whatever fields you'd like to the report, and your settings will be saved for next time.

[[up]](#use) | [[toc]](#table-of-contents)

---
### Integration Tools
![integration](https://github.com/AronGahagan/cpt-dev/blob/develop/images/integration.png?raw=true)

1. _**Export IMS to COBRA**_ - Exports the baselined schedule into *.csv formats for upload to COBRA [versions?]. This is a must-have. See the [Help Doc] (link needed) for more information.

[[up]](#use) | [[toc]](#table-of-contents)

---
## Help
![help](https://github.com/AronGahagan/cpt-dev/blob/develop/images/help.png?raw=true)

1. _**Check for Upgrades**_ - From time to time click _Check for Upgrades_ to get the latest upgrades and hotfixes. _Note: if you do not have an internet connection avaiable from your client computer, download on a non-client asset next time you're in the coffee-shop, then copy the modules over into your Organizer next time you're onsite._
1. _**Submit an Issue**_ - submit your bugs/issues for quick triage and (depending on availability of our developers who are themselves fully deployed) we'll push fixes ASAP.
1. _**Submit a Feature Request**_ - submit your feature requests for review! Suggestions, once reviewed, will be voted on at the next offstite!
1. _**Submit Other Feedback**_ - our entire goal is to help you _sharpen your axe_ so please let us know how we're doing, what you'd like changed, what's broken, etc.
1. _**Uninstall ClearPlan Toolbar**_ - Never do this. _Never_. (Seriously, why would you do this?) ...ok, if you must uninstall (client request, you're rolling off the gig, etc.) then click this button and follow the prompts.

[[up]](#help) | [[toc]](#table-of-contents)

---
## Credits
A special thanks to our developers and contributors, and to the ClearPlan leadership team for inspiring us to `make it better`.

[[up]](#help) | [[toc]](#table-of-contents)

---
<sub>Copyright &#169; 2019 ClearPlan Consulting, LLC. All Rights Reserved.</sub>
