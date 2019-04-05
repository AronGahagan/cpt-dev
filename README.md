![logo](https://github.com/AronGahagan/cpt-dev/blob/master/images/clearplan_avatar_tiny_cropped.jpg?raw=true)

# cpt: The ClearPlan Toolbar

![cpt](https://github.com/AronGahagan/cpt-dev/blob/master/images/cpt.png?raw=true)

_License: This software is provided gratis (free of charge), AS IS and without Warranty. It is free to use, and it is free to distribute **with prior written consent from the contributors/copyright holders** provided **no modifications are made**. Contributors retain their patents. All other rights reserved. Copyright 2019, contributors and ClearPlan Consulting, LLC._

## Telos
* The _telos_ (purpose, goal, ultimate end) of this project is to provide you with a time-saving and error-free tool _to be used in support of industry best practics and **solid processes**_. (Read: don't be like the TSA conveyor belt.)
* It is **not** the _telos_ of this project to absolve you of the responsibility of _scheduling_ responsibly. No tool, and no automation, can replace the mind, experience, judgment, and instincts of a living, breathing human being. Therefore:
* Use only as directed. When in doubt, read the directions.

## Table of Contents
1. [Prerequisites](#prerequisites)
1. [Installation](#installation)
1. [Use](#use)
    1. [View Tools](#view-tools)
    1. [Count Tasks](#count-tasks)
    1. [Text Tools](#text-tools)
    1. [Trace Tools](#trace-tools)
    1. [Status Tools](#status-tools)
    1. [Integration](#integration)
1. [Help](#help)
1. [Credits](#credits)

[[top]](#table-of-contents)

## Prerequisites
1. Microsoft Project Standard 2013+
1. Microsoft Office (Outlook, Excel, Word, PowerPoint)
1. Internet Connection preferred, but not required

[[top]](#table-of-contents)

## Installation
1. Enable Macros
    1. Open Microsoft Project, go to File > Options > Trust Center > Trust Center Settings...
    1. In the **Macro Settings** pane, click **Enable all macros**
    ![installation-01](https://github.com/AronGahagan/cpt-dev/blob/master/images/installation-01.png?raw=true)
    1. In the **Legacy Formats** pane, click **Allow loading files with legacy or non-default formats** _(this will allow import/export of settings required for various features)_ 
    ![installation-02](https://github.com/AronGahagan/cpt-dev/blob/master/images/installation-02.png?raw=true)
    1. Click **OK** a couple of times to close the dialogs
    1. Completely exit Microsoft Project, then restart (this makes the settings above 'stick')
1. Download [cpt.mpp](https://github.com/AronGahagan/cpt-dev/releases/download/v1.3/cpt.mpp)
1. Open the **Organizer** and select the **Modules** tab:
    1. **If you have an internet connection:** copy the `cptSetup_bas` module your Global.MPT
    1. **If you do not have an internet connection:** copy all modules prefixed with `cpt` into your Global.MPT
1. On the Ribbon, click View > Visual Baic > Macros > and run the macro `cptSetup()`
    1. This macro installs necessary core modules if they are not already installed
    1. Changes will be made to your ThisProject module, but if you have exising code it will not be overwritten. cpt-related code will be inserted at the very top of the procedures **Project_Activate** and **Project_Open** and each line is appended with `'</cpt>` for reference.
1. The ClearPlan Toolbar will be added. Click ClearPlan > Help > Help > Check for Upgrades to download the latest hotfixes. If you do not have an internet connection, please contact [cpt@ClearPlanConsulting.com](mailto:cpt@ClearPlanConsulting.com) for the latest hotfixes.
    
[[top]](#table-of-contents)

## Use
### View Tools
1. _Reset All_ - removes all groups and filters, expands all tasks, and reorders by Task.ID
1. _WrapItUp_ - collapses the currently visible outline (or group) of your project starting from the lowest level up.

[[up]](#use) | [[toc]](#table-of-contents)

### Count Tasks
Just what it sounds like - optionally count all tasks, visible (filtered) tasks, or selected tasks. User is prompted with a tally of subtasks vs. summary tasks, and includes inactive tasks separately. External and 'Nothing' Tasks are skipped.

[[up]](#use) | [[toc]](#table-of-contents)

### Text Tools
1. _Advanced Text Tools_ - a single dialog aggregating the text tools below. Presents a preview of changes, and can be applied to any selected tasks. Undo works.
1. The following utilities are also available standalone in the top portion of the splitButton menu:
    1. _Bulk Prepend_ - does just what it says. Select some tasks, click the button, enter a prefix, done.
    1. _Bulk Append_ - same as above but adds a suffix instead of a prefix.
    1. _Enumerate_ - select some tasks, click _enumerate_. Enter how many characters your sequence should be (i.e., '3' = 000), and what number you wish to start by.
    1. _MyReplace_ - a better (and safer) find/replace tool that:
        1. Limits found/replaced texts to the selected tasks
        1. Includes all Text (and Outline Code) fields in the selection also.
1. Other tools:
    1. _Trim Task Names_ - removes all the pesky leading and trailing spaces when folks copy/paste from other tools.
    1. _Replicate a Process_ - WIP (future release)
    1. _Find Duplicate Task Names_ - finds duplicate task names and exports them into an Excel workbook.
    1. _Reset Row Height_ = for the type-A-er in you; when task rows auto-adjust and it gives you chest pains, click this button to restore your mental sanity and enter a limited state of ~ nirvana ~.
1. _Dynamic Filter_ - This is a 'find-as-you-type' tool that allows you to optionally 'pin' a given task while searching for others. Optionally highlight the results instead of filtering. Include Summaries, or related summaries, in your results. Press CTRL+BKSP to clear the filter and return to your originally selected task, or click 'Clear All' to reset the file. This tool is particularly helpful when working with a team to link up a schedule.

[[up]](#use) | [[toc]](#table-of-contents)

### Trace Tools
1. _Driving Path_ - Select any target task, and then click this button. The Primary, Secondary, and Tertiary Driving Path will be displayed. _Note: as of v1.3 "`New Orleans`" this does not yet work with master/subprojects. We are targeting v1.5 "`Coronado`" for upgrade._
1. _Export to PowerPoint_ - _With the target task still selected_, click this button to export the view into a PowerPoint slide deck. Each Driving Path (Primary, Secondary, and Tertiary) will start a new set of slides. If the task count is too high, the network will be split over multiple slides to keep it readable (anti-eye-charting). If multiple slides are requried, the last task on (e.g.,) slide 1 will be repeated as the first task on the following slide to retain context and presentation flow.

[[up]](#use) | [[toc]](#table-of-contents)

### Status Tools
1. _Status Sheet_ - Set a Project Status Date, then click this to create a simple Status Sheet turnaround document. Optionally add custom fields between UID and Task Name, which settings are saved for future runs. Options to create multiple worksheets or workbooks (per user-selected field) and/or send via email, are slated for a future release to be named.
1. _Smart Duration_ - "That task will be done on November 25th." "Ok..uh...let's starg guessing how many days of remainingi duration I need to hit that date..." Ever done that? NO MORE! Select a task, click this button, enter your finish date, and it will Figure Out The Things for you automagically. Please be sure to confirm remaining work after the magic subsides.

[[up]](#use) | [[toc]](#table-of-contents)

### Integration
1. _Export IMS to COBRA_ - Exports the baselined schedule into *.csv formats for upload to COBRA [versions?]. See the [Help Doc] (link needed) for more information.

[[up]](#use) | [[toc]](#table-of-contents)

## Help
1. _Check for Upgrades_ - From time to time click _Check for Upgrades_ to get the latest upgrades and hotfixes. _Note: if you do not have an internet connection avaiable from your client computer, download on a non-client asset next time you're in the coffee-shop, then copy the modules over into your Organizer next time you're onsite._
1. _Submit an Issue_ - submit your bugs/issues for quick triage and (depending on availability of our developers who are themselves fully deployed) we'll push fixes ASAP.
1. _Submit a Feature Request_ - submit your feature requests for review! Suggestions, once reviewed, will be voted on at the next offstite!
1. _Submit Other Feedback_ - our entire goal is to help you _sharpen your axe_ so please let us know how we're doing, what you'd like changed, what's broken, etc.
1. _Uninstall ClearPlan Toolbar_ - Never do this. _Never_. Seriously, why would you do this? ...ok, if you must uninstall (client request, you're rolling off the gig, etc.) then click this button and follow the prompts.

## Credits
A special thanks to our developers and contributors, and to the ClearPlan leadership team for inspiring us to `make it better`.

[[up]](#use) | [[toc]](#table-of-contents)
