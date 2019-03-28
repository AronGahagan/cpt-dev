_License: This software is provided gratis (free of charge), AS IS and without Warranty. It is free to use, and free to distribute with prior written consent from the contributors/copyright holders provided no modifications are made. Contributors retain their patents. All other rights reserved. Copyright 2019, **company-name**._

# cpt: The CP Toolbar

![cpt](/images/cpt.png)

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
1. [Updates and Credits](#updates-and-credits)

[[top]](#table-of-contents)

## Prerequisites
1. Microsoft Project Standard 2013+
1. Microsoft Office (Outlook, Excel, Word, PowerPoint)
1. Microsoft .NET 4.1 (installs with -software- and -operating system-)

[[top]](#table-of-contents)

## Installation
1. Enable Macros
    1. Open Microsoft Project, go to File > Options > Trust Center > Trust Center Settings... > Macros
    1. CLick Enable all macros
    1. ~~Click 'trust access to the vb project object model' (this enables a sort of 'push' update process)~~
    1. Completely exit Microsoft Project, then restart (this makes the settings above 'stick')
1. Download [cptCore.mpp](http://github.com/AronGahagn/cpt) and copy all modules prefixed with "cpt" into your Global.MPT
1. View > Visual Baic > Macros > run the macro **cptSetup**
    1. This macro installs necessary core modules if they are not already installed
    1. Changes will be made to your ThisProject module, but if you have exising code it will not be overwritten. cpt-related code will be inserted at the very top of the procedures *Project_Activate* and *Project_Open* and each line is appended with '</cpt> for reference.
    
[[top]](#table-of-contents)

## Use
## View Tools
1. _Reset All_ - removes all groups and filters, expands all tasks, and reorders by Task.ID
1. _WrapItUp_ - collapses the native outline/summaries (or visible group) of your project startign from the bottom up.

[[up]](#use) | [[toc]](#table-of-contents)

## Count Tasks
Just what it sounds like - optionally count all tasks, visible (filtered) tasks, or selected tasks. Tally reports subtasks vs. summary tasks, and includes inactive tasks separately. External and 'Nothing' Tasks are skipped.

[[up]](#use) | [[toc]](#table-of-contents)

## Text Tools
1. _Bulk Prepend_ - does just what it says. Select some tasks, click the button, enter a prefix, done.
1. _Bulk Append_ - same as above but adds a suffix instead of a prefix.
1. _Enumerate_ - select some tasks, click _enumerate_. Enter how many characters your sequence should be (i.e., '3' = 000), and what number you wish to start by.
1. _MyReplace_ - a better (and safer) find/replace tool that:
    1. Limits found/replaced texts to the selected tasks
    1. Includes all Text (and Outline Code) fields in the selection also.
1. _Find Duplicates_ - finds duplicate task names and spits them into excel.
1. _Trim Task Names_ - removes all the pesky leading and trailing spaces when folks copy/paste from other tools.
1. _Dynamic Filter_ - This is a 'find-as-you-type' tool that allows you to optionally 'pin' a given task while searching for others. Optionally highlight the results instead of filtering. Include Summaries, or related summaries, in your results. Press CTRL+BKSP to clear the filter and return to your originally selected task, or click 'Clear All' to reset the file.
1. _Duplicate a Process_ - WIP - stay tuned!
1. _Advanced Text Tools_ - WIP - stay tuned!

[[up]](#use) | [[toc]](#table-of-contents)

## Trace Tools
1. Driving Path
1. Export to PowerPoint

[[up]](#use) | [[toc]](#table-of-contents)

## Status Tools
1. Status Sheet
1. Smart Duration

[[up]](#use) | [[toc]](#table-of-contents)

## Integration
1. Export IMS to COBRA
1. _Export IMS to MPM WIP_

[[up]](#use) | [[toc]](#table-of-contents)

## Updates and Credits
1. _Upgrades_ - From time to time click _Check for Upgrades_ to get the latest version.
1. _Feature Requests_ - submit your feature requests.
1. _Bug Reports_ - submit your feature requests.

_Credits: TBA_

[[up]](#use) | [[toc]](#table-of-contents)
