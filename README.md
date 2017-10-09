# Sample Scripts

The scripts in this repository are intended to serve as examples and are not supported.  They are intended to be used by those comfortable enough with scripting to modify them and incorporate them into your own project.

## Background
---
Normally scripting used to set a computer name in a task sequence sets the OSDComputerName variable, which is consumed during the Mini-Setup.

When preparing a task sequence for consumption in the factory it often becomes necessary to move computer naming out to Post-Delivery Configuration to overcome the restriction on interaction during the portion which runs in the DellEMC factory.  The challenge is that at this point, the OSDComputerName variable has already been consumed, so setting it will not have the desirtred effect.  In these cases, an alternate method is needed to set the computer name.

These sample scripts demonstrate how WMI can be used to accomplish the rename.  Scripts using this approach should be placed in the Post-Delivery Configuration section of a task sequence prior to the domain join, and should be followed by a restart to the currently installed OS.  (When using WMI to rename, the new name only goes into effect after a reboot.)
