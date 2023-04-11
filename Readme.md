Excel-DNA - WiXInstaller
------------------------

This is based on a user-contributed template (thank you very much to Lee Zeitz!) for making a WiX-based installer for an Excel-DNA add-in.

Use this project as a template - fork it and update from there.
The Test add-in is just as a sample - the add-in project need not be part of the installer solution.

In the ExcelAddInDeploy project, there are some customizations to make:
* EnglishLoc.wxl - the strings for the installer to be fixed up
* Product.wxs - set the ProductCode and UpgradeCode to be unique for your add-in, and the other Guids and paths - look for the TODOs.
* Fix the Resources - Banner, Icon end EULA, as required.

Version for making a Local Machine installer using Active Setup
---

Benoit Patra has created a version of the Excel-DNA WiX installer that uses the Active Setup feature in Windows to install for all users. The repository is here: https://github.com/bpatra/ExcelDNAWixInstallerLM and a very detailed write-up here: http://benoitpatra.com/2014/07/26/a-sample-wix-installer-using-the-activesetup-feature/

Any help to merge the installers, making the per-user or per-machine installations an installation option, would be greatly appreciated.



