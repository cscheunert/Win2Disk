# Win2Disk
A script to install Windows to an attached drive without a bootable installer


This script installs Windows Images such as ISO, WIM or ESD files to attached internal and external drives
It is based on the Prepare-VirtualWinVhd.ps1 script that does a similar job for virtual hard disks (https://gist.github.com/milolav/3e296ed6a9f8a6c431a8553060f7514b).

The script can be used interactive or with the Imagefile, Disknumber and Indexnumber parameters for automation.



|Tested Windows version|Status|
|---|---|
|Windows 7|problem setting the bootloader|
|Windows 8.1|works|
|Windows 10|works|
|Windows 10 Insider|works|
|Windows 11 leak|works|
