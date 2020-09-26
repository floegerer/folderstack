# Introduction

Folderstack is an application written in VB.net that displays files and folders in a popup window for Windows. It can be launched from the taskbar, startmenu or anywhere else.

Folderstack works by passing some simple parameters to it. The most simple way to use is by just passing the path of the folder it should open to it.

## Usage

Create a shortcut to folderstack.exe, right click the shortcut file, select properties and add the path you want to open to the end of the "Run" parameter eg:

~~~
"folderstack.exe C:\Tools"
~~~

When you click the icon, FolderStack will list the content of that folder, any executable or link file within and all the subfolders within. You can click items to launch the file, or open another subfolder to open this folder in another popup window.


## Todos

- Rename all files necessary from folderstack > to folderstack, the new name of the application.
- Clean up the codebase and add comments to all vital subclasses.
