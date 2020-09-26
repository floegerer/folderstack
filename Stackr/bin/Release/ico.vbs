Const DESKTOP = &H10&

Set objShell = CreateObject("Shell.Application")
Set objFolder = objShell.NameSpace(DESKTOP)

Set objFolderItem = objFolder.ParseName("Test Shortcut.lnk")
Set objShortcut = objFolderItem.GetLink

objShortcut.SetIconLocation "C:\Windows\System32\SHELL32.dll", 13
objShortcut.Save