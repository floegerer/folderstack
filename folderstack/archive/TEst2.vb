Dim hImgSmall As IntPtr  'The handle to the system image list.
Dim hImgLarge As IntPtr  'The handle to the system image list.
Dim fName As String      'The file name to get the icon from.
Dim shinfo As SHFILEINFO
shinfo = New SHFILEINFO()
Dim openFileDialog1 As OpenFileDialog
openFileDialog1 = New OpenFileDialog()

openFileDialog1.InitialDirectory = "c:\temp\"
openFileDialog1.Filter           = "All files (*.*)|*.*"
openFileDialog1.FilterIndex      = 2
openFileDialog1.RestoreDirectory = True

listView1.SmallImageList = imageList1
listView1.LargeImageList = ImageList1

shinfo.szDisplayName = New String(Chr(0), 260)
shinfo.szTypeName    = New String(Chr(0), 80)

If (openFileDialog1.ShowDialog() = DialogResult.OK) Then
     fName = openFileDialog1.FileName

     'Use this to get the small icon.
     hImgSmall = SHGetFileInfo(fName, 0, shinfo, _
                 Marshal.SizeOf(shinfo), _
                 SHGFI_ICON Or SHGFI_SMALLICON)

     'Use this to get the large icon.
     'hImgLarge = SHGetFileInfo(fName, 0,
     'ref shinfo, (uint)Marshal.SizeOf(shinfo),
     'SHGFI_ICON | SHGFI_LARGEICON);

     'The icon is returned in the hIcon member of the
     'shinfo struct.
     Dim myIcon As System.Drawing.Icon
     myIcon = System.Drawing.Icon.FromHandle(shinfo.hIcon)

     imageList1.Images.Add(myIcon)       'Add icon to
                                         'imageList.

     listView1.Items.Add(fName, nIndex)  'Add file name and
                                         'icon to listview.
     nIndex = nIndex + 1
End If
