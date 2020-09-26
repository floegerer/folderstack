'
' Created by SharpDevelop.
' User: code2
' Date: 14.03.2012
' Time: 00:36
'
' To change this template use Tools | Options | Coding | Edit Standard Headers.
'

Imports System.IO
Imports System.Math
Imports Stackr.MSjogren.Samples.ShellLink.VB
Imports Microsoft.Win32
Imports System.Runtime.InteropServices


Public Partial Class MainForm

  Public MousePos As Point
  Public WinHeight As Double
  Public WinWidth As Double
  Public WinFocus As Boolean
  Public WinPath As String
  Public TileCount As Double
  Public SettingTileSpacing As Double
  Public CalcHeight As Double
  Public IconSpacing As Double
  Public FileCount As Double
  Public SettingTileBreak As Double
  Public SettingView As String
  Public SettingText As Boolean
  Public SettingIconSize As Double

  Public Sub New()

    'TODO: Need to set to relative path

    WinPath = "C:\Users\Flo\AppData\Roaming\Microsoft\Windows\Start Menu\Programs"
    SettingText = True
    SettingView = "Icons"
    SettingIconSize = 32
    SettingTileBreak  = 8
    SettingTileSpacing = 18
    TileCount = 1
    FileCount = 1
    ExecuteParams()


    Application.EnableVisualStyles()
    Application.UseWaitCursor = False

    Dim IconSpacingHorizontal As Double
    Dim IconSpacingVertical As Double
    Dim Dir As New DirectoryInfo(WinPath)
    Dim RegKey As RegistryKey
    RegKey = Registry.CurrentUser.OpenSubKey( _
      "Control Panel\Desktop\WindowMetrics")

    IconSpacingHorizontalReg = Split(RegKey.GetValue("IconSpacing").ToString,"-")(1)
    IconSpacingHorizontal = Convert.ToDouble(IconSpacingHorizontalReg)
    IconSpacingHorizontal = (IconSpacingHorizontal-480)/15

    'IconSpacingHorizontal = Ceiling(IconSpacingHorizontal/2) * 2
    'MsgBox(IconSpacingHorizontal)

    IconSpacingVerticalReg = Split(RegKey.GetValue("IconVerticalSpacing").ToString,"-")(1)
    IconSpacingVertical = Convert.ToDouble(IconSpacingVerticalReg)
    IconSpacingVertical = (IconSpacingVertical-480)/15

    'IconSpacingVertical = Ceiling(IconSpacingVertical/2) * 2
    'MsgBox(IconSpacingVertical)

    IconSpacing = IconSpacingHorizontal+SettingIconSize
    IconSpacingVertical = IconSpacingVertical+SettingIconSize

    For Each F As FileInfo In Dir.GetFiles("*.lnk")
    FileCount +=1
  Next

  For Each F As DirectoryInfo In Dir.GetDirectories
  FileCount +=1
Next

If SettingView Is "Tiles" Then


  TileCount =  Convert.ToInt16(Ceiling(FileCount/SettingTileBreak))

  'MsgBox(TileCount)

  If FileCount < 10 Then
    TileCount = 1
  End If

  CalcHeight = Ceiling(FileCount/TileCount)*(SettingIconSize+SettingTileSpacing)+64
  WinWidth = TileCount*240

Else

  WinWidth = Round(Sqrt(FileCount))
  CalcHeight = Ceiling(FileCount/WinWidth)

  CalcHeight= CalcHeight * IconSpacingVertical
  WinWidth*=IconSpacing
  WinWidth+=64


  If SettingText = True Then
    CalcHeight+=SettingIconSize+2
  Else
    CalcHeight+=20
  End If

End If

WinHeight = CalcHeight


MousePos = MousePosition
'MsgBox(Screen.PrimaryScreen.WorkingArea.Height)

If MousePos.Y+100 > Screen.PrimaryScreen.WorkingArea.Height Then
  MousePos.Y -= Convert.ToInt16(CalcHeight+56)
Else
  MousePos.Y -= Convert.ToInt16(CalcHeight/2)
End If

MousePos.X -= Convert.ToInt16(WinWidth/2)


Me.InitializeComponent()
Me.Opacity = 0
Me.Location = MousePos
Me.ListView.Cursor = System.Windows.Forms.Cursors.Arrow
Me.ListContainer.Cursor = System.Windows.Forms.Cursors.Arrow
Me.Cursor = System.Windows.Forms.Cursors.Arrow

AddHandler Me.LostFocus, AddressOf Me.WinFocusLost
AddHandler Me.MouseDown, AddressOf Me.WinClick


FillListView(Dir,"*.lnk")


Startup.Start

Me.ClientSize = New System.Drawing.Size(Convert.ToInt16(WinWidth+2), Convert.ToInt16(WinHeight)+2)

End Sub




Public Sub DefineHeight
End Sub




Private Sub FillListView(ByVal Folder As DirectoryInfo, ByVal Filter As String )

'MsgBox(Folder.GetFiles.Length)

Static ImageList As New ImageList With {.ColorDepth = ColorDepth.Depth32Bit, .ImageSize = New Size(Convert.ToInt16(SettingIconSize), Convert.ToInt16(SettingIconSize))}
Static ImageListSmall As New ImageList With {.ColorDepth = ColorDepth.Depth32Bit, .ImageSize = New Size(16, 16)}
Static ImgIndex As Integer = 2

If SettingView Is "Tiles" Then
  ListView.View = System.Windows.Forms.View.Tile
Else If SettingView Is "Icons" Then
    ListView.View = System.Windows.Forms.View.LargeIcon
  End If


  ListView.LargeImageList = ImageList
  ListView.SmallImageList = ImageListSmall
  ListView.TileSize = New System.Drawing.Size(  Convert.ToInt16((WinWidth-64)/TileCount)  , Convert.ToInt16( (SettingIconSize+SettingTileSpacing) )  )

  ImageList.Images.Add(Icon.ExtractAssociatedIcon(".\shortcut.ico"))
  ImageList.Images.Add(Icon.ExtractAssociatedIcon(".\folder.ico"))
  'ImageList.Images.Add(Icon.ExtractAssociatedIcon(".\up.ico"))

  If SettingText = True Then

    ListView.Items.Add(" All Shortcuts ",0)
    ListView.Items.Add(" Parent Folder ",2)

  Else

    ListView.Items.Add("",0)

  End If


  For Each F As FileInfo In Folder.GetFiles(Filter)

  Try

    Dim IconFile As ShellShortcut
    IconFile = New ShellShortcut(F.FullName)

    'MsgBox(IconFile.IconPath.ToString + IconFile.IconIndex.ToString + F.FullName)

    If IconFile.IconPath.Length > 0 Then

      ImageList.Images.Add(IconFile.Icon)
      ImageListSmall.Images.Add(IconFile.Icon)

      'MsgBox("True")

    Else

      '			hInst = Marshal.GetHINSTANCE(Me.GetType().Module)
      '			hIcon = Native.ExtractIcon(hInst,F.FullName,0)
      '
      '			ico = Icon.FromHandle(hIcon)
      '			clone = CType(ico.Clone(), Icon)
      '			ImageListSmall.Images.Add(ico)

      ImageList.Images.Add(Icon.ExtractAssociatedIcon(F.FullName))
      ImageListSmall.Images.Add(Icon.ExtractAssociatedIcon(F.FullName))

    End If

  Catch

    ImageList.Images.Add(Icon.ExtractAssociatedIcon(".\not_found.ico"))
    ImageListSmall.Images.Add(Icon.ExtractAssociatedIcon(".\not_found.ico"))

  End Try

  If SettingText = True Then

    Dim TextDis As String

    TextDis = F.Name.Remove(F.Name.Length-4,4)

    If SettingView Is "Tiles" Then
      TextDis = "​ " + F.Name.Remove(F.Name.Length-4,4)
    End If

    ListView.Items.Add(TextDis, ImgIndex).Tag = F
  Else
    ListView.Items.Add("", ImgIndex).Tag = F
  End If

  ImgIndex += 1

Next

For Each Dir As DirectoryInfo In Folder.GetDirectories
'FillListView(Dir)
ListView.Items.Add(" " + Dir.Name + " ", 1).Tag = Dir
Next

End Sub




Public Sub ExecuteParams()

Dim args As String()
args = Environment.GetCommandLineArgs()

'bei 1 starten, weil das Programm IMMER seinen eigenen FULLPATH als ersten Parameter erkennt!

For i As Integer = 1 To args.Length - 1
Select Case args(i).ToLower
Case "?", "help", "info"
MessageBox.Show(My.Computer.Name, "? / help / info")

Case "tiles"

SettingView = "Tiles"

Case "notext"

SettingText = False

Case "small"

SettingIconSize = 14
SettingTileBreak = 30
SettingTileSpacing = 12

Case Else
WinPath = args(i)
'MessageBox.Show("Unbekannter Parameter:" & vbCrLf & args(i), "Unbekannt!", MessageBoxButtons.OK, MessageBoxIcon.Information)
End Select
Next

End Sub




Public Sub WinFocusLost(sender As Object, e As EventArgs)

If WinFocus
Application.Exit
End If

End Sub




Public Sub WinClick(sender As Object, e As EventArgs)

Application.Exit
'MsgBox("Test")

End Sub




Sub OpacityStepsTick(sender As Object, e As EventArgs)

Me.Opacity += 0.15

If Me.Opacity >= 1.01 Then
OpacitySteps.Enabled = false
End If

End Sub



End Class
