'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Filename:     ShellLinkTest.vb
' Project:      ShellLinkSample
' Author:       Mattias Sjögren (mattias@mvps.org)
'               http://www.msjogren.net/dotnet/
'
' Description:  Simple shortcut editor to demonstrate how to use
'               the ShellShortcut class.
'
' Public types: Class ShellLinkTest
'
'
' Dependencies: ShellShortcut.vb
'
'
' Copyright ©2001-2002, Mattias Sjögren
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Imports System
Imports System.Collections
Imports System.ComponentModel
Imports System.Diagnostics
Imports System.Drawing
Imports System.IO
Imports System.Windows.Forms
Imports System.Runtime.InteropServices

Imports MSjogren.Samples.ShellLink.VB


Public Class ShellLinkTest
  Inherits System.Windows.Forms.Form

  Private m_Shortcut As ShellShortcut

  Public Sub New()
    MyBase.New()

    InitializeComponent()

    cboWinStyle.Items.Add(ProcessWindowStyle.Normal)
    cboWinStyle.Items.Add(ProcessWindowStyle.Minimized)
    cboWinStyle.Items.Add(ProcessWindowStyle.Maximized)
   End Sub

  Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
    If disposing Then
      If Not components Is Nothing Then components.Dispose()
    End If
    MyBase.Dispose(disposing)
  End Sub

#Region " Windows Form Designer generated code "

  'Required by the Windows Form Designer
  Private components As System.ComponentModel.Container

  Private WithEvents Label7 As System.Windows.Forms.Label
  Private WithEvents Label6 As System.Windows.Forms.Label
  Private WithEvents cboWinStyle As System.Windows.Forms.ComboBox
  Private WithEvents Label5 As System.Windows.Forms.Label
  Private WithEvents Label4 As System.Windows.Forms.Label
  Private WithEvents Label3 As System.Windows.Forms.Label
  Private WithEvents Label2 As System.Windows.Forms.Label
  Private WithEvents Label1 As System.Windows.Forms.Label
  Private WithEvents btnBrowseTarget As System.Windows.Forms.Button
  Private WithEvents btnSave As System.Windows.Forms.Button
  Private WithEvents btnOpen As System.Windows.Forms.Button
  Private WithEvents txtIconIdx As System.Windows.Forms.TextBox
  Private WithEvents txtIconFile As System.Windows.Forms.TextBox
  Private WithEvents txtArgs As System.Windows.Forms.TextBox
  Private WithEvents txtWorkingDir As System.Windows.Forms.TextBox
  Private WithEvents txtTarget As System.Windows.Forms.TextBox
Private WithEvents Label9 As System.Windows.Forms.Label
Private WithEvents txtDescription As System.Windows.Forms.TextBox
Private WithEvents lblCurrentFile As System.Windows.Forms.Label
Private WithEvents lblHotkey As System.Windows.Forms.Label
Private WithEvents picIcon As System.Windows.Forms.PictureBox
Private WithEvents btnBrowseIcon As System.Windows.Forms.Button

  'NOTE: The following procedure is required by the Windows Form Designer
  'It can be modified using the Windows Form Designer.
  'Do not modify it using the code editor.
  <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
Me.Label3 = New System.Windows.Forms.Label()
Me.btnBrowseIcon = New System.Windows.Forms.Button()
Me.Label2 = New System.Windows.Forms.Label()
Me.txtTarget = New System.Windows.Forms.TextBox()
Me.lblHotkey = New System.Windows.Forms.Label()
Me.lblCurrentFile = New System.Windows.Forms.Label()
Me.btnSave = New System.Windows.Forms.Button()
Me.picIcon = New System.Windows.Forms.PictureBox()
Me.txtDescription = New System.Windows.Forms.TextBox()
Me.txtIconFile = New System.Windows.Forms.TextBox()
Me.cboWinStyle = New System.Windows.Forms.ComboBox()
Me.btnOpen = New System.Windows.Forms.Button()
Me.txtArgs = New System.Windows.Forms.TextBox()
Me.txtIconIdx = New System.Windows.Forms.TextBox()
Me.txtWorkingDir = New System.Windows.Forms.TextBox()
Me.Label9 = New System.Windows.Forms.Label()
Me.Label4 = New System.Windows.Forms.Label()
Me.Label5 = New System.Windows.Forms.Label()
Me.Label6 = New System.Windows.Forms.Label()
Me.Label7 = New System.Windows.Forms.Label()
Me.Label1 = New System.Windows.Forms.Label()
Me.btnBrowseTarget = New System.Windows.Forms.Button()
Me.SuspendLayout()
'
'Label3
'
Me.Label3.Location = New System.Drawing.Point(8, 100)
Me.Label3.Name = "Label3"
Me.Label3.TabIndex = 10
Me.Label3.Text = "Arguments:"
'
'btnBrowseIcon
'
Me.btnBrowseIcon.Enabled = False
Me.btnBrowseIcon.Location = New System.Drawing.Point(384, 144)
Me.btnBrowseIcon.Name = "btnBrowseIcon"
Me.btnBrowseIcon.Size = New System.Drawing.Size(24, 20)
Me.btnBrowseIcon.TabIndex = 6
Me.btnBrowseIcon.Text = "..."
'
'Label2
'
Me.Label2.Location = New System.Drawing.Point(8, 76)
Me.Label2.Name = "Label2"
Me.Label2.TabIndex = 9
Me.Label2.Text = "Working dir:"
'
'txtTarget
'
Me.txtTarget.Location = New System.Drawing.Point(80, 48)
Me.txtTarget.Name = "txtTarget"
Me.txtTarget.Size = New System.Drawing.Size(296, 20)
Me.txtTarget.TabIndex = 0
Me.txtTarget.Text = ""
'
'lblHotkey
'
Me.lblHotkey.Location = New System.Drawing.Point(80, 196)
Me.lblHotkey.Name = "lblHotkey"
Me.lblHotkey.Size = New System.Drawing.Size(112, 23)
Me.lblHotkey.TabIndex = 17
'
'lblCurrentFile
'
Me.lblCurrentFile.Font = New System.Drawing.Font("Microsoft Sans Serif", 9!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.lblCurrentFile.ForeColor = System.Drawing.Color.Blue
Me.lblCurrentFile.Location = New System.Drawing.Point(64, 8)
Me.lblCurrentFile.Name = "lblCurrentFile"
Me.lblCurrentFile.Size = New System.Drawing.Size(344, 23)
Me.lblCurrentFile.TabIndex = 16
'
'btnSave
'
Me.btnSave.Enabled = False
Me.btnSave.Location = New System.Drawing.Point(312, 240)
Me.btnSave.Name = "btnSave"
Me.btnSave.Size = New System.Drawing.Size(96, 23)
Me.btnSave.TabIndex = 10
Me.btnSave.Text = "&Save"
'
'picIcon
'
Me.picIcon.Location = New System.Drawing.Point(8, 8)
Me.picIcon.Name = "picIcon"
Me.picIcon.Size = New System.Drawing.Size(32, 32)
Me.picIcon.TabIndex = 18
Me.picIcon.TabStop = False
'
'txtDescription
'
Me.txtDescription.Location = New System.Drawing.Point(80, 120)
Me.txtDescription.Name = "txtDescription"
Me.txtDescription.Size = New System.Drawing.Size(328, 20)
Me.txtDescription.TabIndex = 4
Me.txtDescription.Text = ""
'
'txtIconFile
'
Me.txtIconFile.Location = New System.Drawing.Point(80, 144)
Me.txtIconFile.Name = "txtIconFile"
Me.txtIconFile.Size = New System.Drawing.Size(296, 20)
Me.txtIconFile.TabIndex = 5
Me.txtIconFile.Text = ""
'
'cboWinStyle
'
Me.cboWinStyle.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
Me.cboWinStyle.DropDownWidth = 160
Me.cboWinStyle.Location = New System.Drawing.Point(248, 168)
Me.cboWinStyle.Name = "cboWinStyle"
Me.cboWinStyle.Size = New System.Drawing.Size(160, 21)
Me.cboWinStyle.TabIndex = 8
'
'btnOpen
'
Me.btnOpen.Location = New System.Drawing.Point(208, 240)
Me.btnOpen.Name = "btnOpen"
Me.btnOpen.Size = New System.Drawing.Size(96, 23)
Me.btnOpen.TabIndex = 9
Me.btnOpen.Text = "&Open/Create"
'
'txtArgs
'
Me.txtArgs.Location = New System.Drawing.Point(80, 96)
Me.txtArgs.Name = "txtArgs"
Me.txtArgs.Size = New System.Drawing.Size(328, 20)
Me.txtArgs.TabIndex = 3
Me.txtArgs.Text = ""
'
'txtIconIdx
'
Me.txtIconIdx.Location = New System.Drawing.Point(80, 168)
Me.txtIconIdx.Name = "txtIconIdx"
Me.txtIconIdx.Size = New System.Drawing.Size(64, 20)
Me.txtIconIdx.TabIndex = 7
Me.txtIconIdx.Text = ""
'
'txtWorkingDir
'
Me.txtWorkingDir.Location = New System.Drawing.Point(80, 72)
Me.txtWorkingDir.Name = "txtWorkingDir"
Me.txtWorkingDir.Size = New System.Drawing.Size(328, 20)
Me.txtWorkingDir.TabIndex = 2
Me.txtWorkingDir.Text = ""
'
'Label9
'
Me.Label9.Location = New System.Drawing.Point(8, 124)
Me.Label9.Name = "Label9"
Me.Label9.TabIndex = 11
Me.Label9.Text = "Description:"
'
'Label4
'
Me.Label4.Location = New System.Drawing.Point(8, 148)
Me.Label4.Name = "Label4"
Me.Label4.TabIndex = 11
Me.Label4.Text = "Icon file:"
'
'Label5
'
Me.Label5.Location = New System.Drawing.Point(8, 172)
Me.Label5.Name = "Label5"
Me.Label5.TabIndex = 12
Me.Label5.Text = "Icon index:"
'
'Label6
'
Me.Label6.Location = New System.Drawing.Point(168, 172)
Me.Label6.Name = "Label6"
Me.Label6.TabIndex = 14
Me.Label6.Text = "Window style:"
'
'Label7
'
Me.Label7.Location = New System.Drawing.Point(8, 196)
Me.Label7.Name = "Label7"
Me.Label7.Size = New System.Drawing.Size(64, 23)
Me.Label7.TabIndex = 15
Me.Label7.Text = "Hotkey:"
'
'Label1
'
Me.Label1.Location = New System.Drawing.Point(8, 52)
Me.Label1.Name = "Label1"
Me.Label1.TabIndex = 8
Me.Label1.Text = "Target:"
'
'btnBrowseTarget
'
Me.btnBrowseTarget.Enabled = False
Me.btnBrowseTarget.Location = New System.Drawing.Point(384, 48)
Me.btnBrowseTarget.Name = "btnBrowseTarget"
Me.btnBrowseTarget.Size = New System.Drawing.Size(24, 20)
Me.btnBrowseTarget.TabIndex = 1
Me.btnBrowseTarget.Text = "..."
'
'ShellLinkTest
'
Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
Me.ClientSize = New System.Drawing.Size(418, 271)
Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.btnBrowseIcon, Me.picIcon, Me.txtDescription, Me.lblHotkey, Me.lblCurrentFile, Me.Label7, Me.cboWinStyle, Me.Label6, Me.btnBrowseTarget, Me.txtIconIdx, Me.txtIconFile, Me.txtArgs, Me.txtWorkingDir, Me.txtTarget, Me.Label5, Me.Label4, Me.Label3, Me.Label2, Me.Label1, Me.btnSave, Me.btnOpen, Me.Label9})
Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
Me.MaximizeBox = False
Me.Name = "ShellLinkTest"
Me.Text = "ShellLink sample"
Me.ResumeLayout(False)

  End Sub

#End Region

  Protected Sub btnOpen_Click(ByVal sender As Object, ByVal e As System.EventArgs) _
    Handles btnOpen.Click

    Dim ofd As OpenFileDialog = New OpenFileDialog()

    With ofd
      .CheckFileExists = False
      .DefaultExt = "lnk"
      .DereferenceLinks = False
      .Title = "Select a shortcut file"
      .Filter = "Shortcuts (*.lnk)|*.lnk|All files (*.*)|*.*"
      .FilterIndex = 1
    End With

    If ofd.ShowDialog() = DialogResult.OK Then
      m_Shortcut = New ShellShortcut(ofd.FileName)
      lblCurrentFile.Text = ofd.FileName

      With m_Shortcut
        txtTarget.Text = .Path
        txtWorkingDir.Text = .WorkingDirectory
        txtArgs.Text = .Arguments
        txtDescription.Text = .Description
        txtIconFile.Text = .IconPath
        txtIconIdx.Text = .IconIndex.ToString()

        cboWinStyle.SelectedIndex = cboWinStyle.Items.IndexOf(.WindowStyle)
        lblHotkey.Text = TypeDescriptor.GetConverter(GetType(Keys)).ConvertToString(.Hotkey)

        Dim ico As Icon = .Icon
        If Not (ico Is Nothing) Then
          picIcon.Image = ico.ToBitmap()
          ico.Dispose()
        End If

      End With

      btnSave.Enabled = True
      btnBrowseTarget.Enabled = True
      btnBrowseIcon.Enabled = True
    End If

    ofd.Dispose()

  End Sub

  Protected Sub btnSave_Click(ByVal sender As Object, ByVal e As System.EventArgs) _
    Handles btnSave.Click

    With m_Shortcut
      .Path = txtTarget.Text
      .WorkingDirectory = txtWorkingDir.Text
      .Arguments = txtArgs.Text
      .Description = txtDescription.Text
      .IconPath = txtIconFile.Text
      .IconIndex = CInt(txtIconIdx.Text)
      .WindowStyle = CType(cboWinStyle.SelectedItem, ProcessWindowStyle)
      .Save()
    End With

    Dim ico As Icon = m_Shortcut.Icon
    If ico Is Nothing Then
      picIcon.Image = Nothing
    Else
      picIcon.Image = ico.ToBitmap()
      ico.Dispose()
    End If

  End Sub

  Protected Sub BrowseForFile(ByVal sender As Object, ByVal e As System.EventArgs) _
    Handles btnBrowseTarget.Click, btnBrowseIcon.Click

    Dim ofd As OpenFileDialog = New OpenFileDialog()

    With ofd
      If sender Is btnBrowseTarget Then
        .Title = "Select a target file"
        .Filter = "All files (*.*)|*.*"
      Else
        .Title = "Select an icon file or library"
        .Filter = "Icon files|*.ico;*.dll;*.exe|All files (*.*)|*.*"
      End If
      .FilterIndex = 1
      .CheckFileExists = True
    End With

    If ofd.ShowDialog() = DialogResult.OK Then
      If sender Is btnBrowseTarget Then
        txtTarget.Text = ofd.FileName
      Else
        txtIconFile.Text = ofd.FileName
      End If
    End If

    ofd.Dispose()

  End Sub
End Class
