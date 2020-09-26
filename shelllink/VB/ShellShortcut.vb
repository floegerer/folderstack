'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Filename:     ShellShortcut.vb
' Author:       Mattias Sjögren (mattias@mvps.org)
'               http://www.msjogren.net/dotnet/
'
' Description:  Defines a .NET friendly class, ShellShortcut, for reading
'               and writing shortcuts.
'               Define the conditional compilation symbol UNICODE to use
'               IShellLinkW internally.
'
' Public types: Class ShellShortcut
'
'
' Dependencies: ShellLinkNative.vb
'
'
' Copyright ©2001-2002, Mattias Sjögren
' 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Imports System
Imports System.Diagnostics
Imports System.Drawing
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Windows.Forms


Namespace MSjogren.Samples.ShellLink.VB

  '
  ' .NET friendly wrapper for the ShellLink class
  '
  Public Class ShellShortcut
    Implements IDisposable

    Private Const INFOTIPSIZE As Integer = 1024
    Private Const MAX_PATH As Integer = 260

    Private Const SW_SHOWNORMAL As Integer = 1
    Private Const SW_SHOWMINIMIZED As Integer = 2
    Private Const SW_SHOWMAXIMIZED As Integer = 3
    Private Const SW_SHOWMINNOACTIVE As Integer = 7


#If [UNICODE] Then
    Private m_Link As IShellLinkW
#Else
    Private m_Link As IShellLinkA
#End If
    Private m_sPath As String

    '
    ' linkPath: Path to new or existing shortcut file (.lnk).
    '
    Public Sub New(ByVal linkPath As String)

      Dim pf As IPersistFile
      Dim CLSID_ShellLink As Guid = New Guid("00021401-0000-0000-C000-000000000046")
      Dim tShellLink As Type

      ' Workaround for VB.NET compiler bug with ComImport classes
      '#If [UNICODE] Then
      '      m_Link = CType(New ShellLink(), IShellLinkW)
      '#Else
      '      m_Link = CType(New ShellLink(), IShellLinkA)
      '#End If
      tShellLink = Type.GetTypeFromCLSID(CLSID_ShellLink)
#If [UNICODE] Then
      m_Link = CType(Activator.CreateInstance(tShellLink), IShellLinkW)
#Else
      m_Link = CType(Activator.CreateInstance(tShellLink), IShellLinkA)
#End If


      m_sPath = linkPath

      If File.Exists(linkPath) Then
        pf = CType(m_Link, IPersistFile)
        pf.Load(linkPath, 0)
      End If

    End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
      If m_Link Is Nothing Then Exit Sub
      Marshal.ReleaseComObject(m_Link)
      m_Link = Nothing
    End Sub


    '
    ' Gets or sets the argument list of the shortcut.
    '
    Public Property Arguments() As String
      Get
        Dim sb As StringBuilder = New StringBuilder(INFOTIPSIZE)
        m_Link.GetArguments(sb, sb.Capacity)
        Return sb.ToString()
      End Get
      Set(ByVal Value As String)
        m_Link.SetArguments(Value)
      End Set
    End Property

    '
    ' Gets or sets a description of the shortcut.
    '
    Public Property Description() As String
      Get
        Dim sb As StringBuilder = New StringBuilder(INFOTIPSIZE)
        m_Link.GetDescription(sb, sb.Capacity)
        Return sb.ToString()
      End Get
      Set(ByVal Value As String)
        m_Link.SetDescription(Value)
      End Set
    End Property

    '
    ' Gets or sets the working directory (aka start in directory) of the shortcut.
    '
    Public Property WorkingDirectory() As String
      Get
        Dim sb As StringBuilder = New StringBuilder(MAX_PATH)
        m_Link.GetWorkingDirectory(sb, sb.Capacity)
        Return sb.ToString()
      End Get
      Set(ByVal Value As String)
        m_Link.SetWorkingDirectory(Value)
      End Set
    End Property

    '
    ' If Path returns an empty string, the shortcut is associated with
    ' a PIDL instead, which can be retrieved with IShellLink.GetIDList().
    ' This is beyond the scope of this wrapper class.
    '
    ' Gets or sets the target path of the shortcut.
    '
    Public Property Path() As String
      Get
#If [UNICODE] Then
        Dim wfd As WIN32_FIND_DATAW
#Else
        Dim wfd As WIN32_FIND_DATAA
#End If
        Dim sb As StringBuilder = New StringBuilder(MAX_PATH)
        m_Link.GetPath(sb, sb.Capacity, wfd, SLGP_FLAGS.SLGP_UNCPRIORITY)
        Return sb.ToString()
      End Get
      Set(ByVal Value As String)
        m_Link.SetPath(Value)
      End Set
    End Property

    '
    ' Gets or sets the path of the Icon assigned to the shortcut.
    '
    Public Property IconPath() As String
      Get
        Dim sb As StringBuilder = New StringBuilder(MAX_PATH)
        Dim nIconIdx As Integer
        m_Link.GetIconLocation(sb, sb.Capacity, nIconIdx)
        Return sb.ToString()
      End Get
      Set(ByVal Value As String)
        m_Link.SetIconLocation(Value, IconIndex)
      End Set
    End Property

    '
    ' Gets or sets the index of the Icon assigned to the shortcut.
    ' Set to zero when the IconPath property specifies a .ICO file.
    '
    Public Property IconIndex() As Integer
      Get
        Dim sb As StringBuilder = New StringBuilder(MAX_PATH)
        Dim nIconIdx As Integer
        m_Link.GetIconLocation(sb, sb.Capacity, nIconIdx)
        Return nIconIdx
      End Get
      Set(ByVal Value As Integer)
        m_Link.SetIconLocation(IconPath, Value)
      End Set
    End Property

    '
    ' Retrieves the Icon of the shortcut as it will appear in Explorer.
    ' Use the IconPath and IconIndex properties to change it.
    '
    Public ReadOnly Property Icon() As Icon
      Get
        Dim sb As StringBuilder = New StringBuilder(MAX_PATH)
        Dim nIconIdx As Integer
        Dim hIcon, hInst As IntPtr
        Dim ico, clone As Icon


        m_Link.GetIconLocation(sb, sb.Capacity, nIconIdx)

        hInst = Marshal.GetHINSTANCE(Me.GetType().Module)
        hIcon = Native.ExtractIcon(hInst, sb.ToString(), nIconIdx)
        If hIcon.ToInt32() = 0 Then Return Nothing

        ' Return a cloned Icon, because we have to free the original ourself.
        ico = Icon.FromHandle(hIcon)
        clone = CType(ico.Clone(), Icon)
        ico.Dispose()
        Native.DestroyIcon(hIcon)
        Return clone

      End Get
    End Property

    '
    ' Gets or sets the System.Diagnostics.ProcessWindowStyle value
    ' that decides the initial show state of the shortcut target. Note that
    ' ProcessWindowStyle.Hidden is not a valid property value.
    '
    Public Property WindowStyle() As ProcessWindowStyle
      Get
        Dim nWS As Integer
        m_Link.GetShowCmd(nWS)

        Select Case nWS
          Case SW_SHOWMINIMIZED, SW_SHOWMINNOACTIVE
            Return ProcessWindowStyle.Minimized
          Case SW_SHOWMAXIMIZED
            Return ProcessWindowStyle.Maximized
          Case Else
            Return ProcessWindowStyle.Normal
        End Select
      End Get
      Set(ByVal Value As ProcessWindowStyle)
        Dim nWS As Integer

        Select Case Value
          Case ProcessWindowStyle.Normal
            nWS = SW_SHOWNORMAL
          Case ProcessWindowStyle.Minimized
            nWS = SW_SHOWMINNOACTIVE
          Case ProcessWindowStyle.Maximized
            nWS = SW_SHOWMAXIMIZED
          Case Else ' ProcessWindowStyle.Hidden
            Throw New ArgumentException("Unsupported ProcessWindowStyle value.")
        End Select
        m_Link.SetShowCmd(nWS)
      End Set
    End Property

    '
    ' Gets or sets the hotkey for the shortcut.
    '
    Public Property Hotkey() As Keys
      Get
        Dim wHotkey As Short
        Dim dwHotkey As Integer

        m_Link.GetHotkey(wHotkey)

        '
        ' Convert from IShellLink 16-bit format to Keys enumeration 32-bit value
        ' IShellLink: &HMMVK
        ' Keys:  &H00MM00VK        
        '   MM = Modifier (Alt, Control, Shift)
        '   VK = Virtual key code
        '   
        dwHotkey = (wHotkey And &HFF00I) * &H100I Or (wHotkey And &HFFI)
        Return CType(dwHotkey, Keys)

      End Get
      Set(ByVal Value As Keys)
        Dim wHotkey As Short

        If (Value And Keys.Modifiers) = 0 Then
          Throw New ArgumentException("Hotkey must include a modifier key.")
        End If

        '
        ' Convert from Keys enumeration 32-bit value to IShellLink 16-bit format
        ' IShellLink: &HMMVK
        ' Keys:  &H00MM00VK        
        '   MM = Modifier (Alt, Control, Shift)
        '   VK = Virtual key code
        '   
        wHotkey = CShort(CInt(Value And Keys.Modifiers) \ &H100) Or CShort(Value And Keys.KeyCode)
        m_Link.SetHotkey(wHotkey)

      End Set
    End Property

    '
    ' Saves the shortcut to disk.
    '
    Public Sub Save()
      Dim pf As IPersistFile = CType(m_Link, IPersistFile)
      pf.Save(m_sPath, True)
    End Sub

    '
    ' Returns a reference to the internal ShellLink object,
    ' which can be used to perform more advanced operations
    ' not supported by this wrapper class, by using the
    ' IShellLink interface directly.
    '
    Public ReadOnly Property ShellLink() As Object
      Get
        Return m_Link
      End Get
    End Property


#Region "Native Win32 API functions"
    Private Class Native

      <DllImport("shell32.dll", CharSet:=CharSet.Auto)> _
      Public Shared Function ExtractIcon(ByVal hInst As IntPtr, ByVal lpszExeFileName As String, ByVal nIconIndex As Integer) As IntPtr
      End Function

      <DllImport("user32.dll")> _
      Public Shared Function DestroyIcon(ByVal hIcon As IntPtr) As Boolean
      End Function

    End Class
#End Region

  End Class
End Namespace