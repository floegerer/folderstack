'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Filename:     ShellLinkNative.vb
' Author:       Mattias Sjögren (mattias@mvps.org)
'               http://www.msjogren.net/dotnet/
'
' Description:  Defines the native types used to manipulate shell shortcuts.
'
' Public types: Enum SLR_FLAGS
'               Enum SLGP_FLAGS
'               Structure WIN32_FIND_DATA[A|W]
'               Interface IPersistFile
'               Interface IShellLink[A|W]
'               (Class ShellLink)
'
'
' Copyright ©2001-2002, Mattias Sjögren
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Imports System
Imports System.Text
Imports System.Runtime.InteropServices


Namespace MSjogren.Samples.ShellLink.VB

  ' IShellLink.Resolve fFlags
  <Flags()> _
  Public Enum SLR_FLAGS
    SLR_NO_UI = &H1
    SLR_ANY_MATCH = &H2
    SLR_UPDATE = &H4
    SLR_NOUPDATE = &H8
    SLR_NOSEARCH = &H10
    SLR_NOTRACK = &H20
    SLR_NOLINKINFO = &H40
    SLR_INVOKE_MSI = &H80
  End Enum

  ' IShellLink.GetPath fFlags
  <Flags()> _
  Public Enum SLGP_FLAGS
    SLGP_SHORTPATH = &H1
    SLGP_UNCPRIORITY = &H2
    SLGP_RAWPATH = &H4
  End Enum

  <StructLayoutAttribute(LayoutKind.Sequential, CharSet:=CharSet.Ansi)> Public Structure WIN32_FIND_DATAA
    Public dwFileAttributes As Integer
    Public ftCreationTime As ComTypes.FILETIME
    Public ftLastAccessTime As ComTypes.FILETIME
    Public ftLastWriteTime As ComTypes.FILETIME
    Public nFileSizeHigh As Integer
    Public nFileSizeLow As Integer
    Public dwReserved0 As Integer
    Public dwReserved1 As Integer
    <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=MAX_PATH)> Public cFileName As String
    <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=14)> Public cAlternateFileName As String
    Private Const MAX_PATH As Integer = 260
  End Structure

  <StructLayoutAttribute(LayoutKind.Sequential, CharSet:=CharSet.Unicode)> Public Structure WIN32_FIND_DATAW
    Public dwFileAttributes As Integer
    Public ftCreationTime As ComTypes.FILETIME
    Public ftLastAccessTime As ComTypes.FILETIME
    Public ftLastWriteTime As ComTypes.FILETIME
    Public nFileSizeHigh As Integer
    Public nFileSizeLow As Integer
    Public dwReserved0 As Integer
    Public dwReserved1 As Integer
    <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=MAX_PATH)> Public cFileName As String
    <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=14)> Public cAlternateFileName As String
    Private Const MAX_PATH As Integer = 260
  End Structure

  < _
    ComImport(), _
    InterfaceType(ComInterfaceType.InterfaceIsIUnknown), _
    Guid("0000010B-0000-0000-C000-000000000046") _
  > _
  Public Interface IPersistFile

#Region "Methods inherited from IPersist"

    Sub GetClassID( _
      <Out()> ByRef pClassID As Guid)

#End Region

    <PreserveSig()> _
    Function IsDirty() As Integer

    Sub Load( _
      <MarshalAs(UnmanagedType.LPWStr)> ByVal pszFileName As String, _
      ByVal dwMode As Integer)

    Sub Save( _
      <MarshalAs(UnmanagedType.LPWStr)> ByVal pszFileName As String, _
      <MarshalAs(UnmanagedType.Bool)> ByVal fRemember As Boolean)

    Sub SaveCompleted( _
      <MarshalAs(UnmanagedType.LPWStr)> ByVal pszFileName As String)

    Sub GetCurFile( _
      ByRef ppszFileName As IntPtr)

  End Interface

  < _
    ComImport(), _
    InterfaceType(ComInterfaceType.InterfaceIsIUnknown), _
    Guid("000214EE-0000-0000-C000-000000000046") _
  > _
  Public Interface IShellLinkA

    Sub GetPath( _
      <Out(), MarshalAs(UnmanagedType.LPStr)> ByVal pszFile As StringBuilder, _
      ByVal cchMaxPath As Integer, _
      <Out()> ByRef pfd As WIN32_FIND_DATAA, _
      ByVal fFlags As SLGP_FLAGS)

    Sub GetIDList( _
      ByRef ppidl As IntPtr)

    Sub SetIDList( _
      ByVal pidl As IntPtr)

    Sub GetDescription( _
      <Out(), MarshalAs(UnmanagedType.LPStr)> ByVal pszName As StringBuilder, _
      ByVal cchMaxName As Integer)

    Sub SetDescription( _
      <MarshalAs(UnmanagedType.LPStr)> ByVal pszName As String)

    Sub GetWorkingDirectory( _
      <Out(), MarshalAs(UnmanagedType.LPStr)> ByVal pszDir As StringBuilder, _
      ByVal cchMaxPath As Integer)

    Sub SetWorkingDirectory( _
      <MarshalAs(UnmanagedType.LPStr)> ByVal pszDir As String)

    Sub GetArguments( _
      <Out(), MarshalAs(UnmanagedType.LPStr)> ByVal pszArgs As StringBuilder, _
      ByVal cchMaxPath As Integer)

    Sub SetArguments( _
      <MarshalAs(UnmanagedType.LPStr)> ByVal pszArgs As String)

    Sub GetHotkey( _
      ByRef pwHotkey As Short)

    Sub SetHotkey( _
      ByVal wHotkey As Short)

    Sub GetShowCmd( _
      ByRef piShowCmd As Integer)

    Sub SetShowCmd( _
      ByVal iShowCmd As Integer)

    Sub GetIconLocation( _
      <Out(), MarshalAs(UnmanagedType.LPStr)> ByVal pszIconPath As StringBuilder, _
      ByVal cchIconPath As Integer, _
      ByRef piIcon As Integer)

    Sub SetIconLocation( _
      <MarshalAs(UnmanagedType.LPStr)> ByVal pszIconPath As String, _
      ByVal iIcon As Integer)

    Sub SetRelativePath( _
      <MarshalAs(UnmanagedType.LPStr)> ByVal pszPathRel As String, _
      ByVal dwReserved As Integer)

    Sub Resolve( _
      ByVal hwnd As IntPtr, _
      ByVal fFlags As SLR_FLAGS)

    Sub SetPath( _
      <MarshalAs(UnmanagedType.LPStr)> ByVal pszFile As String)

  End Interface

  < _
    ComImport(), _
    InterfaceTypeAttribute(ComInterfaceType.InterfaceIsIUnknown), _
    Guid("000214F9-0000-0000-C000-000000000046") _
  > _
  Public Interface IShellLinkW

    Sub GetPath( _
      <Out(), MarshalAs(UnmanagedType.LPWStr)> ByVal pszFile As StringBuilder, _
      ByVal cchMaxPath As Integer, _
      <Out()> ByRef pfd As WIN32_FIND_DATAW, _
      ByVal fFlags As SLGP_FLAGS)

    Sub GetIDList( _
      ByRef ppidl As IntPtr)

    Sub SetIDList( _
      ByVal pidl As IntPtr)

    Sub GetDescription( _
      <Out(), MarshalAs(UnmanagedType.LPWStr)> ByVal pszName As StringBuilder, _
      ByVal cchMaxName As Integer)

    Sub SetDescription( _
      <MarshalAs(UnmanagedType.LPWStr)> ByVal pszName As String)

    Sub GetWorkingDirectory( _
      <Out(), MarshalAs(UnmanagedType.LPWStr)> ByVal pszDir As StringBuilder, _
      ByVal cchMaxPath As Integer)

    Sub SetWorkingDirectory( _
      <MarshalAs(UnmanagedType.LPWStr)> ByVal pszDir As String)

    Sub GetArguments( _
      <Out(), MarshalAs(UnmanagedType.LPWStr)> ByVal pszArgs As StringBuilder, _
      ByVal cchMaxPath As Integer)

    Sub SetArguments( _
      <MarshalAs(UnmanagedType.LPWStr)> ByVal pszArgs As String)

    Sub GetHotkey( _
      ByRef pwHotkey As Short)

    Sub SetHotkey( _
      ByVal wHotkey As Short)

    Sub GetShowCmd( _
      ByRef piShowCmd As Integer)

    Sub SetShowCmd( _
      ByVal iShowCmd As Integer)

    Sub GetIconLocation( _
      <Out(), MarshalAs(UnmanagedType.LPWStr)> ByVal pszIconPath As StringBuilder, _
      ByVal cchIconPath As Integer, _
      ByRef piIcon As Integer)

    Sub SetIconLocation( _
      <MarshalAs(UnmanagedType.LPWStr)> ByVal pszIconPath As String, _
      ByVal iIcon As Integer)

    Sub SetRelativePath( _
      <MarshalAs(UnmanagedType.LPWStr)> ByVal pszPathRel As String, _
      ByVal dwReserved As Integer)

    Sub Resolve( _
      ByVal hwnd As IntPtr, _
      ByVal fFlags As SLR_FLAGS)

    Sub SetPath( _
      <MarshalAs(UnmanagedType.LPWStr)> ByVal pszFile As String)

  End Interface


  ' The following does currently not compile correctly. Use
  ' Type.GetTypeFromCLSID() and Activator.CreateInstance() instead.
  '
  '< _
  '  ComImport(), _
  '  Guid("00021401-0000-0000-C000-000000000046") _
  '> _
  'Public Class ShellLink
  '  'Implements IPersistFile
  '  'Implements IShellLinkA
  '  'Implements IShellLinkW
  'End Class

End Namespace
