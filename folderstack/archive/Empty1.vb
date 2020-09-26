		<DllImport("comctl32.dll", SetLastError := True)> _
		Public Shared Function ImageList_GetIcon(himl As IntPtr, i As Integer, flags As Integer) As IntPtr
	 	End Function




		Private Structure SHFILEINFO
		        Public hIcon As IntPtr            ' : icon
		        Public iIcon As Integer           ' : icondex
		        Public dwAttributes As Integer    ' : SFGAO_ flags
		        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=260)> _
		        Public szDisplayName As String
		        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=80)> _
		        Public szTypeName As String
		End Structure

		Private Declare Auto Function SHGetFileInfo Lib "shell32.dll" _
		        (ByVal pszPath As String, _
		         ByVal dwFileAttributes As Integer, _
		         ByRef psfi As SHFILEINFO, _
		         ByVal cbFileInfo As Integer, _
		         ByVal uFlags As Integer) As IntPtr

		Private Const SHGFI_ICON      = &H100
		Private Const SHGFI_SMALLICON = &H1
		Private Const SHGFI_LARGEICON = &H0    ' Large icon
		Private Const SHGFI_SYSICONINDEX = &H4000    ' Large icon
		Private Const SHGFI_ICONLOCATION = &H1000
		Private Const ILD_TRANSPARENT = &H1
		Private nIndex                = 0


		<ComImportAttribute()> _
		<GuidAttribute("000214eb-0000-0000-c000-000000000046")> _
		<InterfaceTypeAttribute(ComInterfaceType.InterfaceIsIUnknown)> _
		Public Interface IExtractIcon
		<PreserveSig()> _
		Function GetIconLocation(ByVal uFlags As Integer, _
		ByVal szIconFile As IntPtr, _
		ByVal cchMax As Integer, _
		ByRef piIndex As Integer, _
		ByRef pwFlags As Integer) As Integer

		<PreserveSig()> _
		Function Extract(ByVal pszFile As IntPtr, _
		ByVal nIconIndex As Integer, _
		ByVal phiconLarge As IntPtr, _
		ByVal phiconSmall As IntPtr, _
		ByVal nSettingIconSize As Integer) As Integer
		End Interface
