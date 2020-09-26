'
' Created by SharpDevelop.
' User: code2
' Date: 14.03.2012
' Time: 00:36
'
' To change this template use Tools | Options | Coding | Edit Standard Headers.
'
Partial Class MainForm
	Inherits System.Windows.Forms.Form

	''' <summary>
	''' Designer variable used to keep track of non-visual components.
	''' </summary>
	Private components As System.ComponentModel.IContainer

	''' <summary>
	''' Disposes resources used by the form.
	''' </summary>
	''' <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
	Protected Overrides Sub Dispose(ByVal disposing As Boolean)
		If disposing Then
			If components IsNot Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(disposing)
	End Sub

	''' <summary>
	''' This method is required for Windows Forms designer support.
	''' Do not change the method contents inside the source code editor. The Forms designer might
	''' not be able to load this method if it was changed manually.
	''' </summary>
	Private Sub InitializeComponent()
		Me.components = New System.ComponentModel.Container()
		Me.ListContainer = New System.Windows.Forms.Panel()
		Me.ListView = New System.Windows.Forms.ListView()
		Me.CloseTimer = New System.Windows.Forms.Timer(Me.components)
		Me.KillTimer = New System.Windows.Forms.Timer(Me.components)
		Me.Startup = New System.Windows.Forms.Timer(Me.components)
		Me.OpacitySteps = New System.Windows.Forms.Timer(Me.components)
		Me.ListContainer.SuspendLayout
		Me.SuspendLayout
		'
		'ListContainer
		'
		Me.ListContainer.BackColor = System.Drawing.Color.White
		Me.ListContainer.Controls.Add(Me.ListView)
		Me.ListContainer.Dock = System.Windows.Forms.DockStyle.Fill
		Me.ListContainer.Location = New System.Drawing.Point(0, 0)
		Me.ListContainer.Margin = New System.Windows.Forms.Padding(0)
		Me.ListContainer.MinimumSize = New System.Drawing.Size(10, 10)
		Me.ListContainer.Name = "ListContainer"
		Me.ListContainer.Padding = New System.Windows.Forms.Padding(32)
		Me.ListContainer.Size = New System.Drawing.Size(224, 224)
		Me.ListContainer.TabIndex = 2
		AddHandler Me.ListContainer.MouseDown, AddressOf Me.WinKill
		AddHandler Me.ListContainer.MouseLeave, AddressOf Me.WinQuit
		'
		'ListView
		'
		Me.ListView.BackColor = System.Drawing.Color.White
		Me.ListView.BackgroundImageTiled = true
		Me.ListView.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.ListView.CausesValidation = false
		Me.ListView.Cursor = System.Windows.Forms.Cursors.Default
		Me.ListView.Dock = System.Windows.Forms.DockStyle.Fill
		Me.ListView.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0,Byte))
		Me.ListView.ForeColor = System.Drawing.Color.FromArgb(CType(CType(64,Byte),Integer), CType(CType(64,Byte),Integer), CType(CType(64,Byte),Integer))
		Me.ListView.FullRowSelect = true
		Me.ListView.GridLines = true
		Me.ListView.ImeMode = System.Windows.Forms.ImeMode.Off
		Me.ListView.Location = New System.Drawing.Point(32, 32)
		Me.ListView.Margin = New System.Windows.Forms.Padding(0)
		Me.ListView.MinimumSize = New System.Drawing.Size(10, 10)
		Me.ListView.MultiSelect = false
		Me.ListView.Name = "ListView"
		Me.ListView.Scrollable = false
		Me.ListView.ShowItemToolTips = true
		Me.ListView.Size = New System.Drawing.Size(160, 160)
		Me.ListView.Sorting = System.Windows.Forms.SortOrder.Ascending
		Me.ListView.TabIndex = 0
		Me.ListView.TabStop = false
		Me.ListView.TileSize = New System.Drawing.Size(32, 32)
		Me.ListView.UseCompatibleStateImageBehavior = false
		Me.ListView.View = System.Windows.Forms.View.List

		AddHandler Me.ListView.MouseClick, AddressOf Me.WinLaunch
		AddHandler Me.ListView.MouseDown, AddressOf Me.WinDelayKill
		AddHandler Me.ListView.MouseEnter, AddressOf Me.WinResume
		'
		'CloseTimer
		'
		Me.CloseTimer.Interval = 200
		AddHandler Me.CloseTimer.Tick, AddressOf Me.WinKill
		'
		'KillTimer
		'
		AddHandler Me.KillTimer.Tick, AddressOf Me.WinKill
		'
		'Startup
		'
		Me.Startup.Interval = 1400
		AddHandler Me.Startup.Tick, AddressOf Me.StartupTick
		'
		'OpacitySteps
		'
		Me.OpacitySteps.Enabled = true
		Me.OpacitySteps.Interval = 15
		AddHandler Me.OpacitySteps.Tick, AddressOf Me.OpacityStepsTick
		'
		'MainForm
		'
		Me.AutoScaleDimensions = New System.Drawing.SizeF(6!, 13!)
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.AutoSize = true
		Me.BackColor = System.Drawing.Color.White
		Me.ClientSize = New System.Drawing.Size(224, 224)
		Me.ControlBox = false
		Me.Controls.Add(Me.ListContainer)
		Me.DoubleBuffered = true
		Me.IsMdiContainer = true
		Me.Location = New System.Drawing.Point(3, 23)
		Me.MaximizeBox = false
		Me.MinimizeBox = false
		Me.MinimumSize = New System.Drawing.Size(240, 240)
		Me.Name = "MainForm"
		Me.ShowIcon = false
		Me.ShowInTaskbar = false
		Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
		Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
		Me.TopMost = true
		Me.TransparencyKey = System.Drawing.Color.Fuchsia
		AddHandler MouseDown, AddressOf Me.WinResume
		AddHandler MouseEnter, AddressOf Me.WinResume
		AddHandler MouseLeave, AddressOf Me.WinQuit
		AddHandler MouseUp, AddressOf Me.WinResume
		Me.ListContainer.ResumeLayout(false)
		Me.ResumeLayout(false)
	End Sub
	Private OpacitySteps As System.Windows.Forms.Timer
	Private Startup As System.Windows.Forms.Timer
	Private KillTimer As System.Windows.Forms.Timer
	Private CloseTimer As System.Windows.Forms.Timer
	Private ListView As System.Windows.Forms.ListView
	Private ListContainer As System.Windows.Forms.Panel

	Sub WinQuit(sender As Object, e As EventArgs)
		CloseTimer.Start
	End Sub

	Sub WinResume(sender As Object, e As EventArgs)
		CloseTimer.Stop
		Startup.Stop
	End Sub

	Sub WinKill(sender As Object, e As EventArgs)
		Application.Exit
	End Sub


	Sub WinLaunch(sender As Object, e As EventArgs)



		KillTimer.Stop


		'MsgBox(WinPath.Remove(WinPath.LastIndexOf("\"),9999999999999))

		Dim TextOrg As String
		TextOrg = ListView.FocusedItem.Text

		Dim Text1 As String
		Text1 = ListView.FocusedItem.Text.Remove(0,1)
		Text1 = Text1.Remove(Text1.Length-1,1)

		Try

		If ListView.FocusedItem.Text.StartsWith(" ") And Not TextOrg.Contains("All Shortcuts")  Then
			System.Diagnostics.Process.Start(Application.ExecutablePath(), ControlChars.Quote + WinPath + "\" + Text1 + ControlChars.Quote)
			'"tiles small " +

		Else If TextOrg.StartsWith(" ") And TextOrg.Contains("All Shortcuts") Then
			System.Diagnostics.Process.Start(Environment.GetEnvironmentVariable("WINDIR") + "\explorer.exe", WinPath)

'		Else If TextOrg.StartsWith(" ") And TextOrg.Contains("Parent Folder") Then
'			System.Diagnostics.Process.Start(Application.ExecutablePath(), ControlChars.Quote + WinPath + "\" + Text1 + ControlChars.Quote)
		Else

			If SettingView Is "Tiles" Then
		'MsgBox(WinPath + "\" +  listView1.FocusedItem.Text.Remove(0,2) + ".lnk" )
		System.Diagnostics.Process.Start(WinPath + "\" +  ListView.FocusedItem.Text.Remove(0,2) + ".lnk")
		Else
		System.Diagnostics.Process.Start(WinPath + "\" +  ListView.FocusedItem.Text + ".lnk")
		End If
		End If

		Catch
		MsgBox("Could not find the specified file/path/shortcut...")
		End Try

		Application.Exit




	End Sub

	Sub WinDelayKill(sender As Object, e As MouseEventArgs)
		KillTimer.Start
	End Sub

	Sub StartupTick(sender As Object, e As EventArgs)
		Application.Exit
	End Sub
End Class
