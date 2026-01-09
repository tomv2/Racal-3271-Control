Option Strict On
Option Explicit On

Imports System
Imports System.IO
Imports System.Drawing
Imports System.Globalization
Imports System.Windows.Forms
Imports System.Collections.Generic

Public Class MainForm
    Inherits Form

    ' ----------------------------
    ' VISA session
    ' ----------------------------
    Private session As VisaMessageSession

    ' ----------------------------
    ' UI Controls (connection)
    ' ----------------------------
    Private cmbResource As ComboBox
    Private btnRefresh As Button
    Private btnConnect As Button
    Private btnDisconnect As Button
    Private txtIdn As TextBox
    Private chkDark As CheckBox

    ' A couple of labels we want to style "muted"
    Private lblHint As Label
    Private lblInfo As Label

    ' ----------------------------
    ' UI Controls (presets + RF controls)
    ' ----------------------------
    Private cmbPreset As ComboBox
    Private btnApplyPreset As Button

    Private cmbMode As ComboBox
    Private nudFreq As NumericUpDown
    Private cmbFreqUnit As ComboBox
    Private nudLevel As NumericUpDown
    Private chkRfOn As CheckBox

    Private btnApply As Button
    Private btnReadBack As Button

    ' ----------------------------
    ' UI Controls (log + manual)
    ' ----------------------------
    Private txtLog As TextBox
    Private txtManual As TextBox
    Private btnSend As Button

    ' ----------------------------
    ' Status bar
    ' ----------------------------
    Private status As StatusStrip
    Private statusLabel As ToolStripStatusLabel

    ' ----------------------------
    ' Theme
    ' ----------------------------
    Private currentTheme As Theme = Themes.Light


    ' ----------------------------
    ' Presets
    ' ----------------------------
    Private presets As List(Of RfPreset)
    Private ReadOnly manualHintText As String = "Manual SCPI command (e.g. *IDN? or CFRQ 2GHZ)"
    Private isManualHintActive As Boolean = False


    ' ----------------------------
    ' Tray icon support
    ' ----------------------------
    Private tray As NotifyIcon
    Private trayMenu As ContextMenuStrip

    ' Simple preset model (kept inside this file so you don't need extra files)
    Private Class RfPreset
        Public Property Name As String
        Public Property FrequencyHz As Double
        Public Property LevelDbm As Double
        Public Property RfOn As Boolean = True
        Public Overrides Function ToString() As String
            Return Name
        End Function
    End Class

    ' ----------------------------
    ' Constructor
    ' ----------------------------
    Public Sub New()
        Me.Text = "Racal 3271 Control (Native VISA)"

        ' --- Set window + taskbar icon (embedded resource method) ---
        Try
            Dim asm = Reflection.Assembly.GetExecutingAssembly()
            Using s = asm.GetManifestResourceStream("Racal3271NativeVisa.racal_rf_logo.ico")
                If s IsNot Nothing Then Me.Icon = New Icon(s)
            End Using
        Catch
            ' Ignore icon errors (app still runs)
        End Try

        Me.MinimumSize = New Size(920, 600)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.Font = New Font("Segoe UI", 10.0F, FontStyle.Regular, GraphicsUnit.Point)

        BuildUi()
        LoadPresets()
        RefreshResources()

        SetupTrayIcon()

        AddHandler Me.FormClosing, AddressOf MainForm_FormClosing
    End Sub

    ' ----------------------------
    ' UI Build
    ' ----------------------------
    Private Sub BuildUi()
        Dim root As New TableLayoutPanel() With {
            .Dock = DockStyle.Fill,
            .ColumnCount = 2,
            .RowCount = 3,
            .Padding = New Padding(16)
        }
        root.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 55.0F))
        root.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 45.0F))
        root.RowStyles.Add(New RowStyle(SizeType.Absolute, 140.0F))
        root.RowStyles.Add(New RowStyle(SizeType.Percent, 55.0F))
        root.RowStyles.Add(New RowStyle(SizeType.Percent, 45.0F))
        Me.Controls.Add(root)

        ' ----------------------------
        ' Connection "card"
        ' ----------------------------
        Dim pnlConn As New Panel() With {.Dock = DockStyle.Fill, .Padding = New Padding(14), .BorderStyle = BorderStyle.None}
        root.Controls.Add(pnlConn, 0, 0)
        root.SetColumnSpan(pnlConn, 2)

        Dim connLayout As New TableLayoutPanel() With {.Dock = DockStyle.Fill, .ColumnCount = 6, .RowCount = 3}
        connLayout.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 90.0F))
        connLayout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
        connLayout.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 110.0F))
        connLayout.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 110.0F))
        connLayout.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 110.0F))
        connLayout.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 120.0F))
        connLayout.RowStyles.Add(New RowStyle(SizeType.Absolute, 40.0F))
        connLayout.RowStyles.Add(New RowStyle(SizeType.Absolute, 40.0F))
        connLayout.RowStyles.Add(New RowStyle(SizeType.Absolute, 40.0F))
        pnlConn.Controls.Add(connLayout)

        connLayout.Controls.Add(New Label() With {.Text = "VISA:", .Dock = DockStyle.Fill, .TextAlign = ContentAlignment.MiddleLeft}, 0, 0)

        cmbResource = New ComboBox() With {.Dock = DockStyle.Fill, .DropDownStyle = ComboBoxStyle.DropDownList}
        connLayout.Controls.Add(cmbResource, 1, 0)
        connLayout.SetColumnSpan(cmbResource, 2)

        btnRefresh = New Button() With {.Text = "Refresh", .Dock = DockStyle.Fill}
        AddHandler btnRefresh.Click, Sub() RefreshResources()
        connLayout.Controls.Add(btnRefresh, 3, 0)

        btnConnect = New Button() With {.Text = "Connect", .Dock = DockStyle.Fill}
        AddHandler btnConnect.Click, Sub() Connect()
        connLayout.Controls.Add(btnConnect, 4, 0)

        btnDisconnect = New Button() With {.Text = "Disconnect", .Dock = DockStyle.Fill, .Enabled = False}
        AddHandler btnDisconnect.Click, Sub() SafeDisconnect()
        connLayout.Controls.Add(btnDisconnect, 5, 0)

        ' Row 1: IDN + Dark toggle (no overlap)
        connLayout.Controls.Add(New Label() With {.Text = "*IDN?:", .Dock = DockStyle.Fill, .TextAlign = ContentAlignment.MiddleLeft}, 0, 1)

        txtIdn = New TextBox() With {.Dock = DockStyle.Fill, .ReadOnly = True, .BorderStyle = BorderStyle.FixedSingle}
        connLayout.Controls.Add(txtIdn, 1, 1)
        connLayout.SetColumnSpan(txtIdn, 4) ' columns 1..4, leaving col 5 for Dark mode

        chkDark = New CheckBox() With {.Text = "Dark mode", .Checked = (currentTheme Is Themes.Dark), .Dock = DockStyle.Fill}
        AddHandler chkDark.CheckedChanged,
            Sub()
                currentTheme = If(chkDark.Checked, Themes.Dark, Themes.Light)
                ApplyTheme(Me)
                RestyleButtons()
                ApplyMutedLabelStyles()
            End Sub
        connLayout.Controls.Add(chkDark, 5, 1)

        ' Row 2: Hint text (muted)
        lblHint = New Label() With {
            .Text = "Native VISA (visa32.dll). Default: GPIB0::10::12::INSTR (E1406A primary=10, secondary=12).",
            .Dock = DockStyle.Fill
        }
        connLayout.Controls.Add(lblHint, 0, 2)
        connLayout.SetColumnSpan(lblHint, 6)

        ' ----------------------------
        ' Controls (left) "card"
        ' ----------------------------
        Dim pnlCtl As New Panel() With {.Dock = DockStyle.Fill, .Padding = New Padding(14), .BorderStyle = BorderStyle.None}
        root.Controls.Add(pnlCtl, 0, 1)
        root.SetRowSpan(pnlCtl, 2)

        ' Option A: presets at the top (new row 0)
        Dim ctlLayout As New TableLayoutPanel() With {.Dock = DockStyle.Top, .AutoSize = True, .ColumnCount = 3, .RowCount = 8}
        ctlLayout.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 130.0F))
        ctlLayout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
        ctlLayout.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 140.0F))
        pnlCtl.Controls.Add(ctlLayout)

        ' Row 0: Presets
        ctlLayout.Controls.Add(New Label() With {.Text = "Preset:", .Dock = DockStyle.Fill, .TextAlign = ContentAlignment.MiddleLeft}, 0, 0)

        cmbPreset = New ComboBox() With {.Dock = DockStyle.Fill, .DropDownStyle = ComboBoxStyle.DropDownList}
        ctlLayout.Controls.Add(cmbPreset, 1, 0)

        btnApplyPreset = New Button() With {.Text = "Load", .Dock = DockStyle.Fill}
        AddHandler btnApplyPreset.Click, Sub() ApplySelectedPreset()
        ctlLayout.Controls.Add(btnApplyPreset, 2, 0)

        ' Row 1: Mode
        ctlLayout.Controls.Add(New Label() With {.Text = "Mode:", .Dock = DockStyle.Fill, .TextAlign = ContentAlignment.MiddleLeft}, 0, 1)
        cmbMode = New ComboBox() With {.Dock = DockStyle.Fill, .DropDownStyle = ComboBoxStyle.DropDownList}
        cmbMode.Items.AddRange(New Object() {"FIXED", "SWEEP (best effort)", "LIST (best effort)"})
        cmbMode.SelectedIndex = 0
        ctlLayout.Controls.Add(cmbMode, 1, 1)
        ctlLayout.SetColumnSpan(cmbMode, 2)

        ' Row 2: Frequency
        ctlLayout.Controls.Add(New Label() With {.Text = "Frequency:", .Dock = DockStyle.Fill, .TextAlign = ContentAlignment.MiddleLeft}, 0, 2)
        nudFreq = New NumericUpDown() With {.Dock = DockStyle.Fill, .DecimalPlaces = 6, .Minimum = 0, .Maximum = Decimal.MaxValue, .ThousandsSeparator = True, .Increment = 1D, .Value = 1D}
        ctlLayout.Controls.Add(nudFreq, 1, 2)

        cmbFreqUnit = New ComboBox() With {.Dock = DockStyle.Fill, .DropDownStyle = ComboBoxStyle.DropDownList}
        cmbFreqUnit.Items.AddRange(New Object() {"Hz", "kHz", "MHz", "GHz"})
        cmbFreqUnit.SelectedItem = "GHz"
        ctlLayout.Controls.Add(cmbFreqUnit, 2, 2)

        ' Row 3: Level
        ctlLayout.Controls.Add(New Label() With {.Text = "Level:", .Dock = DockStyle.Fill, .TextAlign = ContentAlignment.MiddleLeft}, 0, 3)
        nudLevel = New NumericUpDown() With {.Dock = DockStyle.Fill, .DecimalPlaces = 1, .Minimum = -140D, .Maximum = 30D, .Increment = 0.5D, .Value = -20D}
        ctlLayout.Controls.Add(nudLevel, 1, 3)
        ctlLayout.Controls.Add(New Label() With {.Text = "dBm", .Dock = DockStyle.Fill, .TextAlign = ContentAlignment.MiddleLeft}, 2, 3)

        ' Row 4: RF Output
        ctlLayout.Controls.Add(New Label() With {.Text = "RF Output:", .Dock = DockStyle.Fill, .TextAlign = ContentAlignment.MiddleLeft}, 0, 4)
        chkRfOn = New CheckBox() With {.Text = "ON", .Dock = DockStyle.Left, .Checked = True}
        ctlLayout.Controls.Add(chkRfOn, 1, 4)
        ctlLayout.SetColumnSpan(chkRfOn, 2)

        ' Row 5: Apply
        btnApply = New Button() With {.Text = "Apply", .Dock = DockStyle.Top, .Enabled = False}
        AddHandler btnApply.Click, Sub() ApplySettings()
        ctlLayout.Controls.Add(btnApply, 1, 5)
        ctlLayout.SetColumnSpan(btnApply, 2)

        ' Row 6: Readback
        btnReadBack = New Button() With {.Text = "Read Back", .Dock = DockStyle.Top, .Enabled = False}
        AddHandler btnReadBack.Click, Sub() ReadBack()
        ctlLayout.Controls.Add(btnReadBack, 1, 6)
        ctlLayout.SetColumnSpan(btnReadBack, 2)

        ' Row 7: small spacer (optional)
        Dim spacer As New Label() With {.Text = "", .Dock = DockStyle.Top, .Height = 8}
        ctlLayout.Controls.Add(spacer, 0, 7)
        ctlLayout.SetColumnSpan(spacer, 3)

        ' ----------------------------
        ' Log (right) "card"
        ' ----------------------------
        Dim pnlLog As New Panel() With {.Dock = DockStyle.Fill, .Padding = New Padding(14), .BorderStyle = BorderStyle.None}
        root.Controls.Add(pnlLog, 1, 1)

        ' Three-row layout: header + log textbox + manual entry row
        Dim logLayout As New TableLayoutPanel() With {
            .Dock = DockStyle.Fill,
            .ColumnCount = 1,
            .RowCount = 3
}
        logLayout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
        logLayout.RowStyles.Add(New RowStyle(SizeType.Absolute, 34.0F))   ' Header
        logLayout.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))  ' Log
        logLayout.RowStyles.Add(New RowStyle(SizeType.Absolute, 44.0F))  ' Manual entry
        pnlLog.Controls.Add(logLayout)

        Dim lblLog As New Label() With {
            .Text = "Session Log",
            .Dock = DockStyle.Fill,
            .Font = New Font(Me.Font.FontFamily, 11.5F, FontStyle.Bold),
            .TextAlign = ContentAlignment.MiddleLeft
}
        logLayout.Controls.Add(lblLog, 0, 0)

        txtLog = New TextBox() With {
            .Dock = DockStyle.Fill,
            .Multiline = True,
            .ScrollBars = ScrollBars.Vertical,
            .ReadOnly = True,
            .BorderStyle = BorderStyle.FixedSingle
}
        logLayout.Controls.Add(txtLog, 0, 1)

        ' Manual entry row
        Dim manualLayout As New TableLayoutPanel() With {
            .Dock = DockStyle.Fill,
            .ColumnCount = 2,
            .RowCount = 1
}
        manualLayout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
        manualLayout.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 110.0F))
        logLayout.Controls.Add(manualLayout, 0, 2)

        txtManual = New TextBox() With {.Dock = DockStyle.Fill, .BorderStyle = BorderStyle.FixedSingle}
        AddHandler txtManual.Enter, Sub() ClearManualHint()
        AddHandler txtManual.Leave, Sub() SetManualHint()
        AddHandler txtManual.TextChanged,
        Sub()
            ' If user starts typing while hint is active, clear it
            If isManualHintActive AndAlso txtManual.Focused Then
                ClearManualHint()
            End If
        End Sub

        SetManualHint()

        AddHandler txtManual.KeyDown, AddressOf ManualKeyDown
        manualLayout.Controls.Add(txtManual, 0, 0)

        btnSend = New Button() With {.Text = "Send", .Dock = DockStyle.Fill, .Enabled = False}
        AddHandler btnSend.Click, Sub() SendManual()
        manualLayout.Controls.Add(btnSend, 1, 0)



        ' ----------------------------
        ' Info (right bottom) "card"
        ' ----------------------------
        Dim pnlInfo As New Panel() With {.Dock = DockStyle.Fill, .Padding = New Padding(14), .BorderStyle = BorderStyle.None}
        root.Controls.Add(pnlInfo, 1, 2)

        lblInfo = New Label() With {.Dock = DockStyle.Fill}
        lblInfo.Text = "Commands: *IDN?, CFRQ <val><unit>, RFLV <val>DBM, RFLV:ON/OFF, CFRQ?, RFLV?. Manual box supports any command."
        pnlInfo.Controls.Add(lblInfo)

        ' ----------------------------
        ' Status bar
        ' ----------------------------
        status = New StatusStrip()
        statusLabel = New ToolStripStatusLabel("Disconnected")
        status.Items.Add(statusLabel)
        Me.Controls.Add(status)
        status.Dock = DockStyle.Bottom

        ' Apply theme + modern button styles
        ApplyTheme(Me)
        If isManualHintActive Then
            txtManual.ForeColor = Color.Gray
        Else
            txtManual.ForeColor = currentTheme.Text
        End If

        RestyleButtons()
        ApplyMutedLabelStyles()
    End Sub

    ' ----------------------------
    ' Tray icon support
    ' ----------------------------
    Private Sub SetupTrayIcon()
        trayMenu = New ContextMenuStrip()
        trayMenu.Items.Add("Restore", Nothing, Sub() RestoreFromTray())
        trayMenu.Items.Add("Disconnect", Nothing, Sub()
                                                      If session IsNot Nothing AndAlso session.IsOpen Then SafeDisconnect()
                                                  End Sub)
        trayMenu.Items.Add(New ToolStripSeparator())
        trayMenu.Items.Add("Exit", Nothing, Sub() Application.Exit())

        tray = New NotifyIcon() With {
            .Visible = True,
            .Text = "Racal 3271 Control",
            .ContextMenuStrip = trayMenu
        }

        ' Prefer the form icon (already set), fallback to default
        Try
            tray.Icon = If(Me.Icon, SystemIcons.Application)
        Catch
            tray.Icon = SystemIcons.Application
        End Try

        AddHandler tray.DoubleClick, Sub() RestoreFromTray()
    End Sub

    Protected Overrides Sub OnResize(e As EventArgs)
        MyBase.OnResize(e)

        If Me.WindowState = FormWindowState.Minimized Then
            Me.Hide()
            If tray IsNot Nothing Then
                tray.ShowBalloonTip(1200, "Racal 3271 Control", "Running in the background (double-click to restore).", ToolTipIcon.Info)
            End If
        End If
    End Sub

    Private Sub RestoreFromTray()
        Me.Show()
        Me.WindowState = FormWindowState.Normal
        Me.Activate()
    End Sub

    ' ----------------------------
    ' Theme helpers
    ' ----------------------------
    Private Sub ApplyTheme(ctrl As Control)
        ctrl.ForeColor = currentTheme.Text

        If TypeOf ctrl Is Form Then
            ctrl.BackColor = currentTheme.Back
        ElseIf TypeOf ctrl Is Panel OrElse TypeOf ctrl Is TableLayoutPanel OrElse TypeOf ctrl Is StatusStrip Then
            ctrl.BackColor = currentTheme.Panel
        ElseIf TypeOf ctrl Is TextBox OrElse TypeOf ctrl Is ComboBox OrElse TypeOf ctrl Is NumericUpDown Then
            ctrl.BackColor = currentTheme.InputBack
        End If

        For Each c As Control In ctrl.Controls
            ApplyTheme(c)
        Next
    End Sub

    Private Sub ApplyMutedLabelStyles()
        Dim muted As Color =
            If(currentTheme Is Themes.Dark,
               Color.FromArgb(170, 170, 170),
               Color.FromArgb(90, 90, 90))

        If lblHint IsNot Nothing Then lblHint.ForeColor = muted
        If lblInfo IsNot Nothing Then lblInfo.ForeColor = muted
    End Sub

    Private Function BorderColorForTheme() As Color
        If currentTheme Is Themes.Dark Then
            Return Color.FromArgb(90, 90, 90)
        End If
        Return Color.FromArgb(150, 150, 150)
    End Function

    Private Sub StylePrimaryButton(btn As Button)
        If btn Is Nothing Then Return
        btn.UseVisualStyleBackColor = False   ' <-- IMPORTANT
        btn.FlatStyle = FlatStyle.Flat
        btn.FlatAppearance.BorderSize = 0
        btn.BackColor = currentTheme.Accent
        btn.ForeColor = Color.White
        btn.Font = New Font(Me.Font, FontStyle.Bold)
        btn.Height = 40
    End Sub

    Private Sub StyleSecondaryButton(btn As Button)
        If btn Is Nothing Then Return
        btn.UseVisualStyleBackColor = False   ' <-- IMPORTANT
        btn.FlatStyle = FlatStyle.Flat
        btn.FlatAppearance.BorderSize = 1
        btn.FlatAppearance.BorderColor = BorderColorForTheme()
        btn.BackColor = currentTheme.Panel
        btn.ForeColor = currentTheme.Text
        btn.Height = 40
    End Sub

    Private Sub RestyleButtons()
        StyleSecondaryButton(btnRefresh)
        StylePrimaryButton(btnConnect)
        StyleSecondaryButton(btnDisconnect)

        StyleSecondaryButton(btnApplyPreset)

        StylePrimaryButton(btnApply)
        StyleSecondaryButton(btnReadBack)

        StylePrimaryButton(btnSend)
    End Sub

    ' ----------------------------
    ' Logging / status
    ' ----------------------------
    Private Sub Log(msg As String)
        If txtLog Is Nothing Then Return
        txtLog.AppendText($"[{DateTime.Now:HH:mm:ss}] {msg}{Environment.NewLine}")
    End Sub

    Private Sub SetStatus(msg As String)
        If statusLabel Is Nothing Then Return
        statusLabel.Text = msg
    End Sub

    ' ----------------------------
    ' Presets
    ' ----------------------------
    Private Sub LoadPresets()
        presets = New List(Of RfPreset) From {
            New RfPreset With {.Name = "1 GHz @ -20 dBm", .FrequencyHz = 1000000000.0, .LevelDbm = -20, .RfOn = True},
            New RfPreset With {.Name = "2 GHz @ -20 dBm", .FrequencyHz = 2000000000.0, .LevelDbm = -20, .RfOn = True},
            New RfPreset With {.Name = "GPS L1 (1.57542 GHz)", .FrequencyHz = 1575420000.0, .LevelDbm = -110, .RfOn = True},
            New RfPreset With {.Name = "WiFi Ch1 (2.412 GHz)", .FrequencyHz = 2412000000.0, .LevelDbm = -30, .RfOn = True},
            New RfPreset With {.Name = "RF OFF (safe)", .FrequencyHz = 1000000000.0, .LevelDbm = -20, .RfOn = False}
        }

        cmbPreset.DataSource = presets
        cmbPreset.DisplayMember = "Name"
        If presets.Count > 0 Then cmbPreset.SelectedIndex = 0
    End Sub

    Private Sub ApplySelectedPreset()
        Dim p = TryCast(cmbPreset.SelectedItem, RfPreset)
        If p Is Nothing Then Return

        Dim hz As Double = p.FrequencyHz

        If hz >= 1000000000.0 Then
            cmbFreqUnit.SelectedItem = "GHz"
            nudFreq.Value = CDec(hz / 1000000000.0)
        ElseIf hz >= 1000000.0 Then
            cmbFreqUnit.SelectedItem = "MHz"
            nudFreq.Value = CDec(hz / 1000000.0)
        ElseIf hz >= 1000.0 Then
            cmbFreqUnit.SelectedItem = "kHz"
            nudFreq.Value = CDec(hz / 1000.0)
        Else
            cmbFreqUnit.SelectedItem = "Hz"
            nudFreq.Value = CDec(hz)
        End If

        nudLevel.Value = CDec(p.LevelDbm)
        chkRfOn.Checked = p.RfOn

        Log($"Preset loaded: {p.Name}")
    End Sub

    ' ----------------------------
    ' VISA resource scanning
    ' ----------------------------
    Private Sub RefreshResources()
        cmbResource.Items.Clear()
        cmbResource.Items.Add("GPIB0::10::12::INSTR")
        cmbResource.Items.Add("GPIB0::10::0::INSTR")

        Try
            Dim found = VisaMessageSession.FindInstruments("?*INSTR")
            For Each r In found
                If Not cmbResource.Items.Contains(r) Then cmbResource.Items.Add(r)
            Next
            Log($"Found {found.Count} VISA resource(s).")
        Catch ex As Exception
            Log("Resource scan failed: " & ex.Message)
        End Try

        If cmbResource.Items.Count > 0 Then cmbResource.SelectedIndex = 0
    End Sub

    ' ----------------------------
    ' Connect / disconnect
    ' ----------------------------
    Private Sub Connect()
        Dim res As String = If(TryCast(cmbResource.SelectedItem, String), "").Trim()
        If res.Length = 0 Then
            MessageBox.Show("Select a VISA resource first.")
            Return
        End If

        Try
            session = New VisaMessageSession()
            session.Open(res, 5000)

            Dim idn = session.QueryLine("*IDN?")
            txtIdn.Text = idn

            Log("Connected: " & res)
            Log("*IDN? -> " & idn)

            btnConnect.Enabled = False
            btnDisconnect.Enabled = True
            btnApply.Enabled = True
            btnReadBack.Enabled = True
            btnSend.Enabled = True

            RestyleButtons()

            SetStatus("Connected: " & res)
            txtManual.Focus()
        Catch ex As Exception
            SafeDisconnect()
            MessageBox.Show("Connect failed: " & ex.Message, "Racal 3271 Control", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub SafeDisconnect()
        Try
            If session IsNot Nothing Then session.Dispose()
        Catch
        Finally
            session = Nothing
            txtIdn.Text = ""

            btnConnect.Enabled = True
            btnDisconnect.Enabled = False
            btnApply.Enabled = False
            btnReadBack.Enabled = False
            RestyleButtons()
            If btnSend IsNot Nothing Then btnSend.Enabled = False

            SetStatus("Disconnected")
            Log("Disconnected.")
        End Try
    End Sub

    ' ----------------------------
    ' Instrument actions
    ' ----------------------------
    Private Sub ApplySettings()
        If session Is Nothing OrElse Not session.IsOpen Then Return

        Try
            Dim modeText As String = CStr(cmbMode.SelectedItem)
            If modeText.StartsWith("FIXED", StringComparison.OrdinalIgnoreCase) Then
                BestEffort("CFRQ:MODE FIXED")
            ElseIf modeText.StartsWith("SWEEP", StringComparison.OrdinalIgnoreCase) Then
                BestEffort("CFRQ:MODE SWEEP")
            ElseIf modeText.StartsWith("LIST", StringComparison.OrdinalIgnoreCase) Then
                BestEffort("CFRQ:MODE LIST")
            End If

            Dim freqStr As String = nudFreq.Value.ToString(CultureInfo.InvariantCulture)
            Dim unit As String = CStr(cmbFreqUnit.SelectedItem).ToUpperInvariant()
            Dim lvlStr As String = nudLevel.Value.ToString(CultureInfo.InvariantCulture)

            session.WriteLine($"CFRQ {freqStr}{unit}")
            session.WriteLine($"RFLV {lvlStr}DBM")
            session.WriteLine(If(chkRfOn.Checked, "RFLV:ON", "RFLV:OFF"))

            Log($"Applied: CFRQ {freqStr}{unit}, RFLV {lvlStr}DBM, RF {(If(chkRfOn.Checked, "ON", "OFF"))}")
            SetStatus("Applied settings")
        Catch ex As Exception
            MessageBox.Show("Apply failed: " & ex.Message)
        End Try
    End Sub

    Private Sub ReadBack()
        If session Is Nothing OrElse Not session.IsOpen Then Return
        Try
            Dim cf = session.QueryLine("CFRQ?")
            Dim rl = session.QueryLine("RFLV?")
            Log("CFRQ? -> " & cf)
            Log("RFLV? -> " & rl)
            SetStatus("Read back OK")
        Catch ex As Exception
            MessageBox.Show("Readback failed: " & ex.Message)
        End Try
    End Sub

    ' ----------------------------
    ' Manual command send
    ' ----------------------------
    Private Sub ManualKeyDown(sender As Object, e As KeyEventArgs)
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            SendManual()
        End If
    End Sub

    Private Sub SendManual()
        Dim cmd As String = If(txtManual.Text, "").Trim()
        If isManualHintActive Then Return
        If cmd.Length = 0 Then Return


        Try
            Log($">> {cmd}")
            If cmd.EndsWith("?", StringComparison.Ordinal) Then
                Dim resp = session.QueryLine(cmd)
                Log($"<< {resp}")
            Else
                session.WriteLine(cmd)
            End If
        Catch ex As Exception
            Log("(!) Manual command failed: " & ex.Message)
        Finally
            txtManual.SelectAll()
            txtManual.Focus()
        End Try
    End Sub

    Private Sub BestEffort(cmd As String)
        Try
            session.WriteLine(cmd)
        Catch ex As Exception
            Log("(!) Best-effort command failed: " & cmd & " | " & ex.Message)
        End Try
    End Sub

    ' ----------------------------
    ' Cleanup
    ' ----------------------------
    Private Sub MainForm_FormClosing(sender As Object, e As FormClosingEventArgs)
        ' Always disconnect cleanly
        SafeDisconnect()

        ' Dispose tray icon to avoid "ghost" tray icon
        Try
            If tray IsNot Nothing Then
                tray.Visible = False
                tray.Dispose()
            End If
            If trayMenu IsNot Nothing Then trayMenu.Dispose()
        Catch
        End Try
    End Sub

    Private Sub SetManualHint()
        If txtManual Is Nothing Then Return
        If txtManual.TextLength = 0 Then
            isManualHintActive = True
            txtManual.ForeColor = Color.Gray
            txtManual.Text = manualHintText
            txtManual.SelectionStart = 0
            txtManual.SelectionLength = 0
        End If
    End Sub

    Private Sub ClearManualHint()
        If txtManual Is Nothing Then Return
        If isManualHintActive Then
            isManualHintActive = False
            txtManual.Text = ""
            txtManual.ForeColor = currentTheme.Text
        End If
    End Sub


End Class
