Option Explicit On
Option Strict Off

Imports System
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.Runtime.InteropServices
Imports System.Security.Cryptography
Imports System.Xml.Linq
Imports System.Drawing
Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Windows.Forms

Friend NotInheritable Class Connection
    Public FilePath As String
    Public Title As String
    Public Address As String
    Public Kind As String
    Public Edits As ProfileEdits
    Public Settings As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

    Public Function Value(key As String) As String
        Dim result As String = Nothing
        If Settings.TryGetValue(key, result) Then Return result
        Return ""
    End Function

    Public Shared Function Read(file As String) As Connection
        Return FromBytes(file, IO.File.ReadAllBytes(file))
    End Function

    Public Shared Function FromBytes(file As String, bytes As Byte()) As Connection
        Dim item As New Connection With {.FilePath = file}
        If Path.GetExtension(file).Equals(".workspace", StringComparison.OrdinalIgnoreCase) Then
            Using reader As New StreamReader(New MemoryStream(bytes), System.Text.Encoding.UTF8, True)
                item.Address = reader.ReadToEnd().Trim()
            End Using
            item.Title = FeedUri(item.Address).Host
            item.Kind = "Workspace feed"
            Return item
        End If
        Dim encoding As Encoding = Encoding.UTF8
        If bytes.Length > 3 AndAlso bytes(1) = 0 AndAlso bytes(3) = 0 Then encoding = Encoding.Unicode
        If bytes.Length > 3 AndAlso bytes(0) = 0 AndAlso bytes(2) = 0 Then encoding = Encoding.BigEndianUnicode
        Using reader As New StreamReader(New MemoryStream(bytes), encoding, True)
            While Not reader.EndOfStream
                Dim parts = reader.ReadLine().Split(New Char() {":"c}, 3)
                If parts.Length <> 3 Then Continue While
                Dim key = parts(0).Trim()
                If item.Settings.ContainsKey(key) Then
                    If key.Equals("full address", StringComparison.OrdinalIgnoreCase) OrElse
                       key.Equals("remoteapplicationmode", StringComparison.OrdinalIgnoreCase) OrElse
                       key.Equals("remoteapplicationprogram", StringComparison.OrdinalIgnoreCase) Then
                        Throw New FormatException("Duplicate connection setting " & key)
                    End If
                End If
                item.Settings(key) = parts(2).Trim()
            End While
        End Using
        item.Address = item.Value("full address")
        If item.Address.Length = 0 Then Throw New FormatException("The RDP file needs a full address.")
        Dim mode = item.Value("remoteapplicationmode")
        If mode <> "" AndAlso mode <> "0" AndAlso mode <> "1" Then Throw New FormatException("Invalid RemoteApp mode.")
        item.Kind = If(mode = "1", "RemoteApp", "Remote Desktop")
        item.Title = item.Address
        If item.Kind = "RemoteApp" Then
            If item.Value("remoteapplicationprogram") = "" Then Throw New FormatException("The RDP file needs a published application.")
            item.Title = If(item.Value("remoteapplicationname") <> "", item.Value("remoteapplicationname"), item.Value("remoteapplicationprogram"))
        End If
        Return item
    End Function

    Public Shared Function FeedUri(value As String) As Uri
        Dim result As Uri = Nothing
        If Not Uri.TryCreate(value, UriKind.Absolute, result) OrElse result.Scheme <> Uri.UriSchemeHttps OrElse
            result.UserInfo <> "" OrElse result.Fragment <> "" OrElse value.IndexOfAny(New Char() {ChrW(13), ChrW(10)}) >= 0 Then
            Throw New FormatException("Enter an HTTPS workspace feed URL.")
        End If
        Return result
    End Function

    Public Shared Function DesktopEndpoint(value As String, ByRef port As Integer) As String
        Dim match = Regex.Match(value, "\A(?<host>\[[0-9a-fA-F.:]+\]|[\w.-]+)(?::(?<port>[0-9]{1,5}))?\z")
        Dim host = match.Groups("host").Value
        If Not match.Success OrElse Uri.CheckHostName(host) = UriHostNameType.Unknown Then
            Throw New FormatException("Enter a desktop hostname or IP address, with an optional port. Use brackets around IPv6.")
        End If
        If match.Groups("port").Success Then
            port = Integer.Parse(match.Groups("port").Value)
            If port < 1 OrElse port > 65535 Then Throw New FormatException("The port must be between 1 and 65535.")
        End If
        If host.StartsWith("[") Then Return host.Trim("["c, "]"c)
        Return New Globalization.IdnMapping().GetAscii(host)
    End Function

    Public Function NativeRequired() As Boolean
        Return Kind <> "Remote Desktop" OrElse Value("signature") <> "" OrElse Value("gatewayhostname") <> "" OrElse
            Value("loadbalanceinfo") <> "" OrElse Value("alternate shell") <> "" OrElse Value("password 51") <> "" OrElse
            Value("enablerdsaadauth") = "1" OrElse Value("targetisaadjoined") = "1" OrElse Value("administrative session") = "1" OrElse
            Value("use redirection server name") = "1"
    End Function
End Class

Friend NotInheritable Class ProfileEdits
    Public Title As String = ""
    Public Address As String = ""
    Public Resolution As String = ""
    Public User As String = ""
    Public Domain As String = ""
    Public PasswordCipher As String = ""
    Public HasCredentials As Boolean

    Public Shared Function Read(path As String) As ProfileEdits
        If Not File.Exists(path) Then Return New ProfileEdits()
        Dim xml = XElement.Load(path)
        Return New ProfileEdits With {
            .Title = If(CStr(xml.Element("title")), ""), .Address = If(CStr(xml.Element("address")), ""),
            .Resolution = If(CStr(xml.Element("resolution")), ""), .User = If(CStr(xml.Element("user")), ""),
            .Domain = If(CStr(xml.Element("domain")), ""), .PasswordCipher = If(CStr(xml.Element("password")), ""),
            .HasCredentials = If(CStr(xml.Element("credentials")), "") = "true"}
    End Function

    Public Sub Save(path As String)
        Directory.CreateDirectory(IO.Path.GetDirectoryName(path))
        Dim temporary = path & "." & Guid.NewGuid().ToString("N") & ".tmp"
        Try
            Dim xml As New XElement("profile", New XElement("title", Title), New XElement("address", Address),
                New XElement("resolution", Resolution), New XElement("user", User), New XElement("domain", Domain),
                New XElement("password", PasswordCipher), New XElement("credentials", If(HasCredentials, "true", "false")))
            xml.Save(temporary)
            If File.Exists(path) Then
                File.Replace(temporary, path, Nothing)
            Else
                File.Move(temporary, path)
            End If
        Finally
            If File.Exists(temporary) Then File.Delete(temporary)
        End Try
    End Sub

    Public Sub SetPassword(password As String)
        Dim bytes = Encoding.UTF8.GetBytes(password)
        Try
            PasswordCipher = Convert.ToBase64String(ProtectedData.Protect(bytes, Nothing, DataProtectionScope.CurrentUser))
        Finally
            Array.Clear(bytes, 0, bytes.Length)
        End Try
    End Sub

    Public Function ReadPassword() As String
        Dim bytes = ProtectedData.Unprotect(Convert.FromBase64String(PasswordCipher), Nothing, DataProtectionScope.CurrentUser)
        Try
            Return Encoding.UTF8.GetString(bytes)
        Finally
            Array.Clear(bytes, 0, bytes.Length)
        End Try
    End Function
End Class

Friend NotInheritable Class ConnectionBackup
    Private Const MaxProfileBytes As Integer = 4 * 1024 * 1024
    Private Const MaxBackupCharacters As Integer = 64 * 1024 * 1024

    Private NotInheritable Class Entry
        Public Extension As String
        Public Data As Byte()
        Public Edits As ProfileEdits
    End Class

    Public Shared Function EditPath(library As String, file As String) As String
        Using hash = SHA256.Create()
            Dim key = BitConverter.ToString(hash.ComputeHash(Encoding.UTF8.GetBytes(file.ToUpperInvariant()))).Replace("-", "")
            Return Path.Combine(library, "Edits", key & ".xml")
        End Using
    End Function

    Private Shared Iterator Function SavedEntries(library As String) As IEnumerable(Of Entry)
        If Not Directory.Exists(library) Then Return
        For Each file In Directory.EnumerateFiles(library)
            Dim extension = Path.GetExtension(file).ToLowerInvariant()
            If extension <> ".rdp" AndAlso extension <> ".workspace" Then Continue For
            If New FileInfo(file).Length > MaxProfileBytes Then Throw New IOException("A connection profile is larger than 4 MB.")
            Yield New Entry With {.Extension = extension, .Data = IO.File.ReadAllBytes(file),
                                  .Edits = ProfileEdits.Read(EditPath(library, file))}
        Next
    End Function

    Private Shared Function ToXml(item As Entry) As XElement
        Dim edits = item.Edits
        Return New XElement("connection", New XAttribute("type", item.Extension),
            New XElement("original", Convert.ToBase64String(item.Data)),
            New XElement("name", edits.Title), New XElement("address", edits.Address),
            New XElement("resolution", edits.Resolution),
            New XElement("user", If(edits.HasCredentials, edits.User, "")),
            New XElement("domain", If(edits.HasCredentials, edits.Domain, "")),
            New XElement("credentials", If(edits.HasCredentials, "true", "false")))
    End Function

    Private Shared Function Fingerprint(item As Entry) As String
        Using hash = SHA256.Create()
            Return Convert.ToBase64String(hash.ComputeHash(Encoding.UTF8.GetBytes(ToXml(item).ToString(SaveOptions.DisableFormatting))))
        End Using
    End Function

    Public Shared Function Export(library As String, destination As String) As Integer
        Dim backup As New XElement("rd-backup", New XAttribute("version", "1"))
        Dim count As Integer
        Dim size As Long
        For Each item In SavedEntries(library)
            Dim element = ToXml(item)
            Decode(element)
            size += element.ToString(SaveOptions.DisableFormatting).Length
            If size > MaxBackupCharacters - 64 Then Throw New IOException("The connection backup exceeds the size limit.")
            backup.Add(element)
            count += 1
        Next
        Dim temporary = destination & "." & Guid.NewGuid().ToString("N") & ".tmp"
        Try
            backup.Save(temporary, SaveOptions.DisableFormatting)
            If File.Exists(destination) Then
                File.Replace(temporary, destination, Nothing)
            Else
                File.Move(temporary, destination)
            End If
        Finally
            If File.Exists(temporary) Then File.Delete(temporary)
        End Try
        Return count
    End Function

    Private Shared Function Decode(element As XElement) As Entry
        Dim extension = CStr(element.Attribute("type"))
        If element.Name <> "connection" OrElse (extension <> ".rdp" AndAlso extension <> ".workspace") Then
            Throw New FormatException("The backup contains an unsupported connection type.")
        End If
        Dim data = Convert.FromBase64String(If(CStr(element.Element("original")), ""))
        If data.Length = 0 OrElse data.Length > MaxProfileBytes Then Throw New FormatException("Invalid connection profile size.")
        Dim parsed = Connection.FromBytes("profile" & extension, data)
        Dim edits As New ProfileEdits With {.Title = If(CStr(element.Element("name")), ""),
            .Address = If(CStr(element.Element("address")), ""), .Resolution = If(CStr(element.Element("resolution")), ""),
            .User = If(CStr(element.Element("user")), ""), .Domain = If(CStr(element.Element("domain")), ""),
            .HasCredentials = If(CStr(element.Element("credentials")), "") = "true"}
        If edits.Address <> "" Then
            If parsed.Kind = "Workspace feed" Then
                Connection.FeedUri(edits.Address)
            ElseIf Not parsed.NativeRequired() Then
                Dim port = 3389
                Connection.DesktopEndpoint(edits.Address, port)
            Else
                Throw New FormatException("An address override cannot be applied to a Windows-managed profile.")
            End If
        End If
        If edits.Resolution <> "" Then edits.Resolution = ManagerWindow.CanonicalResolution(edits.Resolution)
        Return New Entry With {.Extension = extension, .Data = data, .Edits = edits}
    End Function

    Public Shared Function Import(library As String, source As String, ByRef skipped As Integer) As Integer
        Dim backup As XElement
        Dim readerSettings As New System.Xml.XmlReaderSettings With {
            .DtdProcessing = System.Xml.DtdProcessing.Prohibit, .MaxCharactersInDocument = MaxBackupCharacters, .XmlResolver = Nothing}
        Using reader = System.Xml.XmlReader.Create(source, readerSettings)
            backup = XElement.Load(reader)
        End Using
        If backup.Name <> "rd-backup" OrElse CStr(backup.Attribute("version")) <> "1" Then
            Throw New FormatException("This file is not a supported connection backup.")
        End If

        Dim known As New HashSet(Of String)(StringComparer.Ordinal)
        For Each item In SavedEntries(library)
            known.Add(Fingerprint(item))
        Next
        Dim pending As New List(Of Entry)
        skipped = 0
        For Each element In backup.Elements()
            Dim item = Decode(element)
            If known.Add(Fingerprint(item)) Then
                pending.Add(item)
            Else
                skipped += 1
            End If
        Next
        If pending.Count = 0 Then Return 0

        Directory.CreateDirectory(library)
        Dim created As New List(Of String)
        Try
            For Each item In pending
                Dim file = Path.Combine(library, Guid.NewGuid().ToString("N") & item.Extension)
                Dim sidecar = EditPath(library, file)
                Dim temporary = file & ".tmp"
                Try
                    item.Edits.Save(sidecar)
                    created.Add(sidecar)
                    IO.File.WriteAllBytes(temporary, item.Data)
                    IO.File.Move(temporary, file)
                    created.Add(file)
                Finally
                    If IO.File.Exists(temporary) Then IO.File.Delete(temporary)
                End Try
            Next
        Catch
            For i = created.Count - 1 To 0 Step -1
                File.Delete(created(i))
            Next
            Throw
        End Try
        Return pending.Count
    End Function
End Class

Friend NotInheritable Class DesktopView
    Inherits AxHost

    Private Shared ReadOnly RdpEvents As New Guid("336d5562-efa8-482e-8cb3-c5c0fc7a7db6")
    Private Const RemoteDesktopSizeChange As Integer = 12
    Private ReadOnly sizeChangedHandler As Action(Of Integer, Integer) = AddressOf RecordRemoteSize
    Private eventSource As Object
    Private lastRemoteSize As Size
    Public ReadOnly Property RemoteSize As Size
        Get
            Return lastRemoteSize
        End Get
    End Property

    Protected Overrides Sub CreateSink()
        MyBase.CreateSink()
        Dim client = GetOcx()
        ComEventsHelper.Combine(client, RdpEvents, RemoteDesktopSizeChange, sizeChangedHandler)
        eventSource = client
    End Sub

    Protected Overrides Sub DetachSink()
        Try
            If eventSource IsNot Nothing Then
                ComEventsHelper.Remove(eventSource, RdpEvents, RemoteDesktopSizeChange, sizeChangedHandler)
                eventSource = Nothing
            End If
        Finally
            MyBase.DetachSink()
        End Try
    End Sub

    Private Sub RecordRemoteSize(width As Integer, height As Integer)
        lastRemoteSize = New Size(width, height)
    End Sub

    Public Shared Function ProtocolSize(size As Size) As Size
        Dim width = Math.Min(8192, Math.Max(200, size.Width))
        Return New Size(width - width Mod 2, Math.Min(8192, Math.Max(200, size.Height)))
    End Function

    Public Sub New()
        MyBase.New("A0C63C30-F08D-4AB4-907C-34905D770C7D")
    End Sub

    Public ReadOnly Property ConnectionState As Integer
        Get
            Dim client As Object = GetOcx()
            If client Is Nothing Then Return 0
            Return CInt(client.Connected)
        End Get
    End Property

    Public Sub Configure(item As Connection, size As Size, automatic As Boolean)
        Dim client As Object = GetOcx()
        Dim advanced As Object = client.AdvancedSettings7
        Dim port = 3389
        Dim filePort As Integer
        If Integer.TryParse(item.Value("server port"), filePort) AndAlso filePort > 0 AndAlso filePort <= 65535 Then port = filePort
        client.Server = Connection.DesktopEndpoint(item.Address, port)
        client.UserName = item.Value("username")
        client.Domain = item.Value("domain")
        client.DesktopWidth = size.Width
        client.DesktopHeight = size.Height
        advanced.RDPPort = port
        advanced.EnableCredSspSupport = True
        advanced.AuthenticationLevel = If(item.Value("authentication level") = "1", 1, 2)
        advanced.SmartSizing = automatic
        If item.Edits IsNot Nothing AndAlso item.Edits.HasCredentials Then
            client.UserName = item.Edits.User
            client.Domain = item.Edits.Domain
            If item.Edits.PasswordCipher.Length > 0 Then advanced.ClearTextPassword = item.Edits.ReadPassword()
        End If
    End Sub

    Public Sub SetSmartSizing(enabled As Boolean)
        Dim client As Object = GetOcx()
        Dim advanced As Object = client.AdvancedSettings7
        advanced.SmartSizing = enabled
    End Sub

    Public Sub UpdateDisplay(size As Size, dpi As Integer)
        Dim client As Object = GetOcx()
        Dim physicalWidth = CUInt(Math.Max(10, CInt(size.Width * 25.4 / dpi)))
        Dim physicalHeight = CUInt(Math.Max(10, CInt(size.Height * 25.4 / dpi)))
        Dim scale = CUInt(Math.Max(100, Math.Min(500, CInt(dpi * 100 / 96))))
        Dim deviceScale As UInteger = If(scale >= 180UI, 180UI, If(scale >= 140UI, 140UI, 100UI))
        client.UpdateSessionDisplaySettings(CUInt(size.Width), CUInt(size.Height), physicalWidth, physicalHeight,
                                            0UI, scale, deviceScale)
    End Sub

    Public Sub StartSession()
        Dim client As Object = GetOcx()
        client.Connect()
    End Sub
End Class

Friend NotInheritable Class Session
    Public View As DesktopView
    Public Surface As Panel
    Public ResolutionText As String
    Public Started As DateTime
    Public SeenConnection As Boolean
    Public AppliedSize As Size
    Public AppliedDpi As Integer
    Public RequestedSize As Size
    Public ResolutionDeadline As DateTime
End Class

Friend NotInheritable Class ManagerTheme
    Public Shared ReadOnly Names As String() = {"Midnight", "Daylight", "Zombie Green"}
    Public ReadOnly Name As String
    Public ReadOnly Background As Color
    Public ReadOnly Surface As Color
    Public ReadOnly Workspace As Color
    Public ReadOnly Text As Color
    Public ReadOnly Muted As Color
    Public ReadOnly Accent As Color
    Public ReadOnly Hover As Color

    Private Sub New(name As String, background As String, surface As String, workspace As String,
                    text As String, muted As String, accent As String, hover As String)
        Me.Name = name
        Me.Background = ColorTranslator.FromHtml(background)
        Me.Surface = ColorTranslator.FromHtml(surface)
        Me.Workspace = ColorTranslator.FromHtml(workspace)
        Me.Text = ColorTranslator.FromHtml(text)
        Me.Muted = ColorTranslator.FromHtml(muted)
        Me.Accent = ColorTranslator.FromHtml(accent)
        Me.Hover = ColorTranslator.FromHtml(hover)
    End Sub

    Public Shared Function Named(value As String) As ManagerTheme
        Select Case value
            Case "Daylight"
                Return New ManagerTheme("Daylight", "#F3F6FA", "#DEE7F1", "#FFFFFF", "#18283A", "#425A70", "#285B9E", "#CCDDF2")
            Case "Zombie Green"
                Return New ManagerTheme("Zombie Green", "#111810", "#233221", "#172216", "#EDF7E4", "#B5CDA8", "#A6E76B", "#34492C")
            Case Else
                Return New ManagerTheme("Midnight", "#111824", "#222E40", "#182232", "#EAF1F8", "#ABC0D6", "#8EBBFF", "#334A68")
        End Select
    End Function

    Public Shared Function Load(path As String) As ManagerTheme
        Try
            If File.Exists(path) Then Return Named(File.ReadAllText(path).Trim())
        Catch ex As Exception When TypeOf ex Is IOException OrElse TypeOf ex Is UnauthorizedAccessException
        End Try
        Return Named(Nothing)
    End Function
End Class

Friend NotInheritable Class ThemeMenuColors
    Inherits ProfessionalColorTable

    Private ReadOnly palette As ManagerTheme

    Public Sub New(theme As ManagerTheme)
        palette = theme
        UseSystemColors = False
    End Sub

    Public Overrides ReadOnly Property ToolStripDropDownBackground As Color
        Get
            Return palette.Surface
        End Get
    End Property

    Public Overrides ReadOnly Property ImageMarginGradientBegin As Color
        Get
            Return palette.Surface
        End Get
    End Property

    Public Overrides ReadOnly Property ImageMarginGradientMiddle As Color
        Get
            Return palette.Surface
        End Get
    End Property

    Public Overrides ReadOnly Property ImageMarginGradientEnd As Color
        Get
            Return palette.Surface
        End Get
    End Property

    Public Overrides ReadOnly Property MenuItemSelected As Color
        Get
            Return palette.Hover
        End Get
    End Property

    Public Overrides ReadOnly Property MenuBorder As Color
        Get
            Return palette.Muted
        End Get
    End Property

    Public Overrides ReadOnly Property MenuItemBorder As Color
        Get
            Return palette.Accent
        End Get
    End Property
End Class

Friend NotInheritable Class ManagerWindow
    Inherits Form

    Private Shared ReadOnly appFolder As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "DesktopAndRemoteApp")
    Private ReadOnly library As String = Path.Combine(appFolder, "Profiles")
    Private ReadOnly themePath As String = Path.Combine(appFolder, "theme.txt")
    Private ReadOnly themeMenu As New ContextMenuStrip()
    Private palette As ManagerTheme
    Private ReadOnly tree As New TreeView()
    Private ReadOnly stage As New Panel()
    Private Shared ReadOnly zombieBackground As New Lazy(Of Image)(
        Function()
            Using stream = GetType(ManagerWindow).Assembly.GetManifestResourceStream("ZombieGreen.png")
                If stream Is Nothing Then Return Nothing
                Using artwork = Image.FromStream(stream)
                    Return New Bitmap(artwork)
                End Using
            End Using
        End Function)
    Private ReadOnly rootLayout As New TableLayoutPanel()
    Private ReadOnly split As New SplitContainer()
    Private ReadOnly sidebar As New TableLayoutPanel()
    Private ReadOnly sessionLayout As New TableLayoutPanel()
    Private ReadOnly fullScreenButton As New Button()
    Private ReadOnly exitFullScreenButton As New Button()
    Private ReadOnly pinFullScreenButton As New Button()
    Private ReadOnly revealFullScreenButton As New Button()
    Private ReadOnly fullScreenBar As New FlowLayoutPanel()
    Private ReadOnly fullScreenBarTimer As New Timer With {.Interval = 100}
    Private fullScreenBarPinned As Boolean
    Private hideFullScreenBarAt As DateTime
    Private sessionFullScreen As Boolean
    Private savedBounds As Rectangle
    Private savedWindowState As FormWindowState
    Private savedBorderStyle As FormBorderStyle
    Private ReadOnly heading As New Label()
    Private ReadOnly hint As New Label()
    Private ReadOnly pane As New Panel()
    Private ReadOnly address As New TextBox()
    Private ReadOnly resolution As New ComboBox()
    Private ReadOnly connectButton As New Button()
    Private ReadOnly disconnectButton As New Button()
    Private ReadOnly removeButton As New Button()
    Private ReadOnly applyDisplay As New Button()
    Private ReadOnly displayTimer As New Timer With {.Interval = 250}
    Private ReadOnly connectionMenu As New ContextMenuStrip()
    Private ReadOnly backupMenu As New ContextMenuStrip()

    <DllImport("user32.dll")>
    Private Shared Function GetDpiForWindow(window As IntPtr) As UInteger
    End Function
    Private ReadOnly sessions As New Dictionary(Of String, Session)(StringComparer.OrdinalIgnoreCase)
    Private ReadOnly stateTimer As New Timer With {.Interval = 800}
    Private ReadOnly statusText As New Label()
    Private Const StatusBarHeight As Integer = 21

    Public Sub New()
        Text = "Desktop and RemoteApp"
        Icon = System.Drawing.Icon.ExtractAssociatedIcon(Application.ExecutablePath)
        palette = ManagerTheme.Load(themePath)
        Font = New Font("Segoe UI", 10)
        ClientSize = New Size(1160, 720)
        MinimumSize = New Size(920, 540)
        StartPosition = FormStartPosition.CenterScreen
        BackColor = palette.Background
        ForeColor = palette.Text

        rootLayout.Dock = DockStyle.Fill
        rootLayout.RowCount = 2
        rootLayout.ColumnCount = 1
        rootLayout.Padding = New Padding(12)
        rootLayout.RowStyles.Add(New RowStyle(SizeType.Percent, 100))
        rootLayout.RowStyles.Add(New RowStyle(SizeType.Absolute, StatusBarHeight))
        Me.Controls.Add(rootLayout)
        split.Dock = DockStyle.Fill
        split.SplitterWidth = 7
        split.BackColor = palette.Background
        rootLayout.Controls.Add(split, 0, 0)
        split.FixedPanel = FixedPanel.Panel1
        split.Panel1MinSize = 248
        split.SplitterDistance = 268
        tree.Dock = DockStyle.Fill
        tree.BackColor = palette.Surface
        tree.ForeColor = palette.Text
        tree.LineColor = palette.Muted
        tree.BorderStyle = BorderStyle.None
        tree.HideSelection = False
        tree.ItemHeight = 28
        BuildSidebar()
        split.Panel1.Controls.Add(sidebar)
        split.Panel2.Controls.Add(pane)

        pane.Dock = DockStyle.Fill
        sessionLayout.Dock = DockStyle.Fill
        sessionLayout.ColumnCount = 1
        sessionLayout.RowCount = 3
        sessionLayout.RowStyles.Add(New RowStyle(SizeType.Absolute, 43))
        sessionLayout.RowStyles.Add(New RowStyle(SizeType.Absolute, 38))
        sessionLayout.RowStyles.Add(New RowStyle(SizeType.Percent, 100))
        pane.Controls.Add(sessionLayout)
        heading.Font = New Font("Segoe UI Semibold", 15)
        heading.Dock = DockStyle.Fill
        heading.TextAlign = ContentAlignment.MiddleLeft
        heading.Padding = New Padding(8, 0, 0, 0)
        sessionLayout.Controls.Add(heading, 0, 0)
        Dim controls As New FlowLayoutPanel With {.Dock = DockStyle.Fill, .WrapContents = False}
        sessionLayout.Controls.Add(controls, 0, 1)
        StyleButton(connectButton, "Connect", AddressOf ConnectSelected)
        StyleButton(disconnectButton, "Disconnect", AddressOf DisconnectSelected)
        controls.Controls.Add(connectButton)
        controls.Controls.Add(disconnectButton)
        StyleButton(removeButton, "Remove", AddressOf RemoveSelected)
        controls.Controls.Add(removeButton)
        resolution.Margin = New Padding(8, 5, 0, 0)
        resolution.DropDownStyle = ComboBoxStyle.DropDown
        resolution.AccessibleName = "Desktop resolution"
        resolution.Width = 130
        resolution.Items.AddRange(New Object() {"Auto fit", "1024 × 768", "1280 × 720", "1600 × 900", "1920 × 1080", "2560 × 1440"})
        resolution.SelectedIndex = 0
        controls.Controls.Add(resolution)
        StyleButton(applyDisplay, "Apply display", AddressOf ApplySelectedDisplay)
        controls.Controls.Add(applyDisplay)
        StyleButton(fullScreenButton, "Full screen", AddressOf ToggleSessionFullScreen)
        controls.Controls.Add(fullScreenButton)
        stage.Dock = DockStyle.Fill
        stage.BackgroundImageLayout = ImageLayout.Zoom
        stage.BackColor = palette.Workspace
        sessionLayout.Controls.Add(stage, 0, 2)
        hint.Dock = DockStyle.Fill
        hint.TextAlign = ContentAlignment.MiddleCenter
        hint.ForeColor = palette.Muted
        stage.Controls.Add(hint)
        ConfigureFullScreenBar()
        statusText.TextAlign = ContentAlignment.MiddleLeft
        statusText.Dock = DockStyle.Fill
        statusText.Margin = New Padding(3, 0, 3, 0)
        statusText.AutoEllipsis = True
        statusText.ForeColor = palette.Muted
        rootLayout.Controls.Add(statusText, 0, 1)

        AddHandler tree.AfterSelect, Sub(sender, args)
                                         Render()
                                         LoadResolution()
                                         QueueDisplay()
                                     End Sub
        AddHandler resolution.KeyDown, Sub(sender, args)
                                           If args.KeyCode = Keys.Enter Then
                                               args.SuppressKeyPress = True
                                               Run(AddressOf ApplySelectedDisplay)
                                           End If
                                       End Sub
        AddHandler resolution.SelectionChangeCommitted, Sub(sender, args) Run(AddressOf ApplySelectedDisplay)
        AddHandler stage.SizeChanged, Sub(sender, args)
                                          PositionFullScreenBar()
                                          QueueDisplay()
                                      End Sub
        AddHandler LocationChanged, Sub(sender, args) QueueDisplay()
        AddHandler Resize, Sub(sender, args) QueueDisplay()
        AddHandler displayTimer.Tick, Sub(sender, args)
                                          displayTimer.Stop()
                                          UpdateVisibleAutoDisplay()
                                      End Sub
        AddHandler tree.MouseDown, AddressOf SelectRightClickedNode
        Dim menuConnect = connectionMenu.Items.Add("Connect", Nothing, Sub(sender, args) Run(AddressOf ConnectSelected))
        Dim menuDisconnect = connectionMenu.Items.Add("Disconnect", Nothing, Sub(sender, args) Run(AddressOf DisconnectSelected))
        Dim menuFullScreen = connectionMenu.Items.Add("Full screen", Nothing, Sub(sender, args) Run(AddressOf ToggleSessionFullScreen))
        connectionMenu.Items.Add(New ToolStripSeparator())
        connectionMenu.Items.Add("Properties", Nothing, Sub(sender, args) Run(AddressOf EditProperties))
        connectionMenu.Items.Add("Credentials", Nothing, Sub(sender, args) Run(AddressOf EditCredentials))
        AddHandler connectionMenu.Opening, Sub(sender, args)
                                               Dim item = Selected()
                                               args.Cancel = item Is Nothing
                                               If item IsNot Nothing Then
                                                   menuDisconnect.Enabled = sessions.ContainsKey(item.FilePath)
                                                   menuConnect.Enabled = Not menuDisconnect.Enabled
                                                   Dim active As Session = Nothing
                                                   menuFullScreen.Enabled = sessions.TryGetValue(item.FilePath, active) AndAlso active.SeenConnection
                                               End If
                                           End Sub
        tree.ContextMenuStrip = connectionMenu
        AddHandler tree.NodeMouseDoubleClick, Sub(sender, args) Run(AddressOf ConnectSelected)
        AddHandler tree.KeyDown, Sub(sender, args)
                                    If args.KeyCode = Keys.Enter Then
                                        args.SuppressKeyPress = True
                                        Run(AddressOf ConnectSelected)
                                    End If
                                End Sub
        AddHandler address.KeyDown, Sub(sender, args)
                                       If args.KeyCode = Keys.Enter Then
                                           args.SuppressKeyPress = True
                                           Run(AddressOf AddAddress)
                                       End If
                                   End Sub
        AddHandler Shown, Sub(sender, args)
                              Run(AddressOf ReloadTree)
                              stateTimer.Start()
                          End Sub
        AddHandler stateTimer.Tick, Sub(sender, args) PollSessions()
        AddHandler FormClosed, Sub(sender, args)
                                   stateTimer.Dispose()
                                   displayTimer.Dispose()
                                   fullScreenBarTimer.Dispose()
                                   connectionMenu.Dispose()
                                   backupMenu.Dispose()
                                   themeMenu.Dispose()
                               End Sub
        ApplyTheme()
        Render()
    End Sub

    Protected Overrides Sub OnFormClosing(args As FormClosingEventArgs)
        MyBase.OnFormClosing(args)
        If args.Cancel OrElse args.CloseReason <> CloseReason.UserClosing OrElse sessions.Count = 0 Then Return
        Dim prompt = "Exiting will disconnect " & If(sessions.Count = 1, "the open RDP connection.", sessions.Count & " open RDP connections.")
        args.Cancel = MessageBox.Show(Me, prompt & Environment.NewLine & "Exit the manager?", "Confirm exit",
            MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) <> DialogResult.Yes
    End Sub

    Private Sub BuildSidebar()
        sidebar.Dock = DockStyle.Fill
        sidebar.ColumnCount = 2
        sidebar.RowCount = 6
        sidebar.Padding = New Padding(4, 0, 4, 0)
        sidebar.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 50))
        sidebar.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 50))
        sidebar.RowStyles.Add(New RowStyle(SizeType.Absolute, 25))
        sidebar.RowStyles.Add(New RowStyle(SizeType.Absolute, 32))
        For i = 1 To 3
            sidebar.RowStyles.Add(New RowStyle(SizeType.Absolute, 34))
        Next
        sidebar.RowStyles.Add(New RowStyle(SizeType.Percent, 100))

        Dim addressLabel As New Label With {
            .Text = "Host or feed URL", .Dock = DockStyle.Fill, .TextAlign = ContentAlignment.MiddleLeft}
        sidebar.Controls.Add(addressLabel, 0, 0)
        sidebar.SetColumnSpan(addressLabel, 2)
        address.Dock = DockStyle.Fill
        address.AccessibleName = "Desktop host or workspace feed URL"
        sidebar.Controls.Add(address, 0, 1)
        sidebar.SetColumnSpan(address, 2)

        Dim add = ActionButton("Add", AddressOf AddAddress)
        add.AccessibleDescription = "Save the desktop host or workspace feed entered above."
        Dim refresh = ActionButton("Refresh apps", AddressOf ReloadTree)
        refresh.AccessibleDescription = "Refresh Windows-published RemoteApp shortcuts."
        Dim backupButton As New Button()
        StyleButton(backupButton, "Backup", Sub() backupMenu.Show(backupButton, 0, backupButton.Height))
        backupMenu.Items.Add("Export connections…", Nothing, Sub(sender, args) Run(AddressOf ExportConnections))
        backupMenu.Items.Add("Import connections…", Nothing, Sub(sender, args) Run(AddressOf ImportConnections))
        Dim themeButton As New Button()
        StyleButton(themeButton, "Theme", Sub() themeMenu.Show(themeButton, 0, themeButton.Height))
        For Each themeName In ManagerTheme.Names
            Dim name = themeName
            themeMenu.Items.Add(name, Nothing, Sub(sender, args) Run(Sub() SelectTheme(name)))
        Next

        Dim actions As Button() = {add, ActionButton("Import .rdp", AddressOf ImportFile), refresh,
            ActionButton("Open shortcut", AddressOf PickShortcut), backupButton, themeButton}
        For i = 0 To actions.Length - 1
            Dim button = actions(i)
            button.AutoSize = False
            button.Dock = DockStyle.Fill
            sidebar.Controls.Add(button, i Mod 2, 2 + (i \ 2))
        Next
        tree.Margin = New Padding(3, 6, 3, 0)
        sidebar.Controls.Add(tree, 0, 5)
        sidebar.SetColumnSpan(tree, 2)
    End Sub

    Private Sub SelectTheme(name As String)
        Directory.CreateDirectory(appFolder)
        File.WriteAllText(themePath, name)
        palette = ManagerTheme.Named(name)
        ApplyTheme()
    End Sub

    Private Sub ApplyTheme()
        ColorControls(Me)
        stage.BackgroundImage = If(palette.Name = "Zombie Green", zombieBackground.Value, Nothing)
        If stage.BackgroundImage IsNot Nothing Then hint.BackColor = Color.Transparent
        tree.LineColor = palette.Muted
        heading.ForeColor = palette.Accent
        hint.ForeColor = palette.Muted
        statusText.ForeColor = palette.Muted
        For Each menu As ContextMenuStrip In New ContextMenuStrip() {connectionMenu, backupMenu, themeMenu}
            menu.Renderer = New ToolStripProfessionalRenderer(New ThemeMenuColors(palette))
            menu.ForeColor = palette.Text
            For Each item As ToolStripItem In menu.Items
                item.ForeColor = palette.Text
            Next
        Next
        For Each item As ToolStripMenuItem In themeMenu.Items
            item.Checked = item.Text = palette.Name
        Next
    End Sub

    Private Sub ColorControls(control As Control)
        If TypeOf control Is DesktopView Then Return
        control.ForeColor = palette.Text
        control.BackColor = palette.Background
        If Object.ReferenceEquals(control, stage) OrElse TypeOf control Is TextBox OrElse TypeOf control Is ComboBox Then
            control.BackColor = palette.Workspace
        ElseIf TypeOf control Is Button OrElse TypeOf control Is TreeView OrElse Object.ReferenceEquals(control, fullScreenBar) Then
            control.BackColor = palette.Surface
        End If
        Dim button = TryCast(control, Button)
        If button IsNot Nothing Then
            button.UseVisualStyleBackColor = False
            button.FlatStyle = FlatStyle.Flat
            button.FlatAppearance.BorderSize = 0
            button.FlatAppearance.MouseOverBackColor = palette.Hover
            button.FlatAppearance.MouseDownBackColor = palette.Hover
        End If
        For Each child As Control In control.Controls
            ColorControls(child)
        Next
    End Sub

    Private Function ActionButton(title As String, action As Action) As Button
        Dim button As New Button()
        StyleButton(button, title, action)
        Return button
    End Function

    Private Sub StyleButton(button As Button, title As String, action As Action)
        button.Text = title
        button.AutoSize = True
        button.Height = 30
        button.FlatStyle = FlatStyle.Flat
        button.FlatAppearance.BorderSize = 0
        button.BackColor = palette.Surface
        button.ForeColor = palette.Text
        AddHandler button.Click, Sub(sender, args) Run(action)
    End Sub

    Private Sub Run(action As Action)
        Try
            action()
        Catch ex As Exception
            MessageBox.Show(Me, ex.GetBaseException().Message, Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub ReloadTree()
        Dim current = Selected()
        Dim selectedPath = If(current Is Nothing, "", current.FilePath)
        Dim entries As New List(Of Connection)
        Dim skipped As Integer
        If Directory.Exists(library) Then
            For Each file In Directory.EnumerateFiles(library)
                Dim extension = Path.GetExtension(file).ToLowerInvariant()
                If extension <> ".rdp" AndAlso extension <> ".workspace" Then Continue For
                Try
                    entries.Add(Connection.Read(file))
                Catch ex As FormatException
                    skipped += 1
                Catch ex As IOException
                    skipped += 1
                End Try
            Next
        End If
        Dim programs = Environment.GetFolderPath(Environment.SpecialFolder.Programs)
        If Directory.Exists(programs) Then
            For Each folder In Directory.EnumerateDirectories(programs)
                Dim name = Path.GetFileName(folder)
                If Not name.EndsWith(" (RADC)", StringComparison.OrdinalIgnoreCase) AndAlso
                   Not name.Equals("RemoteApp and Desktop Connections", StringComparison.OrdinalIgnoreCase) Then Continue For
                For Each file In Directory.EnumerateFiles(folder, "*", SearchOption.AllDirectories)
                    Dim extension = Path.GetExtension(file).ToLowerInvariant()
                    If extension <> ".lnk" AndAlso extension <> ".rdp" Then Continue For
                    entries.Add(New Connection With {.FilePath = file, .Title = Path.GetFileNameWithoutExtension(file), .Kind = "Windows published", .Address = name})
                Next
            Next
        End If
        For Each pair In sessions
            Dim path = pair.Key
            If Not entries.Exists(Function(item) String.Equals(item.FilePath, path, StringComparison.OrdinalIgnoreCase)) Then
                entries.Add(ConnectionForSession(pair.Value))
            End If
        Next
        For Each item In entries
            ApplyEdits(item)
        Next
        entries.Sort(Function(a, b) StringComparer.CurrentCultureIgnoreCase.Compare(a.Title, b.Title))
        tree.BeginUpdate()
        Try
            tree.Nodes.Clear()
            Dim desktops = tree.Nodes.Add("Remote desktops")
            Dim workspaces = tree.Nodes.Add("Workspaces and published apps")
            For Each item In entries
                Dim parent = If(item.Kind = "Remote Desktop", desktops, workspaces)
                Dim node = parent.Nodes.Add(NodeTitle(item))
                node.Tag = item
                If String.Equals(item.FilePath, selectedPath, StringComparison.OrdinalIgnoreCase) Then tree.SelectedNode = node
            Next
            desktops.Expand()
            workspaces.Expand()
        Finally
            tree.EndUpdate()
        End Try
        If skipped > 0 Then statusText.Text = skipped & " saved entries could not be read."
        Render()
    End Sub

    Private Shared Function ConnectionForSession(item As Session) As Connection
        Return DirectCast(item.Surface.Tag, Connection)
    End Function

    Private Function NodeTitle(item As Connection) As String
        Dim active As Session = Nothing
        If sessions.TryGetValue(item.FilePath, active) Then Return item.Title & If(active.SeenConnection, "  •", "  …")
        Return item.Title
    End Function

    Private Function Selected() As Connection
        If tree.SelectedNode Is Nothing Then Return Nothing
        Return TryCast(tree.SelectedNode.Tag, Connection)
    End Function

    Private Sub Render()
        Dim item = Selected()
        Dim active As Session = Nothing
        If item IsNot Nothing Then sessions.TryGetValue(item.FilePath, active)
        If sessionFullScreen AndAlso active Is Nothing Then SetSessionFullScreen(False)
        fullScreenButton.Enabled = active IsNot Nothing AndAlso active.SeenConnection
        For Each session In sessions.Values
            session.Surface.Visible = Object.ReferenceEquals(session, active)
        Next
        heading.Text = If(item Is Nothing, "Connections", item.Title)
        hint.Visible = active Is Nothing
        connectButton.Enabled = item IsNot Nothing AndAlso active Is Nothing
        disconnectButton.Enabled = active IsNot Nothing
        removeButton.Enabled = item IsNot Nothing AndAlso active Is Nothing AndAlso item.Kind <> "Windows published"
        resolution.Visible = item IsNot Nothing AndAlso Not item.NativeRequired()
        resolution.Enabled = active Is Nothing OrElse active.SeenConnection
        applyDisplay.Visible = resolution.Visible
        applyDisplay.Enabled = resolution.Enabled
        If item Is Nothing Then
            hint.Text = "Select a connection from the tree or add a desktop host or workspace feed."
        ElseIf item.Kind = "Workspace feed" Then
            connectButton.Text = "Subscribe in Windows"
            hint.Text = item.Address
        ElseIf item.NativeRequired() Then
            connectButton.Text = "Open in Windows"
            hint.Text = If(item.Kind = "Remote Desktop", "Windows opens this profile with its original connection settings.", "Windows opens this published resource in its own window.")
        Else
            connectButton.Text = "Connect"
            hint.Text = "Disconnected"
        End If
        If active IsNot Nothing Then active.Surface.BringToFront()
        If sessionFullScreen Then
            fullScreenBar.BringToFront()
            revealFullScreenButton.BringToFront()
        End If
    End Sub

    Protected Overrides Function ProcessCmdKey(ByRef message As Message, keyData As Keys) As Boolean
        If keyData = Keys.F11 AndAlso (sessionFullScreen OrElse fullScreenButton.Enabled) Then
            Run(AddressOf ToggleSessionFullScreen)
            Return True
        End If
        Return MyBase.ProcessCmdKey(message, keyData)
    End Function

    Private Sub ToggleSessionFullScreen()
        If sessionFullScreen Then
            SetSessionFullScreen(False)
        ElseIf fullScreenButton.Enabled Then
            SetSessionFullScreen(True)
        End If
    End Sub

    Private Sub SetSessionFullScreen(enabled As Boolean)
        If sessionFullScreen = enabled Then Return
        sessionFullScreen = enabled
        SuspendLayout()
        Try
            If enabled Then
                Dim monitorBounds = Screen.FromControl(Me).Bounds
                savedWindowState = WindowState
                savedBounds = If(WindowState = FormWindowState.Normal, Bounds, RestoreBounds)
                savedBorderStyle = FormBorderStyle
                WindowState = FormWindowState.Normal
                FormBorderStyle = FormBorderStyle.None
                Bounds = monitorBounds
            Else
                FormBorderStyle = savedBorderStyle
                Bounds = savedBounds
                WindowState = savedWindowState
            End If

            split.Panel1Collapsed = enabled
            statusText.Visible = Not enabled
            heading.Visible = Not enabled
            sessionLayout.GetControlFromPosition(0, 1).Visible = Not enabled
            rootLayout.Padding = If(enabled, New Padding(0), New Padding(12))
            rootLayout.RowStyles(1).Height = If(enabled, 0, StatusBarHeight)
            sessionLayout.RowStyles(0).Height = If(enabled, 0, 43)
            sessionLayout.RowStyles(1).Height = If(enabled, 0, 38)
            If Not enabled Then
                fullScreenBarTimer.Stop()
                SetFullScreenBarVisible(False)
            End If
        Finally
            ResumeLayout(True)
        End Try
        PositionFullScreenBar()
        If enabled Then
            RevealFullScreenBar()
            fullScreenBarTimer.Start()
            Dim item = Selected()
            Dim active As Session = Nothing
            If item IsNot Nothing AndAlso sessions.TryGetValue(item.FilePath, active) Then active.View.Focus()
        Else
            tree.Focus()
        End If
        QueueDisplay()
    End Sub

#If Smoke Then
    Friend Sub VerifySidebar()
        Dim originalSize = Size
        Dim originalSplit = split.SplitterDistance
        Try
            Size = MinimumSize
            split.SplitterDistance = split.Panel1MinSize
            PerformLayout()
            sidebar.PerformLayout()
            If Not Object.ReferenceEquals(sidebar.Parent, split.Panel1) OrElse sidebar.GetColumnSpan(address) <> 2 OrElse
                tree.Top <= address.Bottom Then
                Throw New InvalidOperationException("Connection controls are outside the sidebar.")
            End If
            For Each control As Control In sidebar.Controls
                If Not sidebar.ClientRectangle.Contains(control.Bounds) Then Throw New InvalidOperationException("A sidebar control is clipped.")
                Dim button = TryCast(control, Button)
                If button IsNot Nothing AndAlso (button.Bottom > tree.Top OrElse
                    TextRenderer.MeasureText(button.Text, button.Font).Width + 6 > button.Width) Then
                    Throw New InvalidOperationException("A sidebar action does not fit above the tree: " & button.Text)
                End If
            Next
        Finally
            Size = originalSize
            split.SplitterDistance = originalSplit
        End Try
    End Sub

    Friend Sub VerifyThemes(preference As String)
        For Each themeName As String In ManagerTheme.Names
            File.WriteAllText(preference, themeName)
            palette = ManagerTheme.Load(preference)
            ApplyTheme()
            If (stage.BackgroundImage IsNot Nothing) <> (themeName = "Zombie Green") Then
                Throw New InvalidOperationException("The background must appear only in Zombie Green.")
            End If
            If stage.BackgroundImage IsNot Nothing AndAlso hint.BackColor <> Color.Transparent Then
                Throw New InvalidOperationException("The idle text hides the theme background.")
            End If
            If palette.Name <> themeName OrElse BackColor <> palette.Background OrElse tree.BackColor <> palette.Surface OrElse
                stage.BackColor <> palette.Workspace OrElse heading.ForeColor <> palette.Accent OrElse
                Not DirectCast(themeMenu.Items(Array.IndexOf(ManagerTheme.Names, themeName)), ToolStripMenuItem).Checked Then
                Throw New InvalidOperationException("Theme did not apply or reload: " & themeName)
            End If
        Next
        Using extracted = System.Drawing.Icon.ExtractAssociatedIcon(Application.ExecutablePath)
            Using expected As New Icon(Icon, extracted.Size), frame = extracted.ToBitmap(), formFrame = expected.ToBitmap()
                If frame.GetPixel(frame.Width \ 2, frame.Height \ 2).ToArgb() <> formFrame.GetPixel(formFrame.Width \ 2, formFrame.Height \ 2).ToArgb() Then
                    Throw New InvalidOperationException("Executable and window icons differ.")
                End If
            End Using
        End Using
    End Sub

    Friend Sub VerifyFullScreen()
        Dim item As New Connection With {.FilePath = "fixture.rdp", .Title = "Display test", .Address = "localhost", .Kind = "Remote Desktop"}
        Dim node = tree.Nodes.Add(item.Title)
        node.Tag = item
        Dim surface As New Panel With {.Dock = DockStyle.Fill, .Tag = item}
        Dim view As New DesktopView With {.Dock = DockStyle.Fill}
        surface.Controls.Add(view)
        stage.Controls.Add(surface)
        sessions.Add(item.FilePath, New Session With {.View = view, .Surface = surface, .SeenConnection = True, .ResolutionText = "Auto fit"})
        tree.SelectedNode = node
        Render()
        Dim normalBounds = Bounds
        Dim normalBorder = FormBorderStyle
        Dim displayBounds = Screen.FromControl(Me).Bounds
        Dim originalHandle = view.Handle
        ToggleSessionFullScreen()
        If Not sessionFullScreen OrElse Not split.Panel1Collapsed OrElse Not exitFullScreenButton.Visible OrElse Bounds <> displayBounds OrElse
            FormBorderStyle <> FormBorderStyle.None OrElse view.IsDisposed OrElse view.Handle <> originalHandle Then
            Throw New InvalidOperationException("Full screen did not retain the RDP view.")
        End If
        If Math.Abs((fullScreenBar.Left * 2 + fullScreenBar.Width) - stage.ClientSize.Width) > 2 Then
            Throw New InvalidOperationException("The full-screen bar was not centered.")
        End If
        ToggleFullScreenBarPin()
        UpdateFullScreenBar(New Point(-1, -1), DateTime.UtcNow.AddSeconds(5))
        If Not fullScreenBar.Visible Then Throw New InvalidOperationException("The pin did not keep the bar visible.")
        ToggleFullScreenBarPin()
        UpdateFullScreenBar(New Point(-1, -1), DateTime.UtcNow.AddSeconds(5))
        If fullScreenBar.Visible OrElse Not revealFullScreenButton.Visible Then
            Throw New InvalidOperationException("The unpinned bar did not hide.")
        End If
        RevealFullScreenBar()
        If Not fullScreenBar.Visible OrElse revealFullScreenButton.Visible Then Throw New InvalidOperationException("The bar did not reveal.")
        ToggleSessionFullScreen()
        If sessionFullScreen OrElse split.Panel1Collapsed OrElse exitFullScreenButton.Visible OrElse Bounds <> normalBounds OrElse
            FormBorderStyle <> normalBorder OrElse view.IsDisposed OrElse view.Handle <> originalHandle Then
            Throw New InvalidOperationException("Exiting full screen did not restore the manager and its RDP view.")
        End If
        Dim active = sessions(item.FilePath)
        active.RequestedSize = New Size(1280, 720)
        active.ResolutionDeadline = DateTime.UtcNow.AddSeconds(-1)
        CheckRequestedDisplay(active)
        If Not active.RequestedSize.IsEmpty OrElse active.ResolutionText <> "Auto fit" OrElse
            view.Dock <> DockStyle.Fill OrElse Not statusText.Text.Contains("did not confirm") Then
            Throw New InvalidOperationException("An unconfirmed resize changed the active display.")
        End If
        If rootLayout.RowStyles(1).Height <> StatusBarHeight Then Throw New InvalidOperationException("The status bar height was not restored.")
        tree.Nodes.Remove(node)
        sessions.Remove(item.FilePath)
        surface.Dispose()
    End Sub
#End If

    Private Sub ConfigureFullScreenBar()
        fullScreenBar.AutoSize = True
        fullScreenBar.AutoSizeMode = AutoSizeMode.GrowAndShrink
        fullScreenBar.WrapContents = False
        fullScreenBar.BackColor = palette.Surface
        fullScreenBar.Padding = New Padding(4, 1, 4, 1)
        fullScreenBar.Visible = False

        StyleButton(exitFullScreenButton, "Exit full screen", AddressOf ToggleSessionFullScreen)
        exitFullScreenButton.AccessibleDescription = "Return to the connection tree. Shortcut F11."
        StyleButton(pinFullScreenButton, "Pin", AddressOf ToggleFullScreenBarPin)
        pinFullScreenButton.AccessibleDescription = "Keep the full-screen controls visible."
        pinFullScreenButton.MinimumSize = New Size(60, 0)
        fullScreenBar.Controls.Add(exitFullScreenButton)
        fullScreenBar.Controls.Add(pinFullScreenButton)

        StyleButton(revealFullScreenButton, "▾", Sub()
                                                   RevealFullScreenBar()
                                                   exitFullScreenButton.Focus()
                                               End Sub)
        revealFullScreenButton.AccessibleName = "Show full-screen controls"
        revealFullScreenButton.AutoSize = False
        revealFullScreenButton.Size = New Size(64, 18)
        revealFullScreenButton.Visible = False
        AddHandler revealFullScreenButton.MouseEnter, Sub(sender, args) RevealFullScreenBar()
        AddHandler fullScreenBarTimer.Tick, Sub(sender, args) UpdateFullScreenBar(stage.PointToClient(Cursor.Position), DateTime.UtcNow)
        stage.Controls.Add(fullScreenBar)
        stage.Controls.Add(revealFullScreenButton)
    End Sub

    Private Sub PositionFullScreenBar()
        If Not sessionFullScreen Then Return
        fullScreenBar.Location = New Point(Math.Max(0, (stage.ClientSize.Width - fullScreenBar.PreferredSize.Width) \ 2), 0)
        revealFullScreenButton.Location = New Point(Math.Max(0, (stage.ClientSize.Width - revealFullScreenButton.Width) \ 2), 0)
    End Sub

    Private Sub RevealFullScreenBar()
        If Not sessionFullScreen Then Return
        hideFullScreenBarAt = DateTime.UtcNow.AddMilliseconds(1400)
        SetFullScreenBarVisible(True)
    End Sub

    Private Sub SetFullScreenBarVisible(visible As Boolean)
        fullScreenBar.Visible = sessionFullScreen AndAlso visible
        revealFullScreenButton.Visible = sessionFullScreen AndAlso Not visible
        PositionFullScreenBar()
        If visible Then
            fullScreenBar.BringToFront()
        ElseIf sessionFullScreen Then
            revealFullScreenButton.BringToFront()
        End If
    End Sub

    Private Sub ToggleFullScreenBarPin()
        fullScreenBarPinned = Not fullScreenBarPinned
        pinFullScreenButton.Text = If(fullScreenBarPinned, "Unpin", "Pin")
        pinFullScreenButton.AccessibleDescription = If(fullScreenBarPinned,
            "Let the full-screen controls hide automatically.", "Keep the full-screen controls visible.")
        RevealFullScreenBar()
        Dim item = Selected()
        Dim active As Session = Nothing
        If item IsNot Nothing AndAlso sessions.TryGetValue(item.FilePath, active) Then active.View.Focus()
    End Sub

    Private Sub UpdateFullScreenBar(pointer As Point, currentTime As DateTime)
        If Not sessionFullScreen OrElse Not fullScreenBar.Visible OrElse fullScreenBarPinned Then Return
        If fullScreenBar.Bounds.Contains(pointer) OrElse fullScreenBar.ContainsFocus Then
            hideFullScreenBarAt = currentTime.AddMilliseconds(1400)
        ElseIf currentTime >= hideFullScreenBarAt Then
            SetFullScreenBarVisible(False)
        End If
    End Sub

    Private Sub SelectFile(file As String)
        For Each group As TreeNode In tree.Nodes
            For Each node As TreeNode In group.Nodes
                If String.Equals(DirectCast(node.Tag, Connection).FilePath, file, StringComparison.OrdinalIgnoreCase) Then
                    tree.SelectedNode = node
                    node.EnsureVisible()
                    Return
                End If
            Next
        Next
    End Sub

    Private Sub Save(bytes As Byte(), extension As String)
        Directory.CreateDirectory(library)
        Dim file = Path.Combine(library, Guid.NewGuid().ToString("N") & extension)
        Dim pending = file & ".tmp"
        Try
            IO.File.WriteAllBytes(pending, bytes)
            If extension = ".rdp" Then Connection.Read(pending)
            IO.File.Move(pending, file)
        Finally
            If IO.File.Exists(pending) Then IO.File.Delete(pending)
        End Try
        ReloadTree()
        SelectFile(file)
    End Sub

    Private Sub AddAddress()
        Dim value = address.Text.Trim()
        If value.Contains("://") Then
            Connection.FeedUri(value)
            Save(Encoding.UTF8.GetBytes(value), ".workspace")
        Else
            Dim port = 3389
            Connection.DesktopEndpoint(value, port)
            Save(Encoding.UTF8.GetBytes("full address:s:" & value & vbCrLf & "remoteapplicationmode:i:0" & vbCrLf), ".rdp")
        End If
        address.Clear()
    End Sub

    Private Sub ImportFile()
        Using dialog As New OpenFileDialog With {.Filter = "Remote Desktop profiles (*.rdp)|*.rdp"}
            If dialog.ShowDialog(Me) = DialogResult.OK Then Save(IO.File.ReadAllBytes(dialog.FileName), ".rdp")
        End Using
    End Sub

    Private Sub ExportConnections()
        Using dialog As New SaveFileDialog With {
            .Filter = "Connection backup (*.rdbackup)|*.rdbackup", .DefaultExt = "rdbackup",
            .AddExtension = True, .FileName = "DesktopAndRemoteApp.rdbackup"}
            If dialog.ShowDialog(Me) <> DialogResult.OK Then Return
            Dim count = ConnectionBackup.Export(library, dialog.FileName)
            statusText.Text = String.Format("Exported {0} saved connection(s).", count)
        End Using
    End Sub

    Private Sub ImportConnections()
        Using dialog As New OpenFileDialog With {.Filter = "Connection backup (*.rdbackup)|*.rdbackup", .CheckFileExists = True}
            If dialog.ShowDialog(Me) <> DialogResult.OK Then Return
            Dim skipped As Integer
            Dim count = ConnectionBackup.Import(library, dialog.FileName, skipped)
            ReloadTree()
            statusText.Text = String.Format("Imported {0} connection(s). Skipped {1} already present.", count, skipped)
        End Using
    End Sub

    Private Sub PickShortcut()
        Using dialog As New OpenFileDialog With {.Filter = "Windows shortcuts and RDP files (*.lnk;*.rdp)|*.lnk;*.rdp",
            .InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Programs), .DereferenceLinks = False}
            If dialog.ShowDialog(Me) <> DialogResult.OK Then Return
            Dim extension = Path.GetExtension(dialog.FileName).ToLowerInvariant()
            If extension <> ".lnk" AndAlso extension <> ".rdp" Then Throw New FormatException("Select a Windows shortcut or RDP file.")
            OpenNative(dialog.FileName)
        End Using
    End Sub

    Private Shared Sub OpenNative(file As String)
        Using launched = Process.Start(New ProcessStartInfo(file) With {.UseShellExecute = True})
        End Using
    End Sub

    Private Sub Subscribe(item As Connection)
        Dim answer = MessageBox.Show(Me, "Copy this URL and open RemoteApp and Desktop Connections? Choose Access RemoteApp and desktops and paste it into Windows." &
                                     vbCrLf & vbCrLf & item.Address, Text, MessageBoxButtons.OKCancel)
        If answer <> DialogResult.OK Then Return
        Clipboard.SetText(item.Address)
        Using launched = Process.Start(Path.Combine(Environment.SystemDirectory, "control.exe"), "/name Microsoft.RemoteAppAndDesktopConnections")
        End Using
    End Sub

    Private Sub SelectRightClickedNode(sender As Object, args As MouseEventArgs)
        If args.Button = MouseButtons.Right Then tree.SelectedNode = tree.GetNodeAt(args.Location)
    End Sub

    Private Function EditPath(item As Connection) As String
        Return ConnectionBackup.EditPath(library, item.FilePath)
    End Function

    Private Sub ApplyEdits(item As Connection)
        Try
            item.Edits = ProfileEdits.Read(EditPath(item))
        Catch ex As Exception When TypeOf ex Is IOException OrElse TypeOf ex Is System.Xml.XmlException OrElse TypeOf ex Is UnauthorizedAccessException
            statusText.Text = "Saved properties could not be read for " & item.Title & "."
            item.Edits = New ProfileEdits()
        End Try
        If item.Edits.Title.Length > 0 Then item.Title = item.Edits.Title
        If item.Edits.Address.Length > 0 AndAlso (Not item.NativeRequired() OrElse item.Kind = "Workspace feed") Then
            item.Address = item.Edits.Address
        End If
    End Sub

    Private Shared Function EditorField(label As String, value As String, Optional masked As Boolean = False) As TextBox
        Return New TextBox With {.AccessibleName = label, .Text = value, .Dock = DockStyle.Fill, .UseSystemPasswordChar = masked}
    End Function

    Private Sub ShowEditor(title As String, note As String, labels As String(), fields As Control(), save As Action)
        Using dialog As New Form With {
            .Text = title, .Font = Font, .ClientSize = New Size(560, 110 + fields.Length * 36),
            .StartPosition = FormStartPosition.CenterParent, .FormBorderStyle = FormBorderStyle.FixedDialog,
            .MaximizeBox = False, .MinimizeBox = False, .ShowInTaskbar = False}
            Dim layout As New TableLayoutPanel With {.Dock = DockStyle.Fill, .Padding = New Padding(12), .ColumnCount = 2, .RowCount = fields.Length + 2}
            layout.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 140))
            layout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100))
            For i = 0 To fields.Length - 1
                layout.RowStyles.Add(New RowStyle(SizeType.Absolute, 36))
                layout.Controls.Add(New Label With {.Text = labels(i), .Dock = DockStyle.Fill, .TextAlign = ContentAlignment.MiddleLeft}, 0, i)
                layout.Controls.Add(fields(i), 1, i)
            Next
            layout.RowStyles.Add(New RowStyle(SizeType.Percent, 100))
            Dim description As New Label With {.Text = note, .Dock = DockStyle.Fill}
            layout.Controls.Add(description, 0, fields.Length)
            layout.SetColumnSpan(description, 2)
            Dim actions As New FlowLayoutPanel With {.Dock = DockStyle.Fill, .FlowDirection = FlowDirection.RightToLeft}
            Dim cancel As New Button With {.Text = "Cancel", .DialogResult = DialogResult.Cancel, .AutoSize = True}
            Dim submit As New Button With {.Text = "Save", .AutoSize = True}
            actions.Controls.Add(cancel)
            actions.Controls.Add(submit)
            layout.RowStyles.Add(New RowStyle(SizeType.Absolute, 35))
            layout.Controls.Add(actions, 0, fields.Length + 1)
            layout.SetColumnSpan(actions, 2)
            dialog.Icon = Icon
            dialog.Controls.Add(layout)
            ColorControls(dialog)
            dialog.CancelButton = cancel
            dialog.AcceptButton = submit
            AddHandler submit.Click, Sub(sender, args)
                                         Try
                                             save()
                                             dialog.DialogResult = DialogResult.OK
                                         Catch ex As Exception
                                             MessageBox.Show(dialog, ex.GetBaseException().Message, title, MessageBoxButtons.OK, MessageBoxIcon.Error)
                                         End Try
                                     End Sub
            dialog.ShowDialog(Me)
        End Using
    End Sub

    Private Sub EditProperties()
        Dim item = Selected()
        If item Is Nothing Then Return
        Dim embedded = Not item.NativeRequired()
        Dim title = EditorField("Name", item.Title)
        Dim host = EditorField("Address", item.Address)
        host.ReadOnly = Not embedded AndAlso item.Kind <> "Workspace feed"
        Dim display = EditorField("Resolution", If(embedded, resolution.Text, "Managed by Windows"))
        display.ReadOnly = Not embedded
        Dim newDisplay As String = Nothing
        Dim saved As Boolean
        ShowEditor("Connection properties", "Address changes take effect the next time you connect.",
            New String() {"Name", "Address", "Resolution"}, New Control() {title, host, display},
            Sub()
                Dim edits = ProfileEdits.Read(EditPath(item))
                edits.Title = title.Text.Trim()
                If edits.Title.Length = 0 Then Throw New FormatException("Enter a connection name.")
                If Not host.ReadOnly Then
                    Dim value = host.Text.Trim()
                    If embedded Then
                        Dim port = 3389
                        Connection.DesktopEndpoint(value, port)
                    Else
                        Connection.FeedUri(value)
                    End If
                    If Not String.Equals(value, item.Address, StringComparison.OrdinalIgnoreCase) Then
                        edits.PasswordCipher = ""
                    End If
                    edits.Address = value
                End If
                If embedded Then edits.Resolution = CanonicalResolution(display.Text)
                edits.Save(EditPath(item))
                If embedded Then newDisplay = edits.Resolution
                saved = True
            End Sub)
        If Not saved Then Return
        ReloadTree()
        If newDisplay IsNot Nothing Then
            resolution.Text = newDisplay
            Dim active As Session = Nothing
            If sessions.TryGetValue(item.FilePath, active) AndAlso active.SeenConnection Then ApplySelectedDisplay()
        End If
    End Sub

    Private Sub EditCredentials()
        Dim item = Selected()
        If item Is Nothing Then Return
        If item.NativeRequired() Then
            If MessageBox.Show(Me, "Windows handles credentials for this connection. Open Windows Credential Manager?",
                               "Credentials", MessageBoxButtons.OKCancel) = DialogResult.OK Then
                Using launched = Process.Start(Path.Combine(Environment.SystemDirectory, "control.exe"), "/name Microsoft.CredentialManager")
                End Using
            End If
            Return
        End If
        Dim edits = ProfileEdits.Read(EditPath(item))
        Dim user = EditorField("User name", If(edits.HasCredentials, edits.User, item.Value("username")))
        Dim domain = EditorField("Domain", If(edits.HasCredentials, edits.Domain, item.Value("domain")))
        Dim originalUser = user.Text
        Dim originalDomain = domain.Text
        Dim password = EditorField("New password", "", True)
        Dim forget As New CheckBox With {.Text = "Forget saved password", .AutoSize = True, .Enabled = edits.PasswordCipher <> ""}
        ShowEditor("Connection credentials", "Credentials apply on the next connection. A blank password keeps the saved password for the same account.",
            New String() {"User name", "Domain", "New password", ""}, New Control() {user, domain, password, forget},
            Sub()
                edits.HasCredentials = True
                edits.User = user.Text.Trim()
                edits.Domain = domain.Text.Trim()
                If password.Text.Length > 0 AndAlso forget.Checked Then Throw New FormatException("Choose a new password or forget the saved password.")
                If password.Text.Length > 0 Then
                    edits.SetPassword(password.Text)
                ElseIf forget.Checked OrElse edits.User <> originalUser OrElse edits.Domain <> originalDomain Then
                    edits.PasswordCipher = ""
                End If
                edits.Save(EditPath(item))
            End Sub)
        ReloadTree()
    End Sub

    Friend Shared Function CanonicalResolution(value As String) As String
        value = value.Trim()
        If value.Equals("Auto fit", StringComparison.OrdinalIgnoreCase) Then Return "Auto fit"
        Dim match = Regex.Match(value, "\A([0-9]{3,4})\s*[xX×]\s*([0-9]{3,4})\z")
        If match.Success Then
            Dim width = Integer.Parse(match.Groups(1).Value)
            Dim height = Integer.Parse(match.Groups(2).Value)
            If width >= 200 AndAlso width <= 8192 AndAlso height >= 200 AndAlso height <= 8192 Then
                Return width & " × " & height
            End If
        End If
        Throw New FormatException("Choose Auto fit or enter a resolution such as 1920 x 1080.")
    End Function

    Private Function DisplaySize(display As String) As Size
        If display = "Auto fit" Then Return DesktopView.ProtocolSize(stage.ClientSize)
        Dim parts = display.Split("×"c)
        Return New Size(Integer.Parse(parts(0).Trim()), Integer.Parse(parts(1).Trim()))
    End Function

    Private Function LiveDisplaySize(display As String) As Size
        Dim size = DisplaySize(display)
        If size.Width Mod 2 <> 0 Then Throw New FormatException("Live RDP resolution needs an even width, such as 1920 × 1080.")
        Return size
    End Function

    Private Sub QueueDisplay()
        If Not IsHandleCreated OrElse IsDisposed OrElse Disposing Then Return
        displayTimer.Stop()
        displayTimer.Start()
    End Sub

    Protected Overrides Sub WndProc(ByRef message As Message)
        MyBase.WndProc(message)
        If message.Msg = &H2E0 OrElse message.Msg = &H7E OrElse message.Msg = &H232 Then QueueDisplay()
    End Sub

    Private Sub UpdateVisibleAutoDisplay()
        If WindowState = FormWindowState.Minimized OrElse Not Visible Then Return
        Dim item = Selected()
        Dim active As Session = Nothing
        If item IsNot Nothing AndAlso sessions.TryGetValue(item.FilePath, active) AndAlso active.SeenConnection AndAlso
            active.ResolutionText = "Auto fit" AndAlso active.RequestedSize.IsEmpty Then
            UpdateAutoDisplay(active)
        End If
    End Sub

    Private Sub ApplySelectedDisplay()
        Dim item = Selected()
        If item Is Nothing OrElse item.NativeRequired() Then Return
        Dim display = CanonicalResolution(resolution.Text)
        Dim active As Session = Nothing
        If Not sessions.TryGetValue(item.FilePath, active) Then
            SaveResolution(item, display)
            resolution.Text = display
            Return
        End If
        If Not active.SeenConnection OrElse active.View.ConnectionState <> 1 Then
            LoadResolution()
            Return
        End If
        Dim size = LiveDisplaySize(display)
        If display <> "Auto fit" AndAlso active.RequestedSize = size Then Return
        active.RequestedSize = Size.Empty
        If display = "Auto fit" Then
            If UpdateAutoDisplay(active) Then SaveResolution(item, display)
        ElseIf active.View.RemoteSize = size Then
            CommitFixedDisplay(item, active, size)
        Else
            Try
                active.View.UpdateDisplay(size, 96)
                active.RequestedSize = size
                active.ResolutionDeadline = DateTime.UtcNow.AddSeconds(6)
                statusText.Text = "Waiting for the remote desktop to confirm " & display & "…"
            Catch ex As COMException
                statusText.Text = "Windows rejected the RDP resize (0x" & ex.ErrorCode.ToString("X8") & ")."
            End Try
        End If
        LoadResolution()
    End Sub

    Private Sub SaveResolution(item As Connection, display As String)
        Dim edits = ProfileEdits.Read(EditPath(item))
        If edits.Resolution = display Then Return
        edits.Resolution = display
        edits.Save(EditPath(item))
        item.Edits = edits
    End Sub

    Private Function UpdateAutoDisplay(active As Session) As Boolean
        Dim size = DisplaySize("Auto fit")
        Dim dpi = Math.Max(96, CInt(GetDpiForWindow(Handle)))
        If active.AppliedSize = size AndAlso active.AppliedDpi = dpi AndAlso active.ResolutionText = "Auto fit" Then Return True
        Try
            If active.View.ConnectionState <> 1 Then Return False
            active.View.SetSmartSizing(True)
        Catch ex As COMException
            statusText.Text = "Windows could not enable Auto fit (0x" & ex.ErrorCode.ToString("X8") & ")."
            Return False
        End Try
        Try
            active.View.UpdateDisplay(size, dpi)
        Catch ex As COMException
            statusText.Text = "Auto fit is scaling locally (RDP resize 0x" & ex.ErrorCode.ToString("X8") & ")."
        End Try
        active.View.Dock = DockStyle.Fill
        active.Surface.AutoScroll = False
        active.ResolutionText = "Auto fit"
        active.AppliedSize = size
        active.AppliedDpi = dpi
        Return True
    End Function

    Private Sub CommitFixedDisplay(item As Connection, active As Session, size As Size)
        Try
            active.View.SetSmartSizing(False)
        Catch ex As COMException
            statusText.Text = "The remote size changed, but Windows could not disable scaling (0x" & ex.ErrorCode.ToString("X8") & ")."
            Return
        End Try
        active.View.Dock = DockStyle.None
        active.Surface.AutoScroll = True
        active.Surface.AutoScrollPosition = Point.Empty
        active.View.Location = Point.Empty
        active.View.Size = size
        active.ResolutionText = size.Width & " × " & size.Height
        active.AppliedSize = size
        active.AppliedDpi = 96
        statusText.Text = "Remote resolution set to " & active.ResolutionText & "."
        SaveResolution(item, active.ResolutionText)
    End Sub

    Private Sub CheckRequestedDisplay(active As Session)
        If active.RequestedSize.IsEmpty Then Return
        Dim selectedItem = Selected()
        Dim sessionItem = ConnectionForSession(active)
        Dim isSelected = selectedItem IsNot Nothing AndAlso String.Equals(selectedItem.FilePath, sessionItem.FilePath, StringComparison.OrdinalIgnoreCase)
        If active.View.RemoteSize = active.RequestedSize Then
            Dim confirmed = active.RequestedSize
            active.RequestedSize = Size.Empty
            Dim item = If(isSelected, selectedItem, sessionItem)
            Run(Sub() CommitFixedDisplay(item, active, confirmed))
        ElseIf DateTime.UtcNow >= active.ResolutionDeadline Then
            active.RequestedSize = Size.Empty
            statusText.Text = sessionItem.Title & " did not confirm the requested resolution."
        Else
            Return
        End If
        If isSelected Then LoadResolution()
        QueueDisplay()
    End Sub

    Private Sub LoadResolution()
        Dim item = Selected()
        If item Is Nothing OrElse item.NativeRequired() Then Return
        Dim active As Session = Nothing
        If sessions.TryGetValue(item.FilePath, active) Then
            resolution.Text = If(active.RequestedSize.IsEmpty, active.ResolutionText,
                active.RequestedSize.Width & " × " & active.RequestedSize.Height)
            Return
        End If
        If item.Edits IsNot Nothing AndAlso item.Edits.Resolution.Length > 0 Then
            resolution.Text = item.Edits.Resolution
            Return
        End If
        Dim width, height As Integer
        If Integer.TryParse(item.Value("desktopwidth"), width) AndAlso Integer.TryParse(item.Value("desktopheight"), height) AndAlso
            width >= 200 AndAlso height >= 200 AndAlso width <= 8192 AndAlso height <= 8192 Then
            resolution.Text = width & " × " & height
        Else
            resolution.Text = "Auto fit"
        End If
    End Sub

    Private Sub RemoveSelected()
        Dim item = Selected()
        If item Is Nothing OrElse item.Kind = "Windows published" OrElse sessions.ContainsKey(item.FilePath) Then Return
        If MessageBox.Show(Me, "Remove " & item.Title & " from the manager?", Text, MessageBoxButtons.YesNo) <> DialogResult.Yes Then Return
        IO.File.Delete(item.FilePath)
        IO.File.Delete(EditPath(item))
        ReloadTree()
    End Sub

    Private Sub ConnectSelected()
        Dim item = Selected()
        If item Is Nothing OrElse sessions.ContainsKey(item.FilePath) Then Return
        If item.Kind = "Workspace feed" Then
            Subscribe(item)
            Return
        End If
        If item.NativeRequired() Then
            OpenNative(item.FilePath)
            Return
        End If
        Dim display = CanonicalResolution(resolution.Text)
        Dim autoFit = display = "Auto fit"
        Dim size = DisplaySize(display)
        Dim surface As New Panel With {.Dock = DockStyle.Fill, .AutoScroll = Not autoFit, .Tag = item}
        Dim view As New DesktopView With {.Size = size}
        If autoFit Then view.Dock = DockStyle.Fill
        surface.Controls.Add(view)
        stage.Controls.Add(surface)
        surface.BringToFront()
        Try
            view.CreateControl()
            view.Configure(item, size, autoFit)
            view.StartSession()
            sessions.Add(item.FilePath, New Session With {.View = view, .Surface = surface, .ResolutionText = display, .Started = DateTime.UtcNow})
        Catch
            surface.Dispose()
            Throw
        End Try
        tree.SelectedNode.Text = NodeTitle(item)
        statusText.Text = "Connecting to " & item.Title & "…"
        Render()
    End Sub

    Private Sub DisconnectSelected()
        Dim item = Selected()
        If item Is Nothing Then Return
        Dim active As Session = Nothing
        If Not sessions.TryGetValue(item.FilePath, active) Then Return
        sessions.Remove(item.FilePath)
        active.Surface.Dispose()
        tree.SelectedNode.Text = NodeTitle(item)
        statusText.Text = "Disconnected from " & item.Title & "."
        Render()
    End Sub

    Private Sub PollSessions()
        Dim completed As New List(Of String)
        Dim changed As Boolean = False
        For Each pair In sessions
            Dim state As Integer
            Try
                state = pair.Value.View.ConnectionState
            Catch ex As InvalidOperationException
                completed.Add(pair.Key)
                Continue For
            Catch ex As Runtime.InteropServices.COMException
                completed.Add(pair.Key)
                Continue For
            End Try
            If state = 1 Then CheckRequestedDisplay(pair.Value)
            If state = 1 AndAlso Not pair.Value.SeenConnection Then
                pair.Value.SeenConnection = True
                statusText.Text = "Connected to " & ConnectionForSession(pair.Value).Title & "."
                changed = True
            ElseIf state = 0 AndAlso (pair.Value.SeenConnection OrElse (DateTime.UtcNow - pair.Value.Started).TotalSeconds > 4) Then
                completed.Add(pair.Key)
            End If
        Next
        For Each key In completed
            Dim entry = sessions(key)
            statusText.Text = "Session ended for " & ConnectionForSession(entry).Title & "."
            sessions.Remove(key)
            entry.Surface.Dispose()
        Next
        If changed OrElse completed.Count > 0 Then
            QueueDisplay()
            For Each group As TreeNode In tree.Nodes
                For Each node As TreeNode In group.Nodes
                    node.Text = NodeTitle(DirectCast(node.Tag, Connection))
                Next
            Next
            Render()
        End If
    End Sub
End Class

Friend Module Program
    <DllImport("user32.dll")>
    Private Function SetProcessDpiAwarenessContext(context As IntPtr) As Boolean
    End Function
    <STAThread>
    Public Sub Main(args As String())
        SetProcessDpiAwarenessContext(New IntPtr(-4))
        Application.EnableVisualStyles()
        Application.SetCompatibleTextRenderingDefault(False)
#If Smoke Then
        If args.Length = 2 AndAlso args(0) = "--verify-control" Then
            Try
                Using form As New Form With {.ShowInTaskbar = False, .Opacity = 0}
                    Using manager As New ManagerWindow With {.ShowInTaskbar = False, .Opacity = 0}
                        manager.Show()
                        manager.VerifySidebar()
                        manager.VerifyFullScreen()
                        manager.VerifyThemes(args(1) & ".theme")
                    End Using
                    form.Show()
                    Dim view As New DesktopView()
                    form.Controls.Add(view)
                    view.CreateControl()
                    view.Configure(New Connection With {.Address = "localhost"}, New Size(1024, 768), True)
                    If view.ConnectionState <> 0 Then Throw New InvalidOperationException("Unexpected connection state.")
                    view.SetSmartSizing(False)
                    If DesktopView.ProtocolSize(New Size(1025, 769)) <> New Size(1024, 769) OrElse
                        DesktopView.ProtocolSize(New Size(8193, 199)) <> New Size(8192, 200) Then
                        Throw New InvalidOperationException("RDP monitor dimensions were not normalized.")
                    End If
                    Dim disconnectedResize As String = "returned successfully"
                    Try
                        view.UpdateDisplay(New Size(1280, 800), 144)
                    Catch failure As COMException
                        If failure.ErrorCode <> &H80004005 AndAlso failure.ErrorCode <> &H8007139F AndAlso failure.ErrorCode <> &H8000FFFF Then Throw
                        disconnectedResize = "HRESULT 0x" & failure.ErrorCode.ToString("X8")
                    End Try
                    Dim fixture = args(1) & ".xml"
                    Dim edits As New ProfileEdits With {.Title = "Fixture", .Address = "example.test:3390", .Resolution = "1280 × 800",
                                                       .User = "fixture", .Domain = "EXAMPLE", .HasCredentials = True}
                    edits.SetPassword("fixture-password")
                    edits.Save(fixture)
                    Dim loaded = ProfileEdits.Read(fixture)
                    If loaded.Title <> edits.Title OrElse loaded.Address <> edits.Address OrElse loaded.Resolution <> edits.Resolution OrElse
                        loaded.User <> edits.User OrElse loaded.Domain <> edits.Domain OrElse Not loaded.HasCredentials OrElse
                        loaded.ReadPassword() <> "fixture-password" OrElse IO.File.ReadAllText(fixture).Contains("fixture-password") Then
                        Throw New InvalidOperationException("Profile and credential round trip failed.")
                    End If
                    loaded.Address = "other.example.test"
                    loaded.PasswordCipher = ""
                    loaded.Save(fixture)
                    If ProfileEdits.Read(fixture).PasswordCipher <> "" Then Throw New InvalidOperationException("Credential clear failed.")
                    VerifyBackup(args(1) & ".backup-test")
                    IO.File.WriteAllText(args(1), "ActiveX display binding and monitor bounds verified. Disconnected resize: " & disconnectedResize & ". Profile options and Windows DPAPI credentials round-tripped. Full screen centered, pinned, hid and revealed its controls while retaining the same RDP control. Backup restored desktop, RemoteApp and workspace profiles and skipped duplicates. All three palettes applied and reloaded, and the executable and window icons matched. Sidebar actions fit at minimum width. An unconfirmed resize kept Auto fit and the current layout.")
                End Using
            Catch ex As Exception
                IO.File.WriteAllText(args(1), ex.ToString())
                Environment.ExitCode = 1
            End Try
            Return
        End If
#End If
        Application.Run(New ManagerWindow())
    End Sub
#If Smoke Then
    Private Sub VerifyBackup(folder As String)
        Dim original = Path.Combine(folder, "origin")
        Dim restored = Path.Combine(folder, "restored")
        Directory.CreateDirectory(original)
        Directory.CreateDirectory(restored)
        Dim desktop = Path.Combine(original, "desktop.rdp")
        Dim remoteApp = Path.Combine(original, "remoteapp.rdp")
        Dim workspace = Path.Combine(original, "work.workspace")
        Dim feed = "https://app.hostedflex.net/RDWeb/Feed/webfeed.aspx"
        Dim rdp = "full address:s:broker.example.test" & vbCrLf & "remoteapplicationmode:i:1" & vbCrLf &
                  "remoteapplicationprogram:s:||Calculator" & vbCrLf & "signature:s:opaque-test-data" & vbCrLf
        IO.File.WriteAllText(desktop, "full address:s:desktop.example.test" & vbCrLf)
        IO.File.WriteAllText(remoteApp, rdp, Encoding.Unicode)
        IO.File.WriteAllText(workspace, feed)
        IO.File.WriteAllText(Path.Combine(restored, "existing.rdp"), "full address:s:keep.example.test" & vbCrLf)

        Dim settings As New ProfileEdits With {.Title = "Office desktop", .Address = "desktop.example.test:3390",
            .Resolution = "1920 × 1080", .User = "office-test", .Domain = "EXAMPLE", .HasCredentials = True}
        settings.SetPassword("nonportable-fixture")
        settings.Save(ConnectionBackup.EditPath(original, desktop))
        Dim backup = Path.Combine(folder, "connections.rdbackup")
        If ConnectionBackup.Export(original, backup) <> 3 Then Throw New InvalidOperationException("Export count mismatch.")
        Dim exported = IO.File.ReadAllText(backup)
        If exported.Contains(settings.PasswordCipher) OrElse exported.Contains("nonportable-fixture") Then
            Throw New InvalidOperationException("Export leaked a password.")
        End If

        Dim skipped As Integer
        If ConnectionBackup.Import(restored, backup, skipped) <> 3 OrElse skipped <> 0 Then
            Throw New InvalidOperationException("Import count mismatch.")
        End If
        If ConnectionBackup.Import(restored, backup, skipped) <> 0 OrElse skipped <> 3 Then
            Throw New InvalidOperationException("Duplicate import was not skipped.")
        End If
        Dim desktopFound, appFound, workspaceFound As Boolean
        For Each file In Directory.GetFiles(restored)
            Dim profile = Connection.Read(file)
            If profile.Kind = "RemoteApp" Then
                appFound = Convert.ToBase64String(IO.File.ReadAllBytes(file)) = Convert.ToBase64String(IO.File.ReadAllBytes(remoteApp))
            ElseIf profile.Kind = "Workspace feed" Then
                workspaceFound = profile.Address = feed
            ElseIf profile.Address = "desktop.example.test" Then
                Dim edits = ProfileEdits.Read(ConnectionBackup.EditPath(restored, file))
                desktopFound = edits.Title = settings.Title AndAlso edits.Address = settings.Address AndAlso
                    edits.Resolution = settings.Resolution AndAlso edits.User = settings.User AndAlso
                    edits.Domain = settings.Domain AndAlso edits.PasswordCipher = ""
            End If
        Next
        If Not desktopFound OrElse Not appFound OrElse Not workspaceFound OrElse Directory.GetFiles(restored).Length <> 4 Then
            Throw New InvalidOperationException("Restored connection data mismatch.")
        End If

        Dim broken = Path.Combine(folder, "bad.rdbackup")
        Dim xml = XElement.Load(backup)
        xml.Add(New XElement("connection", New XAttribute("type", ".exe")))
        xml.Save(broken)
        Dim empty = Path.Combine(folder, "empty")
        Try
            ConnectionBackup.Import(empty, broken, skipped)
            Throw New InvalidOperationException("Invalid backup was accepted.")
        Catch ex As FormatException
            If Directory.Exists(empty) Then Throw New InvalidOperationException("Invalid backup left partial profiles behind.")
        End Try
    End Sub
#End If
End Module
