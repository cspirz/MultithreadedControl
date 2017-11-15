Imports System.Text
Imports System.IO

<Assembly: log4net.Config.XmlConfigurator(ConfigFile:="bulklogger.config", Watch:=True)>


<ComClass(BackgroundWorker.ClassId, BackgroundWorker.InterfaceId, BackgroundWorker.EventsId)>
Public Class BackgroundWorker


#Region "VB6 Interop Code"

#If COM_INTEROP_ENABLED Then

#Region "COM Registration"

    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.

    Public Const ClassId As String = "f69a40ea-6ca8-4c63-a9d0-33c7e7434618"
    Public Const InterfaceId As String = "69670011-d5f6-4fe9-8ae6-a20de46cea44"
    Public Const EventsId As String = "7dc9f69f-a241-4ba3-bbe7-2ba69f56e04d"

    'These routines perform the additional COM registration needed by ActiveX controls
    <EditorBrowsable(EditorBrowsableState.Never)>
    <ComRegisterFunction()>
    Private Shared Sub Register(ByVal t As Type)
        ComRegistration.RegisterControl(t)
    End Sub

    <EditorBrowsable(EditorBrowsableState.Never)>
    <ComUnregisterFunction()>
    Private Shared Sub Unregister(ByVal t As Type)
        ComRegistration.UnregisterControl(t)
    End Sub

#End Region

#Region "VB6 Events"

    'This section shows some examples of exposing a UserControl's events to VB6.  Typically, you just
    '1) Declare the event as you want it to be shown in VB6
    '2) Raise the event in the appropriate UserControl event.

    Public Shadows Event Click() 'Event must be marked as Shadows since .NET UserControls have the same name.
    Public Event DblClick()

    Private Sub InteropUserControl_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Click
        RaiseEvent Click()
    End Sub

    Private Sub InteropUserControl_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.DoubleClick
        RaiseEvent DblClick()
    End Sub

#End Region

#Region "VB6 Properties"

    'The following are examples of how to expose typical form properties to VB6.  
    'You can also use these as examples on how to add additional properties.

    'Must Shadow this property as it exists in Windows.Forms and is not overridable
    Public Shadows Property Visible() As Boolean
        Get
            Return MyBase.Visible
        End Get
        Set(ByVal value As Boolean)
            MyBase.Visible = value
        End Set
    End Property

    Public Shadows Property Enabled() As Boolean
        Get
            Return MyBase.Enabled
        End Get
        Set(ByVal value As Boolean)
            MyBase.Enabled = value
        End Set
    End Property

    Public Shadows Property ForegroundColor() As Integer
        Get
            Return ActiveXControlHelpers.GetOleColorFromColor(MyBase.ForeColor)
        End Get
        Set(ByVal value As Integer)
            MyBase.ForeColor = ActiveXControlHelpers.GetColorFromOleColor(value)
        End Set
    End Property

    Public Shadows Property BackgroundColor() As Integer
        Get
            Return ActiveXControlHelpers.GetOleColorFromColor(MyBase.BackColor)
        End Get
        Set(ByVal value As Integer)
            MyBase.BackColor = ActiveXControlHelpers.GetColorFromOleColor(value)
        End Set
    End Property

    Public Overrides Property BackgroundImage() As System.Drawing.Image
        Get
            Return Nothing
        End Get
        Set(ByVal value As System.Drawing.Image)
            If value IsNot Nothing Then
                MsgBox("Setting the background image of an Interop UserControl is not supported, please use a PictureBox instead.", MsgBoxStyle.Information)
            End If
            MyBase.BackgroundImage = Nothing
        End Set
    End Property

#End Region

#Region "VB6 Methods"

    Public Overrides Sub Refresh()
        MyBase.Refresh()
    End Sub

    'Ensures that tabbing across VB6 and .NET controls works as expected
    Private Sub UserControl_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.LostFocus
        ActiveXControlHelpers.HandleFocus(Me)
    End Sub

    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

        'Raise Load event
        Me.OnCreateControl()
    End Sub

    <SecurityPermission(SecurityAction.LinkDemand, Flags:=SecurityPermissionFlag.UnmanagedCode)> _
    Protected Overrides Sub WndProc(ByRef m As System.Windows.Forms.Message)

        Const WM_SETFOCUS As Integer = &H7
        Const WM_PARENTNOTIFY As Integer = &H210
        Const WM_DESTROY As Integer = &H2
        Const WM_LBUTTONDOWN As Integer = &H201
        Const WM_RBUTTONDOWN As Integer = &H204

        If m.Msg = WM_SETFOCUS Then
            'Raise Enter event
            Me.OnEnter(New System.EventArgs)

        ElseIf m.Msg = WM_PARENTNOTIFY AndAlso _
            (m.WParam.ToInt32 = WM_LBUTTONDOWN OrElse _
             m.WParam.ToInt32 = WM_RBUTTONDOWN) Then

            If Not Me.ContainsFocus Then
                'Raise Enter event
                Me.OnEnter(New System.EventArgs)
            End If

        ElseIf m.Msg = WM_DESTROY AndAlso Not Me.IsDisposed AndAlso Not Me.Disposing Then
            'Used to ensure that VB6 will cleanup control properly
            Me.Dispose()
        End If

        MyBase.WndProc(m)
    End Sub

    'This event will hook up the necessary handlers
    Private Sub InteropUserControl_ControlAdded(ByVal sender As Object, ByVal e As ControlEventArgs) Handles Me.ControlAdded
        ActiveXControlHelpers.WireUpHandlers(e.Control, AddressOf ValidationHandler)
    End Sub

    'Ensures that the Validating and Validated events fire appropriately
    Friend Sub ValidationHandler(ByVal sender As Object, ByVal e As EventArgs)

        If Me.ContainsFocus Then Return

        'Raise Leave event
        Me.OnLeave(e)

        If Me.CausesValidation Then
            Dim validationArgs As New CancelEventArgs
            Me.OnValidating(validationArgs)

            If validationArgs.Cancel AndAlso Me.ActiveControl IsNot Nothing Then
                Me.ActiveControl.Focus()
            Else
                'Raise Validated event
                Me.OnValidated(e)
            End If
        End If

    End Sub

#End Region

#End If

#End Region

    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

    Public Delegate Sub StartEventHandler(ByVal simpleEventText As String)
    Public Delegate Sub FinishAsyncEventHandler(ByVal asyncEventText As String)



    Public Event StartEvent As StartEventHandler
    Public Event FinishAsyncEvent As FinishAsyncEventHandler

    Public Sub StartProcessing(ByVal StartDirectory As String)
        Try
            RaiseEvent StartEvent("Distribution of " & StartDirectory & " in Process...")
            Me.BackgroundWorker1.RunWorkerAsync(StartDirectory)
        Catch
        End Try
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object,
    ByVal e As DoWorkEventArgs) _
    Handles BackgroundWorker1.DoWork
        'wait
        Dim count As Integer = 0
        Dim dir As String = e.Argument.ToString()
        Dim root As String = Path.GetDirectoryName(dir)
        Dim Versions As String() = Directory.GetDirectories(root, Path.GetFileName(dir), SearchOption.AllDirectories)
        Dim fileCount As Integer
        Try
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
        Dim files = Directory.EnumerateFiles(dir, "Properties.ini", IO.SearchOption.AllDirectories)
        For Each currentFile As String In files
            Dim fileName = Path.GetFullPath(currentFile)
            count = count + 1
            ' Do the project compile here
            Dim percentComplete As Integer = count / fileCount * 100
            DistributeProject(fileName, percentComplete)
            Me.BackgroundWorker1.ReportProgress(percentComplete, fileName)
        Next
    End Sub

    Private Sub BackgroundWorker1_ProgressChanged(ByVal sender As System.Object,
    ByVal e As ProgressChangedEventArgs) _
    Handles BackgroundWorker1.ProgressChanged
        Me.LabelWarningMessage.ForeColor = Color.Red
        'Me.LabelWarningMessage.Text = "Working in background..."
        Me.LabelWarningMessage.Text = e.UserState.ToString()
        Me.LabelWarningMessage.Visible = True
        Me.ProgressBar1.Value = e.ProgressPercentage
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As System.Object,
    ByVal e As RunWorkerCompletedEventArgs) _
    Handles BackgroundWorker1.RunWorkerCompleted
        Me.LabelWarningMessage.Visible = False
        RaiseEvent FinishAsyncEvent("Bulk Bistribution Completed.")
    End Sub

    Private Sub BackgroundWorker_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.ProgressBar1.Value = 0
        Me.LabelWarningMessage.Visible = False
    End Sub

    Public Sub DistributeProject(ByVal ProjectDirectory As String, ByVal PercentComplete As Integer, Optional ByVal IsInTST As Boolean = True, Optional ByVal UpdateWebsite As Boolean = True)
        Dim sModules As String = ""
        Dim sTmp As String
        Dim bCreateUpdateDelta As Boolean
        Dim OnNetwork As Boolean
        Dim sTargetModuleFolder As String
        Dim ClassInstall As Boolean
        Dim sPath As String
        Dim sFile As String
        Dim sUtilityMenuFolder As String
        Dim sDestination As String
        Dim sPrgFile As String
        Dim FChn As StreamWriter
        Dim reader As StreamReader
        Dim iTmp As DialogResult
        Dim x As Integer
        Dim proc As Process
        Dim DistributionDirectory As String
        Dim TestBaseDirectory As String = ""
        Dim TestHomeDirectory As String = ""
        Dim ProjectBaseDirectory As String = ""

        Dim TemplateLibrary As String = "D:\dev\templates"
        Dim ManualLibrary As String = "D:\DEV\manuals"
        Dim UploadStaging As String = "D:\DEV\HTML"
        Dim IDEPath As String = "D:\DEV\IDE"

        Dim WinZipProgram As String = ""
        Dim WinZipAvailable As Boolean
        Dim NSISAvailable As Boolean = False
        Dim INSTAvailable As Boolean = False
        Dim MakeNSIS As String = ""
        Dim Properties As Properties


        If My.Computer.FileSystem.FileExists(ProjectDirectory) Then
            'Ini file Loading and Parsing
            Properties = New Properties(ProjectDirectory)
            If String.IsNullOrEmpty(Properties.CodeDescription) Then
                Exit Sub
            ElseIf Not ProjectDirectory.ToUpper().Contains("\" & Properties.CodeDescription.ToUpper() & "." & Properties.Version.Remove(Properties.Version.IndexOf("."), 1)) Then
                Exit Sub
            End If
        End If

        Dim StdFilesPath As String = GetSetting("IDE", "Options", "StdFilesPath", IDEPath & "\StdFiles")
        Dim ToolPath As String = GetSetting("IDE", "Options", "ToolPath")

        ToolPath = "D:\DEV\" & Properties.Version.Remove(Properties.Version.IndexOf("."), 1) & "pxBasic"

        If My.Computer.FileSystem.DirectoryExists(ToolPath) Then
            TestBaseDirectory = ToolPath
            TestHomeDirectory = TestBaseDirectory & "\HOME"
        Else
            Exit Sub
        End If

        Properties = New Properties(ProjectDirectory)

        Try
            DistributionDirectory = Properties.OutputFolder
            ProjectBaseDirectory = ProjectDirectory.Remove(ProjectDirectory.IndexOf("\Properties.ini"))

            If Properties.ReleaseDate <> Date.Today.ToShortDateString() Then
                Properties.ReleaseDate = Date.Today.ToShortDateString()
                Properties.Save(ProjectDirectory)
                CheckReleaseDate(TestHomeDirectory, ProjectBaseDirectory)
            End If




            ' Set global param
            If IsInTST And My.Computer.FileSystem.DirectoryExists(TemplateLibrary) _
            Or IsInTST And My.Computer.FileSystem.DirectoryExists("C:\DEV\TEMPLATES") Then
                OnNetwork = True
            Else
                OnNetwork = False
            End If

            WinZipAvailable = False
            If My.Computer.FileSystem.FileExists("C:\Program Files\WinZip Self-Extractor\Wzipse32.exe") Then
                WinZipAvailable = True
                WinZipProgram = "C:\Program Files\WinZip Self-Extractor\Wzipse32.exe"
            End If

            NSISAvailable = False
            INSTAvailable = False
            If My.Computer.FileSystem.FileExists(IDEPath & "\NSIS\makensis.exe") Then
                NSISAvailable = True
                MakeNSIS = IDEPath & "\NSIS\makensis.exe"
                If My.Computer.FileSystem.FileExists(IDEPath & "\NSIS\Plugins\pvxinst.dll") Then
                    INSTAvailable = True
                End If
            Else
                If My.Computer.FileSystem.FileExists(Environment.SpecialFolder.ProgramFiles.ToString() & "\NSIS\makensis.exe") Then
                    NSISAvailable = True
                    MakeNSIS = Environment.SpecialFolder.ProgramFiles.ToString() & "\NSIS\makensis.exe"
                    If My.Computer.FileSystem.FileExists(Environment.SpecialFolder.ProgramFiles.ToString() & "\NSIS\Plugins\pvxinst.dll") Then
                        INSTAvailable = True
                    End If
                End If

            End If


            'Trap the 4.00 version, for now
            If My.Computer.FileSystem.FileExists(TestBaseDirectory & "\INSTALL.PVX") Then
                ClassInstall = True
            End If

            If ClassInstall Then
                CreateInstallationClass(TestBaseDirectory, ProjectBaseDirectory)
                If Val(Properties.Version) >= 5 Then
                    If Not My.Computer.FileSystem.FileExists(ProjectBaseDirectory & "\" & Properties.PrimaryModule & "\" & Properties.PrimaryModule & "Manifest." & Properties.CodeIdentifier) Then
                        CreateDeveloperManifest(TestBaseDirectory, ProjectBaseDirectory)
                    End If
                End If
            End If

            ' Perform various checks
            Me.BackgroundWorker1.ReportProgress(PercentComplete, Properties.ZipExeFilename & " Initialization`")


            If DistributionDirectory = "" Then
                log.Error("Image Destination Folder must be completed. " & Properties.CodeDescription)
                Exit Sub
            End If
            ' Check for existence of CTL file - warn if it is not there.
            If Not My.Computer.FileSystem.FileExists(CTLFilename(ProjectBaseDirectory)) Then
                log.Error("Control file does not exist for Project's Primary Module. Primary Module: " & Properties.PrimaryModule)
                Exit Sub
            End If
            ' TST Only: If on net, if no WinZipSE, issue warning.
            If OnNetwork Then
                If Not My.Computer.FileSystem.DirectoryExists(Properties.ZipExeOutputFolder) Then
                    log.Error("The target Zip/Exe directory does not exist. " & Properties.ZipExeOutputFolder)
                    Exit Sub
                End If
            End If

            ' If: on TST net; is an Enhancement;
            '     UpdateAvailable flag set; Update website chkbox is set
            ' Then if no Original Release Date, warn and exit
            bCreateUpdateDelta = False
            If Properties.UpdateAvailable = 1 _
            And OnNetwork Then
                If Properties.OriginDate = "" Then
                    log.Error("The Origin Date has not been entered.")
                    Exit Sub
                End If
                bCreateUpdateDelta = True
            End If


            ' Erase Distribution output folder
            If My.Computer.FileSystem.DirectoryExists(DistributionDirectory) Then
                Me.BackgroundWorker1.ReportProgress(PercentComplete, Properties.ZipExeFilename & " Erase output folder")
                My.Computer.FileSystem.DeleteDirectory(DistributionDirectory, FileIO.DeleteDirectoryOption.DeleteAllContents)
            End If


            ' Recursively copy Project files to Distribution folder.
            ' This also creates the output folder.
            Me.BackgroundWorker1.ReportProgress(PercentComplete, Properties.ZipExeFilename & " Copy Project file to Distribution folder")

            My.Computer.FileSystem.CopyDirectory(ProjectBaseDirectory, DistributionDirectory, True)
            ' If Patch Install of Libraries, erase all parsed libraries here.
            If Properties.IsPatchInstall Then
                If My.Computer.FileSystem.FileExists(ProjectBaseDirectory & "\PATCH.TXT") Then
                    reader = My.Computer.FileSystem.OpenTextFileReader(ProjectBaseDirectory & "\PATCH.TXT", Encoding.ASCII)
                    Do
                        sTmp = reader.ReadLine()
                        If Mid$(sTmp, 1, 2) = ".." Then
                            sTmp = Mid$(sTmp, 3) ' Remove leading ".."
                            sPath = Mid$(sTmp, 1, InStrRev(sTmp, "\") - 1) ' Parse out path with leading "\"
                            sFile = Mid$(sTmp, InStrRev(sTmp, "\") + 1) ' Parse out filename
                            If My.Computer.FileSystem.FileExists(DistributionDirectory & sPath & "\" & sFile) Then
                                My.Computer.FileSystem.DeleteFile(DistributionDirectory & sPath & "\" & sFile)
                            End If
                        End If
                    Loop Until reader.EndOfStream
                    reader.Close()
                End If
            End If

            Dim PxPlusCode As String
            PxPlusCode = GetSSN(TestHomeDirectory).Trim()


            ' Check whether Destination directory was created. If not,
            ' error-out with warning.
            If Not My.Computer.FileSystem.DirectoryExists(DistributionDirectory) Then
                iTmp = MessageBox.Show("Distribution Image Destination folder could not be created." _
                              & Environment.NewLine & Environment.NewLine _
                              & "Directory that could not be created: " _
                              & DistributionDirectory,
                              "Error", MessageBoxButtons.OK)
                Exit Sub
            End If

            ' Recursively copy IDE Standard Installation files to output folder.
            Me.BackgroundWorker1.ReportProgress(PercentComplete, Properties.ZipExeFilename & " Copy Installation files to Distribution")
            If Properties.AddlFilesFolder <> "" Then
                My.Computer.FileSystem.CopyDirectory(Properties.AddlFilesFolder, DistributionDirectory, True)
            Else
                My.Computer.FileSystem.CopyDirectory(StdFilesPath, DistributionDirectory, True)
            End If

            ' If selected, include TST Utility Menu in Distribution
            If OnNetwork And
            Properties.IncludeMainUtilityMenu = "1" Then
                sTmp = Mid$(Properties.Version, 1, 1) & Mid$(Properties.Version, 3, 2)
                sUtilityMenuFolder = "D:\dev\version." & sTmp
                My.Computer.FileSystem.CopyDirectory(sUtilityMenuFolder, DistributionDirectory, True)
            End If

            ' If "readme.htm" exists in the project folder, copy it to the root of the distribution.
            If My.Computer.FileSystem.FileExists(ProjectBaseDirectory & "\AdditionalInstructions.htm") Then
                File.Copy(ProjectBaseDirectory & "\AdditionalInstructions.htm", DistributionDirectory & "\AdditionalInstructions.htm", True)
            End If

            ' Delete Properties.ini file
            Me.BackgroundWorker1.ReportProgress(PercentComplete, Properties.ZipExeFilename & " Delete Properties.ini file")
            If My.Computer.FileSystem.FileExists(DistributionDirectory & "\PROPERTIES.INI") Then _
                My.Computer.FileSystem.DeleteFile(DistributionDirectory & "\PROPERTIES.INI")
            If My.Computer.FileSystem.FileExists(DistributionDirectory & "\PROPER~1.INI") Then _
                My.Computer.FileSystem.DeleteFile(DistributionDirectory & "\PROPER~1.INI")

            If OnNetwork Then
                Select Case Val(Properties.DistributionVersion)
                    Case 3 ' 
                        Me.BackgroundWorker1.ReportProgress(PercentComplete, Properties.ZipExeFilename & " Copy Latest templates from Template Library: " & TemplateLibrary & "\TEMPLATES_3")
                        If Properties.PrimaryModule = "XX" Then
                            sTargetModuleFolder = "SYSTEM"
                        Else
                            sTargetModuleFolder = Properties.PrimaryModule
                        End If
                        sDestination = DistributionDirectory & "\" & sTargetModuleFolder
                        File.Copy(TemplateLibrary & "\TEMPLATES\TEMPLATE.TXT", sDestination & "\TEMPLATE.XXX", True)
                        sTmp = sDestination & "\" & Properties.PrimaryModule & "TEMPLATE." & Properties.CodeIdentifier
                        If My.Computer.FileSystem.FileExists(sTmp) Then My.Computer.FileSystem.DeleteFile(sTmp)
                    Case Is < 3 ' 
                        Me.BackgroundWorker1.ReportProgress(PercentComplete, Properties.ZipExeFilename & " Copy Latest version from Template Library: " & TemplateLibrary & "\OLD_STUFF")
                        If Properties.PrimaryModule = "XX" Then
                            sTargetModuleFolder = "SYSTEM"
                        Else
                            sTargetModuleFolder = Properties.PrimaryModule
                        End If
                        sDestination = DistributionDirectory & "\" & sTargetModuleFolder
                        File.Copy(TemplateLibrary & "\TEMPLATES_PRE\PRE.LIB", sDestination & "\PRE.LIB", True)
                End Select
            End If

            ' Recursively rename all 8.3 files as uppercase.
            Me.BackgroundWorker1.ReportProgress(PercentComplete, Properties.ZipExeFilename & " Rename all 8.3 std files as uppercase")

            ' For TST Enhancements, copy TST Enhancement catalog to Distribution.
            If OnNetwork _
            And Val(Properties.Version) >= 3 _
            And Val(Properties.DoNotIncludeCatalog) <> 1 Then
                Call My.Computer.FileSystem.CopyDirectory("D:\DEV\TEMPLATE\DISTRIBUTION", DistributionDirectory, True)
            End If

            ' Password-protect ProvideX programs.
            If Properties.Password <> "" Then
                Me.BackgroundWorker1.ReportProgress(PercentComplete, Properties.ZipExeFilename & " Password-protect ProxideX programs")
                sPrgFile = "$Prg$"
                Call CopyFromLibrary("PVX", TestBaseDirectory & "\PVX", "PVXPPR")
                FChn = My.Computer.FileSystem.OpenTextFileWriter(sPrgFile, False, Encoding.ASCII)
                ' Pass Data to Pvx, Initialize Vars
                FChn.WriteLine("10 PASSWORD$ = """ & Properties.Password & """")
                FChn.WriteLine("20 OUTDIR$ = """ & DistributionDirectory & """")
                FChn.WriteLine("25 PASSWORDAPPLYTO$ = """ & Properties.PasswordApplyTo & """")
                If bCreateUpdateDelta Then
                    FChn.WriteLine("30 OUTDIRD$ = """ & DistributionDirectory & "U" & """")
                End If
                FChn.WriteLine("40 CALL ""..\PVX\PVXPPR;PROTECT"", OUTDIR$, PASSWORD$, PASSWORDAPPLYTO$ ")
                If bCreateUpdateDelta Then
                    FChn.WriteLine("50 CALL ""..\PVX\PVXPPR;PROTECT"", OUTDIRD$, PASSWORD$, PASSWORDAPPLYTO$ ")
                End If
                FChn.WriteLine("60 BYE")
                FChn.Flush()
                FChn.Close()

                ' Run temp prg and wait for it to exit
                Dim info As New ProcessStartInfo(TestBaseDirectory & "\HOME\PVXWIN32.EXE")
                info.WorkingDirectory = TestBaseDirectory & "\HOME\"
                info.Arguments = "-HD ..\HOME\" & sPrgFile
                info.CreateNoWindow = True
                info.UseShellExecute = False
                Dim Passproc As Process = Process.Start(info)
                Passproc.WaitForExit()

                My.Computer.FileSystem.DeleteFile(sPrgFile)

            End If

            ' Create INST file, if applicable.
            Me.BackgroundWorker1.ReportProgress(PercentComplete, Properties.ZipExeFilename & " Create INST.TXT file, if applicable")
            If (InStr("CUSTOM", Properties.PrimaryModule) _
            And Val(Properties.Version) < 3.31) _
            Or (InStr("EXTRA", Properties.PrimaryModule) _
            And Val(Properties.Version) < 3.03) Then
                FChn = My.Computer.FileSystem.OpenTextFileWriter(DistributionDirectory & "\INST.TXT", False, Encoding.ASCII)
                FChn.WriteLine("")
                FChn.Flush()
                FChn.Close()
            End If

            ' Create MDINST file
            Me.BackgroundWorker1.ReportProgress(PercentComplete, Properties.ZipExeFilename & " Create MDINST file")
            FChn = My.Computer.FileSystem.OpenTextFileWriter(DistributionDirectory & "\SOA\MDINST", False, Encoding.ASCII)
            If Properties.PrimaryModule = "XXX" Then
                FChn.WriteLine("SYSSTEM")
            Else
                FChn.WriteLine("" & Properties.PrimaryModule & "")
            End If
            FChn.Flush()
            FChn.Close()

            ' Create VERSIONS.HTM and VERSIONS.TXT
            ' Only if CTL file exists
            If My.Computer.FileSystem.FileExists(CTLFilename(ProjectBaseDirectory)) Then
                Me.BackgroundWorker1.ReportProgress(PercentComplete, Properties.ZipExeFilename & " Create VERSIONS.TXT")
                ' Get list of modules/versions modified.
                sModules = GetModuleLevels(TestHomeDirectory, ProjectBaseDirectory)
                ' TST Only: Create VERSIONS.HTM
                If OnNetwork Then
                    FChn = My.Computer.FileSystem.OpenTextFileWriter(DistributionDirectory & "\VERSIONS.HTM", False, Encoding.ASCII)
                    FChn.WriteLine("<html>")
                    FChn.WriteLine("")
                    FChn.WriteLine("<head>")
                    FChn.WriteLine("<title>Version Compatibility</title>")
                    FChn.WriteLine("</head>")
                    FChn.WriteLine("<body>")
                    If Not Properties.IsNonTest Then
                        FChn.WriteLine("<p><img src=""logo.jpg"" alt=""Logo"" align=""left"" >")
                        FChn.WriteLine("<br/><br/><br/><br/><br/>")
                        FChn.WriteLine("<strong>")
                    End If
                    FChn.WriteLine("Version Compatibility</strong><br>")
                    FChn.WriteLine("<p>&nbsp;</p>")
                    FChn.WriteLine("<p><b>Product: </b>" + Properties.ProjectName + "<br>")
                    FChn.WriteLine("<b>Version: </b>" + Properties.Version + "</p>")
                    FChn.WriteLine("<p><i>Compatible with these options:</i></p>")
                    FChn.WriteLine("<blockquote>")
                    FChn.WriteLine("<p>")
                    ' Modules/Versions modified by Project
                    sTmp = ""
                    For x = 1 To Len(sModules)
                        If Mid$(sModules, x, 1) = ";" Then
                            If sTmp <> "" Then
                                FChn.WriteLine(sTmp & "<br>")
                                sTmp = ""
                            End If
                        Else
                            sTmp = sTmp & Mid$(sModules, x, 1)
                        End If
                    Next x
                    FChn.WriteLine("</p></blockquote>")
                    FChn.WriteLine("<hr>")
                    If Not Properties.IsNonTest Then
                        FChn.WriteLine("<p><a href=""install.htm"">Installation Procedure</a><br>")
                        FChn.WriteLine("<a href=""license.htm"">License Agreement</a><br>")
                        FChn.WriteLine("<a href=""manual.htm"">Product Manual</a><br>")
                        FChn.WriteLine("<a href=""support.htm"">Product Support</a><br>")
                        FChn.WriteLine("<a href=""readme.htm"">Readme Document</a></p>")
                    End If
                    FChn.WriteLine("</body>")
                    FChn.WriteLine("</html>")
                    FChn.Flush()
                    FChn.Close()
                Else
                    FChn = My.Computer.FileSystem.OpenTextFileWriter(DistributionDirectory & "\VERSIONS.HTM", False, Encoding.ASCII)
                    FChn.WriteLine("<html>")
                    FChn.WriteLine("<head>")
                    FChn.WriteLine("<title>Version Compatibility</title>")
                    FChn.WriteLine("</head>")
                    FChn.WriteLine("<body>")
                    FChn.WriteLine("<p><strong>Version Compatibility</strong></p>")
                    FChn.WriteLine("<p>&nbsp;</p>")
                    FChn.WriteLine("<p><b>Product: </b>" + Properties.ProjectName + "<br>")
                    FChn.WriteLine("<b>Version: </b>" + Properties.Version + "</p>")
                    FChn.WriteLine("<p><i>Compatible with these options:</i></p>")
                    FChn.WriteLine("<blockquote>")
                    FChn.WriteLine("<p>")
                    sTmp = ""
                    For x = 1 To Len(sModules)
                        If Mid$(sModules, x, 1) = ";" Then
                            If sTmp <> "" Then
                                FChn.WriteLine(sTmp & "<br>")
                                sTmp = ""
                            End If
                        Else
                            sTmp = sTmp & Mid$(sModules, x, 1)
                        End If
                    Next x
                    FChn.WriteLine("</p></blockquote>")
                    FChn.WriteLine("<hr>")
                    FChn.WriteLine("<p><a href=""install.htm"">Installation Procedure</a><br>")
                    FChn.WriteLine("<a href=""license.htm"">License Agreement</a><br>")
                    FChn.WriteLine("<a href=""support.htm"">Product Support</a></p>")
                    FChn.WriteLine("</body>")
                    FChn.WriteLine("</html>")
                    FChn.Flush()
                    FChn.Close()
                End If
            End If

            ' Create VERSIONS.TXT
            If My.Computer.FileSystem.FileExists(CTLFilename(ProjectBaseDirectory)) Then
                FChn = My.Computer.FileSystem.OpenTextFileWriter(DistributionDirectory & "\VERSIONS.TXT", False, Encoding.ASCII)
                FChn.WriteLine("Product: " & Properties.ProjectName)
                FChn.WriteLine("Version: " & Properties.Version)
                FChn.WriteLine("")
                FChn.WriteLine("Compatible with these options:")
                FChn.WriteLine("")
                ' Modules/Versions modified by Project
                sTmp = ""
                For x = 1 To Len(sModules)
                    If Mid$(sModules, x, 1) = ";" Then
                        If sTmp <> "" Then
                            FChn.WriteLine(Space(10), sTmp)
                            sTmp = ""
                        End If
                    Else
                        sTmp = sTmp & Mid$(sModules, x, 1)
                    End If
                Next x
                FChn.Flush()
                FChn.Close()

            End If

            ' Create INST_ID.TXT
            Me.BackgroundWorker1.ReportProgress(PercentComplete, Properties.ZipExeFilename & " Create INSTALLID.TXT")
            FChn = My.Computer.FileSystem.OpenTextFileWriter(DistributionDirectory & "\INSTALLID.TXT", False, Encoding.ASCII)
            FChn.WriteLine(Properties.ProjectName)
            FChn.WriteLine(Properties.Version)
            FChn.WriteLine("Disk 1 of 1")
            If Properties.IsSystem Then
                FChn.WriteLine(Properties.ZipExeFilename)
            End If
            If Properties.PostEXECommand <> "" Then
                FChn.WriteLine(Properties.PostEXECommand)
            End If
            FChn.Flush()
            FChn.Close()

            If Not NSISAvailable Then
                Me.BackgroundWorker1.ReportProgress(PercentComplete, Properties.ZipExeFilename & " Rename STUBSETUP.EXE as SETUP.EXE")
                If My.Computer.FileSystem.FileExists(DistributionDirectory & "\" & "STUBSetup.exe") Then
                    My.Computer.FileSystem.RenameFile(DistributionDirectory & "\" & "STUBSetup.exe", "Setup.exe")
                End If
            End If

            If NSISAvailable And Not OnNetwork Then
                Me.BackgroundWorker1.ReportProgress(PercentComplete, Properties.ZipExeFilename & " Create Installer " & Properties.ZipExeFilename)
                FChn = My.Computer.FileSystem.OpenTextFileWriter(DistributionDirectory & "\installer.nsi", False, Encoding.ASCII)
                FChn.WriteLine(";Generated Install file for NSIS")
                FChn.WriteLine("")
                FChn.WriteLine("!define PRODUCT_NAME """ & Properties.ProjectName & """")
                FChn.WriteLine("!define PRODUCT_VERSION """ & Properties.Version & """")
                FChn.WriteLine("")
                FChn.WriteLine("; MUI 1.67 compatible ------")
                FChn.WriteLine("!include " & """MUI.nsh""")
                FChn.WriteLine("")
                FChn.WriteLine("!define MUI_ABORTWARNING")
                If My.Computer.FileSystem.FileExists(Properties.AddlFilesFolder & "\publisher.nsh") Then
                    FChn.WriteLine("!include """ & Properties.AddlFilesFolder & "\publisher.nsh""")
                Else
                    FChn.WriteLine("!define MUI_ICON """ & "${NSISDIR}\Contrib\Graphics\Icons\modern-install.ico""")
                End If
                FChn.WriteLine("!insertmacro MUI_PAGE_WELCOME")
                If My.Computer.FileSystem.FileExists(Properties.AddlFilesFolder & "\license.rtf") Then
                    FChn.WriteLine("!define MUI_LICENSEPAGE_RADIOBUTTONS")
                    FChn.WriteLine("!insertmacro MUI_PAGE_LICENSE """ & Properties.AddlFilesFolder & "\license.rtf""")
                End If
                FChn.WriteLine("!define MUI_DIRECTORYPAGE_VERIFYONLEAVE")
                FChn.WriteLine("!define MUI_PAGE_CUSTOMFUNCTION_LEAVE dirLeave")

                FChn.WriteLine("!insertmacro MUI_PAGE_DIRECTORY")
                FChn.WriteLine("!insertmacro MUI_PAGE_INSTFILES")
                FChn.WriteLine("!insertmacro MUI_PAGE_FINISH")
                FChn.WriteLine("!insertmacro MUI_LANGUAGE """ & "English""")
                FChn.WriteLine("Name " & """${PRODUCT_NAME} ${PRODUCT_VERSION}""")
                If OnNetwork Then
                    FChn.WriteLine("OutFile """ & Properties.ZipExeFilename & ".exe""")
                ElseIf Trim$(Properties.ZipExeFilename) <> "" Then
                    FChn.WriteLine("OutFile """ & Properties.ZipExeFilename & ".exe""")
                Else
                    FChn.WriteLine("OutFile " & """setup.exe""")
                End If
                FChn.WriteLine("")
                FChn.WriteLine("ShowInstDetails nevershow")
                FChn.WriteLine("ShowUnInstDetails nevershow")
                FChn.WriteLine("")
                ' main section will always need to be built
                FChn.WriteLine("Section " & """Main""" & " SEC01")
                If INSTAvailable And (Val(Properties.Version) >= 4) Then
                    FChn.WriteLine(vbTab & "InitPluginsDir")
                End If
                FChn.WriteLine(vbTab & "SetOutPath " & """$INSTDIR""")
                FChn.WriteLine(vbTab & "SetOverwrite on")
                If INSTAvailable And (Val(Properties.Version) >= 4) Then
                    FChn.WriteLine("InitPluginsDir")
                    FChn.WriteLine("DetailPrint " & """Checking ProvideX COM Object...""")
                    FChn.WriteLine("PVXINST::CheckCom")
                    FChn.WriteLine("Pop $R0")
                    FChn.WriteLine("DetailPrint $R0")
                    FChn.WriteLine("StrCmp $R0 " & """1""" & " +2")
                    FChn.WriteLine("Abort " & """The Providex COM object is not registered on this workstation-contact your reseller""")
                    FChn.WriteLine("DetailPrint " & """Initializing ProvideX COM object...""")
                    FChn.WriteLine("PVXINST::LoadCom")
                    FChn.WriteLine("Pop $R0")
                    FChn.WriteLine("DetailPrint $R0")
                    FChn.WriteLine("StrCmp $R0 " & """1""" & " +2")
                    FChn.WriteLine("Abort " & """There was an error initializing the ProvideX COM object-contact your reseller""")
                End If
                ' Add Additional Files to the setup
                If Trim$(Properties.AddlFilesFolder) <> "" Then
                    FChn.WriteLine(vbTab & "File /r " & Properties.AddlFilesFolder & "\*.*""")
                End If
                FChn.WriteLine("")
                If INSTAvailable And (Val(Properties.Version) >= 4) Then
                    FChn.WriteLine("DetailPrint " & """Preparing installation...""")
                    FChn.WriteLine("PVXINST::InstallPrepare")
                    FChn.WriteLine("Pop $R0")
                    FChn.WriteLine("DetailPrint $R0")
                    FChn.WriteLine("StrCmp $R0 " & """1""" & " +2")
                    FChn.WriteLine("MessageBox MB_OK " & """The Prepare method failed""")
                    FChn.WriteLine("DetailPrint " & """Installing...""")
                    FChn.WriteLine("PVXINST::InstallModules")
                    FChn.WriteLine("Pop $R0")
                    FChn.WriteLine("Pop $R1")
                    FChn.WriteLine("DetailPrint " & """$R0 $R1""")
                    FChn.WriteLine("StrCmp $R0 " & """1""" & " +2")
                    FChn.WriteLine("MessageBox MB_OK $R1")
                    FChn.WriteLine("DetailPrint " & """Finalizing the installation...""")
                    FChn.WriteLine("PVXINST::GradualProgress /NOUNLOAD 1 2 100 " & """Installation Finished.""")
                    FChn.WriteLine("PVXINST::InstallFinalize /NOUNLOAD")
                    FChn.WriteLine("DetailPrint " & """Installation Finished.""")
                End If
                If Not INSTAvailable Then
                    If My.Computer.FileSystem.FileExists(StdFilesPath & "\PVXINST.EXE") And (Val(Properties.Version) >= 4) Then
                        FChn.WriteLine(vbTab & "File /oname=$TEMP\PVXINST.EXE """ & StdFilesPath & "\PVXINST.EXE""")
                        FChn.WriteLine(vbTab & "ExecWait '""" & "$TEMP\PVXINST.EXE""" & " $INSTDIR'")
                        FChn.WriteLine(vbTab & "Delete $TEMP\PVXINST.EXE")
                    End If
                End If
                FChn.WriteLine("SectionEnd")
                ' We could include optional section information here that would never change
                '
                ' We could include a function header in right here
                FChn.WriteLine("Function .onInit")
                FChn.WriteLine(vbTab & "ReadRegStr $INSTDIR HKCU " & """Software\ODBC\ODBC.INI\SOTAMAS90""" & " " & """MAS90RootDirectory""")
                FChn.WriteLine("FunctionEnd")
                FChn.WriteLine("Function dirLeave")
                FChn.WriteLine(vbTab & "IfFileExists $INSTDIR\Home\pvxwin32.exe PathGood")
                FChn.WriteLine(vbTab & "Abort ; if $INSTDIR is not a ProvideX directory, don't let us install there")
                FChn.WriteLine("PathGood:")
                FChn.WriteLine(vbTab & "PVXINST::CheckVersion")
                FChn.WriteLine(vbTab & "Pop $R0")
                FChn.WriteLine(vbTab & "StrCmp $R0 ${PRODUCT_VERSION} +3")
                FChn.WriteLine(vbTab & "MessageBox MB_OK " & """The version being installed is not compatible with the current version.""")
                FChn.WriteLine(vbTab & "Abort ; ")
                FChn.WriteLine("FunctionEnd")
                If INSTAvailable And (Val(Properties.Version) >= 4) Then
                    FChn.WriteLine("Function .onGUIEnd")
                    FChn.WriteLine(vbTab & "PVXINST::Unload")
                    FChn.WriteLine("FunctionEnd")
                End If

                FChn.Flush()
                FChn.Close()

                If My.Computer.FileSystem.FileExists(DistributionDirectory & "\STUBSETUP.EXE") Then
                    My.Computer.FileSystem.DeleteFile(DistributionDirectory & "\STUBSETUP.EXE")
                End If
                If My.Computer.FileSystem.FileExists(DistributionDirectory & "\Installation.txt") Then
                    My.Computer.FileSystem.DeleteFile(DistributionDirectory & "\Installation.txt")
                End If
                If My.Computer.FileSystem.FileExists(DistributionDirectory & "\publisher.nsh") Then
                    My.Computer.FileSystem.DeleteFile(DistributionDirectory & "\publisher.nsh")
                End If

                ' Should have enough info to make installer, without extra files
                proc = New Process()
                proc.StartInfo.FileName = Chr(34) & MakeNSIS & Chr(34) & " /O" & DistributionDirectory & "\install.log " & Chr(34) & DistributionDirectory & "\installer.nsi"""
                proc.Start()
                proc.WaitForExit()


            End If


            ' TST Only: Create SE/Inst
            If OnNetwork And (WinZipAvailable And Not NSISAvailable) Then
                Me.BackgroundWorker1.ReportProgress(PercentComplete, Properties.ZipExeFilename & " Create SE program file")
                ' Create DIALOG.TXT
                FChn = My.Computer.FileSystem.OpenTextFileWriter(Properties.ZipExeOutputFolder & "\DIALOG.TXT", False, Encoding.ASCII)
                FChn.WriteLine("Installation for:")
                FChn.WriteLine("")
                FChn.WriteLine("   Product: " & Chr(9) & Properties.ProjectName)
                FChn.WriteLine("   Version: " & Chr(9) & Properties.Version)
                FChn.WriteLine("   Release: " & Chr(9) & Properties.ReleaseDate)
                FChn.Flush()
                FChn.Close()

                ' Create ABOUT.TXT
                FChn = My.Computer.FileSystem.OpenTextFileWriter(Properties.ZipExeOutputFolder & "\ABOUT.TXT", False, Encoding.ASCII)
                FChn.WriteLine("Test Distribution System for ProvideX programs")
                FChn.Flush()
                FChn.Close()


                ' Delete old .ZIP file, if present.
                If My.Computer.FileSystem.FileExists(Properties.ZipExeOutputFolder & "\" _
                & Properties.ZipExeFilename & ".zip") Then
                    My.Computer.FileSystem.DeleteFile(Properties.ZipExeOutputFolder & "\" &
                         Properties.ZipExeFilename & ".zip")
                End If

                ' Optionally include manual PDF in SE/Zip distribution.
                If Properties.IncludePDFinSEZip = "1" Then
                    If My.Computer.FileSystem.FileExists(ManualLibrary & "\" &
                                      Mid(Properties.ZipExeFilename, 1, 7) & ".pdf") Then
                        File.Copy(ManualLibrary & "\" & Mid(Properties.ZipExeFilename, 1, 7) & ".pdf",
                             DistributionDirectory & "\pdf" & "\" & Mid(Properties.ZipExeFilename, 1, 7) & ".pdf", True)
                    End If
                End If

                ' Create .ZIP file.


                ' Create SEINPUT.TXT
                ' This should be replaced with the .nsi file generation
                FChn = My.Computer.FileSystem.OpenTextFileWriter(Properties.ZipExeOutputFolder & "\SEINPUT.TXT", False, Encoding.ASCII)
                FChn.WriteLine("-setup")
                FChn.WriteLine("-t" & Properties.ZipExeOutputFolder & "\DIALOG.TXT")
                FChn.WriteLine("-a" & Properties.ZipExeOutputFolder & "\ABOUT.TXT")
                'FChn.WriteLine( "-stTST Enhancement Installation for " & Properties.Code4 & " " & Properties.Version
                FChn.WriteLine("-3")
                FChn.WriteLine("-win32")
                If Properties.AddlFilesFolder <> "" Then
                    FChn.WriteLine("-i " + Properties.AddlFilesFolder + "\ICON.ICO")
                Else
                    FChn.WriteLine("-i " + StdFilesPath + "\ICON.ICO")
                End If
                FChn.WriteLine("-c SETUP.EXE")
                FChn.Flush()
                FChn.Close()

                ' Delete old S/E Exe, if present.
                If My.Computer.FileSystem.FileExists(Properties.ZipExeOutputFolder & "\" _
                & Properties.ZipExeFilename & ".exe") Then
                    My.Computer.FileSystem.DeleteFile(Properties.ZipExeOutputFolder & "\" &
                         Properties.ZipExeFilename & ".exe")
                End If

                ' Create SE/Zip Exe.
                ' This should be replaced with the new NSIS Installer
                sTmp = Chr(34) & WinZipProgram _
                       & Chr(34) & " " & Properties.ZipExeOutputFolder & "\" _
                       & Properties.ZipExeFilename _
                       & " @" & Properties.ZipExeOutputFolder & "\SEINPUT.TXT"
                proc = New Process()
                proc.StartInfo.FileName = sTmp
                proc.Start()
                proc.WaitForExit()

                ' Delete temporary files.
                If My.Computer.FileSystem.FileExists(Properties.ZipExeOutputFolder _
                & "\" & "DIALOG.TXT") Then _
                    My.Computer.FileSystem.DeleteFile(Properties.ZipExeOutputFolder _
                         & "\" & "DIALOG.TXT")
                If My.Computer.FileSystem.FileExists(Properties.ZipExeOutputFolder _
                & "\" & "ABOUT.TXT") Then _
                    My.Computer.FileSystem.DeleteFile(Properties.ZipExeOutputFolder _
                         & "\" & "ABOUT.TXT")
                If My.Computer.FileSystem.FileExists(Properties.ZipExeOutputFolder _
                & "\" & "SEINPUT.TXT") Then _
                    My.Computer.FileSystem.DeleteFile(Properties.ZipExeOutputFolder _
                         & "\" & "SEINPUT.TXT")
            End If

            ' New Methods here, retaining old methods
            If (OnNetwork And NSISAvailable) Then
                'Make new Installer here

                ' Delete old .ZIP file, if present.
                If My.Computer.FileSystem.FileExists(Properties.ZipExeOutputFolder & "\" _
                & Properties.ZipExeFilename & ".zip") Then
                    My.Computer.FileSystem.DeleteFile(Properties.ZipExeOutputFolder & "\" &
                         Properties.ZipExeFilename & ".zip")
                End If
                If bCreateUpdateDelta Then
                    If My.Computer.FileSystem.FileExists(Properties.ZipExeOutputFolder & "\" _
                    & Properties.ZipExeFilename & "U.zip") Then
                        My.Computer.FileSystem.DeleteFile(Properties.ZipExeOutputFolder & "\" &
                             Properties.ZipExeFilename & "U.zip")
                    End If
                End If

                ' Create .ZIP file.

                ' Delete old S/E Exe, if present.
                If My.Computer.FileSystem.FileExists(Properties.ZipExeOutputFolder & "\" _
                & Properties.ZipExeFilename & ".exe") Then
                    My.Computer.FileSystem.DeleteFile(Properties.ZipExeOutputFolder & "\" &
                         Properties.ZipExeFilename & ".exe")
                End If
                If bCreateUpdateDelta Then
                    If My.Computer.FileSystem.FileExists(Properties.ZipExeOutputFolder & "\" _
                    & Properties.ZipExeFilename & "U.exe") Then
                        My.Computer.FileSystem.DeleteFile(Properties.ZipExeOutputFolder & "\" &
                             Properties.ZipExeFilename & "U.exe")
                    End If
                End If

                Me.BackgroundWorker1.ReportProgress(PercentComplete, Properties.ZipExeFilename & " Create Installer " & Properties.ZipExeFilename)

                FChn = My.Computer.FileSystem.OpenTextFileWriter(Properties.ZipExeOutputFolder & "\installer.nsi", False, Encoding.ASCII)
                FChn.WriteLine(";Generated Install file for NSIS")
                FChn.WriteLine("")
                FChn.WriteLine("!define PRODUCT_NAME """ & Properties.ProjectName & """")
                FChn.WriteLine("!define PRODUCT_VERSION """ & Properties.Version & """")
                FChn.WriteLine("")
                FChn.WriteLine("; MUI 1.67 compatible ------")
                FChn.WriteLine("!include " & """MUI.nsh""")
                FChn.WriteLine("")
                FChn.WriteLine("!define MUI_ABORTWARNING")
                FChn.WriteLine("!include """ & StdFilesPath & "\standard.nsh""")
                FChn.WriteLine("!define MUI_WELCOMEPAGE_TEXT ""${PRE_WELCOME_TEXT}Product: ${PRODUCT_NAME}\r\nVersion: ${PRODUCT_VERSION}\r\nRelease Date: " & Properties.ReleaseDate & "\r\n\r\n${POST_WELCOME_TEXT}$_CLICK""")
                FChn.WriteLine("!insertmacro MUI_PAGE_WELCOME")
                If My.Computer.FileSystem.FileExists(StdFilesPath & "\license.rtf") Then
                    FChn.WriteLine("!define MUI_LICENSEPAGE_RADIOBUTTONS")
                    FChn.WriteLine("!insertmacro MUI_PAGE_LICENSE """ & StdFilesPath & "\license.rtf""")
                End If
                FChn.WriteLine("!define MUI_DIRECTORYPAGE_VERIFYONLEAVE")
                FChn.WriteLine("!define MUI_PAGE_CUSTOMFUNCTION_LEAVE dirLeave")

                FChn.WriteLine("!insertmacro MUI_PAGE_DIRECTORY")
                FChn.WriteLine("!insertmacro MUI_PAGE_INSTFILES")
                FChn.WriteLine("!insertmacro MUI_PAGE_FINISH")
                FChn.WriteLine("!insertmacro MUI_LANGUAGE """ & "English""")
                FChn.WriteLine("Name " & """${PRODUCT_NAME} ${PRODUCT_VERSION}""")
                If OnNetwork Then
                    FChn.WriteLine("OutFile """ & Properties.ZipExeOutputFolder & "\" & Properties.ZipExeFilename & ".exe""")
                Else
                    FChn.WriteLine("OutFile " & """setup.exe""")
                End If
                FChn.WriteLine("")
                FChn.WriteLine("ShowInstDetails nevershow")
                FChn.WriteLine("ShowUnInstDetails nevershow")
                FChn.WriteLine("")
                ' main section will always need to be built
                FChn.WriteLine("Section " & """Main""" & " SEC01")
                If INSTAvailable And (Val(Properties.Version) >= 4) Then
                    FChn.WriteLine(vbTab & "InitPluginsDir")
                End If
                FChn.WriteLine(vbTab & "SetOutPath " & """$INSTDIR""")
                FChn.WriteLine(vbTab & "SetOverwrite on")
                ' Optionally include manual PDF in SE/Zip distribution.
                If Properties.IncludePDFinSEZip = "1" Then
                    If My.Computer.FileSystem.FileExists(ManualLibrary & "\" &
                                      Mid(Properties.ZipExeFilename, 1, 7) & ".pdf") Then
                        My.Computer.FileSystem.CreateDirectory(DistributionDirectory & "\pdf")
                        File.Copy(ManualLibrary & "\" & Mid(Properties.ZipExeFilename, 1, 7) & ".pdf",
                             DistributionDirectory & "\pdf" & "\" & Mid(Properties.ZipExeFilename, 1, 7) & ".pdf", True)
                    End If
                End If
                If INSTAvailable And (Val(Properties.Version) >= 4) Then
                    FChn.WriteLine(vbTab & "DetailPrint " & """Checking ProvideX COM Object...""")
                    FChn.WriteLine(vbTab & "PVXINST::CheckCom")
                    FChn.WriteLine(vbTab & "Pop $R0")
                    FChn.WriteLine(vbTab & "DetailPrint $R0")
                    FChn.WriteLine(vbTab & "StrCmp $R0 " & """1""" & " +2")
                    FChn.WriteLine(vbTab & "Abort " & """The Providex COM object is not registered""")
                    FChn.WriteLine(vbTab & "DetailPrint " & """Initializing ProvideX COM object...""")
                    FChn.WriteLine(vbTab & "PVXINST::LoadCom")
                    FChn.WriteLine(vbTab & "Pop $R0")
                    FChn.WriteLine(vbTab & "DetailPrint $R0")
                    FChn.WriteLine(vbTab & "StrCmp $R0 " & """1""" & " +2")
                    FChn.WriteLine(vbTab & "Abort " & """There was an error initializing the ProvideX COM object""")
                End If
                FChn.WriteLine(vbTab & "File /r """ & DistributionDirectory & "\*.*""")
                ' Add Additional Files to the setup
                If Not INSTAvailable Then
                    FChn.WriteLine("")
                End If
                If INSTAvailable And (Val(Properties.Version) >= 4) Then
                    FChn.WriteLine(vbTab & "DetailPrint " & """Preparing installation...""")
                    FChn.WriteLine(vbTab & "PVXINST::PreInstall")
                    FChn.WriteLine(vbTab & "Pop $R0")
                    FChn.WriteLine(vbTab & "DetailPrint $R0")
                    FChn.WriteLine(vbTab & "StrCmp $R0 " & """1""" & " +2")
                    FChn.WriteLine(vbTab & "MessageBox MB_OK " & """Preparation failed.""")
                    FChn.WriteLine(vbTab & "DetailPrint " & """Installing...""")
                    FChn.WriteLine(vbTab & "PVXINST::Install")
                    FChn.WriteLine(vbTab & "Pop $R0")
                    FChn.WriteLine(vbTab & "Pop $R1")
                    FChn.WriteLine(vbTab & "DetailPrint " & """$R0 $R1""")
                    FChn.WriteLine(vbTab & "StrCmp $R0 " & """1""" & " +2")
                    FChn.WriteLine(vbTab & "MessageBox MB_OK $R1")
                    FChn.WriteLine(vbTab & "DetailPrint " & """Finalizing...""")
                    FChn.WriteLine(vbTab & "PVXINST::GradualProgress /NOUNLOAD 1 2 100 " & """Installation Finished.""")
                    FChn.WriteLine(vbTab & "PVXINST::Finalize /NOUNLOAD")
                    FChn.WriteLine(vbTab & "DetailPrint " & """Installation Finished.""")
                End If
                If Not INSTAvailable Then
                    If My.Computer.FileSystem.FileExists(Properties.AddlFilesFolder & "\PVXINST.EXE") And (Val(Properties.Version) >= 4) Then
                        FChn.WriteLine(vbTab & "File /oname=$TEMP\PVXINST.EXE """ & Properties.AddlFilesFolder & "\PVXINST.EXE""")
                        FChn.WriteLine(vbTab & "ExecWait '""" & "$TEMP\PVXINST.EXE""" & " $INSTDIR'")
                        FChn.WriteLine(vbTab & "Delete $TEMP\PVXINST.EXE")
                    End If
                End If
                FChn.WriteLine("SectionEnd")
                ' We could include optional section information here that would never change
                '
                ' We could include a function header in right here
                FChn.WriteLine("Function .onInit")
                FChn.WriteLine(vbTab & "ReadRegStr $INSTDIR HKCU " & """Software\ODBC\ODBC.INI\SOTAMAS90""" & " " & """MAS90RootDirectory""")
                FChn.WriteLine("FunctionEnd")
                FChn.WriteLine("Function dirLeave")
                FChn.WriteLine(vbTab & "IfFileExists $INSTDIR\Home\pvxwin32.exe PathGood")
                FChn.WriteLine(vbTab & "Abort ; ")
                FChn.WriteLine("PathGood:")
                If Properties.PartNumber = "TST LMAV-A" Then
                    FChn.WriteLine(vbTab & "PVXINST::CheckLMAVVersion")
                    FChn.WriteLine(vbTab & "Pop $R0")
                    FChn.WriteLine(vbTab & "StrCmp $R0 " & """0""" & " +3")
                    FChn.WriteLine(vbTab & "MessageBox MB_OK " & """This Avatax connector version is incompatible with the version currently installed.$\r$\nThe installation cannot continue.""")
                    FChn.WriteLine(vbTab & "Abort ; if LMAV is not the right veriion and there is a module, do not install")
                End If
                FChn.WriteLine(vbTab & "PVXINST::CheckVersion")
                FChn.WriteLine(vbTab & "Pop $R0")
                FChn.WriteLine(vbTab & "StrCmp $R0 ${PRODUCT_VERSION} +3")
                FChn.WriteLine(vbTab & "MessageBox MB_OK " & """The version being installed is not compatible with the current version.""")
                FChn.WriteLine(vbTab & "Abort ; ")
                FChn.WriteLine("FunctionEnd")
                If INSTAvailable And (Val(Properties.Version) >= 4) Then
                    FChn.WriteLine("Function .onGUIEnd")
                    FChn.WriteLine(vbTab & "PVXINST::Unload")
                    FChn.WriteLine("FunctionEnd")
                End If


                FChn.Flush()
                FChn.Close()

                If bCreateUpdateDelta Then
                    FChn = My.Computer.FileSystem.OpenTextFileWriter(Properties.ZipExeOutputFolder & "\Uinstaller.nsi", False, Encoding.ASCII)
                    FChn.WriteLine(";Generated Install file for NSIS")
                    FChn.WriteLine("")
                    FChn.WriteLine("!define PRODUCT_NAME """ & Properties.ProjectName & """")
                    FChn.WriteLine("!define PRODUCT_VERSION """ & Properties.Version & """")
                    FChn.WriteLine("")
                    FChn.WriteLine("; MUI 1.67 compatible ------")
                    FChn.WriteLine("!include " & """MUI.nsh""")
                    FChn.WriteLine("")
                    FChn.WriteLine("!define MUI_ABORTWARNING")
                    FChn.WriteLine("!include """ & StdFilesPath & "\standard.nsh""")
                    FChn.WriteLine("!define MUI_WELCOMEPAGE_TEXT "" ${PRE_WELCOME_TEXT}Product: ${PRODUCT_NAME}\r\nVersion: ${PRODUCT_VERSION}\r\nRelease Date: " & Properties.ReleaseDate & "\r\n\r\n${POST_WELCOME_TEXT}$_CLICK""")
                    FChn.WriteLine("!insertmacro MUI_PAGE_WELCOME")
                    If My.Computer.FileSystem.FileExists(StdFilesPath & "\license.rtf") Then
                        FChn.WriteLine("!define MUI_LICENSEPAGE_RADIOBUTTONS")
                        FChn.WriteLine("!insertmacro MUI_PAGE_LICENSE """ & StdFilesPath & "\license.rtf""")
                    End If
                    FChn.WriteLine("!define MUI_DIRECTORYPAGE_VERIFYONLEAVE")
                    FChn.WriteLine("!define MUI_PAGE_CUSTOMFUNCTION_LEAVE dirLeave")
                    FChn.WriteLine("!insertmacro MUI_PAGE_DIRECTORY")
                    FChn.WriteLine("!insertmacro MUI_PAGE_INSTFILES")
                    FChn.WriteLine("!insertmacro MUI_PAGE_FINISH")
                    FChn.WriteLine("!insertmacro MUI_LANGUAGE """ & "English""")
                    FChn.WriteLine("Name " & """${PRODUCT_NAME} ${PRODUCT_VERSION}""")
                    FChn.WriteLine("OutFile """ & Properties.ZipExeOutputFolder & "\" & Properties.ZipExeFilename & "U.exe""")
                    FChn.WriteLine("")
                    FChn.WriteLine("ShowInstDetails nevershow")
                    FChn.WriteLine("ShowUnInstDetails nevershow")
                    FChn.WriteLine("")
                    ' main section will always need to be built
                    FChn.WriteLine("Section " & """Main""" & " SEC01")
                    If INSTAvailable And (Val(Properties.Version) >= 1) Then
                        FChn.WriteLine(vbTab & "InitPluginsDir")
                    End If
                    FChn.WriteLine(vbTab & "SetOutPath " & """$INSTDIR""")
                    FChn.WriteLine(vbTab & "SetOverwrite on")
                    If INSTAvailable And (Val(Properties.Version) >= 1) Then
                        FChn.WriteLine(vbTab & "DetailPrint " & """Checking ProvideX COM Object...""")
                        FChn.WriteLine(vbTab & "PVXINST::CheckCom")
                        FChn.WriteLine(vbTab & "Pop $R0")
                        FChn.WriteLine(vbTab & "DetailPrint $R0")
                        FChn.WriteLine(vbTab & "StrCmp $R0 " & """1""" & " +2")
                        FChn.WriteLine(vbTab & "Abort " & """The Providex COM object is not registered""")
                        FChn.WriteLine(vbTab & "DetailPrint " & """Initializing ProvideX COM object...""")
                        FChn.WriteLine(vbTab & "PVXINST::LoadCom")
                        FChn.WriteLine(vbTab & "Pop $R0")
                        FChn.WriteLine(vbTab & "DetailPrint $R0")
                        FChn.WriteLine(vbTab & "StrCmp $R0 " & """1""" & " +2")
                        FChn.WriteLine(vbTab & "Abort " & """There was an error initializing the ProvideX COM object""")
                    End If
                    FChn.WriteLine(vbTab & "File /r """ & DistributionDirectory & "U\*.*""")
                    ' Add Additional Files to the setup
                    If Not INSTAvailable Then
                        FChn.WriteLine("")
                    End If
                    If INSTAvailable And (Val(Properties.Version) >= 4) Then
                        FChn.WriteLine(vbTab & "DetailPrint " & """Preparing installation...""")
                        FChn.WriteLine(vbTab & "PVXINST::PreInstall")
                        FChn.WriteLine(vbTab & "Pop $R0")
                        FChn.WriteLine(vbTab & "DetailPrint $R0")
                        FChn.WriteLine(vbTab & "StrCmp $R0 " & """1""" & " +2")
                        FChn.WriteLine(vbTab & "MessageBox MB_OK " & """The Preparation failed""")
                        FChn.WriteLine(vbTab & "DetailPrint " & """Installing...""")
                        FChn.WriteLine(vbTab & "PVXINST::Install")
                        FChn.WriteLine(vbTab & "Pop $R0")
                        FChn.WriteLine(vbTab & "Pop $R1")
                        FChn.WriteLine(vbTab & "DetailPrint " & """$R0 $R1""")
                        FChn.WriteLine(vbTab & "StrCmp $R0 " & """1""" & " +2")
                        FChn.WriteLine(vbTab & "MessageBox MB_OK $R1")
                        FChn.WriteLine(vbTab & "DetailPrint " & """Finalizing the installation...""")
                        FChn.WriteLine(vbTab & "PVXINST::GradualProgress /NOUNLOAD 1 2 100 " & """Installation Finished.""")
                        FChn.WriteLine(vbTab & "PVXINST::InstallFinalize /NOUNLOAD")
                        FChn.WriteLine(vbTab & "DetailPrint " & """Installation Finished.""")
                    End If
                    If Not INSTAvailable Then
                        If My.Computer.FileSystem.FileExists(Properties.AddlFilesFolder & "\PVXINST.EXE") And (Val(Properties.Version) >= 4) Then
                            FChn.WriteLine(vbTab & "File /oname=$TEMP\PVXINST.EXE """ & Properties.AddlFilesFolder & "\PVXINST.EXE""")
                            FChn.WriteLine(vbTab & "ExecWait '""" & "$TEMP\PVXINST.EXE""" & " $INSTDIR'")
                            FChn.WriteLine(vbTab & "Delete $TEMP\PVXINST.EXE")
                        End If
                    End If
                    FChn.WriteLine("SectionEnd")
                    ' We could include optional section information here that would never change
                    '
                    ' We could include a function header in right here
                    FChn.WriteLine("Function .onInit")
                    FChn.WriteLine(vbTab & "ReadRegStr $INSTDIR HKCU " & """Software\ODBC\ODBC.INI\SOTAMAS90""" & " " & """MAS90RootDirectory""")
                    FChn.WriteLine("FunctionEnd")
                    FChn.WriteLine("Function dirLeave")
                    FChn.WriteLine(vbTab & "IfFileExists $INSTDIR\Home\pvxwin32.exe PathGood")
                    FChn.WriteLine(vbTab & "Abort ; ")
                    FChn.WriteLine("PathGood:")
                    FChn.WriteLine(vbTab & "PVXINST::CheckVersion")
                    FChn.WriteLine(vbTab & "Pop $R0")
                    FChn.WriteLine(vbTab & "StrCmp $R0 ${PRODUCT_VERSION} +3")
                    FChn.WriteLine(vbTab & "MessageBox MB_OK " & """The Enhancement version being installed is not compatible with this version of Sage 100.$\r$\nThe installation cannot continue.""")
                    FChn.WriteLine(vbTab & "Abort ; ")
                    FChn.WriteLine("FunctionEnd")
                    If INSTAvailable And (Val(Properties.Version) >= 4) Then
                        FChn.WriteLine("Function .onGUIEnd")
                        FChn.WriteLine(vbTab & "PVXINST::Unload")
                        FChn.WriteLine("FunctionEnd")
                    End If

                    FChn.Flush()
                    FChn.Close()

                End If
                If My.Computer.FileSystem.FileExists(Properties.ZipExeOutputFolder & "\STUBSETUP.EXE") Then
                    My.Computer.FileSystem.DeleteFile(Properties.ZipExeOutputFolder & "\STUBSETUP.EXE")
                End If
                If My.Computer.FileSystem.FileExists(Properties.ZipExeOutputFolder & "\Installation.txt") Then
                    My.Computer.FileSystem.DeleteFile(Properties.ZipExeOutputFolder & "\Installation.txt")
                End If
                ' Should have enough info to make installer, without extra files

                Dim info As ProcessStartInfo = New ProcessStartInfo(MakeNSIS)
                info.UseShellExecute = False
                info.CreateNoWindow = True
                info.Arguments = "/O" &
                        Properties.ZipExeOutputFolder & "\install.log " &
                        Properties.ZipExeOutputFolder & "\installer.nsi"
                Dim Nsisproc As Process = Process.Start(info)
                Nsisproc.WaitForExit()


                If My.Computer.FileSystem.FileExists(Properties.ZipExeOutputFolder & "\installer.nsi") Then
                    My.Computer.FileSystem.DeleteFile(Properties.ZipExeOutputFolder & "\installer.nsi")
                End If
                If My.Computer.FileSystem.FileExists(Properties.ZipExeOutputFolder & "\standard.nsh") Then
                    My.Computer.FileSystem.DeleteFile(Properties.ZipExeOutputFolder & "\standard.nsh")
                End If

            End If

            If OnNetwork _
            And (WinZipAvailable Or NSISAvailable) _
            And UpdateWebsite Then
                Me.BackgroundWorker1.ReportProgress(PercentComplete, Properties.ZipExeFilename & " Create HTML download page")
                FChn = My.Computer.FileSystem.OpenTextFileWriter(UploadStaging & "\" & LCase$(Mid$(Properties.ZipExeFilename, 1, 4)) & ".html", False, Encoding.ASCII)
                FChn.WriteLine("<html>")
                FChn.WriteLine("<head><title>" & Properties.ProjectName & " Download</title></head>")
                FChn.WriteLine("<body bgcolor=""#FFFFFF""")
                FChn.WriteLine("<p><b><font size=""2"" face=""Verdana"" color=""#000080"">" & Properties.ProjectName & " Product Download Page</font></b><br>")
                FChn.WriteLine("")
                FChn.WriteLine("<font face=""Verdana"" color=""#000080""><font size=""-2"">Version: " & Properties.Version & "</font><br>")
                FChn.WriteLine("<font face=""Verdana"" color=""#000080""><font size=""-2"">Date Updated: " & Properties.ReleaseDate & "</font><br>")
                FChn.WriteLine("<font face=""Verdana"" color=""#000080""><font size=""-2"">Date 1st Released: " & Properties.OriginDate & "</font></p>")
                FChn.WriteLine("")

                ' Installation D/L Line
                FChn.WriteLine("<p><font face=""Verdana"" color=""#000080""><font size=""-2"">Download Installation Program: </font>")
                FChn.WriteLine("")
                FChn.WriteLine("<b><a href=""../downloads/" & Properties.ZipExeFilename & "" & ".EXE" & """><font face=""Verdana"" size=""2"" color=""#9999FF"">" & Properties.ZipExeFilename & ".EXE" & "</font></a></b>")
                ' Manual D/L Line
                FChn.WriteLine("<br><font face=""Verdana"" color=""#000080""><font size=""-2"">Manual in Word format: </font>")
                FChn.WriteLine("")
                FChn.WriteLine("<b><a href=""../downloads/" & Properties.ZipExeFilename & "" & ".DOC" & """><font face=""Verdana"" size=""2"" color=""#9999FF"">" & Properties.ZipExeFilename & ".DOC" & "</font></a></b>")
                If My.Computer.FileSystem.FileExists(ManualLibrary & "\" & Mid(Properties.ZipExeFilename, 1, 7) & ".pdf") Then
                    FChn.WriteLine("<br><font face=""Verdana"" color=""#000080""><font size=""-2"">Manual in Adobe Acrobat (PDF) format: </font>")
                    FChn.WriteLine("")
                    FChn.WriteLine("<b><a href=""../downloads/" & Properties.ZipExeFilename & "" & ".PDF" & """><font face=""Verdana"" size=""2"" color=""#9999FF"">" & Properties.ZipExeFilename & ".PDF" & "</font></a></b>")
                    FChn.WriteLine("<a href=""http://www.adobe.com/products/acrobat/readstep.html""><img border=""0"" src=""http://www.adobe.com/images/getacro.gif"" align=""right"" width=""88"" height=""31""></a>")
                End If

                ' versions this distribution is compatible with
                FChn.WriteLine("<p><font size=""-2"" face=""Verdana"" color=""#000080""> Compatible with these PXPlus versions:</font></p>")
                FChn.WriteLine("<blockquote><p><font size=""-2"" face=""Verdana"" color=""#000080"">")
                sTmp = ""
                For x = 1 To Len(sModules)
                    If Mid$(sModules, x, 1) = ";" Then
                        If sTmp <> "" Then
                            FChn.WriteLine(sTmp & "<br>")
                            sTmp = ""
                        End If
                    Else
                        sTmp = sTmp & Mid$(sModules, x, 1)
                    End If
                Next x
                FChn.WriteLine("</p></blockquote>")
                FChn.WriteLine("<p><b><font size=""-2"" face=""Verdana"" color=""#000080"">To install, run the downloaded program <b>" _
                             & Properties.ZipExeFilename & ".EXE" & "</b> and follow on-screen instructions. Continue as instructed in the Manual.</font></p>")
                FChn.WriteLine("<p><b><font size=""-2"" face=""Verdana"" color=""#000080"">Other versions:</b> If you require a different version of this product, e-mail us at <a href=""mailto:cspirz@gmail.com""><font size=""-2"" face=""Verdana"" color=""#9999FF"">cspirz@gmail.com</font></a>.")

                FChn.WriteLine("</body>")
                FChn.WriteLine("</html>")
                FChn.Flush()
                FChn.Close()

            End If

            ' TST Only: Create Product Version Build Date File
            If OnNetwork _
            And (WinZipAvailable Or NSISAvailable) _
            And UpdateWebsite Then
                sTmp = UploadStaging & "\" & Properties.ZipExeFilename & "BUILDFILE.BLD"
                FChn = My.Computer.FileSystem.OpenTextFileWriter(sTmp, False, Encoding.ASCII)
                FChn.WriteLine(Properties.ReleaseDate)
                FChn.WriteLine("## This is the Release Date of the currently posted Version.")
                FChn.Flush()
                FChn.Close()

            End If

            If OnNetwork _
            And (WinZipAvailable Or NSISAvailable) _
            And UpdateWebsite Then
                Me.BackgroundWorker1.ReportProgress(PercentComplete, Properties.ZipExeFilename & " Copy to upload directory")
                File.Copy(Properties.ZipExeOutputFolder & "\" & Properties.ZipExeFilename & ".exe",
                                   UploadStaging & "\" & Properties.ZipExeFilename & ".exe", True)

            End If

            ' TST Only: Copy Manual (.doc and .pdf) to upload staging directory
            If OnNetwork _
            And (WinZipAvailable Or NSISAvailable) _
            And UpdateWebsite Then
                Me.BackgroundWorker1.ReportProgress(PercentComplete, Properties.ZipExeFilename & "Copy Manual to upload staging directory")
                If My.Computer.FileSystem.FileExists(ManualLibrary & "\" &
                                  Mid(Properties.ZipExeFilename, 1, 7) & ".doc") Then
                    File.Copy(ManualLibrary & "\" &
                                  Mid(Properties.ZipExeFilename, 1, 7) & ".doc",
                                   UploadStaging & "\" & Mid(Properties.ZipExeFilename, 1, 7) & ".doc", True)
                Else
                    MessageBox.Show(Properties.ZipExeFilename & ".doc" & " (manual) does not exist." _
                        & Environment.NewLine & Environment.NewLine & "Unable to copy manual to the Enhancement Upload Staging Folder.")
                End If
                If My.Computer.FileSystem.FileExists(ManualLibrary & "\" &
                                  Mid(Properties.ZipExeFilename, 1, 7) & ".pdf") Then
                    File.Copy(ManualLibrary & "\" &
                                  Mid(Properties.ZipExeFilename, 1, 7) & ".pdf",
                                   UploadStaging & "\" & Mid(Properties.ZipExeFilename, 1, 7) & ".pdf", True)
                End If
            End If

            If Properties.IsSystem _
            And OnNetwork Then
                If My.Computer.FileSystem.FileExists(ManualLibrary & "\" &
                                  Mid(Properties.ZipExeFilename, 1, 7) & ".pdf") Then
                    File.Copy(ManualLibrary & "\" & Mid(Properties.ZipExeFilename, 1, 7) & ".pdf",
                         DistributionDirectory & "\pdf\" & Mid(Properties.ZipExeFilename, 1, 7) & ".pdf", True)
                End If
            End If
            log.Info("Successful Distribution of " & Properties.CodeDescription & " Version: " & Properties.Version & " " & Properties.ZipExeFilename)

            Exit Sub
        Catch ex As Exception
            log.Error("Failed Distribution of " & Properties.CodeDescription & " Version: " & Properties.Version & " " & Properties.ZipExeFilename, ex)

        End Try

    End Sub

    Public Sub CreateInstallationClass(ByVal TestBaseDirectory As String, ByVal ProjectBaseDirectory As String)
        Dim sPrgFile As String
        Dim sInstallFile As String
        Dim FChn As StreamWriter
        Dim sTemplateFile As String
        Dim FTpl As StreamReader
        Dim sBuffer As String
        Dim CodeZillaPath As String = "Q:\DEV\VB\IDE"


        sPrgFile = "TempPrg"
        If Properties.PrimaryModule = "LM" Then
            sInstallFile = ProjectBaseDirectory & "\SOA\IN_Installation.pvc"
        Else
            sInstallFile = ProjectBaseDirectory & "\" & Properties.PrimaryModule & "\" & Properties.PrimaryModule & "_Installation.pvc"
        End If
        sTemplateFile = CodeZillaPath + "\xx_Installation.txt"
        My.Computer.FileSystem.CurrentDirectory = TestBaseDirectory & "\HOME"
        FChn = My.Computer.FileSystem.OpenTextFileWriter(sInstallFile, False, System.Text.Encoding.ASCII)
        FChn.WriteLine("! xx_Installation class definition Template")
        FChn.WriteLine("! Created 06/28/04 By TST Business Systems")
        FChn.WriteLine("! This template is used for 4.0 Installation")
        FChn.WriteLine("! Rename to the Primary Module (i.e. AP_Installation.pvc for APCB.400)")
        FChn.WriteLine("!")
        If Properties.PrimaryModule = "LM" Then
            FChn.WriteLine("def class """ & "IN_Installation" & """")
        Else
            FChn.WriteLine("def class " & """" & Properties.PrimaryModule & "_Installation" & """")
        End If
        FChn.WriteLine("like """ & "SY_Installation" & """")
        FChn.WriteLine("function POSTINSTALL()POST_INSTALL")
        FChn.WriteLine("function PREINSTALL()PRE_INSTALL")
        FChn.WriteLine("function PREFINALIZE()PRE_FINALIZE")
        FChn.WriteLine("function POSTFINALIZE()POST_FINALIZE")
        FChn.WriteLine("function POSTUNINSTALL()POST_UNINSTALL")
        FChn.WriteLine("function PREUNINSTALL()PRE_UNINSTALL")
        FChn.WriteLine("end def")
        FChn.WriteLine("!")
        FChn.WriteLine("ON_CREATE:")
        FChn.WriteLine("! Required to function, set TERM$ for SUMDI4")
        FChn.WriteLine("enter COSESSION")
        FChn.WriteLine("Return")
        If My.Computer.FileSystem.FileExists(sTemplateFile) Then
            FTpl = My.Computer.FileSystem.OpenTextFileReader(sTemplateFile, Encoding.ASCII)
            Do While Not FTpl.EndOfStream
                sBuffer = FTpl.ReadLine()
                FChn.WriteLine(sBuffer)
            Loop
            FTpl.Close()
        Else
            FChn.WriteLine("!")
            FChn.WriteLine("!")
            FChn.WriteLine("POST_INSTALL:")
            FChn.WriteLine("! Run after the files have been copied from xxINST in MAS90")
            FChn.WriteLine("! This should be called to run SUMDI4, after any new Files have been placed in xxINST")
            FChn.WriteLine("! These are the only calls required to install")
            FChn.WriteLine("TERM$ = COSESSION'LEGACYTERM$")
            FChn.WriteLine("call " & """" & "..\SOA\SUMDI4" & """" & ",err=*next,CMP_STRING$,TERM$")
            FChn.WriteLine("if not(nul(CMP_STRING$)) then {")
            FChn.WriteLine("FOR COMPILE=1 TO (LEN(CMP_STRING$)-1);")
            FChn.WriteLine("App$ = CMP_STRING$(COMPILE, 2)")
            FChn.WriteLine("%RECOMPILE_FLG=1,V$=APP$(1,1)+""/""+ App$(2, 1); ")
            FChn.WriteLine("TMP=_OBJ'COMPILEDICTIONARY(V$)")
            FChn.WriteLine("Next COMPILE")
            FChn.WriteLine("}")
            FChn.WriteLine("Return")
            FChn.WriteLine("!")
            FChn.WriteLine("PRE_INSTALL:")
            FChn.WriteLine("!")
            FChn.WriteLine("! This hook is called from SY_Installation'Install before any files are merged from xxINST")
            FChn.WriteLine("!")
            FChn.WriteLine("Return")
            FChn.WriteLine("!")
            FChn.WriteLine("PRE_FINALIZE:")
            FChn.WriteLine("!")
            FChn.WriteLine("! This hook is called from SY_Installation after mergeing, but before compiling menus")
            FChn.WriteLine("Return")
            FChn.WriteLine("!")
            FChn.WriteLine("POST_FINALIZE:")
            FChn.WriteLine("! Called from SY_Installation after all installation related task completed.")
            FChn.WriteLine("! Menus are compiled here.")
            FChn.WriteLine("!")
            FChn.WriteLine("Return")
            FChn.WriteLine("POST_UNINSTALL:")
            FChn.WriteLine("! Called from SY_Installation after all uninstallation related task completed.")
            FChn.WriteLine("!")
            FChn.WriteLine("Return")
            FChn.WriteLine("PRE_UNINSTALL:")
            FChn.WriteLine("! Called from SY_Installation before all uninstallation related task completed.")
            FChn.WriteLine("!")
            FChn.WriteLine("Return")
        End If
        FChn.Flush()
        FChn.Close()

        FChn = My.Computer.FileSystem.OpenTextFileWriter(TestBaseDirectory & "\Home\" & sPrgFile, False, System.Text.Encoding.ASCII)
        ' Pass Data to Pvx, Initialize Vars
        FChn.WriteLine("01 FILENAME$ = """ & sInstallFile & """")
        FChn.WriteLine("20 PREINPUT """ & "LOAD ""+quo+FILENAME$+quo")
        FChn.WriteLine("30 PREINPUT """ & "ERASE ""+quo+FILENAME$+quo")
        FChn.WriteLine("40 PREINPUT """ & "PROGRAM ""+quo+FILENAME$+quo")
        FChn.WriteLine("50 PREINPUT """ & "SAVE ""+quo+FILENAME$+quo")
        FChn.WriteLine("60 PREINPUT """ & "QUIT""")
        FChn.WriteLine("70 ESCAPE")
        FChn.Flush()
        FChn.Close()

        Dim info As New ProcessStartInfo(TestBaseDirectory & "\HOME\PVXWIN32.EXE")
        info.WorkingDirectory = TestBaseDirectory & "\HOME\"
        info.Arguments = "-HD ..\HOME\" & sPrgFile
        info.CreateNoWindow = True
        info.UseShellExecute = False

        Dim proc As Process = Process.Start(info)
        proc.WaitForExit()

        My.Computer.FileSystem.DeleteFile(TestBaseDirectory & "\Home\" & sPrgFile)

    End Sub

    Public Function GetSSN(ByVal TestSystemHome As String) As String ' Returns PvX SNN function result
        Dim sPrgFile As String
        Dim reader As StreamReader
        Dim FChn As StreamWriter
        Dim sReadString As String
        Dim sReturnFile As String

        sPrgFile = "TempPrg"
        sReturnFile = TestSystemHome & "\Return.DAT"
        My.Computer.FileSystem.CurrentDirectory = TestSystemHome
        FChn = My.Computer.FileSystem.OpenTextFileWriter(TestSystemHome & "\" & sPrgFile, False, Encoding.ASCII)
        FChn.WriteLine("20 RETURNFILENAME$ = """ & sReturnFile & """")
        FChn.WriteLine("210 ERASE RETURNFILENAME$,ERR=*NEXT")
        FChn.WriteLine("220 SERIAL RETURNFILENAME$")
        FChn.WriteLine("230 RCHN=UNT,E1=RCHN;OPEN LOCK(E1,ISZ=-1) RETURNFILENAME$")
        FChn.WriteLine("390 E1=RCHN;WRITE RECORD (E1) SSN")
        FChn.WriteLine("500 E1=RCHN;CLOSE(E1)")
        FChn.WriteLine("64998 QUIT")
        FChn.Flush()
        FChn.Close()


        Dim info As New ProcessStartInfo(TestSystemHome & "\PVXWIN32.EXE")
        info.WorkingDirectory = TestSystemHome
        info.Arguments = "-HD ..\HOME\" & sPrgFile
        info.CreateNoWindow = True
        info.UseShellExecute = False
        Dim proc As Process = Process.Start(info)
        proc.WaitForExit()

        reader = My.Computer.FileSystem.OpenTextFileReader(sReturnFile, Encoding.ASCII)
        sReadString = reader.ReadToEnd()
        reader.Close()

        My.Computer.FileSystem.DeleteFile(sPrgFile)
        My.Computer.FileSystem.DeleteFile(sReturnFile)
        Return sReadString.Substring(sReadString.Length - 3, 3)

    End Function

    Public Sub CreateDeveloperManifest(ByVal TestBaseDirectory As String, ByVal ProjectBaseDirectory As String)
        Dim sPrgFile As String, sMDFolder As String
        Dim FChn As StreamWriter
        ' Copy program needed from library to {test system}..\MD
        sMDFolder = TestBaseDirectory & "\MD"
        Call CopyFromLibrary("MD", sMDFolder, "MD_DeveloperManifest.M4P")
        My.Computer.FileSystem.CurrentDirectory = sMDFolder
        sPrgFile = "MDZTMP"
        FChn = My.Computer.FileSystem.OpenTextFileWriter(sPrgFile, False, Encoding.ASCII)
        FChn.WriteLine("10 BEGIN")
        FChn.WriteLine("20 %SY0ENH$ = """ & GetSY0ENH(ProjectBaseDirectory) & """")
        FChn.WriteLine("30 %ISMDPROJECT = -1")
        FChn.WriteLine("35 %ENH_4CODE$ = """ & Properties.CodeDescription & """")
        FChn.WriteLine("40 %VERSION$ = """ & Properties.Version & """")
        FChn.WriteLine("45 %PROJECT_DIR$ = """ & ProjectBaseDirectory & """")
        FChn.WriteLine("50 PERFORM ""MD_DeveloperManifest.M4P""")
        FChn.WriteLine("64998 BYE")
        FChn.Flush()
        FChn.Close()

        Dim info As New ProcessStartInfo(TestBaseDirectory & "\HOME\PVXWIN32.EXE")
        info.WorkingDirectory = TestBaseDirectory & "\HOME"
        info.Arguments = "-HD ../MD/MDZTMP"
        info.CreateNoWindow = True
        info.UseShellExecute = False
        Dim proc As Process = Process.Start(info)
        proc.WaitForExit()

        My.Computer.FileSystem.DeleteFile(sPrgFile)
    End Sub

    Public Sub CopyFromLibrary(ByVal SourceFolder As String,
                           ByVal DestinationFolder As String,
                           ByVal FileName As String)
        Dim CodeZillaPath As String = "Q:\DEV\VB\IDE"

        ' Copy Filename from IDE\{SourceFolder} to
        ' (full path) DestinationFolder
        File.Copy(CodeZillaPath & "\" & SourceFolder & "\" & FileName, DestinationFolder & "\" & FileName, True)

    End Sub

    Public Function GetSY0ENH(ByVal ProjectBaseDirectory As String) As String
        ' Returns a composed version of SY0ENH$, as used in Doug's
        ' file edit programs.
        Dim sSY0ENH As String
        Dim sTemp As String = ""
        Dim sAllModules As String = ""
        Dim sAllModulesOrig As String
        Dim Properties As New Properties(ProjectBaseDirectory & "\Properties.ini")
        sSY0ENH = Space(765) 'expanded this
        Mid$(sSY0ENH, 1, 3) = Properties.CodeIdentifier
        Mid$(sSY0ENH, 4) = Properties.ProjectName
        Mid$(sSY0ENH, 509) = ProjectBaseDirectory
        Mid$(sSY0ENH, 39, 2) = Properties.PrimaryModule
        Mid$(sSY0ENH, 1, 3) = Properties.CodeIdentifier
        ' Build string of all modules like GLAPPO...etc
        sAllModulesOrig = Properties.AllModules
        sAllModulesOrig = Replace(sAllModulesOrig, "-y", "")
        sAllModulesOrig = Replace(sAllModulesOrig, "-n", "")
        Dim x As Integer
        For x = 1 To Len(sAllModulesOrig)
            If Mid$(sAllModulesOrig, x, 1) = ";" Then
                If sTemp <> "" Then
                    sAllModules = sAllModules + sTemp
                    sTemp = ""
                End If
            Else
                sTemp = sTemp & Mid$(sAllModulesOrig, x, 1)
            End If
        Next x
        Mid$(sSY0ENH, 87) = sAllModules
        Mid$(sSY0ENH, 51) = Properties.Version
        Mid$(sSY0ENH, 41) = Properties.PartNumber
        ' Regularize date, pack and format as MAS90 6-char date $
        sTemp = Format(DateValue(Properties.ReleaseDate), "mm/dd/yyyy")
        Mid$(sSY0ENH, 61) = (Chr(Int(Val(Mid$(sTemp, 7, 4)) / 64) + 32)) & (Chr((Val(Mid$(sTemp, 7, 4)) Mod 64) + 32)) & Mid$(sTemp, 1, 2) & Mid$(sTemp, 4, 2)
        GetSY0ENH = sSY0ENH
    End Function

    Public Function CTLFilename(ByVal ProjectFolderBase As String) As String
        ' Retreives/constructs pathed filename of the project's CTL file.
        Dim sFilenamePrefix As String
        Dim sFilename As String
        Dim sPrimaryModuleFolder As String
        Dim Properties As New Properties(ProjectFolderBase & "\Properties.ini")
        If UCase$(Properties.PrimaryModule) = "LM" Then
            sFilenamePrefix = "IN"
            sPrimaryModuleFolder = "SOA"
        Else
            sFilenamePrefix = Properties.PrimaryModule
            sPrimaryModuleFolder = Properties.PrimaryModule
        End If
        sFilename = ProjectFolderBase & "\" & sPrimaryModuleFolder _
                & "\" & sFilenamePrefix & "0CTL." + Properties.CodeIdentifier
        If sFilename <> "" Then
            Return sFilename
        Else
            Return ""
        End If
    End Function

    Public Function GetModuleLevels(ByVal TestHomeDirectory As String, ByVal ProjectBaseDirectory As String) As String
        ' Returns $ containing each module in project with version,
        ' separated by ";" delimiter, with a trailing delimiter. ie:
        ' A/P 3.41;P/O 3.41;
        Dim sPrgFile As String
        Dim sReturnFile As String
        Dim sCTLFile As String
        Dim FChn As StreamWriter
        Dim sReadString As String

        Dim Properties As New Properties(ProjectBaseDirectory & "\Properties.ini")

        sPrgFile = "$Prg$"
        sReturnFile = TestHomeDirectory & "\$Return$.DAT"
        If Properties.PrimaryModule <> "LM" Then
            sCTLFile$ = ProjectBaseDirectory & "\" &
                        Properties.PrimaryModule & "\" &
                        Properties.PrimaryModule & "0CTL." &
                        Properties.CodeIdentifier
        Else
            sCTLFile$ = ProjectBaseDirectory & "\" &
                        "SOA\IN0CTL." & Properties.CodeIdentifier
        End If
        My.Computer.FileSystem.CurrentDirectory = TestHomeDirectory
        FChn = My.Computer.FileSystem.OpenTextFileWriter(TestHomeDirectory & "\" & sPrgFile, False, Encoding.ASCII)
        ' Pass Data to Pvx, Initialize Vars
        FChn.WriteLine("20 RETURNFILENAME$ = """ & sReturnFile & """")
        FChn.WriteLine("30 PRIMARYMODULE$ = """ & Properties.PrimaryModule & """")
        FChn.WriteLine("40 ENHCODE$ = """ & Properties.CodeIdentifier & """")
        FChn.WriteLine("50 CTLFILE$ = """ & sCTLFile$ & """")

        ' Open/Lock Return File Channel
        FChn.WriteLine("210 ERASE RETURNFILENAME$,ERR=*NEXT")
        FChn.WriteLine("220 SERIAL RETURNFILENAME$")
        FChn.WriteLine("230 RCHN=UNT,E1=RCHN;OPEN LOCK(E1,ISZ=-1) RETURNFILENAME$")
        ' Open CTL File
        FChn.WriteLine("240 CTLCHN=UNT,E1=CTLCHN,E1$=CTLFILE$; OPEN (E1)E1$")
        ' Goto 1st record
        FChn.WriteLine("250 E1=CTLCHN,E2$=""l""; READ (E1,KEY=E2$,DOM=*NEXT)")
        ' Loop through "l" records
        FChn.WriteLine("310 TOP_CTL_LOOP:")
        FChn.WriteLine("320 E1=CTLCHN,E2$=KEY(E1,END=EXIT_CTL_LOOP); READ (E1,KEY=E2$)CTL$")
        FChn.WriteLine("330 IF CTL$(1,1)<>""l"" THEN GOTO EXIT_CTL_LOOP")
        ' Reject non-Enh records
        FChn.WriteLine("350 IF LEN(E2$)<5 THEN GOTO TOP_CTL_LOOP")
        FChn.WriteLine("360 EMODULE$=CTL$(2,3)")
        FChn.WriteLine("370 EVERSION$=CTL$(50,10)")
        FChn.WriteLine("380 IF EMODULE$=""SYS"" THEN EMODULE$=""L/M""")
        ' Write line of data for this module/version
        FChn.WriteLine("390 E1=RCHN;WRITE RECORD (E1) EMODULE$ + DIM(1) + STP(EVERSION$,2) + CHR(59)")
        FChn.WriteLine("400 GOTO TOP_CTL_LOOP")
        FChn.WriteLine("410 EXIT_CTL_LOOP:")

        ' Close Return Chn and adios
        FChn.WriteLine("500 E1=RCHN;CLOSE(E1)")
        FChn.WriteLine("510 E1=CTLCHN;CLOSE(E1)")
        FChn.WriteLine("64998 QUIT")
        FChn.Flush()
        FChn.Close()

        ' Run the Providex program and wait until it is done.
        Dim proc As New Process()
        proc.StartInfo.WorkingDirectory = TestHomeDirectory
        proc.StartInfo.FileName = TestHomeDirectory & "\PVXWIN32.EXE"
        proc.StartInfo.Arguments = "-HD ..\HOME\" & sPrgFile
        proc.StartInfo.WindowStyle = ProcessWindowStyle.Hidden
        proc.Start()
        proc.WaitForExit()

        sReadString = My.Computer.FileSystem.ReadAllText(sReturnFile)

        My.Computer.FileSystem.DeleteFile(sPrgFile)
        My.Computer.FileSystem.DeleteFile(sReturnFile)
        Return sReadString

    End Function

    Public Sub ChangeCTLData(ByVal sDate As String,
                         ByVal sPartNumber As String,
                         ByVal sDesc As String,
                         ByVal sVersion As String,
                         ByVal ProjectBaseDirectory As String,
                         ByVal TestHomeDirectory As String)
        ' sDate is a MAS90 packed 6-digit string
        ' Updates the "l" (L) records in the CTL file with the passed date
        ' and passed Part Number
        Dim sFilenamePrefix As String
        Dim sPrimaryModuleFolder As String
        Dim sPrgFile As String
        Dim FChn As StreamWriter
        Dim sFile As String
        Dim sReturnFile As String

        Dim Properties As New Properties(ProjectBaseDirectory & "\Properties.ini")

        ' Ensure that Part Number is 10-characters in length
        sPartNumber = sPartNumber.PadRight(10)
        '
        sDesc = sDesc.PadRight(31)
        '
        sVersion = Mid$(sVersion, 1, 5)

        ' If CTL file exists then get project date from
        ' main module "l" record
        If UCase$(Properties.PrimaryModule) = "LM" Then
            sFilenamePrefix = "IN"
            sPrimaryModuleFolder = "SOA"
        Else
            sFilenamePrefix = Properties.PrimaryModule
            sPrimaryModuleFolder = Properties.PrimaryModule
        End If
        sFile = ProjectBaseDirectory & "\" & sPrimaryModuleFolder _
                & "\" & sFilenamePrefix & "0CTL." + Properties.CodeIdentifier

        sPrgFile = "$Prg$"
        sReturnFile = TestHomeDirectory & "\$Return$.DAT"
        My.Computer.FileSystem.CurrentDirectory = TestHomeDirectory
        FChn = My.Computer.FileSystem.OpenTextFileWriter(TestHomeDirectory & "\" & sPrgFile, False, Encoding.ASCII)
        ' Pass Data to Pvx, Initialize Vars
        FChn.WriteLine("10 FILENAME$ = """ & sFile & """")
        FChn.WriteLine("30 NEWDATE$ = """ & sDate & """")
        FChn.WriteLine("40 PARTNUMBER$ = """ & sPartNumber & """")
        FChn.WriteLine("50 VERSION$ = """ & sVersion & """")
        FChn.WriteLine("60 DESC$ = """ & sDesc & """")
        ' Loop through "l" recs and change date for each one found.
        FChn.WriteLine("110 E1=UNT;OPEN (E1)FILENAME$")
        FChn.WriteLine("120 READ(E1,KEY=""l"",DOM=*NEXT)")
        FChn.WriteLine("130 TOP_LOOP: KEY$ = KEY(E1,END=EXIT_LOOP)")
        FChn.WriteLine("140 READ(E1,KEY=KEY$)DATA$")
        FChn.WriteLine("150 IF DATA$(1,1) <> ""l"" THEN GOTO EXIT_LOOP")
        FChn.WriteLine("160 DATA$(60,6) = NEWDATE$")
        FChn.WriteLine("170 DATA$(9,10) = PARTNUMBER$")
        FChn.WriteLine("180 DATA$(19,31) = DESC$")
        FChn.WriteLine("190 DATA$(50,5) = VERSION$")
        FChn.WriteLine("200 WRITE(E1,KEY=KEY$)DATA$;GOTO TOP_LOOP")
        FChn.WriteLine("210 EXIT_LOOP: CLOSE(E1)")
        FChn.WriteLine("64998 BYE")
        FChn.Flush()
        FChn.Close()
        FChn.Dispose()

        Dim proc As New Process()
        proc.StartInfo.WorkingDirectory = TestHomeDirectory
        proc.StartInfo.FileName = TestHomeDirectory & "\PVXWIN32.EXE"
        proc.StartInfo.Arguments = "-HD ..\HOME\" & sPrgFile
        proc.StartInfo.WindowStyle = ProcessWindowStyle.Hidden
        proc.Start()
        proc.WaitForExit()


        My.Computer.FileSystem.DeleteFile(sPrgFile)

    End Sub

    Public Sub CheckReleaseDate(ByVal TestHomeDirectory As String, ByVal ProjectBaseDirectory As String)
        Dim sDate, sDescription, sCTLDate, sRec As String

        Dim Properties As New Properties(ProjectBaseDirectory & "\Properties.ini")

        sRec = GetPrimaryModuleCTLRec(TestHomeDirectory, ProjectBaseDirectory)
        If Len(sRec) <> 0 Then
            sDate = Mid$(sRec, 60, 6)
            Dim sPartNumber As String
            sPartNumber = Mid$(sRec, 9, 10)
            sDescription = Mid$(sRec, 19, 31)
            sCTLDate = Mid$(sDate, 3, 2) & "/" &
                       Mid$(sDate, 5, 2) & "/" &
                       Unpack(Mid$(sDate, 1, 2))
            If DateValue(sCTLDate) <> DateValue(Properties.ReleaseDate) _
            Or Trim$(sDescription) <> Trim$(Properties.ProjectName) _
            Or Trim$(sPartNumber) <> Trim$(Properties.PartNumber) Then
                'Version 4 Sy_Enhancement Tag
                Dim Date40 As Date
                DateTime.TryParse(Properties.ReleaseDate.ToString(), Date40)
                Dim sDate40 As String = Date40.ToString("yyyyMMdd")
                If Val(Properties.Version) >= 4 Then
                    ChangeSY_EnhancementData(sDate40, Properties.PartNumber, Properties.ProjectName, Properties.Version, ProjectBaseDirectory, TestHomeDirectory)
                End If
                ' Build MAS90 6-digit packed date.
                sDate = Format(DateValue(Properties.ReleaseDate), "MM/dd/yyyy")
                sDate = Pack(Mid$(sDate, 7, 4)) & Mid$(sDate, 1, 2) & Mid$(sDate, 4, 2)
                ChangeCTLData(sDate, Properties.PartNumber, Properties.ProjectName, Properties.Version, ProjectBaseDirectory, TestHomeDirectory)
                'End If
            End If
        End If


    End Sub

    Public Function Unpack(ByVal Packed As String) As String
        ' Duplicates the function of FNUNPACK in MAS90
        ' Unpacks a 2-character MAS90 packed year
        Unpack = Trim$(Str$(
                 ((Asc(Mid$(Packed, 1, 1)) - 32) * 64) _
                 + (Asc(Mid$(Packed, 2, 1)) - 32)
                 ))
    End Function

    Public Function Pack(ByVal D As String) As String
        ' Duplicates the FNPACK function in MAS90
        ' Packs a 4-character date string to a 2-char packed string
        Pack = (Chr(Int(Val(D) / 64) + 32)) _
               & (Chr((Val(D) Mod 64) + 32))
    End Function

    Public Sub ChangeSY_EnhancementData(ByVal sDate As String,
                         ByVal sPartNumber As String,
                         ByVal sDesc As String,
                         ByVal sVersion As String,
                         ByVal ProjectBaseDirectory As String,
                         ByVal TestHomeDirectory As String)
        ' sDate is a yyyymmdd
        ' Updates the records in the SY_Enhancement file with the passed date
        ' and passed Part Number
        Dim sFilenamePrefix As String
        Dim sPrimaryModuleFolder As String
        Dim sPrgFile As String
        Dim FChn As StreamWriter
        Dim sFile As String
        Dim sModule As String
        Dim sDeveloperCode As String

        Dim Properties As New Properties(ProjectBaseDirectory & "\Properties.ini")

        ' Ensure that Part Number is 10-characters in length
        sPartNumber = sPartNumber.PadRight(10)
        sPartNumber = Mid$(sPartNumber, 1, 10)

        '
        sDesc = sDesc.PadRight(31)
        sDesc = Mid$(sDesc, 1, 31)
        '
        sVersion = Mid$(sVersion, 1, 5)

        ' If SY_Enhancement file exists then get project date from
        ' main module record
        If UCase$(Properties.PrimaryModule) = "LM" Then
            sFilenamePrefix = "IN"
            sPrimaryModuleFolder = "SOA"
        Else
            sFilenamePrefix = Properties.PrimaryModule
            sPrimaryModuleFolder = Properties.PrimaryModule
        End If
        sFile = ProjectBaseDirectory & "\" & sPrimaryModuleFolder _
                & "\" & sFilenamePrefix & "_Enhancement." + Properties.CodeIdentifier

        sDeveloperCode = "234"
        sModule = Properties.PrimaryModule.Substring(0, 1) & "/" & Properties.PrimaryModule.Substring(1, 1)

        sPrgFile = "$Prg$"
        My.Computer.FileSystem.CurrentDirectory = TestHomeDirectory
        FChn = My.Computer.FileSystem.OpenTextFileWriter(TestHomeDirectory & "\" & sPrgFile, False, Encoding.ASCII)
        ' Pass Data to Pvx, Initialize Vars
        FChn.WriteLine("10 FILENAME$ = """ & sFile & """")
        FChn.WriteLine("30 NEWDATE$ = """ & sDate & """")
        FChn.WriteLine("40 PARTNUMBER$ = """ & sPartNumber & """")
        FChn.WriteLine("50 VERSION$ = """ & sVersion & """")
        FChn.WriteLine("60 DESC$ = """ & sDesc & """")
        FChn.WriteLine("70 MODULECODE$ = """ & sModule & """")
        FChn.WriteLine("80 DEVELOPERCODE$ = """ & sDeveloperCode & """")
        ' Loop through recs and change date for each one found.
        FChn.WriteLine("110 E1=UNT;OPEN (E1,IOL=*)FILENAME$")
        FChn.WriteLine("115 E2=UNT;OPEN (E2,IOL=*)FILENAME$")
        FChn.WriteLine("120 SELECT * FROM E1 BEGIN $$ END $FE$")
        'FChn.WriteLine( "150 REMOVE (E1)"
        FChn.WriteLine("160 RELEASEDATE$ = NEWDATE$")
        'FChn.WriteLine( "170 ENHANCEMENTCODE$ = PARTNUMBER$"
        FChn.WriteLine("180 ENHANCEMENTNAME$ = DESC$")
        FChn.WriteLine("190 ENHANCEMENTLEVEL = NUM(VERSION$)")
        FChn.WriteLine("200 WRITE(E2)")
        FChn.WriteLine("210 NEXT RECORD")
        FChn.WriteLine("220 CLOSE(E1)")
        FChn.WriteLine("220 CLOSE(E2)")
        FChn.WriteLine("64998 BYE")
        FChn.Flush()
        FChn.Close()
        FChn.Dispose()

        Dim proc As New Process()
        proc.StartInfo.WorkingDirectory = TestHomeDirectory
        proc.StartInfo.FileName = TestHomeDirectory & "\PVXWIN32.EXE"
        proc.StartInfo.Arguments = "-HD ..\HOME\" & sPrgFile
        proc.StartInfo.WindowStyle = ProcessWindowStyle.Hidden
        proc.Start()
        proc.WaitForExit()

        My.Computer.FileSystem.DeleteFile(sPrgFile)

    End Sub

    Public Function GetPrimaryModuleCTLRec(ByVal TestHomeDirectory As String, ByVal ProjectBaseDirectory As String) As String
        ' Retreives the project's primary module's "l" (L) record date
        ' in MAS90-type packed format.
        Dim sKey As String, sData As String, sFilename As String, sKeyPrefix As String

        Dim Properties As New Properties(ProjectBaseDirectory & "\Properties.ini")

        If UCase$(Properties.PrimaryModule) = "LM" Then
            sKeyPrefix = "SYS"
        Else
            sKeyPrefix = Mid$(Properties.PrimaryModule, 1, 1) & "/" _
                         & Mid$(Properties.PrimaryModule, 2, 1)
        End If
        sKey = "l" & sKeyPrefix & "_" & Properties.CodeIdentifier
        sFilename = CTLFilename(ProjectBaseDirectory)
        If sFilename <> "" Then
            sData = ReadString(sFilename, sKey, ProjectBaseDirectory, TestHomeDirectory)
            If sData <> "" _
            And sData <> "ErrOpening" _
            And sData <> "ErrKey" _
            And sData <> "ErrReading" Then
                Return sData
            Else
                Return ""
                log.Error("Warning: A correct l (lowercase-L) record does not exist in the project's CTL file!")
            End If
        Else
            Return ""
        End If
    End Function

    Public Function ReadString(ByVal FileName As String,
                         ByVal Key As String,
                         ByVal ProjectBaseDirectory As String,
                         ByVal TestHomeDirectory As String) As String
        ' Make sure file exists before calling.
        ' Returns "" if key not found.
        ' If error while reading, returns "ErrReading"
        ' If open error, returns "ErrOpening"
        ' If key can't be found, returns "ErrKey"

        Dim Properties As New Properties(ProjectBaseDirectory & "\Properties.ini")

        Dim sPrgFile As String
        Dim sReturnFile As String
        Dim FChn As StreamWriter

        If Not My.Computer.FileSystem.FileExists(FileName) Then
            Return "ErrOpening"
            Exit Function
        End If

        sPrgFile = "$Prg$"
        sReturnFile = TestHomeDirectory & "\$Return$.DAT"
        My.Computer.FileSystem.CurrentDirectory = TestHomeDirectory

        Try
            FChn = My.Computer.FileSystem.OpenTextFileWriter(TestHomeDirectory & "\" & sPrgFile, False, Encoding.ASCII)
            ' Pass Data to Pvx, Initialize Vars
            FChn.WriteLine("10 FILENAME$ = """ & FileName & """")
            FChn.WriteLine("20 RETURNFILENAME$ = """ & sReturnFile & """")
            FChn.WriteLine("30 KEY$ = """ & Key & """")
            FChn.WriteLine("40 DATA$ = DIM(0)")
            FChn.WriteLine("50 SETERR GENERROR")
            ' Grab data from MAS90
            FChn.WriteLine("110 E1=UNT;OPEN (E1,ERR=OPENERROR)FILENAME$")
            FChn.WriteLine("120 READ(E1,KEY=KEY$,DOM=NOTFOUND,ERR=READERROR)DATA$")
            FChn.WriteLine("130 CLOSE(E1)")
            FChn.WriteLine("140 GOTO WRITEDATA")
            ' Key Not Found
            FChn.WriteLine("200 NOTFOUND:")
            FChn.WriteLine("210 DATA$ = ""ErrKey""")
            FChn.WriteLine("220 GOTO WRITEDATA")
            ' Read Error
            FChn.WriteLine("300 READERROR:")
            FChn.WriteLine("310 DATA$ = ""ErrReading""")
            FChn.WriteLine("320 GOTO WRITEDATA")
            ' Error Opening
            FChn.WriteLine("400 OPENERROR:")
            FChn.WriteLine("410 DATA$ = ""ErrOpening""")
            FChn.WriteLine("420 GOTO WRITEDATA")
            ' Return Data to IDE
            FChn.WriteLine("500 WRITEDATA:")
            FChn.WriteLine("505 CLOSE(E1,ERR=*NEXT)")
            FChn.WriteLine("510 ERASE RETURNFILENAME$,ERR=*NEXT")
            FChn.WriteLine("520 SERIAL RETURNFILENAME$")
            FChn.WriteLine("530 OPEN LOCK(E1,ISZ=-1) RETURNFILENAME$")
            FChn.WriteLine("540 WRITE RECORD (E1)DATA$")
            FChn.WriteLine("600 CLOSE(E1)")
            FChn.WriteLine("64000 GENERROR:")
            FChn.WriteLine("64998 BYE")
            FChn.Flush()
            FChn.Close()
            FChn.Dispose()

            Dim proc As New Process()
            proc.StartInfo.WorkingDirectory = TestHomeDirectory
            proc.StartInfo.FileName = TestHomeDirectory & "\PVXWIN32.EXE"
            proc.StartInfo.Arguments = "-HD ..\HOME\" & sPrgFile
            proc.StartInfo.WindowStyle = ProcessWindowStyle.Hidden
            proc.Start()
            proc.WaitForExit()

            Dim sReadString As String = My.Computer.FileSystem.ReadAllText(sReturnFile)

            My.Computer.FileSystem.DeleteFile(sPrgFile)
            My.Computer.FileSystem.DeleteFile(sReturnFile)

            Return sReadString

            Exit Function
        Catch
            Return "ErrNotFound"
        End Try

    End Function

End Class