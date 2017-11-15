
Imports MultithreadedControl.IniFile

Public Class Properties

    Public Shared Ini As New IniFile()

    Public Sub New(ByVal ProjectIni As String)
        Ini.Load(ProjectIni)
    End Sub

    Public Sub Save(ByVal ProjectIni As String)
        Ini.Save(ProjectIni)
    End Sub

    Public Shared Property IsNonTest As Boolean
        Get
            Try
                Return Convert.ToBoolean(Ini.GetSection("Main").GetKey("IsNonTest").GetValue().ToString())
            Catch
                Return False
            End Try

        End Get
        Set(value As Boolean)
            Ini.GetSection("Main").GetKey("IsNonTest").SetValue(Convert.ToString(value))
        End Set
    End Property

    Public Shared Property PasswordApplyTo As String
        Set(value As String)
            Ini.GetSection("Installation").GetKey("PasswordApplyTo").SetValue(value)
        End Set
        Get
            Try
                Return Ini.GetSection("Installation").GetKey("PasswordApplyTo").GetValue()
            Catch
                Return ""
            End Try

        End Get
    End Property

    Public Shared Property IsPatchInstall As Boolean
        Set(value As Boolean)
            Ini.GetSection("Main").GetKey("IsPatchInstall").SetValue(Convert.ToString(value))
        End Set
        Get
            Try
                Return Convert.ToBoolean(Ini.GetSection("Main").GetKey("IsPatchInstall").GetValue())
            Catch
                Return False
            End Try

        End Get
    End Property

    Public Shared Property LastControlSystemFolder As String
        Set(value As String)
            Ini.GetSection(Environment.UserName).GetKey("LastControlSystemFolder").SetValue(value)
        End Set
        Get
            Try
                Return Ini.GetSection(Environment.UserName).GetKey("LastControlSystemFolder").GetValue()
            Catch
                Return ""
            End Try

        End Get
    End Property

    Public Shared Property LastTestSystemFolder As String
        Set(value As String)
            Ini.GetSection(Environment.UserName).GetKey("LastTestSystemFolder").SetValue(value)
        End Set
        Get
            Try
                Return Ini.GetSection(Environment.UserName).GetKey("LastTestSystemFolder").GetValue()
            Catch
                Return ""
            End Try

        End Get
    End Property
    ' Additions for Security Headaches 07/09/04 
    Public Shared Property PVXLoginName As String
        Set(value As String)
            Ini.GetSection(Environment.UserName).GetKey("PVXLoginName").SetValue(value)
        End Set
        Get
            Try
                Return Ini.GetSection(Environment.UserName).GetKey("PVXLoginName").GetValue()
            Catch
                Return ""
            End Try

        End Get
    End Property
    Public Shared Property PVXLoginPass As String
        Set(value As String)
            Ini.GetSection(Environment.UserName).GetKey("PVXLoginPass").SetValue(value)

        End Set
        Get
            Try
                Return Ini.GetSection(Environment.UserName).GetKey("PVXLoginPass").GetValue()
            Catch
                Return ""
            End Try

        End Get
    End Property
    Public Shared Property PVXLoginCompany As String
        Set(value As String)
            Ini.GetSection(Environment.UserName).GetKey("PVXLoginCompany").SetValue(value)
        End Set
        Get
            Try
                Return Ini.GetSection(Environment.UserName).GetKey("PVXLoginCompany").GetValue()
            Catch
                Return ""
            End Try

        End Get
    End Property
    ' End Additions 07/09/04

    Public Shared Property AddlFilesFolder As String
        Set(value As String)
            Ini.GetSection("Installation").GetKey("AddlFilesFolder").SetValue(value)

        End Set
        Get
            Try
                Return Ini.GetSection("Installation").GetKey("AddlFilesFolder").GetValue()
            Catch
                Return ""
            End Try

        End Get
    End Property

    Public Shared Property CodeDescription As String
        Set(value As String)
            Ini.GetSection("Main").GetKey("CodeDescription").SetValue(value)
        End Set
        Get
            Try
                Return Ini.GetSection("Main").GetKey("CodeDescription").GetValue()
            Catch ex As Exception
                Return ""
            End Try
        End Get
    End Property


    Public Shared Property PasswordSeedString As String
        Set(value As String)
            Ini.GetSection("Main").GetKey("PasswordSeedString").SetValue(value)

        End Set
        Get
            Try
                Return Ini.GetSection("Main").GetKey("PasswordSeedString").GetValue()
            Catch
                Return ""
            End Try

        End Get
    End Property


    Public Shared Property IsSystem As Boolean
        Set(value As Boolean)
            Ini.GetSection("Main").GetKey("IsSystem").SetValue(Convert.ToString(value))
            Ini.GetSection("Main").GetKey("IsCustom").SetValue(Convert.ToString(Not value))
        End Set
        Get
            Try
                Return Convert.ToBoolean(Ini.GetSection("Main").GetKey("IsEnhancement").GetValue())
            Catch
                Return False
            End Try

        End Get
    End Property


    Public Shared Property IsCustom As Boolean
        Set(value As Boolean)
            Ini.GetSection("Main").GetKey("IsCustom").SetValue(Convert.ToString(value))
            Ini.GetSection("Main").GetKey("IsSystem").SetValue(Convert.ToString(Not value))

        End Set
        Get
            Try
                Return Convert.ToBoolean(Ini.GetSection("Main").GetKey("IsCustom").GetValue())
            Catch
                Return False
            End Try

        End Get
    End Property


    Public Shared Property ZipExeFilename As String
        Set(value As String)
            Ini.GetSection("Installation").GetKey("ZipExeFilename").SetValue(value)
        End Set
        Get
            Try
                Return Ini.GetSection("Installation").GetKey("ZipExeFilename").GetValue()
            Catch
                Return ""
            End Try

        End Get
    End Property


    Public Shared Property ZipExeOutputFolder As String
        Set(value As String)
            Ini.GetSection("Installation").GetKey("ZipExeOutputFolder").SetValue(value)
        End Set
        Get
            Try
                Return Ini.GetSection("Installation").GetKey("ZipExeOutputFolder").GetValue()
            Catch
                Return ""
            End Try

        End Get
    End Property


    Public Shared Property PostEXECommand As String
        Set(value As String)
            Ini.GetSection("Installation").GetKey("PostEXECommand").SetValue(value)

        End Set
        Get
            Try
                Return Ini.GetSection("Installation").GetKey("PostEXECommand").GetValue()
            Catch
                Return ""
            End Try

        End Get
    End Property


    Public Shared Property OutputFolder As String
        Set(value As String)
            Ini.GetSection("Installation").GetKey("OutputFolder").SetValue(value)

        End Set
        Get
            Try
                Return Ini.GetSection("Installation").GetKey("OutputFolder").GetValue()
            Catch
                Return ""
            End Try
        End Get
    End Property


    Public Shared Property Password As String
        Set(value As String)
            Ini.GetSection("Installation").GetKey("Password").SetValue(value)

        End Set
        Get
            Try
                Return Ini.GetSection("Installation").GetKey("Password").GetValue()
            Catch
                Return ""
            End Try

        End Get
    End Property


    Public Shared Property AllModules As String
        Set(value As String)
            Ini.GetSection("Main").GetKey("AllModules").SetValue(value)

        End Set
        Get
            Try
                Return Ini.GetSection("Main").GetKey("AllModules").GetValue()
            Catch
                Return ""
            End Try


        End Get
    End Property


    Public Shared Property PrimaryModule As String
        Set(value As String)
            Ini.GetSection("Main").GetKey("PrimaryModule").SetValue(value)

        End Set
        Get
            Try
                Return Ini.GetSection("Main").GetKey("PrimaryModule").GetValue()
            Catch
                Return ""
            End Try

        End Get
    End Property

    Public Shared Property PartNumber As String
        Set(value As String)
            Ini.GetSection("Main").GetKey("PartNumber").SetValue(value)

        End Set
        Get
            Try
                Return Ini.GetSection("Main").GetKey("PartNumber").GetValue()
            Catch
                Return ""
            End Try

        End Get
    End Property


    Public Shared Property ProjectNumber As String
        Set(value As String)
            Ini.GetSection("Main").GetKey("ProjectNumber").SetValue(value)

        End Set
        Get
            Try
                Return Ini.GetSection("Main").GetKey("ProjectNumber").GetValue()
            Catch
                Return ""
            End Try


        End Get
    End Property


    Public Shared Property ReleaseDate As String
        Set(value As String)
            Ini.GetSection("Main").GetKey("ReleaseDate").SetValue(value)

        End Set
        Get
            Try
                Return Ini.GetSection("Main").GetKey("ReleaseDate").GetValue()
            Catch
                Return ""
            End Try

        End Get
    End Property


    Public Shared Property UpdateAvailable As String
        Set(value As String)
            Ini.GetSection("Main").GetKey("UpdateAvailable").SetValue(value)

        End Set
        Get
            Try
                Return Ini.GetSection("Main").GetKey("UpdateAvailable").GetValue()
            Catch
                Return ""
            End Try

        End Get
    End Property


    Public Shared Property DoNotIncludeCatalog As String
        Set(value As String)
            Ini.GetSection("Main").GetKey("DoNotIncludeCatalog").SetValue(value)

        End Set
        Get
            Try
                Return Ini.GetSection("Main").GetKey("DoNotIncludeCatalog").GetValue()
            Catch
                Return ""
            End Try

        End Get
    End Property


    Public Shared Property Version As String
        Set(value As String)
            Ini.GetSection("Main").GetKey("Version").SetValue(value)

        End Set
        Get
            Try
                Return Ini.GetSection("Main").GetKey("Version").GetValue()
            Catch
                Return ""
            End Try

        End Get
    End Property


    Public Shared Property CodeIdentifier As String
        Set(value As String)
            Ini.GetSection("Main").GetKey("CodeIdentifier").SetValue(value)

        End Set
        Get
            Try
                Return Ini.GetSection("Main").GetKey("CodeIdentifier").GetValue()
            Catch
                Return ""
            End Try

        End Get
    End Property


    Public Shared Property IncludeMainUtilityMenu As String
        Set(value As String)
            Ini.GetSection("Main").GetKey("IncludeMainUtilityMenu").SetValue(value)

        End Set
        Get
            Try
                Return Ini.GetSection("Main").GetKey("IncludeMainUtilityMenu").GetValue()
            Catch
                Return ""
            End Try

        End Get
    End Property


    Public Shared Property IncludePDFinSEZip As String
        Set(value As String)
            Ini.GetSection("Main").GetKey("IncludePDFinSEZip").SetValue(value)

        End Set
        Get
            Try
                Return Ini.GetSection("Main").GetKey("IncludePDFinSEZip").GetValue()
            Catch
                Return ""
            End Try

        End Get
    End Property

    Public Shared Property ProjectName As String
        Set(value As String)
            If value.Length > 35 Then value = value.Substring(1, 35)
            Ini.GetSection("Main").GetKey("ProjectName").SetValue(value)

        End Set
        Get
            Try
                Return Ini.GetSection("Main").GetKey("ProjectName").GetValue()
            Catch
                Return ""
            End Try

        End Get
    End Property


    Public Shared Property AddlModule As String
        Set(value As String)
            Ini.GetSection("Main").GetKey("AddlModule").SetValue(value)

        End Set
        Get
            Try
                Return Ini.GetSection("Main").GetKey("AddlModule").GetValue()
            Catch
                Return ""
            End Try

        End Get
    End Property

    Public Shared Property DistributionVersion As String
        Set(value As String)
            Ini.GetSection("Main").GetKey("DistributionVersion").SetValue(value)

        End Set
        Get
            Try
                Return Ini.GetSection("Main").GetKey("DistributionVersion").GetValue()
            Catch
                Return ""
            End Try

        End Get
    End Property

    Public Shared Property OriginDate As String
        Set(value As String)
            Ini.GetSection("Main").GetKey("OriginDate").SetValue(value)

        End Set
        Get
            Try
                Return Ini.GetSection("Main").GetKey("OriginDate").GetValue()
            Catch
                Return ""
            End Try

        End Get
    End Property

End Class
