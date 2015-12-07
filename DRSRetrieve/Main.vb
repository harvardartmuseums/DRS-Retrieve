'#
'# The input file may contain a list of Rendition Numbers or Object Numbers
'# There can only be one number per line
'# The order of command line parameters is critical
'#

Imports System.Data.Odbc
Imports System.Data.SqlClient
Imports System.Text.RegularExpressions
Imports System.Collections

Module Main
    Const FILENAMEFORMAT_1 As String = "{0}-{1}"
    Const FILENAMEFORMAT_2 As String = "{1}"
    Const FILEEXTENSION_TIFF As String = ".tif"
    Const FILEEXTENSION_JPEG As String = ".jpg"

    Public FILENAMEFORMAT As String
    Public DBCONNECTION As String
    Public SERVER As String
    Public OUTPUTPATH As String
    Public SOURCEFILE As String
    Public LOOKUPMODULE As String
    Public FILESIZE As String
    Public GETALLFILES As Boolean = False

    Public Structure rendition
        Public renditionNumber As String
        Public primaryDisplay As Integer
    End Structure

    Sub Main()
        ' Parse the command line parameters
        If My.Application.CommandLineArgs.Count < 2 Then
            Console.WriteLine("Error: Invalid number of parameters")
            Exit Sub
        End If

        ReadParameters()

        ' Create the full connection string
        'DBCONNECTION = "Data Source=" & SERVER & ";database=TMS;Integrated Security=SSPI"
        DBCONNECTION = "DSN=" & SERVER

        ' Test for the existance of the source file
        If FileIO.FileSystem.FileExists(SOURCEFILE) = False Then
            Console.WriteLine("Error: unable to find the source file")
            Exit Sub
        End If

        ' Create the destination folder
        Dim fileInfo As System.IO.FileInfo
        fileInfo = My.Computer.FileSystem.GetFileInfo(SOURCEFILE)

        OUTPUTPATH = fileInfo.FullName.Replace(fileInfo.Extension, String.Empty) & "\"

        If FileIO.FileSystem.DirectoryExists(OUTPUTPATH) = False Then
            FileIO.FileSystem.CreateDirectory(OUTPUTPATH)
        Else
            Console.WriteLine("Error: Unable to create the output folder. It already exists.")
            Exit Sub
        End If

        Console.WriteLine("About to retrieve some files from the DRS...")

        ' Start reading the input file and retrieve the images
        Dim inputFile As System.IO.StreamReader
        inputFile = New System.IO.StreamReader(SOURCEFILE)
        Do While Not inputFile.EndOfStream
            Dim lookupNumber As String = inputFile.ReadLine.TrimEnd

            Console.WriteLine("Working on " & lookupNumber)

            Dim renditions As New List(Of rendition)

            Select Case LOOKUPMODULE
                Case "objects"
                    renditions = LookupRenditions(lookupNumber)
                Case "media"
                    Dim r As rendition
                    r.renditionNumber = lookupNumber
                    r.primaryDisplay = 1
                    renditions.Add(r)
                Case Else
                    'default to the Objects module if no lookup module switch is supplied
                    renditions = LookupRenditions(lookupNumber)
            End Select

            For Each r As rendition In renditions
                If String.IsNullOrEmpty(r.renditionNumber) = False Then
                    If GETALLFILES Or (Not GETALLFILES And r.primaryDisplay = 1) Then
                        Dim fileName As String = LookUpFileName(r.renditionNumber)
                        If String.IsNullOrEmpty(fileName) = False Then
                            fileName = TranslateFileSize(fileName)

                            'Get IDS Response
                            Dim IDSResponse As System.Net.HttpWebResponse = LookupIDS2(fileName)
                            Dim IDS As String = IDSResponse.ResponseUri.AbsoluteUri

                            'Determine reponse type (jpeg/tiff/unknown)
                            'Construct file name
                            Dim outputFilename As String = OUTPUTPATH & _
                                        String.Format(FILENAMEFORMAT, lookupNumber.Trim, r.renditionNumber.Trim)
                            Select Case IDSResponse.ContentType
                                Case "image/jpeg"
                                    outputFilename &= FILEEXTENSION_JPEG
                                Case "image/tiff"
                                    outputFilename &= FILEEXTENSION_TIFF
                                Case Else
                                    outputFilename = String.Empty
                            End Select

                            'Save file
                            If String.IsNullOrEmpty(outputFilename) = False Then
                                SaveFile(IDS, outputFilename)
                            End If

                            'Log info to output file
                            Dim out As System.IO.StreamWriter = New System.IO.StreamWriter(OUTPUTPATH & "out.txt", True)
                            out.WriteLine(String.Format("{0}, {1}, {2}, {3}, {4}, {5}, {6}", _
                                lookupNumber.Trim, _
                                r.renditionNumber.Trim, _
                                r.primaryDisplay, _
                                outputFilename, _
                                IDSResponse.ContentType.ToString, _
                                IDSResponse.ResponseUri.AbsoluteUri.ToString, _
                                IDSResponse.StatusCode.ToString))
                            out.Close()
                        End If
                    End If
                End If
            Next

        Loop
    End Sub

    Private Sub SaveFile(ByVal URN As String, ByVal OutputFileName As String)
        Try
            My.Computer.Network.DownloadFile(URN, OutputFileName)
        Catch ex As Exception
            'Log the error somewhere
        End Try
    End Sub

    Private Function LookupIDS2(ByVal URN As String) As System.Net.HttpWebResponse
        Dim http As System.Net.HttpWebRequest = System.Net.WebRequest.Create(URN)
        Dim response As System.Net.HttpWebResponse

        Try
            response = CType(http.GetResponse(), System.Net.HttpWebResponse)
        Catch ex As Net.WebException
            response = CType(ex.Response, System.Net.HttpWebResponse)
        End Try

        response.Close()

        Return response
    End Function

    Private Function LookupIDS(ByVal URN As String) As String
        Dim out As System.IO.StreamWriter = New System.IO.StreamWriter(OUTPUTPATH & "out.txt", True)
        Dim http As System.Net.WebRequest = System.Net.WebRequest.Create(URN)
        Try
            Dim response As System.Net.HttpWebResponse = CType(http.GetResponse(), System.Net.HttpWebResponse)

            out.WriteLine(String.Format("{0}, {1}, {2}", _
                        response.ContentType.ToString, _
                        response.ResponseUri.AbsoluteUri.ToString, _
                        response.StatusCode.ToString))

            Dim URI As String = response.ResponseUri.AbsoluteUri.ToString
            response.Close()

            Return URI

        Catch ex As Net.WebException
            out.WriteLine(ex.Message)
            http.Abort()
            Return String.Empty

        Catch ex As Exception
            http.Abort()
            Return String.Empty

        Finally
            out.Close()
        End Try

    End Function

    Private Function LookUpFileName(ByVal RenditionNumber As String) As String
        'Dim cn As SqlConnection = New SqlConnection(DBCONNECTION)
        Dim cn As OdbcConnection = New OdbcConnection(DBCONNECTION)
        cn.Open()

        'Dim sql As SqlCommand = New SqlCommand( _
        '                "SELECT MP.Path + '/' + MF.FileName " & _
        '                "FROM MediaRenditions MR " & _
        '                "INNER JOIN MediaMaster MM ON MR.MediaMasterID = MM.MediaMasterID " & _
        '                "INNER JOIN MediaRenditions DMR ON MM.DisplayRendID = DMR.RenditionID " & _
        '                "INNER JOIN MediaFiles MF ON DMR.PrimaryFileID = MF.FileID " & _
        '                "INNER JOIN MediaPaths MP ON MF.PathID = MP.PathID " & _
        '                "WHERE MR.RenditionNumber = @RenditionNumber", _
        '                cn)
        'sql.Parameters.AddWithValue("@RenditionNumber", RenditionNumber)

        Dim sql As OdbcCommand = New OdbcCommand( _
                        "SELECT MP.Path + '/' + MF.FileName " & _
                        "FROM MediaRenditions MR " & _
                        "INNER JOIN MediaMaster MM ON MR.MediaMasterID = MM.MediaMasterID " & _
                        "INNER JOIN MediaRenditions DMR ON MM.DisplayRendID = DMR.RenditionID " & _
                        "INNER JOIN MediaFiles MF ON DMR.PrimaryFileID = MF.FileID " & _
                        "INNER JOIN MediaPaths MP ON MF.PathID = MP.PathID " & _
                        "WHERE MR.RenditionNumber = ?", _
                        cn)
        sql.Parameters.AddWithValue("?", RenditionNumber)

        Dim output As String = CType(sql.ExecuteScalar, String)

        cn.Close()

        Return output
    End Function

    Private Function LookupPrimaryRendition(ByVal ObjectNumber As String) As String
        'Dim cn As SqlConnection = New SqlConnection(DBCONNECTION)
        Dim cn As OdbcConnection = New OdbcConnection(DBCONNECTION)
        cn.Open()

        'Dim sql As SqlCommand = New SqlCommand( _
        '            "SELECT PMR.RenditionNumber " & _
        '            "FROM Objects O " & _
        '            "INNER JOIN MediaXrefs MX ON O.ObjectID = MX.ID AND MX.TableID = 108 AND MX.PrimaryDisplay = 1 " & _
        '            "INNER JOIN MediaMaster MM ON MX.MediaMasterID = MM.MediaMasterID " & _
        '            "INNER JOIN MediaRenditions PMR ON MM.PrimaryRendID = PMR.RenditionID " & _
        '            "WHERE O.ObjectNumber = @ObjectNumber", _
        '            cn)
        'sql.Parameters.AddWithValue("@ObjectNumber", ObjectNumber)

        Dim sql As OdbcCommand = New OdbcCommand( _
            "SELECT PMR.RenditionNumber " & _
            "FROM Objects O " & _
            "INNER JOIN MediaXrefs MX ON O.ObjectID = MX.ID AND MX.TableID = 108 AND MX.PrimaryDisplay = 1 " & _
            "INNER JOIN MediaMaster MM ON MX.MediaMasterID = MM.MediaMasterID " & _
            "INNER JOIN MediaRenditions PMR ON MM.PrimaryRendID = PMR.RenditionID " & _
            "WHERE O.ObjectNumber = ?", _
            cn)
        sql.Parameters.AddWithValue("?", ObjectNumber)

        Dim output As String = CType(sql.ExecuteScalar, String)
        cn.Close()

        Return output
    End Function

    Private Function LookupRenditions(ByVal ObjectNumber As String) As List(Of rendition)
        'Dim cn As SqlConnection = New SqlConnection(DBCONNECTION)
        Dim cn As OdbcConnection = New OdbcConnection(DBCONNECTION)
        Dim reader As OdbcDataReader

        cn.Open()

        'Dim sql As SqlCommand = New SqlCommand( _
        '            "SELECT PMR.RenditionNumber " & _
        '            "FROM Objects O " & _
        '            "INNER JOIN MediaXrefs MX ON O.ObjectID = MX.ID AND MX.TableID = 108 AND MX.PrimaryDisplay = 1 " & _
        '            "INNER JOIN MediaMaster MM ON MX.MediaMasterID = MM.MediaMasterID " & _
        '            "INNER JOIN MediaRenditions PMR ON MM.PrimaryRendID = PMR.RenditionID " & _
        '            "WHERE O.ObjectNumber = @ObjectNumber", _
        '            cn)
        'sql.Parameters.AddWithValue("@ObjectNumber", ObjectNumber)

        Dim sql As OdbcCommand = New OdbcCommand( _
            "SELECT PMR.RenditionNumber, MX.PrimaryDisplay " & _
            "FROM Objects O " & _
            "INNER JOIN MediaXrefs MX ON O.ObjectID = MX.ID AND MX.TableID = 108 " & _
            "INNER JOIN MediaMaster MM ON MX.MediaMasterID = MM.MediaMasterID " & _
            "INNER JOIN MediaRenditions PMR ON MM.PrimaryRendID = PMR.RenditionID " & _
            "WHERE O.ObjectNumber = ?", _
            cn)
        sql.Parameters.AddWithValue("?", ObjectNumber)

        Dim output As New List(Of rendition)
        reader = sql.ExecuteReader()
        Do While reader.Read
            Dim r As rendition
            r.renditionNumber = reader("RenditionNumber").ToString
            r.primaryDisplay = reader("PrimaryDisplay")
            output.Add(r)
        Loop

        reader.Close()
        cn.Close()

        Return output
    End Function

    Private Function TranslateFileSize(ByVal Filename As String) As String
        If FILESIZE = "PRD" Then
            If Regex.IsMatch(Filename, "_dynmc*", RegexOptions.IgnoreCase) Then
                Return Regex.Replace(Filename, "_dynmc*", "_prdwork")
            ElseIf Regex.IsMatch(Filename, "_lgdl*", RegexOptions.IgnoreCase) Then
                Return Regex.Replace(Filename, "_lgdl*", "_prdwork")
            Else
                Return Filename 'String.Empty
            End If
        Else
            Return Filename
        End If
    End Function

    Private Sub ReadParameters()
        SOURCEFILE = My.Application.CommandLineArgs(0)
        SERVER = My.Application.CommandLineArgs(1).Split(":")(1)

        If My.Application.CommandLineArgs.Count > 2 Then
            LOOKUPMODULE = My.Application.CommandLineArgs(2).Split(":")(1)
        End If
        If My.Application.CommandLineArgs.Count > 3 Then
            FILESIZE = My.Application.CommandLineArgs(3).Split(":")(1)
        End If
        If My.Application.CommandLineArgs.Count > 4 Then
            FILENAMEFORMAT = IIf(My.Application.CommandLineArgs(4).Split(":")(1) = "2", FILENAMEFORMAT_2, FILENAMEFORMAT_1)
        End If
        If My.Application.CommandLineArgs.Count > 5 Then
            GETALLFILES = True
        End If
    End Sub

    Private Sub ShowUsage()
        Console.WriteLine("Retrieves a set of files from the DRS repository.")
        Console.WriteLine("")
        Console.WriteLine("DRSRetrieve filename [/S:server] [/M:module] [/F:filesize] [/N:filenameformat] [/A]")
        Console.WriteLine("")
        Console.WriteLine("filename" & Space(6) & "Specifies the input file containing items to retrieve.")
        Console.WriteLine("/S" & Space(10) & "Specifies the ODBC connection to the database server.")
        Console.WriteLine("/M" & Space(10) & "Specifies the module to search.")
        Console.WriteLine("/F" & Space(10) & "Specifies the filesize to retrieve.")
        Console.WriteLine("/N" & Space(10) & "Specifies the format of the filename.")
        Console.WriteLine("/A" & Space(10) & "Retrieves all files for each item in the input file.")
        Console.WriteLine("")
        Console.WriteLine("")
        Console.WriteLine("Example:")
        Console.WriteLine("DRSRetrieve MyObjectList.txt /s:Museum /m:Objects /f:full /n:2")
    End Sub
End Module
