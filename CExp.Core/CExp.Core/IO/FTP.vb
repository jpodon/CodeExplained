Imports System.Net
Imports System.IO

Namespace IO

    Public Class FTP
        ''' <summary>
        ''' Shared method which will return a list of files from a
        ''' FTP Location
        ''' </summary>
        ''' <param name="ftpAddress"></param>
        ''' <param name="ftpUser"></param>
        ''' <param name="ftpPassword"></param>
        ''' <param name="ExceptionInfo"></param>
        ''' <returns>List(of string)</returns>
        ''' <remarks>ExceptionInfo will return any exception messages should a failure occur</remarks>
        Public Shared Function ListRemoteFiles(ftpAddress As String, ftpUser As String, ftpPassword As String, ExceptionInfo As Exception) As List(Of String)

            Dim ListOfFilesOnFTPSite As New List(Of String)

            Dim ftpRequest As FtpWebRequest = Nothing
            Dim ftpResponse As FtpWebResponse = Nothing

            Dim strReader As StreamReader = Nothing
            Dim sline As String = ""

            Try
                ftpRequest = CType(WebRequest.Create(ftpAddress), FtpWebRequest)

                With ftpRequest
                    .Credentials = New NetworkCredential(ftpUser, ftpPassword)
                    .Method = WebRequestMethods.Ftp.ListDirectory
                End With

                ftpResponse = CType(ftpRequest.GetResponse, FtpWebResponse)

                strReader = New StreamReader(ftpResponse.GetResponseStream)

                If strReader IsNot Nothing Then sline = strReader.ReadLine

                While sline IsNot Nothing
                    ListOfFilesOnFTPSite.Add(sline)
                    sline = strReader.ReadLine
                End While

            Catch ex As Exception
                ExceptionInfo = ex

            Finally
                If ftpResponse IsNot Nothing Then
                    ftpResponse.Close()
                    ftpResponse = Nothing
                End If

                If strReader IsNot Nothing Then
                    strReader.Close()
                    strReader = Nothing
                End If
            End Try

            ListRemoteFiles = ListOfFilesOnFTPSite

            ListOfFilesOnFTPSite = Nothing
        End Function
        ''' <summary>
        ''' Shared method which will download a single file to a target location
        ''' </summary>
        ''' <param name="ftpAddress"></param>
        ''' <param name="ftpUser"></param>
        ''' <param name="ftpPassword"></param>
        ''' <param name="fileToDownload"></param>
        ''' <param name="downloadTargetFolder"></param>
        ''' <param name="deleteAfterDownload"></param>
        ''' <param name="ExceptionInfo"></param>
        ''' <returns></returns>
        ''' <remarks>ExceptionInfo will return any exception messages should a failure occur</remarks>
        Public Shared Function DownloadSingleFile(ftpAddress As String, ftpUser As String, ftpPassword As String, fileToDownload As String, _
                                                  downloadTargetFolder As String, deleteAfterDownload As Boolean, ExceptionInfo As Exception) As Boolean

            Dim FileDownloaded As Boolean = False

            Try

                Dim sFtpFile As String = ftpAddress & fileToDownload

                Dim sTargetFileName = System.IO.Path.GetFileName(sFtpFile)
                sTargetFileName = sTargetFileName.Replace("/", "\")
                sTargetFileName = downloadTargetFolder & sTargetFileName

                My.Computer.Network.DownloadFile(sFtpFile, sTargetFileName, ftpUser, ftpPassword)

                If deleteAfterDownload Then
                    Dim ftpRequest As FtpWebRequest = Nothing

                    ftpRequest = CType(WebRequest.Create(sFtpFile), FtpWebRequest)

                    With ftpRequest
                        .Credentials = New NetworkCredential(ftpUser, ftpPassword)
                        .Method = WebRequestMethods.Ftp.DeleteFile
                    End With

                    Dim response As FtpWebResponse = CType(ftpRequest.GetResponse(), FtpWebResponse)
                    response.Close()

                    ftpRequest = Nothing

                    FileDownloaded = True

                End If

            Catch ex As Exception
                ExceptionInfo = ex
            End Try

            Return FileDownloaded
        End Function


        Public Shared Function UploadFTPFiles(ftpAddress As String, ftpUser As String, ftpPassword As String, fileToUpload As String, _
                                                  targetFileName As String, deleteAfterUpload As Boolean, ExceptionInfo As Exception) As Boolean

            Dim credential As NetworkCredential

            Try
                Dim s As String

                credential = New NetworkCredential(ftpUser, ftpPassword)

                If ftpAddress.EndsWith("/") = False Then ftpAddress = ftpAddress & "/"

                Dim request As FtpWebRequest = DirectCast(WebRequest.Create(d), FtpWebRequest)

                request.KeepAlive = False
                request.Method = WebRequestMethods.Ftp.UploadFile
                request.Credentials = credential
                request.UsePassive = False
                request.Timeout = (60 * 1000) * 3 '3 mins

                Using reader As New FileStream(SourceFile.Get(context), FileMode.Open)


                    Dim buffer(Convert.ToInt32(reader.Length - 1)) As Byte
                    reader.Read(buffer, 0, buffer.Length)
                    reader.Close()

                    request.ContentLength = buffer.Length
                    Dim stream As Stream = request.GetRequestStream
                    stream.Write(buffer, 0, buffer.Length)
                    stream.Close()


                    Using response As FtpWebResponse = DirectCast(request.GetResponse, FtpWebResponse)


                        If deleteAfterUpload.Get(context) = True Then
                            My.Computer.FileSystem.DeleteFile(SourceFile.Get(context))
                        End If

                        response.Close()
                    End Using

                End Using

            Catch ex As Exception
                Exceptions.HandleException(ex, Me.DisplayName)
            Finally

            End Try
        End Function

    End Class

End Namespace


