Imports System.Net
Imports System.IO

Namespace IO

    Public Class FTP
        Public Shared Function ListFTPFiles(ftpAddress As String, ftpUser As String, ftpPassword As String, ExceptionInfo As Exception) As List(Of String)

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

            ListFTPFiles = ListOfFilesOnFTPSite

            ListOfFilesOnFTPSite = Nothing
        End Function
    End Class

End Namespace


