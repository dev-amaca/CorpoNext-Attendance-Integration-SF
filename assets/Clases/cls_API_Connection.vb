Imports System.Text
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel
Imports System.Net
Imports System.IO

Public Class cls_API_Connection

    Public strURL As String
    Public strUser As String
    Public strPassword As String
    Public strUserCompount As String
    Public strCompany As String
    Public strFunction As String
    Public strEnviroment As String
    Public strAuthBasic As String
    Public oResult As Object

    Public Function GetAutnBasic(strUser, strPassword, strCompany) As String

        Dim credentials As String = strUser & "@" & strCompany & ":" & strPassword

        strAuthBasic = Convert.ToBase64String(Encoding.ASCII.GetBytes(credentials))

        Return strAuthBasic

    End Function


End Class
