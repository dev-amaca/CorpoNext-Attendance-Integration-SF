Imports System.Net
Imports System.Net.Http
Imports System.Net.Http.Headers
Imports System.Text
Imports Newtonsoft.Json
Imports System.Data.OleDb
Imports zkemkeeper
Imports System.IO
Imports System.Diagnostics.Metrics
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.Button
Imports System.Runtime.Intrinsics.Arm
Imports System.Threading
Imports System.Reflection.Emit
Imports System.Timers
Imports System.Reflection.Metadata
Imports System.Buffers
Imports Newtonsoft


Public Class Home

    Dim DB_Conn As cls_DataBase = New cls_DataBase
    Dim oDispTimer As cls_Disp_Timer = New cls_Disp_Timer
    Dim oDispTime As cls_Data_Time_User = New cls_Data_Time_User
    Dim oConn_Api As cls_API_Connection = New cls_API_Connection

    Dim CounterProgressBar As Integer
    Dim blnRegistrosEliminados_FechaSelected As Boolean = False

    Public axCZKEM1 As New zkemkeeper.CZKEM
    Public Property API_Content As StringContent

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            ProgressBar1.Value = 0

            Me.Enabled = False

            Dim blnSuccessSaveUserTime As Boolean = True
            Dim DateSelected As Date = DateTimePicker1.Text
            blnRegistrosEliminados_FechaSelected = False

            'INIT PROCESS
            EscribeLog("** INIT APP **")

            EscribeLog("VALIDAR CONEXION CON DISPOSITIVO TIMER CORPO")

            oDispTimer.blnSuccessConn = Connect_Disp_Timer(oDispTimer)

            If oDispTimer.blnSuccessConn Then
                EscribeLog("Conexion Exitosa")

                oDispTimer.intMachineNumer = 1
                axCZKEM1.RegEvent(oDispTimer.intMachineNumer, 65535)

                Cursor = Cursors.WaitCursor

                EscribeLog("Desactivando Dospositivo")
                axCZKEM1.EnableDevice(oDispTimer.intMachineNumer, False) 'DESACTIVAR DISPOSITIVO PARA SER CONSUMIDO


                EscribeLog("Leyendo todos los registros almacenados en memoriria.")
                If axCZKEM1.EnableDevice(oDispTimer.intMachineNumer, False) Then 'READ ALL DATA IN DISPOSITIVE

                    EscribeLog("Recorriendo los registros obtenidos.")

                    Dim CountRows As Integer = 0
                    Dim Status As Integer = 0

                    axCZKEM1.GetDeviceStatus(oDispTimer.intMachineNumer, Status, CountRows)

                    If CountRows = 0 Then
                        EscribeLog("No Hay registros que enviar. Numero de registros: " & CountRows)
                        EscribeLog("Process Exit")
                        'Exit Sub
                    Else
                        EscribeLog("Numero de registros por enviar: " & CountRows)
                    End If

                    ProgressBar1.Maximum = 100

                    While axCZKEM1.SSR_GetGeneralLogData(oDispTimer.intMachineNumer, oDispTime.strEnrrollNumber, oDispTime.intIdVerifyMode, oDispTime.intInOutMode, oDispTime.intYear, oDispTime.intMonth, oDispTime.intDay, oDispTime.intHour, oDispTime.intMinute, oDispTime.intSecond, oDispTime.intWorkCode)

                        ProgressBar1.Value = 1

                        Dim oUserTime As cls_User_Time = New cls_User_Time

                        If oDispTimer.strEquipo = "CNXT_NORTE" And oDispTime.strEnrrollNumber = "291" Then
                            oDispTime.strEnrrollNumber = "296"
                        End If

                        oUserTime.strEnrollNumber = oDispTime.strEnrrollNumber
                        oUserTime.intVerifyMode = oDispTime.intIdVerifyMode
                        oUserTime.intInOutMode = oDispTime.intInOutMode
                        oUserTime.strCompleteDate = oDispTime.intDay.ToString() & "/" & oDispTime.intMonth.ToString() & "/" & oDispTime.intYear.ToString()
                        oUserTime.strHour = oDispTime.intHour.ToString() & ":" & oDispTime.intMinute.ToString() & ":" & oDispTime.intSecond.ToString()
                        oUserTime.strGroup = oDispTime.intYear.ToString() & "/" & oDispTime.intMonth.ToString() & "/" & oDispTime.intDay.ToString() & "-" & oDispTime.strEnrrollNumber
                        oUserTime.strEquipo = oDispTimer.strEquipo

                        If (oDispTime.intYear >= DateSelected.Year And oDispTime.intMonth = DateSelected.Month And oDispTime.intDay = (DateSelected.Day)) Then

                            '############################################################################################################################
                            '############################################# DEBUG MOOD ###################################################################

                            If oUserTime.strEnrollNumber = 277 Then

                                EscribeLog("Usuario Luis Manzanares:")
                                EscribeLog("")

                                EscribeLog("EnrollNumer: " & oUserTime.strEnrollNumber)
                                EscribeLog("intVerifyMode: " & oUserTime.intVerifyMode)
                                EscribeLog("intInOutMode: " & oUserTime.intInOutMode)
                                EscribeLog("strCompleteDate: " & oUserTime.strCompleteDate)
                                EscribeLog("strHour: " & oUserTime.strHour)
                                EscribeLog("strGroup: " & oUserTime.strGroup)
                                EscribeLog("strEquipo: " & oUserTime.strEquipo)

                                EscribeLog("")

                                EscribeLog("**** **** **** **** **** **** **** **** **** **** **** **** **** ")

                                EscribeLog("Las fechas del registro permiten guardarlo en la tabla dinamica.")

                                If Not Save_Reg_UserTime(oUserTime) Then
                                    blnSuccessSaveUserTime = False
                                End If
                                EscribeLog("**** **** **** **** **** **** **** **** **** **** **** **** **** ")

                            End If

                            '############################################################################################################################
                            '############################################# DEBUG MOOD ###################################################################

                        End If


                    End While

                    Cursor = Cursors.Default

                    'En este punto, la informacion ya se obtuvo de la terminal y ya se guardo en la base de datos dinamica. 
                    'Ahora se procede a obtener los registros de la base de datos dinamica.

                    'Validamos que no halla habido errores
                    If blnSuccessSaveUserTime Then
                        Dim DB_Data As OleDbDataReader
                        Dim CountX_Control As Integer
                        Dim CicloX As String = "A"
                        Dim strGroupX As String
                        Dim strGroupY As String

                        EscribeLog("Proceso para obtener los registros de BD Dinamica.")

                        DB_Data = Get_Reg_UserTime()

                        'Validarque si hay registros

                        If DB_Data.HasRows Then

                            'SI HAY REGISTROS
                            Dim oUser As cls_oJSON_UserTime = New cls_oJSON_UserTime

                            While DB_Data.Read

                                Try
                                    ProgressBar1.Value = 1

                                    'Variables para separar tiempos de entrada y salida
                                    Dim InHour, InMinute, InSecond, OutHour, OutMinute, OutSecond As String

                                    oUser.dtDateDB = DB_Data.Item("iDate").ToString()
                                    oUser.dtTimeDB = DB_Data.Item("iTime").ToString()
                                    strGroupX = DB_Data.Item("iGrupo").ToString()
                                    oUser.strParam2 = DB_Data.Item("sf_id").ToString()
                                    oUser.strEnroll = DB_Data.Item("EnrollNumber").ToString()

                                    If strGroupY <> strGroupX Then
                                        CicloX = "A"
                                        strGroupY = ""
                                        oUser.intCountx = 0
                                        CountX_Control = 0
                                    End If

                                    If CicloX = "A" Then
                                        'Esto parace ser la entrada

                                        InHour = oUser.dtTimeDB.Hour
                                        InMinute = oUser.dtTimeDB.Minute
                                        InSecond = oUser.dtTimeDB.Second

                                        oUser.intCountx = oUser.intCountx + 1
                                        CountX_Control = CountX_Control + 1

                                        oUser.strParam1 = strGroupX & "-" & oUser.strParam2 & "-" & CountX_Control

                                        oUser.strParam4 = "PT" & InHour & "H" & InMinute & "M" & InSecond & "S"
                                        oUser.strParam6 = "'" & oUser.strParam1 & "'"

                                        If CountX_Control = 1 Then
                                            oUser.strParam7 = "TT_Work"
                                        Else
                                            oUser.strParam7 = "TT_Work1"
                                        End If

                                        CicloX = "B"

                                        strGroupY = strGroupX

                                    Else
                                        'Esto parace ser la salida

                                        If strGroupX = strGroupY Then

                                            If oUser.strParam2 <> "" Then

                                                OutHour = oUser.dtTimeDB.Hour
                                                OutMinute = oUser.dtTimeDB.Minute
                                                OutSecond = oUser.dtTimeDB.Second

                                                If InHour = OutHour And InMinute = OutMinute Then

                                                    oUser.strParam5 = "PT" & OutHour & "H" & OutMinute + 1 & "M" & OutSecond & "S"
                                                    oUser.strParam3 = oUser.dtDateDB

                                                Else

                                                    oUser.strParam5 = "PT" & OutHour & "H" & OutMinute & "M" & OutSecond & "S"
                                                    oUser.strParam3 = oUser.dtDateDB

                                                End If

                                                Dim respDate As List(Of String) = DateFormat(oUser.dtDateDB)

                                                Dim intYear As Integer = respDate(2)
                                                Dim intMonth As Integer = respDate(1)
                                                Dim intDay As Integer = respDate(0)

                                                Dim epoch As DateTime = New DateTime(1970, 1, 1, 0, 0, 0)
                                                Dim currentTime As DateTime = New DateTime(intYear, intMonth, intDay, 0, 0, 0)
                                                Dim elapsedTime As TimeSpan = currentTime - epoch


                                                oUser.strParam3 = elapsedTime.TotalMilliseconds.ToString()

                                                'Armar Json REQUEST
                                                Dim oJSON As cls_JSON_Request = New cls_JSON_Request
                                                Dim oUri As cls_Metadata = New cls_Metadata

                                                oUri.uri = "ExternalTimeData(" + oUser.strParam6 + ")"

                                                oJSON.__metadata = New cls_Metadata
                                                oJSON.__metadata = oUri

                                                oJSON.userId = oUser.strParam2
                                                oJSON.externalCode = oUser.strParam1
                                                oJSON.startDate = "/Date(" + oUser.strParam3 + ")/"
                                                oJSON.startTime = oUser.strParam4
                                                oJSON.endTime = oUser.strParam5
                                                oJSON.timeType = oUser.strParam7

                                                Dim JSON_REQUEST As String = JsonConvert.SerializeObject(oJSON, Formatting.Indented)

                                                EscribeLog("########## JSON REQUEST ############")
                                                EscribeLog("Reviar log JSON REQUEST")
                                                EscribeLog_JSON_Request(JSON_REQUEST)

                                                EscribeLog("####################################")


                                                Conn_API_Send_ExternalTimeData(JSON_REQUEST)

                                                CicloX = "A"

                                            Else
                                                oUser.strCicloX = "A"
                                                CicloX = "A"
                                            End If

                                        Else

                                            strGroupY = ""
                                            oUser.strCicloX = "A"
                                            CicloX = "A"
                                            oUser.intCountx = 0
                                            CountX_Control = 0

                                        End If

                                        CicloX = "A"

                                    End If
                                Catch ex As Exception
                                    MessageBox.Show("Ocurrio un error al enviar las asistencias. Verifica el archivo log para conusltar el detalle del proceso.", " Proceso Terminado", MessageBoxButtons.OK)
                                    Exit While
                                    Exit Sub
                                End Try

                            End While

                            MessageBox.Show("Las asistencias fueron enviadas correctamente.", " Proceso Terminado", MessageBoxButtons.OK)
                        Else

                            'NO HAY REGISTROS
                            EscribeLog("No hay registros que enviar a SSFF")
                            blnRegistrosEliminados_FechaSelected = False

                            MessageBox.Show("No hay registros pendientes por enviar.", " Proceso Terminado", MessageBoxButtons.OK)

                        End If


                    End If

                    'Cerrar Conexion DB
                    DB_Conn.DB.Close()

                Else
                    blnRegistrosEliminados_FechaSelected = False
                    EscribeLog("No fue posible leer los registros almacenados en memoriria. Probablemente no hay registros u ocurrio un error. Validar posibles codigo de errores obtenidos.")
                    EscribeLog("Reactivando Dospositivo")
                    axCZKEM1.EnableDevice(oDispTimer.intMachineNumer, True) 'ACTIVAR DISPOSITIVO 
                End If




            Else
                blnRegistrosEliminados_FechaSelected = False
                EscribeLog("ERROR: No fue posible conectarse al dispositivo timer: " + oDispTimer.strEquipo)
                EscribeLog("Reactivando Dospositivo")
                axCZKEM1.EnableDevice(oDispTimer.intMachineNumer, True) 'ACTIVAR DISPOSITIVO 
            End If

        Catch ex As Exception

            blnRegistrosEliminados_FechaSelected = False
            EscribeLog("ERROR GRAL: " + ex.Message)
            EscribeLog("Reactivando Dospositivo")
            axCZKEM1.EnableDevice(oDispTimer.intMachineNumer, True) 'ACTIVAR DISPOSITIVO 

            MessageBox.Show("Ocurrio un error al enviar las asistencias. Mas detalles en " & Application.StartupPath & "\Logs Aplication", "Proceso Terminado", MessageBoxButtons.OK)

        End Try

        ProgressBar1.Value = 100



        Me.Enabled = True

    End Sub

    Private Function Connect_Disp_Timer(cls_DispTimer As cls_Disp_Timer) As Boolean
        Dim blnResp As Boolean = True
        Try

            blnResp = axCZKEM1.Connect_Net(cls_DispTimer.strIPEquipo.Trim(), Convert.ToInt32(cls_DispTimer.strPort.Trim()))
            blnResp = True

        Catch ex As Exception
            'ERROR
            blnResp = False
        End Try

        Return blnResp

    End Function

    Private Function Get_Data_Disp_Timer(cls_DispTimer As cls_Disp_Timer) As String

        Try

        Catch ex As Exception

        End Try

        Return ""
    End Function

    Private Function Save_Reg_UserTime(oUserTime As cls_User_Time) As Boolean
        Dim resp As Boolean = True

        Try

            EscribeLog("Abriendo conexion BD")

            'Esto es para eliminar los registros del dia en que se enviara la asistencia. Esto evitara que se dupliquen los registros.
            If Not blnRegistrosEliminados_FechaSelected Then
                Delete_Reg_Date_Selected(oUserTime)
            End If

            DB_Conn.DB.Open()

            If oUserTime.strEquipo = "CNXT_NORTE" And oUserTime.strEnrollNumber = "291" Then
                oUserTime.strEnrollNumber = "296"
            End If

            'Insertar Registro
            DB_Conn.cmd.CommandText = " INSERT INTO att_log (EnrollNumber, VerifyMode, InOutMode, iDate, iTime, iGrupo,iEquipo) VALUES('" & oUserTime.strEnrollNumber.ToString() & "', '" & oUserTime.intVerifyMode.ToString() & "', '" & oUserTime.intInOutMode.ToString() & "', '" & oUserTime.strCompleteDate.ToString() & "', '" & oUserTime.strHour.ToString() & "', '" & oUserTime.strGroup & "', '" & oUserTime.strEquipo & "')"

            EscribeLog("INSERT QUERY" + DB_Conn.cmd.CommandText)

            DB_Conn.cmd.ExecuteNonQuery()

            DB_Conn.cmd.CommandText = " INSERT INTO Registros_Historicos (EnrollNumber, VerifyMode, InOutMode, iDate, iTime, iGrupo,iEquipo) VALUES('" & oUserTime.strEnrollNumber.ToString() & "', '" & oUserTime.intVerifyMode.ToString() & "', '" & oUserTime.intInOutMode.ToString() & "', '" & oUserTime.strCompleteDate.ToString() & "', '" & oUserTime.strHour.ToString() & "', '" & oUserTime.strGroup & "', '" & oUserTime.strEquipo & "')"

            EscribeLog("INSERT QUERY RESPALDO" + DB_Conn.cmd.CommandText)

            DB_Conn.cmd.ExecuteNonQuery()

            DB_Conn.DB.Close()

            EscribeLog("Proceso terminado correctamente.")

        Catch ex As Exception

            EscribeLog("Ocurrio un error al insertar el registro: Error: " + ex.Message)
            resp = False

        End Try

        Return resp
    End Function

    Private Function Delete_Reg_Date_Selected(oUserTime As cls_User_Time) As Boolean
        Dim resp As Boolean = True

        Try

            EscribeLog("Abriendo conexion BD")

            DB_Conn.DB.Open()

            'Eliminar Registro Para No Duplicarlos. 

            DB_Conn.cmd.CommandText = " Delete FROM att_log"

            EscribeLog("DELETE QUERY: " + DB_Conn.cmd.CommandText)

            DB_Conn.cmd.ExecuteNonQuery()

            DB_Conn.DB.Close()

            EscribeLog("Proceso terminado correctamente.")

            blnRegistrosEliminados_FechaSelected = True

        Catch ex As Exception

            EscribeLog("Ocurrio un error al insertar el registro: Error: " + ex.Message)
            blnRegistrosEliminados_FechaSelected = False
            resp = False

        End Try

        Return resp
    End Function

    Private Function Get_Reg_UserTime() As OleDbDataReader

        Dim RestReg As OleDbDataReader

        Try

            EscribeLog("Abriendo conexion BD")

            DB_Conn.cmd.CommandText = "SELECT EnrollNumber,iDate, iTime, iGrupo, (SELECT sf_id FROM sf_user WHERE att_id = EnrollNumber AND active = """ & 1 & """) As sf_id FROM att_log WHERE iDate = Format(""" & DateTimePicker1.Text & """,""Short Date"") ORDER BY iGrupo, iTime ASC"

            EscribeLog("QUERY: " + DB_Conn.cmd.CommandText)

            EscribeLog("Ejecutando QUERY")

            DB_Conn.DB.Open()

            DB_Conn.cmd.ExecuteNonQuery()

            RestReg = DB_Conn.cmd.ExecuteReader()

            EscribeLog("Proceso terminado correctamente.")

        Catch ex As Exception

            EscribeLog("Ocurrio un error al obtener los registros: Error: " + ex.Message)

        End Try

        Return RestReg

    End Function

    Private Async Sub Conn_API_Send_ExternalTimeData(strJSON_Resquest As String)

        Try

            Await SendPostRequest(strJSON_Resquest)

        Catch ex As Exception

            EscribeLog_JSON_Request("Error al conectar a API: " & ex.Message)

        End Try

    End Sub

    Async Function SendPostRequest(strJSON_Resquest As String) As Task

        Using httpClient As New HttpClient()

            httpClient.BaseAddress = New Uri(oConn_Api.strURL)

            httpClient.DefaultRequestHeaders.Clear()

            httpClient.DefaultRequestHeaders.Authorization = New AuthenticationHeaderValue("Basic", oConn_Api.strAuthBasic)

            Dim postData As New StringContent(strJSON_Resquest, Encoding.UTF8, "application/json")

            Dim response As HttpResponseMessage = Await httpClient.PostAsync("", postData)

            If response.IsSuccessStatusCode Then

                Dim content As String = Await response.Content.ReadAsStringAsync()

                EscribeLog_JSON_Request("")
                EscribeLog_JSON_Request("Response: " & content)

            Else

                EscribeLog_JSON_Request("Request failed with status code: " & response.StatusCode)

            End If

        End Using
    End Function

    Private Sub EscribeLog(sMensaje)

        Dim file As System.IO.StreamWriter
        Dim prefix As String

        prefix = (DateTime.Now).Year.ToString & "." & (DateTime.Now).Month.ToString & "." & (DateTime.Now).Day.ToString
        file = My.Computer.FileSystem.OpenTextFileWriter(Application.StartupPath & "\Logs Aplication\" & prefix & " - Log_App.log", True)
        file.WriteLine(prefix + " - " + sMensaje)
        file.Close()
        sMensaje = ""

    End Sub

    Private Sub Home_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Try

            If SetDB_Settings() Then

                axCZKEM1 = New zkemkeeper.CZKEM

                oDispTimer.intNumEquipo = 1
                oDispTimer.strPort = "4370"
                oDispTimer.strIPEquipo = "192.168.8.9"
                oDispTimer.strEquipo = "CNXT_NORTE"

                txtIPDevice.Text = oDispTimer.strIPEquipo
                txtPortDevice.Text = oDispTimer.strPort
                txtNameDevice.Text = oDispTimer.strEquipo
                txtNumberDevice.Text = oDispTimer.intNumEquipo

                txtIPDevice.Enabled = False
                txtPortDevice.Enabled = False
                txtNameDevice.Enabled = False
                txtNumberDevice.Enabled = False
                btnSaveDispInfo.Enabled = False

                oConn_Api.strUser = "luis.manzanares"
                oConn_Api.strPassword = "manzana10"
                oConn_Api.strCompany = "C0017755703P"
                oConn_Api.strFunction = "upsert"
                oConn_Api.strURL = "https://api8.successfactors.com/odata/v2/" & oConn_Api.strFunction
                oConn_Api.strAuthBasic = oConn_Api.GetAutnBasic(oConn_Api.strUser, oConn_Api.strPassword, oConn_Api.strCompany)

                ProgressBar1.Minimum = 0
                ProgressBar1.Maximum = 3

                ComboBox1.SelectedText = "Productivo"
                ComboBox2.SelectedItem = "Norte"

                lblVersion.Text = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString()

            Else
                EscribeLog("Error init conn BD. Exit proccess.")
            End If

        Catch ex As Exception

            EscribeLog("Error gral. Error: " + ex.Message)

        End Try

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged

        Select Case ComboBox1.SelectedItem
            Case "Productivo"
                oConn_Api.strURL = "https://api19.sapsf.com/odata/v2/"
                oConn_Api.strCompany = "C0017755703P"
            Case "Test"
                oConn_Api.strURL = "https://api19.sapsf.com/odata/v2/"
                oConn_Api.strCompany = "C0017755703P"
            Case "Desarrollo"
                oConn_Api.strURL = "https://api19.sapsf.com/odata/v2/"
                oConn_Api.strCompany = "C0017755703P"
        End Select

    End Sub

    Private Function SetDB_Settings() As Boolean

        EscribeLog("Iniciar configs BD")

        Try

            DB_Conn.DB.ConnectionString = "Provider = Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "/DB_Control_Timers_User.accdb"
            DB_Conn.cmd.Connection = DB_Conn.DB
            DB_Conn.DB.Close()

            EscribeLog("Configuracion asignada correctamente")

            Return True

        Catch ex As Exception

            EscribeLog("Ocurrio un error al asignar la condiguracion: Error: " + ex.Message)

            Return False

        End Try

    End Function

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        ' Establece una conexión con el dispositivo
        If axCZKEM1.Connect_Net(oDispTimer.strIPEquipo, oDispTimer.strPort) Then
            ' Conexión exitosa

            ' Intenta reiniciar el dispositivo
            If axCZKEM1.RestartDevice(oDispTimer.intMachineNumer) Then
                ' El dispositivo se reinició con éxito
                MsgBox("El dispositivo se ha reiniciado con éxito.")
            Else
                ' Manejar el error de reinicio
                MsgBox("Error al reinciar el dispositivo.")
            End If

            ' Cierra la conexión
            axCZKEM1.Disconnect()
        Else
            ' Manejar el error de conexión
            MsgBox("No se pudo establecer la conexión con el dispositivo.")
        End If
    End Sub

    Private Sub btnSaveInfoDevice_Click(sender As Object, e As EventArgs) Handles btnSaveInfoDevice.Click

        txtIPDevice.Enabled = True
        txtPortDevice.Enabled = True
        txtNameDevice.Enabled = True
        txtNumberDevice.Enabled = True
        btnSaveDispInfo.Enabled = True

    End Sub

    Private Sub btnSaveDispInfo_Click(sender As Object, e As EventArgs) Handles btnSaveDispInfo.Click

        oDispTimer.intNumEquipo = txtNumberDevice.Text
        oDispTimer.strPort = txtPortDevice.Text
        oDispTimer.strIPEquipo = txtIPDevice.Text
        oDispTimer.strEquipo = txtNameDevice.Text

        txtIPDevice.Enabled = False
        txtPortDevice.Enabled = False
        txtNameDevice.Enabled = False
        txtNumberDevice.Enabled = False
        btnSaveDispInfo.Enabled = False

    End Sub

    Private Function DateFormat(strDateOriginal As String) As List(Of String)
        Dim respDate As List(Of String)

        Try
            Dim arrDate As List(Of String)

            arrDate = strDateOriginal.Split("/").ToList()

            respDate = arrDate

        Catch ex As Exception

        End Try

        Return respDate
    End Function

    Private Sub EscribeLog_JSON_Request(sMensaje)

        Dim file As System.IO.StreamWriter
        Dim prefix As String

        prefix = (DateTime.Now).Year.ToString & "." & (DateTime.Now).Month.ToString & "." & (DateTime.Now).Day.ToString
        file = My.Computer.FileSystem.OpenTextFileWriter(Application.StartupPath & "\Logs Aplication\" & prefix & " - JSON_Request.log", True)
        file.WriteLine(prefix + " - " + sMensaje)
        file.Close()
        sMensaje = ""

    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged

        Select Case ComboBox2.SelectedItem
            Case "Norte"
                oDispTimer.intNumEquipo = 1
                oDispTimer.strPort = "4370"
                oDispTimer.strIPEquipo = "192.168.8.9"
                oDispTimer.strEquipo = "CNXT_NORTE"

                txtIPDevice.Text = oDispTimer.strIPEquipo
                txtPortDevice.Text = oDispTimer.strPort
                txtNameDevice.Text = oDispTimer.strEquipo
                txtNumberDevice.Text = oDispTimer.intNumEquipo
            Case "Centro"
                oDispTimer.intNumEquipo = 2
                oDispTimer.strPort = "4370"
                oDispTimer.strIPEquipo = "189.208.208.4"
                oDispTimer.strEquipo = "CNEXT_WTC_27"

                txtIPDevice.Text = oDispTimer.strIPEquipo
                txtPortDevice.Text = oDispTimer.strPort
                txtNameDevice.Text = oDispTimer.strEquipo
                txtNumberDevice.Text = oDispTimer.intNumEquipo
        End Select

    End Sub

End Class
