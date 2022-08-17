
Const MaximaCantidadVeribox As Integer = 25
Const THE_NAME = 1
Const THE_TAGID = 2
Dim Aux_R110_LASTRESPONSE As String
Private Sub CFixScheduler_Activated()

End Sub

Private Sub CFixScheduler_Initialize()

End Sub

Private Sub CFixScheduler_InitializeConfigure()

End Sub

Private Sub Diario_OnTimeOut(ByVal lTimerId As Long)

'If (readValue("Fix32.ECOGAS.NSD.F_CURACTIVENODE_0") = 0 And System.MyNodeName = "SRV01") Or (readValue("Fix32.ECOGAS.NSD.F_CURACTIVENODE_0") = 1 And System.MyNodeName = "SRV02") Then
If (System.MyNodeName = "SRV01") Or (System.MyNodeName = "SRV02") Then
    writeReports "Diario"
    writeReportsGU "Diario"
    'writeReportsVeribox "Diario"
End If

End Sub




Private Sub Estadisticas_OnTimeOut(ByVal lTimerId As Long)

If (System.MyNodeName = "SRV01") Then

    writeReportsOPCStatics "Horario"
End If

End Sub

Public Function writeReportsGU(tipo As String)
Dim oExcel As Object
Dim xlApp As Excel.Application
Dim xlLibro As Excel.Workbook
Dim oBook As Object
Dim oSheet As Object
Dim DIA
Dim hora
Dim Dia_a
Dim hora_a

Dim i, j, k As Integer



Dim sUserID As String
Dim sUserName As String
Dim sGroupName As String
Dim guData As Variant

Dim value As Integer
Dim randomText As String
  tiempo1 = Now
  DIA = Format(Date, "YYYY-MM-dd")
  hora = Format(Time, "HH:mm")
  'Start a new workbook in Excel
  Set xlApp = New Excel.Application
  Set xlLibro = xlApp.Workbooks.Open("C:\ECOGAS\REPORTE\Plantillas\GU.xlsx", True, True, , "")

  Set oSheet = xlApp.Worksheets("DATOS")


  k = 5
  
    For i = 3 To 4
            If i = 3 Then
                maxStations = 96
            ElseIf i = 4 Then
                maxStations = 108
            End If
            
            For j = 1 To maxStations
                auxZona = Format(i, "Z000")
                auxRTU = Format(j, "E000")
                auxID = auxZona + auxRTU
                guData = readGUValues(auxZona + "_" + auxRTU)
                If guData(1) = 0 Then
                    oSheet.Range("B" + CStr(k)).value = guData(3)
                    oSheet.Range("C" + CStr(k)).value = guData(4)
                    oSheet.Range("A" + CStr(k)).value = guData(2)
                    oSheet.Range("D" + CStr(k)).value = guData(5)
                    oSheet.Range("E" + CStr(k)).value = guData(6)
                    oSheet.Range("F" + CStr(k)).value = guData(7)
                    oSheet.Range("G" + CStr(k)).value = guData(8)
                    oSheet.Range("H" + CStr(k)).value = guData(9)
                    oSheet.Range("I" + CStr(k)).value = guData(10)
                    oSheet.Range("J" + CStr(k)).value = guData(11)
                    oSheet.Range("K" + CStr(k)).value = guData(12)
                    oSheet.Range("L" + CStr(k)).value = guData(13)
                    oSheet.Range("M" + CStr(k)).value = guData(14)
                    oSheet.Range("N" + CStr(k)).value = guData(15)
                    oSheet.Range("O" + CStr(k)).value = guData(16)
                    oSheet.Range("P" + CStr(k)).value = guData(17)
                    oSheet.Range("Q" + CStr(k)).value = guData(18)
                    oSheet.Range("R" + CStr(k)).value = guData(19)
                    oSheet.Range("S" + CStr(k)).value = guData(20)
                    oSheet.Range("T" + CStr(k)).value = guData(21)
                    oSheet.Range("U" + CStr(k)).value = guData(22)
                    oSheet.Range("V" + CStr(k)).value = guData(23)
                    oSheet.Range("W" + CStr(k)).value = guData(24)
                    oSheet.Range("X" + CStr(k)).value = guData(25)
                    oSheet.Range("Y" + CStr(k)).value = guData(26)
                    oSheet.Range("Z" + CStr(k)).value = guData(27)
                    oSheet.Range("AA" + CStr(k)).value = guData(28)
                    oSheet.Range("AB" + CStr(k)).value = guData(29)
                    oSheet.Range("AC" + CStr(k)).value = guData(30)
                    oSheet.Range("AD" + CStr(k)).value = guData(31)
                    oSheet.Range("AE" + CStr(k)).value = guData(32)
                    oSheet.Range("AF" + CStr(k)).value = guData(33)
                    oSheet.Range("AG" + CStr(k)).value = guData(34)
                Else
                    'oSheet.Range("B" + CStr(k)).value = "FALLA"
                    'oSheet.Range("C" + CStr(k)).value = "FALLA"
                    'oSheet.Range("A" + CStr(k)).value = "FALLA"
                    oSheet.Range("B" + CStr(k)).value = guData(3)
                    oSheet.Range("C" + CStr(k)).value = guData(4)
                    oSheet.Range("A" + CStr(k)).value = guData(2)
                    oSheet.Range("D" + CStr(k)).value = "FALLA"
                    oSheet.Range("E" + CStr(k)).value = "FALLA"
                    oSheet.Range("F" + CStr(k)).value = "FALLA"
                    oSheet.Range("G" + CStr(k)).value = "FALLA"
                    oSheet.Range("H" + CStr(k)).value = "FALLA"
                    oSheet.Range("I" + CStr(k)).value = "FALLA"
                    oSheet.Range("J" + CStr(k)).value = "FALLA"
                    oSheet.Range("K" + CStr(k)).value = "FALLA"
                    oSheet.Range("L" + CStr(k)).value = "FALLA"
                    oSheet.Range("M" + CStr(k)).value = "FALLA"
                    oSheet.Range("N" + CStr(k)).value = "FALLA"
                    oSheet.Range("O" + CStr(k)).value = "FALLA"
                    oSheet.Range("P" + CStr(k)).value = "FALLA"
                    oSheet.Range("Q" + CStr(k)).value = "FALLA"
                    oSheet.Range("R" + CStr(k)).value = "FALLA"
                    oSheet.Range("S" + CStr(k)).value = "FALLA"
                    oSheet.Range("T" + CStr(k)).value = "FALLA"
                    oSheet.Range("U" + CStr(k)).value = "FALLA"
                    oSheet.Range("V" + CStr(k)).value = "FALLA"
                    oSheet.Range("W" + CStr(k)).value = "FALLA"
                    oSheet.Range("X" + CStr(k)).value = "FALLA"
                    oSheet.Range("Y" + CStr(k)).value = "FALLA"
                    oSheet.Range("Z" + CStr(k)).value = "FALLA"
                    oSheet.Range("AA" + CStr(k)).value = "FALLA"
                    oSheet.Range("AB" + CStr(k)).value = "FALLA"
                    oSheet.Range("AC" + CStr(k)).value = "FALLA"
                    oSheet.Range("AD" + CStr(k)).value = "FALLA"
                    oSheet.Range("AE" + CStr(k)).value = "FALLA"
                    oSheet.Range("AF" + CStr(k)).value = "FALLA"
                    oSheet.Range("AG" + CStr(k)).value = "FALLA"
                End If
                k = k + 1
                Next
        Next
  

    System.FixGetUserInfo sUserID, sUserName, sGroupName
    'oSheet.Range("A1").value = "Usuario: " & sUserName
    'oSheet.Range("A2").value = "Usuario: " & sUserName
     'Acá comienzo a escribir datos Veribox en Hoja
    If tipo = "Diario" Then
       
        Set oSheet = xlApp.Worksheets("Veribox")
        maxStations = 25
        For j = 1 To maxStations
            If j < 10 Then
                auxRTU = "00" + CStr(j)
            Else
                auxRTU = "0" + CStr(j)
            End If
            k = j + 4
            auxID = auxZona + auxRTU
            If readBDValue("Fix32.ECOGAS.V" + auxRTU + "COM.F_CV") = 0 Then
                'oSheet.Range("A" + CStr(k)).value = readBDValue("Fix32.ECOGAS.V" + auxRTU + "_FCDIA.F_CV") & "-" & readBDValue("Fix32.ECOGAS.V" + auxRTU + "_FCMES.F_CV") & "-2020"
                oSheet.Range("A" + CStr(k)).value = readBDValue("Fix32.ECOGAS.V" + auxRTU + "DATETIME.A_DESC")
                oSheet.Range("B" + CStr(k)).value = readBDValue("Fix32.ECOGAS.V" + auxRTU + "SQLID.A_DESC")
                oSheet.Range("C" + CStr(k)).value = readBDValue("Fix32.ECOGAS.V" + auxRTU + "SQLID.F_CV")
                oSheet.Range("D" + CStr(k)).value = readBDValue("Fix32.ECOGAS.V" + auxRTU + "FQ011A.A_CV")
                oSheet.Range("E" + CStr(k)).value = readBDValue("Fix32.ECOGAS.V" + auxRTU + "FQ011.A_CV")
                oSheet.Range("F" + CStr(k)).value = readBDValue("Fix32.ECOGAS.V" + auxRTU + "FUQ011A.A_CV")
                oSheet.Range("G" + CStr(k)).value = readBDValue("Fix32.ECOGAS.V" + auxRTU + "FUQ011.A_CV")
                oSheet.Range("H" + CStr(k)).value = readBDValue("Fix32.ECOGAS.V" + auxRTU + "PT011.F_CV")
                oSheet.Range("I" + CStr(k)).value = readBDValue("Fix32.ECOGAS.V" + auxRTU + "TT011.F_CV")
                oSheet.Range("J" + CStr(k)).value = readBDValue("Fix32.ECOGAS.V" + auxRTU + "ET001.F_CV")
                oSheet.Range("K" + CStr(k)).value = readBDValue("Fix32.ECOGAS.V" + auxRTU + "ET002.F_CV")
                oSheet.Range("L" + CStr(k)).value = readBDValue("Fix32.ECOGAS.V" + auxRTU + "AUCV.A_CV")
            Else
                'oSheet.Range("A" + CStr(k)).value = readBDValue("Fix32.ECOGAS.V" + auxRTU + "DATETIME.A_DESC")
                oSheet.Range("B" + CStr(k)).value = readBDValue("Fix32.ECOGAS.V" + auxRTU + "SQLID.A_DESC")
                oSheet.Range("C" + CStr(k)).value = "FALLA"
                oSheet.Range("D" + CStr(k)).value = "FALLA"
                oSheet.Range("E" + CStr(k)).value = "FALLA"
                oSheet.Range("F" + CStr(k)).value = "FALLA"
                oSheet.Range("G" + CStr(k)).value = "FALLA"
                oSheet.Range("H" + CStr(k)).value = "FALLA"
                oSheet.Range("I" + CStr(k)).value = "FALLA"
                oSheet.Range("J" + CStr(k)).value = "FALLA"
                oSheet.Range("K" + CStr(k)).value = "FALLA"
                oSheet.Range("L" + CStr(k)).value = "FALLA"
            End If
            
        Next
    End If
   'SALVO,CIERRO E IMPRIMO ARCHIVO EXCEL TODO EN 1
   
   
    Dia_a = Format(Date, "_YYYY-MM-dd")
    hora_a = Format(Time, "_HH-mm")

    value = CInt(Int((99 * Rnd()) + 1))
    randomText = CStr(value)
    'Calculo el tiempo de ejecución de esta planilla
    tiempocalculado = Now - tiempo1
    calculo = Format(tiempocalculado, "sss.s")
    oSheet.Range("A2").value = "Generado: " & Now & " en " & calculo & "ms"
   
    If tipo = "Horario" Then
        xlLibro.SaveAs ("C:\ECOGAS\REPORTE\GU\Horarios\Horarios" & Dia_a & hora_a & "_" & randomText & ".xlsx")
    Else
        xlLibro.SaveAs ("C:\ECOGAS\REPORTE\GU\Diarios\Diarios" & Dia_a & hora_a & "_" & randomText & ".xlsx")
    End If
   xlApp.Quit
End Function

Private Sub Horarios_OnTimeOut(ByVal lTimerId As Long)

    '2017-05-02 Se reemplaza la condición para que se graben los datos horarios ininterrumpidamente en los nodos SRV01 y SRV02
    'If (readValue("Fix32.ECOGAS.NSD.F_CURACTIVENODE_0") = 0 And System.MyNodeName = "SRV01") Or (readValue("Fix32.ECOGAS.NSD.F_CURACTIVENODE_0") = 1 And System.MyNodeName = "SRV02") Then
    If (System.MyNodeName = "SRV01") Or (System.MyNodeName = "SRV02") Then
        writeReports "Horario"
        writeReportsGU "Horario"
    
    End If

End Sub


Private Function readBDValue(TAG As String) As String
    On Error GoTo ErrorHandler
    readBDValue = readValue(TAG, 1)
    Exit Function
ErrorHandler:
    readBDValue = " "
End Function

Private Function readBDValueXLS(TAG As String) As String
    On Error GoTo ErrorHandler
    readBDValueXLS = readValue(TAG, 1)
    Exit Function
ErrorHandler:
    readBDValueXLS = "0"
End Function



Private Sub postgresGU_OnTimeOut(ByVal lTimerId As Long)
    Dim cnn As New ADODB.Connection
    Dim rst As New ADODB.Recordset
    
    Dim i As Integer
    Dim maxStations As Integer
    
    Dim auxRTU As String
    
    
    Dim sqlCMD As String
  
    

    On Error Resume Next
 
        'CONTROLAR al copiar archivo para SRV02 postgresGU.TimerStart =  12:30:00
    If (readValue("Fix32.ECOGAS.NSD.F_CURACTIVENODE_0") = 0 And System.MyNodeName = "SRV01") Or (readValue("Fix32.ECOGAS.NSD.F_CURACTIVENODE_0") = 1 And System.MyNodeName = "SRV02") Then

        cnn.ConnectionString = "Driver=PostgreSQL;" & _
                               "Server=" & "143.4.12.52" & ";" & _
                               "Port=" & "5434" & ";" & _
                               "User Id=" & "postgres" & ";" & _
                               "Password=" & "DAC5D9DECDD8CFD9" & ";" & _
                               "Database=" & "EcoGas_GNC" & ";"
                               
        cnn.Open
    
        If (cnn.State) Then
            For i = 3 To 4
                If i = 3 Then
                    maxStations = 96
                ElseIf i = 4 Then
                    maxStations = 108
                End If
                
                For j = 1 To maxStations
                    auxZona = Format(i, "Z000")
                    auxRTU = Format(j, "E000")
                    sqlCMD = sqlCmdGU(auxZona + "_" + auxRTU)
                    If sqlCMD <> "" And sqlCMD <> "NO RTU" Then
                        cnn.Execute sqlCMD
                    End If
                Next j
            Next i
      
        End If
      
        cnn.Close
    End If
End Sub

Private Sub postgresMedCroma_OnTimeOut(ByVal lTimerId As Long)
    Dim cnn As New ADODB.Connection
    Dim rst As New ADODB.Recordset
    
    Dim i As Integer
    Dim maxStations As Integer
    
    Dim auxRTU As String
    
    
    Dim sqlCMD As String
  
    
  
  On Error Resume Next
 

    If (readValue("Fix32.ECOGAS.NSD.F_CURACTIVENODE_0") = 0 And System.MyNodeName = "SRV01") Or (readValue("Fix32.ECOGAS.NSD.F_CURACTIVENODE_0") = 1 And System.MyNodeName = "SRV02") Then
    
    
        cnn.ConnectionString = "Driver=PostgreSQL;" & _
                               "Server=" & "143.4.12.52" & ";" & _
                               "Port=" & "5434" & ";" & _
                               "User Id=" & "postgres" & ";" & _
                               "Password=" & "DAC5D9DECDD8CFD9" & ";" & _
                               "Database=" & "EcoGas_GNC" & ";"
                               
        cnn.Open
    
        If (cnn.State) Then
            maxStations = 200
            For i = 2 To maxStations
                auxRTU = Format(i, "000")
                
                If auxRTU = "044" Or auxRTU = "004" Or auxRTU = "002" Or auxRTU = "003" Then 'Sólo datos de Cromatografía
                    sqlCMD = sqlCmdCROMA("R" + auxRTU)
                    If sqlCMD <> "" And sqlCMD <> "NO RTU" Then
                        cnn.Execute sqlCMD
                    End If
                Else
                    
                        
                            If Not (auxRTU = "045" Or auxRTU = "045") Then
                                
                                    sqlCMD = sqlCmdRTU("R" + auxRTU)
                                    If sqlCMD <> "" And sqlCMD <> "NO RTU" Then
                                        cnn.Execute sqlCMD
                                    End If
                               
                            End If
                        
                    
                End If
  

            Next i
      
        End If
      
        cnn.Close

    End If
End Sub

Private Function writeReportsOPCStatics(tipo As String)
Dim oExcel As Object
Dim xlApp As Excel.Application
Dim xlLibro As Excel.Workbook
Dim oBook As Object
Dim oSheet As Object
Dim DIA
Dim hora

Dim i, j, k As Integer

Static diaAnterior

Dim sUserID As String
Dim sUserName As String
Dim sGroupName As String

  DIA = Format(Date, "YYYY-MM-dd")
  hora = Format(Time, "HH:mm")
  Dia_a = Format(Date, "_YYYY-MM-dd")
  hora_a = Format(Time, "_HH-mm")



  'Start a new workbook in Excel
  Set xlApp = New Excel.Application
  Set xlLibro = xlApp.Workbooks.Open("C:\ECOGAS\REPORTE\Plantillas\OPCEstadisticas.xlsx", True, False, , "")

  Set oSheet = xlApp.Worksheets("HOJA1")

    'oSheet.Rows("3:3").Select
    oSheet.Rows(4).Insert
    
    
    oSheet.Range("A4").value = DIA + " " + hora
    'oSheet.Range("B3").Value = readBDValue("Fix32.ECOGAS.MS001_RADIO_PERCENTVALID.A_DESC")
    oSheet.Range("B4").value = readBDValue("Fix32.ECOGAS.MS001_RADIO_PERCENTVALID.F_CV")
    oSheet.Range("C4").value = readBDValue("Fix32.ECOGAS.MS001_RADIO_PERCENTRETURN.F_CV")
    oSheet.Range("D4").value = readBDValue("Fix32.ECOGAS.MS001_MODEM_PERCENTVALID.F_CV")
    oSheet.Range("E4").value = readBDValue("Fix32.ECOGAS.MS001_MODEM_PERCENTRETURN.F_CV")
    oSheet.Range("F4").value = readBDValue("Fix32.ECOGAS.MS001_GPRS_PERCENTVALID.F_CV")
    oSheet.Range("G4").value = readBDValue("Fix32.ECOGAS.MS001_GPRS_PERCENTRETURN.F_CV")
    oSheet.Range("H4").value = readBDValue("Fix32.ECOGAS.MS001_GPRS2_PERCENTVALID.F_CV")
    oSheet.Range("I4").value = readBDValue("Fix32.ECOGAS.MS001_GPRS2_PERCENTRETURN.F_CV")
    oSheet.Range("L4").value = readBDValue("Fix32.ECOGAS.MS001_IP_RADIO_PERCENTVALID.F_CV")
    oSheet.Range("M4").value = readBDValue("Fix32.ECOGAS.MS001_IP_RADIO_PERCENTRETURN.F_CV")
    oSheet.Range("J4").value = readBDValue("Fix32.ECOGAS.MS001_IP_CROMA_PERCENTVALID.F_CV")
    oSheet.Range("K4").value = readBDValue("Fix32.ECOGAS.MS001_IP_CROMA_PERCENTRETURN.F_CV")
    oSheet.Range("N4").value = readBDValue("Fix32.ECOGAS.MS001_GU_GPRS_PERCENTVALID.F_CV")
    oSheet.Range("O4").value = readBDValue("Fix32.ECOGAS.MS001_GU_GPRS_PERCENTRETURN.F_CV")
    oSheet.Range("P4").value = readBDValue("Fix32.ECOGAS.MS001_GU_MODEM_PERCENTVALID.F_CV")
    oSheet.Range("Q4").value = readBDValue("Fix32.ECOGAS.MS001_GU_MODEM_PERCENTRETURN.F_CV")
    oSheet.Range("R4").value = readBDValue("Fix32.ECOGAS.MS001_GPRS3_PERCENTVALID.F_CV")
    oSheet.Range("S4").value = readBDValue("Fix32.ECOGAS.MS001_GPRS3_PERCENTRETURN.F_CV")
    oSheet.Range("T4").value = readBDValue("Fix32.ECOGAS.MS001_VPNTGN_PERCENTVALID.F_CV")
    oSheet.Range("U4").value = readBDValue("Fix32.ECOGAS.MS001_VPNTGN_PERCENTRETURN.F_CV")
    oSheet.Range("V4").value = readBDValue("Fix32.ECOGAS.MS001_DATOS_DISPONIBLE.F_CV")
    oSheet.Range("W4").value = readBDValue("Fix32.ECOGAS.MS001_DATOS_CORRECTOS.F_CV")
    oSheet.Range("X4").value = readBDValue("Fix32.ECOGAS.MS001_DATOS_OEE.F_CV")
    
      
    xlLibro.Save
    
 
    'If tipo = "Horario" Then
    '    xlLibro.SaveAs ("C:\ECOGAS\REPORTE\DATOS\Horarios\estadistica" & Dia_a & hora_a & ".xlsx")
   ' Else
    '    xlLibro.SaveAs ("C:\ECOGAS\REPORTE\DATOS\Diarios\stadistica" & Dia_a & hora_a & ".xlsx")
   ' End If
   xlApp.Quit
End Function


Private Sub PostgreVeriboxDataloggerToPDB_OnTimeOut(ByVal lTimerId As Long)
    Dim cnn As New ADODB.Connection
    Dim rst As New ADODB.Recordset
    Static iAlreadyHere As Integer

        
    If iAlreadyHere = 1 Then 'Si ya estamos aquí salgo
        Exit Sub
    End If
    iAlreadyHere = "1" ' bloqueo a otras llamadas a esta función
    
    strSQL = "Select datetime, presion, batt_vb from datalogger.veribox_data WHERE datetime  >= (NOW() - INTERVAL '24 hours' )  AND veribox_sn = 107905 ORDER BY datetime DESC LIMIT 1;"
    On Error Resume Next
    cnn.ConnectionString = "Driver=PostgreSQL UNICODE;" & _
                           "Server=" & "10.2.8.30" & ";" & _
                           "Port=" & "5432" & ";" & _
                           "User Id=" & "postgres" & ";" & _
                           "Password=" & "ecopass" & ";" & _
                           "Database=" & "gas" & ";"
                           
    cnn.Open
    rst.Open strSQL, cnn

    If (cnn.State) Then
                If Not rst.EOF Then

                    While (Not rst.EOF)  'leo hasta el final de la consulta
                        
                        If Not IsNull(rst.Fields(1).value) Then   'si es el valor no es nulo es nombre
                            R110_LASTRESPONSE = rst.Fields(0).value
                            
                            
                            If R110_LASTRESPONSE <> Aux_R110_LASTRESPONSE Then
                                writevalue rst.Fields(1).value, "FIX32.ECOGAS.R110PT002.F_CV", 1 'comunicación ok
                                writevalue rst.Fields(2).value / 1000, "FIX32.ECOGAS.R110ET001.F_CV", 1 'comunicación ok
                                writevalue rst.Fields(0).value, "FIX32.ECOGAS.R110_LASTRESPONSE.A_DESC", 1 'comunicación ok
                                writevalue 0, "FIX32.ECOGAS.R110COM.F_CV", 1 'comunicación OK
                                Aux_R110_LASTRESPONSE = R110_LASTRESPONSE
                                
                            End If
                            
                            
                         Else
                            If readBDValue("Fix32.ECOGAS.R110COM.F_CV") <> 1 Then
                               writevalue 1, "FIX32.ECOGAS.R110COM.F_CV", 1 'como la consulta es de las últimas 24 horas si es nulo entonces comunicación FALLA
                            End If
                         End If
                            

                        rst.MoveNext

                    Wend
                    
                Else
                    writevalue 1, "FIX32.ECOGAS.R110COM.F_CV", 1 'comunicación FALLA
                End If
                rst.Close
            
    End If
    
    cnn.Close
    iAlreadyHere = 0    'Habilito para una nueva llamda a la función
    'writevalue cantidadVeriboxConectadosDia, "FIX32.ECOGAS.MS001_VERIBOX_CONNECTED.F_CV", 1 'comunicación ok
    
    
End Sub

Private Sub PostgreVerivoxToPDB_OnTimeOut(ByVal lTimerId As Long)
    'Escribe sólo los datos de las 6AM en pantalla
    Dim cnn As New ADODB.Connection
    Dim rst As New ADODB.Recordset
    Dim i, j, pcaNumber, tagComValor As Integer
    Dim txtWhere As String
    Dim mapVeribox(1 To 2, 1 To 30) As String
    Static diaAnterior
    Static iAlreadyHere As Integer
    Dim cantidadVeriboxConectadosDia As Integer
        
    If iAlreadyHere = 1 Then 'Si ya estamos aquí salgo
        Exit Sub
    End If
    iAlreadyHere = "1" ' bloqueo a otras llamadas a esta función

    
    
    
    i = 1: mapVeribox(THE_NAME, i) = "Acabio C. L.":                    mapVeribox(THE_TAGID, i) = "V001"
    i = 2: mapVeribox(THE_NAME, i) = "Allevard Rejna Argentina":        mapVeribox(THE_TAGID, i) = "V002"
    i = 3: mapVeribox(THE_NAME, i) = "Arcor - Caroya":                  mapVeribox(THE_TAGID, i) = "V003"
    i = 4: mapVeribox(THE_NAME, i) = "Arcor - Recreo":                  mapVeribox(THE_TAGID, i) = "V004"
    i = 5: mapVeribox(THE_NAME, i) = "Bagley Argentina S.A.":           mapVeribox(THE_TAGID, i) = "V005"
    i = 6: mapVeribox(THE_NAME, i) = "Canteras Cerro Negro S.A.":       mapVeribox(THE_TAGID, i) = "V006"
    i = 7: mapVeribox(THE_NAME, i) = "Cartocor S.A.":                   mapVeribox(THE_TAGID, i) = "V007"
    i = 8: mapVeribox(THE_NAME, i) = "Cerámica Catamarca S.R.L.":       mapVeribox(THE_TAGID, i) = "V008"
    i = 9: mapVeribox(THE_NAME, i) = "Colortex":                        mapVeribox(THE_TAGID, i) = "V009"
    i = 10: mapVeribox(THE_NAME, i) = "D. G. F. Militares RIII DPM":       mapVeribox(THE_TAGID, i) = "V010"
    i = 11: mapVeribox(THE_NAME, i) = "D. G. F. Militares RIII DPQ":       mapVeribox(THE_TAGID, i) = "V011"
    i = 12: mapVeribox(THE_NAME, i) = "D. G. F. Militares VM":          mapVeribox(THE_TAGID, i) = "V012"
    i = 13: mapVeribox(THE_NAME, i) = "DPAMA S.A.":                     mapVeribox(THE_TAGID, i) = "V013"
   ' i = 14: mapVeribox(THE_NAME, i) = "F. de alimentos Santa Clara S.A.":       mapVeribox(THE_TAGID, i) = "V014"
    i = 14: mapVeribox(THE_NAME, i) = "Gas Carbónico Chiantore":        mapVeribox(THE_TAGID, i) = "V014"
    i = 15: mapVeribox(THE_NAME, i) = "José Guma S.A.":                 mapVeribox(THE_TAGID, i) = "V015"
    i = 16: mapVeribox(THE_NAME, i) = "Lácteos La Cristina S.A.":       mapVeribox(THE_TAGID, i) = "V016"
    i = 17: mapVeribox(THE_NAME, i) = "Metal Veneta S.A.":              mapVeribox(THE_TAGID, i) = "V017"
    i = 18: mapVeribox(THE_NAME, i) = "Molfino Hnos. S.A.":             mapVeribox(THE_TAGID, i) = "V018"
    i = 19: mapVeribox(THE_NAME, i) = "Olca S.A.I.C.":                  mapVeribox(THE_TAGID, i) = "V019"
    i = 20: mapVeribox(THE_NAME, i) = "Porta Hnos. S. A.":              mapVeribox(THE_TAGID, i) = "V020"
    i = 21: mapVeribox(THE_NAME, i) = "Refinería del Centro":           mapVeribox(THE_TAGID, i) = "V021"
    i = 22: mapVeribox(THE_NAME, i) = "Ricoltex":                       mapVeribox(THE_TAGID, i) = "V022"
    i = 23: mapVeribox(THE_NAME, i) = "Tecotex":                        mapVeribox(THE_TAGID, i) = "V023"
    i = 24: mapVeribox(THE_NAME, i) = "Vitopel S.A.":                   mapVeribox(THE_TAGID, i) = "V024"
    i = 25: mapVeribox(THE_NAME, i) = "Volkswagen Argentina S.A.":      mapVeribox(THE_TAGID, i) = "V025"
    
    On Error Resume Next
    strSQL = "select datetime,stationname,vol_c,vol_nc from ucv.ucv_data order by datetime desc"
    
    'strSQL = "select t.datetime, t.stationname,t.vol_c, t.vol_nc from ucv.ucv_data as t where date_part('hour',t.datetime)= 6 and date_part('day',now()-t.datetime)<1;"
    txtSelect = "SELECT t.datetime, t.stationname,pca,t.vol_c, t.vol_nc ,t.vol_c_an, t.vol_nc_an, t.presion, t.temperatura, t.batt_vc ,t.batt_vb, t.Alarma , t.veribox_sn "
    txtFrom = "FROM ucv.ucv_data as t "
    txtWhere = "WHERE date_part('hour',t.datetime)= 6 and date_part('day',now()-t.datetime)<1 "
    'txtWhere = "WHERE date_part('hour',t.datetime)= 6 and date_part('day',now())<1 "
    txtOrderBy = "ORDER BY t.stationname "
    
    
    strSQL = txtSelect & txtFrom & txtWhere & txtOrderBy
    
    
    
    Dim a, b As String
    a = """"
    strSQL = Replace(strSQL, a, " ")
    strSQL = Replace(strSQL, "Where", "WHERE date_part('hour',datetime)= 6 AND ")
    
    Dim diaActual As Integer
    Dim mesActual As Integer
    
    
    horaActual = Hour(Now)
    diaActual = Day(Now)
    mesActual = Month(Now)
    
    cantidadVeriboxConectadosDia = readBDValue("FIX32.ECOGAS.MS001_VERIBOX_CONNECTED.F_CV")
                                
    If horaActual > 5 And diaActual <> diaAnterior Then
        cantidadRespuestas = readBDValue("FIX32.ECOGAS.MS001_VERIBOX_NUM_RESP.F_CV")
        cantidadPreguntas = readBDValue("FIX32.ECOGAS.MS001_VERIBOX_NUM_PREG.F_CV")
        cantidadOperativos = readBDValue("FIX32.ECOGAS.MS001_VERIBOX_OPERATIVOS.F_CV")
        writevalue cantidadRespuestas + cantidadVeriboxConectadosDia, "FIX32.ECOGAS.MS001_VERIBOX_NUM_RESP.F_CV", 1 'comunicación ok
        writevalue cantidadPreguntas + cantidadOperativos, "FIX32.ECOGAS.MS001_VERIBOX_NUM_PREG.F_CV", 1 'comunicación ok
    
    
        
    'pongo en falla una sola véz al día a todos, si estamos en un nuevo día de gas
    'para luego sacarlo de falla a medida que tenga los datos
        diaAnterior = diaActual
        k = MaximaCantidadVeribox 'me sirve para mapear todos los datos que tengo
        cantidadVeriboxConectadosDia = 0
        Do
             tagNumber = mapVeribox(THE_TAGID, k)
             writevalue 1, "FIX32.ECOGAS." & tagNumber & "COM.F_CV", 1  'cargo en BDTR
             writevalue 1111111, "FIX32.ECOGAS." & tagNumber & "FQ011A.F_CV", 1 'comunicación ok
             writevalue 1111111, "FIX32.ECOGAS." & tagNumber & "FUQ011.F_CV", 1   'comunicación ok

            k = k - 1
        Loop While k > 0
    End If

    
    cnn.ConnectionString = "Driver=PostgreSQL UNICODE;" & _
                           "Server=" & "10.2.8.30" & ";" & _
                           "Port=" & "5432" & ";" & _
                           "User Id=" & "postgres" & ";" & _
                           "Password=" & "ecopass" & ";" & _
                           "Database=" & "gas" & ";"
                           
    cnn.Open
    rst.Open strSQL, cnn

    If (cnn.State) Then
                If Not rst.EOF Then
                    columnas = rst.Fields.Count 'MSFlexGrid1.Cols = rst.Fields.Count
                    j = 0
                    i = 1
                    While (Not rst.EOF)  'leo hasta el final de la consulta
                        tagNumber = ""
                        If Not IsNull(rst.Fields(1).value) Then   'si es el valor no es nulo es nombre
                            resultadoNombreLinea = rst.Fields(1).value
                            
                            k = MaximaCantidadVeribox
                            Do
                                If resultadoNombreLinea = mapVeribox(THE_NAME, k) Then
                                'si el tag de la línea coincide con algunos del mapa
                                'proceso el dato
                                    tagNumber = mapVeribox(THE_TAGID, k)
                                    tagComValor = readBDValue("FIX32.ECOGAS." & tagNumber & "COM.F_CV") + 0
                                    'me fijo si ya obtuve el dato del día
                                    If tagComValor > 0 Then
                                    ' si está con falla lo pongo ok y lo proceso
                                        writevalue 0, "FIX32.ECOGAS." & tagNumber & "COM.F_CV", 1 'nuevo dato
                                        cantidadVeriboxConectadosDia = cantidadVeriboxConectadosDia + 1 'sumo un dato
                                    Else
                                    ' ya tengo el dato, no lo proceso
                                        tagNumber = ""
                                    End If
                                    k = 0 'encontrado y salgo
                                Else
                                    k = k - 1 'sigo buscando
                                End If
                                
                            Loop While k > 0
                        End If
                        
                        If Not tagNumber = "" Then  'si obtuve su tag lo escribo
                            'MsgBox rst.Fields(1).Value & " " & tagNumber & " " & rst.Fields(3).Value
                            Dim bb As Long
                            
                            bb = rst.Fields(3).value
                            
                            
                            writevalue rst.Fields(0).value, "FIX32.ECOGAS." & tagNumber & "DATETIME.A_DESC", 1  'comunicación ok
                            'writevalue rst.Fields(1).value, "FIX32.ECOGAS." & tagNumber & "SQLID.A_DESC", 1  'comunicación ok
                            writevalue CStr(diaActual), "FIX32.ECOGAS." & tagNumber & "_FCDIA.F_CV", 1 'comunicación ok
                            writevalue CStr(mesActual), "FIX32.ECOGAS." & tagNumber & "_FCMES.F_CV", 1 'comunicación ok
                            writevalue rst.Fields(2).value, "FIX32.ECOGAS." & tagNumber & "SQLID.A_CV", 1 'comunicación ok
                            
                            writevalue rst.Fields(3).value, "FIX32.ECOGAS." & tagNumber & "FQ011.A_CV", 1  'comunicación ok
                            writevalue rst.Fields(4).value, "FIX32.ECOGAS." & tagNumber & "FUQ011.A_CV", 1 'comunicación ok
                            writevalue rst.Fields(7).value, "FIX32.ECOGAS." & tagNumber & "PT011.F_CV", 1 'comunicación ok
                            writevalue rst.Fields(8).value, "FIX32.ECOGAS." & tagNumber & "TT011.F_CV", 1 'comunicación ok
                            writevalue rst.Fields(9).value, "FIX32.ECOGAS." & tagNumber & "ET001.F_CV", 1 'comunicación ok
                            writevalue rst.Fields(10).value, "FIX32.ECOGAS." & tagNumber & "ET002.F_CV", 1 'comunicación ok
                            writevalue rst.Fields(11).value, "FIX32.ECOGAS." & tagNumber & "AUCV.A_CV", 1 'comunicación ok
                            a = rst.Fields(5).value


                            'si el dato es nulo lo pongo en -1
                            If Not IsNull(rst.Fields(5).value) Then  ' reviso si no es nulo
                                writevalue rst.Fields(5).value, "FIX32.ECOGAS." & tagNumber & "FQ011A.F_CV", 1   'comunicación ok
                            Else
                                writevalue -1, "FIX32.ECOGAS." & tagNumber & "FQ011A.F_CV", 1   'sin dato
                            End If
                            'si el dato es nulo lo pongo en -1
                            If Not IsNull(rst.Fields(6).value) Then  ' reviso si no es nulo
                                writevalue rst.Fields(6).value, "FIX32.ECOGAS." & tagNumber & "FUQ011A.F_CV", 1   'comunicación ok
                            Else
                                writevalue -1, "FIX32.ECOGAS." & tagNumber & "FUQ011A.F_CV", 1    'sin dato
                            End If
                        End If

                        rst.MoveNext

                    Wend
            
            End If
            rst.Close
            
    End If
    
    cnn.Close
    iAlreadyHere = 0    'Habilito para una nueva llamda a la función
    'writevalue cantidadVeriboxConectadosDia, "FIX32.ECOGAS.MS001_VERIBOX_CONNECTED.F_CV", 1 'comunicación ok
    

End Sub
