
Option Explicit

Const THE_NAME = 1
Const THE_DESC = 2
Const THE_UNIT = 3
Const THE_WIDH = 4
Const THE_COMMA = ", "
Const THE_COMMA_CRLF = ", " & vbCrLf

Const GU_FIELDS = 33
Const MED_FIELDS = 20
Const CROMA_FIELDS = 18
Const VERIBOX_FIELDS = 13

Const GU_STATIONS = 193
Const CAMARAS_STATIONS = 191
Const CROMAS_STATIONS = 3
Const VERIBOX_STATIONS = 27

Const DT_FORMAT = "d/MMM/yy H:mm:ss"
Const NO_QUERY = "¡¡NO QUERY!!"
Const FieldQty As Integer = 50
Const SortCritQty As Integer = 3

Const TABU = "     "

Dim FieldNamesGU(1 To 4, 1 To FieldQty) As String
Dim FieldNamesMedicion(1 To 4, 1 To FieldQty) As String
Dim FieldNamesCromatografia(1 To 4, 1 To FieldQty) As String
Dim FieldNamesVeribox(1 To 4, 1 To FieldQty) As String


Dim SortCrit(1 To 2, 1 To SortCritQty) As String

Dim StationNamesGU(1 To 2, 1 To 200) As String
Dim StationNamesCamaras(1 To 2, 1 To 200) As String
Dim StationNamesCromas(1 To 2, 1 To 10) As String
Dim StationNamesVeribox(1 To 2, 1 To 50) As String

Dim bOrden1Asc As Boolean


Private Sub CFixPicture_Initialize()
 InitializeArrays
 
 cmbSortList1.Clear
 cmbSortList1.AddItem "Cromatografia"
 cmbSortList1.AddItem "Medicion"
 cmbSortList1.AddItem "Grandes Usuarios"
 cmbSortList1.AddItem "Veribox"
 cmbSortList1.ListIndex = 1
 dtFechaFinal.Value = Now()
 dtFechaInicial = dtFechaFinal.Value - CDate("1:00")
 
 
 bOrden1Asc = True
 'CargarListaCampos lbCampos, FieldNamesVeribox, VERIBOX_FIELDS
 CargarListaCampos lbCampos, FieldNamesMedicion, MED_FIELDS
 CargarListaCampos cbOrden1, SortCrit, SortCritQty
 cbOrden1.ListIndex = 0
End Sub

Private Sub CFixPicture_KeyDown(ByVal KeyCode As Long, ByVal Shift As Long, ContinueProcessing As Boolean)

End Sub






Private Function readBDValue(TAG As String) As String
    On Error GoTo ErrorHandler
    readBDValue = readValue(TAG, 1)
    Exit Function
ErrorHandler:
    readBDValue = "_"
End Function



Private Sub btnConsultar_Click()
Dim sql As String
    sql = GetCurrentSQLString
    If sql <> NO_QUERY Then
        
        If cmbSortList1.Value = "Veribox" Then
            cargarDatosPostgres2 GetCurrentSQLString
        Else
            cargarDatosPostgres GetCurrentSQLString
        End If
    End If
End Sub









Private Sub cmbSortList1_Change()
    refrescarCombo
End Sub


Private Sub btnExportar_Click()
    exportGridData
End Sub



Private Sub dtFechaFinal_Change()
'Dim diferencia As Long

'diferencia = DateDiff("d", dtFechaInicial.value, dtFechaFinal.value)

'If diferencia >= 30 Or diferencia <= 0 Then
    
'    dtFechaInicial.value = DateDiff("d", 30, dtFechaFinal.value)
    'MsgBox "El intervalo máximo de visualización es de 30 días.", , "Error"
    
'End If

End Sub

Private Sub cbOrden1_Change()

End Sub



Private Sub dtFechaInicial_Change()

'Dim diferencia As Long'

'diferencia = DateDiff("d", dtFechaInicial.value, dtFechaFinal.value)

'If diferencia >= 30 Or diferencia <= 0 Then
    
'    dtFechaFinal.value = DateAdd("d", 30, dtFechaInicial.value)
    
    'MsgBox "El intervalo máximo de visualización es de 30 días", , "Error"
    
'End If

End Sub





Public Function cargarDatosPostgres(strSQL As String)

    Dim cnn As New ADODB.Connection
    Dim rst As New ADODB.Recordset
    Dim i, j, answer As Integer

    Dim xDate, xYear, xMonth, xDay, yDate, xResto As String

    
    
    cnn.ConnectionString = "Driver=PostgreSQL;" & _
                           "Server=" & "143.4.12.52" & ";" & _
                           "Port=" & "5434" & ";" & _
                           "User Id=" & "postgres" & ";" & _
                           "Password=" & "DAC5D9DECDD8CFD9" & ";" & _
                           "Database=" & "EcoGas_GNC" & ";"
                           
    cnn.Open
     
    answer = MsgBox("¿Sólo un dato diario?", vbYesNo, "Consulta")
    If answer = vbYes Then
        strSQL = Replace(strSQL, "Where", "WHERE date_part('hour',""DateTime"")= 7 AND ")
    End If
    
    rst.Open strSQL, cnn
    MSFlexGrid1.Clear
    MSFlexGrid1.Font.Size = 9
    MSFlexGrid1.Rows = 3
    If (cnn.State) Then
                MSFlexGrid1.FixedRows = 2
                MSFlexGrid1.FixedCols = 0
                
                If Not rst.EOF Then
                    MSFlexGrid1.Cols = rst.Fields.Count
                    
                    j = 0
                    
                    With lbCampos
                        For i = 0 To .ListCount - 1
                            If .Selected(i) Then
                                If cmbSortList1.Value = "Grandes Usuarios" Then
                                    MSFlexGrid1.TextMatrix(0, j) = FieldNamesGU(THE_DESC, i + 1)
                                    MSFlexGrid1.TextMatrix(1, j) = FieldNamesGU(THE_UNIT, i + 1)
                                    MSFlexGrid1.ColWidth(j) = CInt(FieldNamesGU(THE_WIDH, i + 1))
                                Else
                                    If cmbSortList1.Value = "Medicion" Then
                                        MSFlexGrid1.TextMatrix(0, j) = FieldNamesMedicion(THE_DESC, i + 1)
                                        MSFlexGrid1.TextMatrix(1, j) = FieldNamesMedicion(THE_UNIT, i + 1)
                                        MSFlexGrid1.ColWidth(j) = CInt(FieldNamesMedicion(THE_WIDH, i + 1))
                                    
                                        Else
                                            MSFlexGrid1.TextMatrix(0, j) = FieldNamesCromatografia(THE_DESC, i + 1)
                                            MSFlexGrid1.TextMatrix(1, j) = FieldNamesCromatografia(THE_UNIT, i + 1)
                                            MSFlexGrid1.ColWidth(j) = CInt(FieldNamesCromatografia(THE_WIDH, i + 1))
                                    End If
                                End If
                                
                                j = j + 1
                            End If
                        Next i
                    End With

                    If j = 0 Then
                        For i = 0 To lbCampos.ListCount - 1
                            If cmbSortList1.Value = "Grandes Usuarios" Then
                                MSFlexGrid1.TextMatrix(0, j) = FieldNamesGU(THE_DESC, i + 1)
                                MSFlexGrid1.TextMatrix(1, j) = FieldNamesGU(THE_UNIT, i + 1)
                                MSFlexGrid1.ColWidth(j) = CInt(FieldNamesGU(THE_WIDH, i + 1))
                            Else
                                If cmbSortList1.Value = "Medicion" Then
                                    MSFlexGrid1.TextMatrix(0, j) = FieldNamesMedicion(THE_DESC, i + 1)
                                    MSFlexGrid1.TextMatrix(1, j) = FieldNamesMedicion(THE_UNIT, i + 1)
                                    MSFlexGrid1.ColWidth(j) = CInt(FieldNamesMedicion(THE_WIDH, i + 1))
                                Else
                                    MSFlexGrid1.TextMatrix(0, j) = FieldNamesCromatografia(THE_DESC, i + 1)
                                    MSFlexGrid1.TextMatrix(1, j) = FieldNamesCromatografia(THE_UNIT, i + 1)
                                    MSFlexGrid1.ColWidth(j) = CInt(FieldNamesCromatografia(THE_WIDH, i + 1))
                                End If
                            End If
                            
                            j = j + 1
                        Next i
                    End If
                    
                    i = 2
                
                    While (Not rst.EOF)
                       MSFlexGrid1.AddItem " "
                       For j = 0 To rst.Fields.Count - 1
                            If Not IsNull(rst.Fields(j).Value) Then
                                'MsgBox (rst.Fields(j).Value)
                                ' si es fecha lo formateo
                                
                         '       If (IsDate(rst.Fields(j).Value)) Then
                         '           xDate = rst.Fields(j).Value
                          '          xDay = Left$(xDate, 2)
                          '          xMonth = Mid$(xDate, 4, 2)
                          '          xYear = Mid$(xDate, 7, 4)
                          '          xResto = Mid$(xDate, 12, 20)
                          '          yDate = xYear & "/" & xMonth & "/" & xDay & " " & xResto
                                    'MsgBox (yDate)
                                    'MsgBox (IsDate(yDate))
                         '           MSFlexGrid1.TextMatrix(i, j) = yDate
                         '       Else
                                    MSFlexGrid1.TextMatrix(i, j) = rst.Fields(j).Value
                         '       End If
                                

                            End If
                        Next
                        
                        
                    rst.MoveNext
                            
                        i = i + 1
                Wend
                MSFlexGrid1.RemoveItem i
            Else
                MSFlexGrid1.Clear
                MSFlexGrid1.Rows = 3
                MSFlexGrid1.Cols = 2
                
            End If
        rst.Close
            
    End If
    cnn.Close
   
    
End Function


Public Function cargarDatosPostgres2(strSQL As String)

    Dim cnn As New ADODB.Connection
    Dim rst As New ADODB.Recordset
    Dim i, j, answer As Integer
    'strSQL = "select datetime,stationname,vol_c,vol_nc from ucv.ucv_data order by datetime desc"
    
    'strSQL = "select t.datetime, t.stationname,t.vol_c, t.vol_nc from ucv.ucv_data as t where date_part('hour',t.datetime)= 6 and date_part('day',now()-t.datetime)<1;"
    Dim a As String
    a = """"
    strSQL = Replace(strSQL, a, " ")
    answer = MsgBox("¿Sólo datos de cierre diario?", vbYesNo, "Consulta")
    If answer = vbYes Then
        strSQL = Replace(strSQL, "Where", "WHERE date_part('hour',DateTime)= 6 AND ")
    End If
    
    
     
    cnn.ConnectionString = "Driver=PostgreSQL UNICODE;" & _
                           "Server=" & "10.2.8.30" & ";" & _
                           "Port=" & "5432" & ";" & _
                           "User Id=" & "postgres" & ";" & _
                           "Password=" & "ecopass" & ";" & _
                           "Database=" & "gas" & ";"
                           
    cnn.Open
    rst.Open strSQL, cnn
    MSFlexGrid1.Clear
    MSFlexGrid1.Font.Size = 9
    MSFlexGrid1.Rows = 3
    If (cnn.State) Then
                MSFlexGrid1.FixedRows = 2
                MSFlexGrid1.FixedCols = 0
                If Not rst.EOF Then
                    MSFlexGrid1.Cols = rst.Fields.Count
                    j = 0
                    With lbCampos
                        For i = 0 To .ListCount - 1
                            If .Selected(i) Then
                                If cmbSortList1.Value = "Grandes Usuarios" Then
                                    MSFlexGrid1.TextMatrix(0, j) = FieldNamesGU(THE_DESC, i + 1)
                                    MSFlexGrid1.TextMatrix(1, j) = FieldNamesGU(THE_UNIT, i + 1)
                                    MSFlexGrid1.ColWidth(j) = CInt(FieldNamesGU(THE_WIDH, i + 1))
                                Else
                                    If cmbSortList1.Value = "Medicion" Then
                                        MSFlexGrid1.TextMatrix(0, j) = FieldNamesMedicion(THE_DESC, i + 1)
                                        MSFlexGrid1.TextMatrix(1, j) = FieldNamesMedicion(THE_UNIT, i + 1)
                                        MSFlexGrid1.ColWidth(j) = CInt(FieldNamesMedicion(THE_WIDH, i + 1))
                                    Else
                                        If cmbSortList1.Value = "Cromatografia" Then
                                            MSFlexGrid1.TextMatrix(0, j) = FieldNamesCromatografia(THE_DESC, i + 1)
                                            MSFlexGrid1.TextMatrix(1, j) = FieldNamesCromatografia(THE_UNIT, i + 1)
                                            MSFlexGrid1.ColWidth(j) = CInt(FieldNamesCromatografia(THE_WIDH, i + 1))
                                        Else
                                            If cmbSortList1.Value = "Veribox" Then
                                                MSFlexGrid1.TextMatrix(0, j) = FieldNamesVeribox(THE_DESC, i + 1)
                                                MSFlexGrid1.TextMatrix(1, j) = FieldNamesVeribox(THE_UNIT, i + 1)
                                                MSFlexGrid1.ColWidth(j) = CInt(FieldNamesVeribox(THE_WIDH, i + 1))
                                            End If
                                        End If
                                    End If
                                End If
                                
                                j = j + 1
                            End If
                        Next i
                    End With

                    If j = 0 Then
                        For i = 0 To lbCampos.ListCount - 1
                            If cmbSortList1.Value = "Grandes Usuarios" Then
                                MSFlexGrid1.TextMatrix(0, j) = FieldNamesGU(THE_DESC, i + 1)
                                MSFlexGrid1.TextMatrix(1, j) = FieldNamesGU(THE_UNIT, i + 1)
                                MSFlexGrid1.ColWidth(j) = CInt(FieldNamesGU(THE_WIDH, i + 1))
                            Else
                                If cmbSortList1.Value = "Medicion" Then
                                    MSFlexGrid1.TextMatrix(0, j) = FieldNamesMedicion(THE_DESC, i + 1)
                                    MSFlexGrid1.TextMatrix(1, j) = FieldNamesMedicion(THE_UNIT, i + 1)
                                    MSFlexGrid1.ColWidth(j) = CInt(FieldNamesMedicion(THE_WIDH, i + 1))
                                Else
                                    If cmbSortList1.Value = "Cromatografia" Then
                                        MSFlexGrid1.TextMatrix(0, j) = FieldNamesCromatografia(THE_DESC, i + 1)
                                        MSFlexGrid1.TextMatrix(1, j) = FieldNamesCromatografia(THE_UNIT, i + 1)
                                        MSFlexGrid1.ColWidth(j) = CInt(FieldNamesCromatografia(THE_WIDH, i + 1))
                                    Else
                                    If cmbSortList1.Value = "Veribox" Then
                                        MSFlexGrid1.TextMatrix(0, j) = FieldNamesVeribox(THE_DESC, i + 1)
                                        MSFlexGrid1.TextMatrix(1, j) = FieldNamesVeribox(THE_UNIT, i + 1)
                                        MSFlexGrid1.ColWidth(j) = CInt(FieldNamesVeribox(THE_WIDH, i + 1))
                                        
                                   End If
                                    End If
                                End If
                            End If
                            
                            j = j + 1
                        Next i
                    End If
                    
                    i = 2
                
                    While (Not rst.EOF)
                       MSFlexGrid1.AddItem " "
                       For j = 0 To rst.Fields.Count - 1
                            If Not IsNull(rst.Fields(j).Value) Then
                                MSFlexGrid1.TextMatrix(i, j) = rst.Fields(j).Value
                            End If
                        Next
                        
                        
                    rst.MoveNext
                            
                        i = i + 1
                Wend
                MSFlexGrid1.RemoveItem i
            Else
                MSFlexGrid1.Clear
                MSFlexGrid1.Rows = 3
                MSFlexGrid1.Cols = 2
                
            End If
        rst.Close
            
    End If
    cnn.Close
  

End Function



Private Function refrescarCombo()
Dim i, j, maxStations As Integer
Dim auxZona, auxRTU, auxID, auxName As String
 

 
 
 lbEstaciones.Clear
 If cmbSortList1.Text = "Grandes Usuarios" Then
     CargarListaCampos lbCampos, FieldNamesGU, GU_FIELDS
     CargarListaCampos lbEstaciones, StationNamesGU, GU_STATIONS
 
 End If


 If cmbSortList1.Text = "Medicion" Then
     CargarListaCampos lbCampos, FieldNamesMedicion, MED_FIELDS
     CargarListaCampos lbEstaciones, StationNamesCamaras, CAMARAS_STATIONS
 End If
 
 If cmbSortList1.Text = "Cromatografia" Then
     CargarListaCampos lbCampos, FieldNamesCromatografia, CROMA_FIELDS
     CargarListaCampos lbEstaciones, StationNamesCromas, CROMAS_STATIONS
 End If
 
 If cmbSortList1.Text = "Veribox" Then
     CargarListaCampos lbCampos, FieldNamesVeribox, VERIBOX_FIELDS
     CargarListaCampos lbEstaciones, StationNamesVeribox, VERIBOX_STATIONS
 End If
 
 
End Function






Private Sub gComandos_Click()

End Sub

Private Sub gListaEstaciones_Click()

End Sub

Private Sub gOrdenamiento_Click()

End Sub

Private Sub gPanelControl_Click()

End Sub

Private Sub lbEstaciones_Click()

End Sub

Private Sub MSFlexGrid1_Click()

End Sub

Private Sub Orden1Asc_Click()
    ToggleSortOrder 1
End Sub

Private Sub Orden1Desc_Click()
    ToggleSortOrder 1
End Sub








Private Sub CargarListaCampos(cb As Object, ar() As String, max As Integer)
Dim i As Integer
    With cb
        .Clear
        For i = 1 To max
            .AddItem ar(THE_DESC, i)
            .List(.ListCount - 1, 1) = ar(THE_NAME, i)
        Next i
    End With
End Sub

Private Sub lbCampos_Change()

End Sub

Private Sub lbCampos_Click()

End Sub






Private Sub InitializeArrays()
Dim i As Integer
    i = 1:       FieldNamesGU(THE_NAME, i) = "DateTime":             FieldNamesGU(THE_DESC, i) = "Hora Entrada":            FieldNamesGU(THE_UNIT, i) = "":     FieldNamesGU(THE_WIDH, i) = "2400"
    i = 2:       FieldNamesGU(THE_NAME, i) = "rtu_id":               FieldNamesGU(THE_DESC, i) = "ID IFIX":                  FieldNamesGU(THE_UNIT, i) = "":     FieldNamesGU(THE_WIDH, i) = "1000"
    i = 3:       FieldNamesGU(THE_NAME, i) = "StationName":          FieldNamesGU(THE_DESC, i) = "Nombre Estación":         FieldNamesGU(THE_UNIT, i) = "":     FieldNamesGU(THE_WIDH, i) = "3000"
    i = 4:       FieldNamesGU(THE_NAME, i) = "VAcc_Today":           FieldNamesGU(THE_DESC, i) = "Vol. Acumulado Hoy":      FieldNamesGU(THE_UNIT, i) = "m3":  FieldNamesGU(THE_WIDH, i) = "1500"
    i = 5:       FieldNamesGU(THE_NAME, i) = "VAcc_Yesterday":       FieldNamesGU(THE_DESC, i) = "Vol. Acumulado Ayer":     FieldNamesGU(THE_UNIT, i) = "m3":  FieldNamesGU(THE_WIDH, i) = "1500"
    i = 6:       FieldNamesGU(THE_NAME, i) = "VAcc_PrevMonth":       FieldNamesGU(THE_DESC, i) = "Vol. Acumulado Previo":   FieldNamesGU(THE_UNIT, i) = "m3":  FieldNamesGU(THE_WIDH, i) = "1500"
    i = 7:       FieldNamesGU(THE_NAME, i) = "DailyMaxMeasure":      FieldNamesGU(THE_DESC, i) = "Medición Máxima 1":       FieldNamesGU(THE_UNIT, i) = "m3":  FieldNamesGU(THE_WIDH, i) = "1500"
    i = 8:       FieldNamesGU(THE_NAME, i) = "VHP00":                FieldNamesGU(THE_DESC, i) = "Vol. Acumulado 00":       FieldNamesGU(THE_UNIT, i) = "m3":  FieldNamesGU(THE_WIDH, i) = "1500"
    i = 9:       FieldNamesGU(THE_NAME, i) = "VHP01":                FieldNamesGU(THE_DESC, i) = "Vol. Acumulado 01":       FieldNamesGU(THE_UNIT, i) = "m3":  FieldNamesGU(THE_WIDH, i) = "1500"
    i = 10:      FieldNamesGU(THE_NAME, i) = "VHP02":                FieldNamesGU(THE_DESC, i) = "Vol. Acumulado 02":       FieldNamesGU(THE_UNIT, i) = "m3":  FieldNamesGU(THE_WIDH, i) = "1500"
    i = 11:      FieldNamesGU(THE_NAME, i) = "VHP03":                FieldNamesGU(THE_DESC, i) = "Vol. Acumulado 03":       FieldNamesGU(THE_UNIT, i) = "m3":  FieldNamesGU(THE_WIDH, i) = "1500"
    i = 12:      FieldNamesGU(THE_NAME, i) = "VHP04":                FieldNamesGU(THE_DESC, i) = "Vol. Acumulado 04":       FieldNamesGU(THE_UNIT, i) = "m3":  FieldNamesGU(THE_WIDH, i) = "1500"
    i = 13:      FieldNamesGU(THE_NAME, i) = "VHP05":                FieldNamesGU(THE_DESC, i) = "Vol. Acumulado 05":       FieldNamesGU(THE_UNIT, i) = "m3":  FieldNamesGU(THE_WIDH, i) = "1500"
    i = 14:      FieldNamesGU(THE_NAME, i) = "VHP06":                FieldNamesGU(THE_DESC, i) = "Vol. Acumulado 06":       FieldNamesGU(THE_UNIT, i) = "m3":  FieldNamesGU(THE_WIDH, i) = "1500"
    i = 15:      FieldNamesGU(THE_NAME, i) = "VHP07":                FieldNamesGU(THE_DESC, i) = "Vol. Acumulado 07":       FieldNamesGU(THE_UNIT, i) = "m3":  FieldNamesGU(THE_WIDH, i) = "1500"
    i = 16:      FieldNamesGU(THE_NAME, i) = "VHP08":                FieldNamesGU(THE_DESC, i) = "Vol. Acumulado 08":       FieldNamesGU(THE_UNIT, i) = "m3":  FieldNamesGU(THE_WIDH, i) = "1500"
    i = 17:      FieldNamesGU(THE_NAME, i) = "VHP09":                FieldNamesGU(THE_DESC, i) = "Vol. Acumulado 09":       FieldNamesGU(THE_UNIT, i) = "m3":  FieldNamesGU(THE_WIDH, i) = "1500"
    i = 18:      FieldNamesGU(THE_NAME, i) = "VHP10":                FieldNamesGU(THE_DESC, i) = "Vol. Acumulado 10":       FieldNamesGU(THE_UNIT, i) = "m3":  FieldNamesGU(THE_WIDH, i) = "1500"
    i = 19:      FieldNamesGU(THE_NAME, i) = "VHP11":                FieldNamesGU(THE_DESC, i) = "Vol. Acumulado 11":       FieldNamesGU(THE_UNIT, i) = "m3":  FieldNamesGU(THE_WIDH, i) = "1500"
    i = 20:      FieldNamesGU(THE_NAME, i) = "VHP12":                FieldNamesGU(THE_DESC, i) = "Vol. Acumulado 12":       FieldNamesGU(THE_UNIT, i) = "m3":  FieldNamesGU(THE_WIDH, i) = "1500"
    i = 21:      FieldNamesGU(THE_NAME, i) = "VHP13":                FieldNamesGU(THE_DESC, i) = "Vol. Acumulado 13":       FieldNamesGU(THE_UNIT, i) = "m3":  FieldNamesGU(THE_WIDH, i) = "1500"
    i = 22:      FieldNamesGU(THE_NAME, i) = "VHP14":                FieldNamesGU(THE_DESC, i) = "Vol. Acumulado 14":       FieldNamesGU(THE_UNIT, i) = "m3":  FieldNamesGU(THE_WIDH, i) = "1500"
    i = 23:      FieldNamesGU(THE_NAME, i) = "VHP15":                FieldNamesGU(THE_DESC, i) = "Vol. Acumulado 15":       FieldNamesGU(THE_UNIT, i) = "m3":  FieldNamesGU(THE_WIDH, i) = "1500"
    i = 24:      FieldNamesGU(THE_NAME, i) = "VHP16":                FieldNamesGU(THE_DESC, i) = "Vol. Acumulado 16":       FieldNamesGU(THE_UNIT, i) = "m3":  FieldNamesGU(THE_WIDH, i) = "1500"
    i = 25:      FieldNamesGU(THE_NAME, i) = "VHP17":                FieldNamesGU(THE_DESC, i) = "Vol. Acumulado 17":       FieldNamesGU(THE_UNIT, i) = "m3":  FieldNamesGU(THE_WIDH, i) = "1500"
    i = 26:      FieldNamesGU(THE_NAME, i) = "VHP18":                FieldNamesGU(THE_DESC, i) = "Vol. Acumulado 18":       FieldNamesGU(THE_UNIT, i) = "m3":  FieldNamesGU(THE_WIDH, i) = "1500"
    i = 27:      FieldNamesGU(THE_NAME, i) = "VHP19":                FieldNamesGU(THE_DESC, i) = "Vol. Acumulado 19":       FieldNamesGU(THE_UNIT, i) = "m3":  FieldNamesGU(THE_WIDH, i) = "1500"
    i = 28:      FieldNamesGU(THE_NAME, i) = "VHP20":                FieldNamesGU(THE_DESC, i) = "Vol. Acumulado 20":       FieldNamesGU(THE_UNIT, i) = "m3":  FieldNamesGU(THE_WIDH, i) = "1500"
    i = 29:      FieldNamesGU(THE_NAME, i) = "VHP21":                FieldNamesGU(THE_DESC, i) = "Vol. Acumulado 21":       FieldNamesGU(THE_UNIT, i) = "m3":  FieldNamesGU(THE_WIDH, i) = "1500"
    i = 30:      FieldNamesGU(THE_NAME, i) = "VHP22":                FieldNamesGU(THE_DESC, i) = "Vol. Acumulado 22":       FieldNamesGU(THE_UNIT, i) = "m3":  FieldNamesGU(THE_WIDH, i) = "1500"
    i = 31:      FieldNamesGU(THE_NAME, i) = "VHP23":                FieldNamesGU(THE_DESC, i) = "Vol. Acumulado 23":       FieldNamesGU(THE_UNIT, i) = "m3":  FieldNamesGU(THE_WIDH, i) = "1500"
    i = 32:      FieldNamesGU(THE_NAME, i) = "VAcc_ThisMonth":       FieldNamesGU(THE_DESC, i) = "Vol. Acumulado Mes":      FieldNamesGU(THE_UNIT, i) = "m3":  FieldNamesGU(THE_WIDH, i) = "1500"
    i = 33:      FieldNamesGU(THE_NAME, i) = "DailyMaxMeasure2":     FieldNamesGU(THE_DESC, i) = "Medicion Máxima 2":       FieldNamesGU(THE_UNIT, i) = "m3":  FieldNamesGU(THE_WIDH, i) = "1500"
    
    
    i = 1:       FieldNamesMedicion(THE_NAME, i) = "DateTime":             FieldNamesMedicion(THE_DESC, i) = "Hora Entrada":        FieldNamesMedicion(THE_UNIT, i) = "":              FieldNamesMedicion(THE_WIDH, i) = "2400"
    i = 2:       FieldNamesMedicion(THE_NAME, i) = "StationName":          FieldNamesMedicion(THE_DESC, i) = "Nombre Estacion":     FieldNamesMedicion(THE_UNIT, i) = "":              FieldNamesMedicion(THE_WIDH, i) = "3000"
    i = 3:       FieldNamesMedicion(THE_NAME, i) = "rtu_id":               FieldNamesMedicion(THE_DESC, i) = "ID RTU":              FieldNamesMedicion(THE_UNIT, i) = "":              FieldNamesMedicion(THE_WIDH, i) = "1000"
    i = 4:       FieldNamesMedicion(THE_NAME, i) = "PEnt":                 FieldNamesMedicion(THE_DESC, i) = "Pres. Entrada":       FieldNamesMedicion(THE_UNIT, i) = "bar":        FieldNamesMedicion(THE_WIDH, i) = "1500"
    i = 5:       FieldNamesMedicion(THE_NAME, i) = "PSal":                 FieldNamesMedicion(THE_DESC, i) = "Pres. Salida":        FieldNamesMedicion(THE_UNIT, i) = "bar":        FieldNamesMedicion(THE_WIDH, i) = "1500"
    i = 6:       FieldNamesMedicion(THE_NAME, i) = "Pmed_1":               FieldNamesMedicion(THE_DESC, i) = "Pres. Medicion 1":    FieldNamesMedicion(THE_UNIT, i) = "bar":        FieldNamesMedicion(THE_WIDH, i) = "1500"
    i = 7:       FieldNamesMedicion(THE_NAME, i) = "Tmed_1":               FieldNamesMedicion(THE_DESC, i) = "Temp. Medicion 1":    FieldNamesMedicion(THE_UNIT, i) = "°C":            FieldNamesMedicion(THE_WIDH, i) = "1500"
    i = 8:       FieldNamesMedicion(THE_NAME, i) = "Qinst_1":              FieldNamesMedicion(THE_DESC, i) = "Caudal 1":            FieldNamesMedicion(THE_UNIT, i) = "Mm3/h":         FieldNamesMedicion(THE_WIDH, i) = "1500"
    i = 9:       FieldNamesMedicion(THE_NAME, i) = "AcDia_1":              FieldNamesMedicion(THE_DESC, i) = "Acumulado 1":         FieldNamesMedicion(THE_UNIT, i) = "Mm3":           FieldNamesMedicion(THE_WIDH, i) = "1500"
    i = 10:      FieldNamesMedicion(THE_NAME, i) = "AcAnt_1":              FieldNamesMedicion(THE_DESC, i) = "Ac. Anterior 1":      FieldNamesMedicion(THE_UNIT, i) = "Mm3":           FieldNamesMedicion(THE_WIDH, i) = "1500"
    i = 11:      FieldNamesMedicion(THE_NAME, i) = "Pmed_2":               FieldNamesMedicion(THE_DESC, i) = "Pres. Medicion 2":    FieldNamesMedicion(THE_UNIT, i) = "bar":        FieldNamesMedicion(THE_WIDH, i) = "1500"
    i = 12:      FieldNamesMedicion(THE_NAME, i) = "Tmed_2":               FieldNamesMedicion(THE_DESC, i) = "Temp. Medicion 2":    FieldNamesMedicion(THE_UNIT, i) = "°C":            FieldNamesMedicion(THE_WIDH, i) = "1500"
    i = 13:      FieldNamesMedicion(THE_NAME, i) = "Qinst_2":              FieldNamesMedicion(THE_DESC, i) = "Caudal 2":            FieldNamesMedicion(THE_UNIT, i) = "Mm3/h":         FieldNamesMedicion(THE_WIDH, i) = "1500"
    i = 14:      FieldNamesMedicion(THE_NAME, i) = "AcDia_2":              FieldNamesMedicion(THE_DESC, i) = "Acumulado 2":         FieldNamesMedicion(THE_UNIT, i) = "Mm3":           FieldNamesMedicion(THE_WIDH, i) = "1500"
    i = 15:      FieldNamesMedicion(THE_NAME, i) = "AcAnt_2":              FieldNamesMedicion(THE_DESC, i) = "Ac. Anterior 2":      FieldNamesMedicion(THE_UNIT, i) = "Mm3":           FieldNamesMedicion(THE_WIDH, i) = "1500"
    i = 16:      FieldNamesMedicion(THE_NAME, i) = "Pmed_3":               FieldNamesMedicion(THE_DESC, i) = "Pres. Medicion 3":    FieldNamesMedicion(THE_UNIT, i) = "bar":        FieldNamesMedicion(THE_WIDH, i) = "1500"
    i = 17:      FieldNamesMedicion(THE_NAME, i) = "Tmed_3":               FieldNamesMedicion(THE_DESC, i) = "Temp. Medicion 3":    FieldNamesMedicion(THE_UNIT, i) = "°C":            FieldNamesMedicion(THE_WIDH, i) = "1500"
    i = 18:      FieldNamesMedicion(THE_NAME, i) = "Qinst_3":              FieldNamesMedicion(THE_DESC, i) = "Caudal 3":            FieldNamesMedicion(THE_UNIT, i) = "Mm3/h":         FieldNamesMedicion(THE_WIDH, i) = "1500"
    i = 19:      FieldNamesMedicion(THE_NAME, i) = "AcDia_3":              FieldNamesMedicion(THE_DESC, i) = "Acumulado 3":         FieldNamesMedicion(THE_UNIT, i) = "Mm3":           FieldNamesMedicion(THE_WIDH, i) = "1500"
    i = 20:      FieldNamesMedicion(THE_NAME, i) = "AcAnt_3":              FieldNamesMedicion(THE_DESC, i) = "Ac. Anterior 3":      FieldNamesMedicion(THE_UNIT, i) = "Mm3":           FieldNamesMedicion(THE_WIDH, i) = "1500"
    
    
    i = 1:       FieldNamesCromatografia(THE_NAME, i) = "DateTime":            FieldNamesCromatografia(THE_DESC, i) = "Hora Entrada":       FieldNamesCromatografia(THE_UNIT, i) = "":              FieldNamesCromatografia(THE_WIDH, i) = "2400"
    i = 2:       FieldNamesCromatografia(THE_NAME, i) = "StationName":         FieldNamesCromatografia(THE_DESC, i) = "Nombre Estacion":    FieldNamesCromatografia(THE_UNIT, i) = "":              FieldNamesCromatografia(THE_WIDH, i) = "3000"
    i = 3:       FieldNamesCromatografia(THE_NAME, i) = "rtu_id":              FieldNamesCromatografia(THE_DESC, i) = "ID RTU":             FieldNamesCromatografia(THE_UNIT, i) = "":              FieldNamesCromatografia(THE_WIDH, i) = "1000"
    i = 4:       FieldNamesCromatografia(THE_NAME, i) = "DensR":               FieldNamesCromatografia(THE_DESC, i) = "Dens. relativa":     FieldNamesCromatografia(THE_UNIT, i) = "sgu":           FieldNamesCromatografia(THE_WIDH, i) = "1500"
    i = 5:       FieldNamesCromatografia(THE_NAME, i) = "N2":                  FieldNamesCromatografia(THE_DESC, i) = "Nitrogeno":          FieldNamesCromatografia(THE_UNIT, i) = "%mol":          FieldNamesCromatografia(THE_WIDH, i) = "1500"
    i = 6:       FieldNamesCromatografia(THE_NAME, i) = "CO2":                 FieldNamesCromatografia(THE_DESC, i) = "Dioxido de Carbono": FieldNamesCromatografia(THE_UNIT, i) = "%mol":          FieldNamesCromatografia(THE_WIDH, i) = "1500"
    i = 7:       FieldNamesCromatografia(THE_NAME, i) = "C1":                  FieldNamesCromatografia(THE_DESC, i) = "Metano":             FieldNamesCromatografia(THE_UNIT, i) = "%mol":          FieldNamesCromatografia(THE_WIDH, i) = "1500"
    i = 8:       FieldNamesCromatografia(THE_NAME, i) = "C2":                  FieldNamesCromatografia(THE_DESC, i) = "Etano":              FieldNamesCromatografia(THE_UNIT, i) = "%mol":          FieldNamesCromatografia(THE_WIDH, i) = "1500"
    i = 9:       FieldNamesCromatografia(THE_NAME, i) = "C3":                  FieldNamesCromatografia(THE_DESC, i) = "Propano":            FieldNamesCromatografia(THE_UNIT, i) = "%mol":          FieldNamesCromatografia(THE_WIDH, i) = "1500"
    i = 10:      FieldNamesCromatografia(THE_NAME, i) = "iC4":                 FieldNamesCromatografia(THE_DESC, i) = "i-Butano":           FieldNamesCromatografia(THE_UNIT, i) = "%mol":          FieldNamesCromatografia(THE_WIDH, i) = "1500"
    i = 11:      FieldNamesCromatografia(THE_NAME, i) = "nC4":                 FieldNamesCromatografia(THE_DESC, i) = "n-Butano":           FieldNamesCromatografia(THE_UNIT, i) = "%mol":          FieldNamesCromatografia(THE_WIDH, i) = "1500"
    i = 12:      FieldNamesCromatografia(THE_NAME, i) = "neo-C5":                 FieldNamesCromatografia(THE_DESC, i) = "Hexano+":        FieldNamesCromatografia(THE_UNIT, i) = "%mol":          FieldNamesCromatografia(THE_WIDH, i) = "1500"
    i = 13:      FieldNamesCromatografia(THE_NAME, i) = "i-C5":                FieldNamesCromatografia(THE_DESC, i) = "i-Pentano":          FieldNamesCromatografia(THE_UNIT, i) = "%mol":          FieldNamesCromatografia(THE_WIDH, i) = "1500"
    i = 14:      FieldNamesCromatografia(THE_NAME, i) = "n-C5":                FieldNamesCromatografia(THE_DESC, i) = "n-Pentano":          FieldNamesCromatografia(THE_UNIT, i) = "%mol":          FieldNamesCromatografia(THE_WIDH, i) = "1500"
    i = 15:      FieldNamesCromatografia(THE_NAME, i) = "n-C6":                FieldNamesCromatografia(THE_DESC, i) = "n-Hexano":           FieldNamesCromatografia(THE_UNIT, i) = "%mol":          FieldNamesCromatografia(THE_WIDH, i) = "1500"
    i = 16:      FieldNamesCromatografia(THE_NAME, i) = "n-C7":                FieldNamesCromatografia(THE_DESC, i) = "n-Heptano":          FieldNamesCromatografia(THE_UNIT, i) = "%mol":          FieldNamesCromatografia(THE_WIDH, i) = "1500"
    i = 17:      FieldNamesCromatografia(THE_NAME, i) = "n-C8":                FieldNamesCromatografia(THE_DESC, i) = "n-Octano":           FieldNamesCromatografia(THE_UNIT, i) = "%mol":          FieldNamesCromatografia(THE_WIDH, i) = "1500"
    i = 18:       FieldNamesCromatografia(THE_NAME, i) = "Pcal":                FieldNamesCromatografia(THE_DESC, i) = "Poder Calorifico":   FieldNamesCromatografia(THE_UNIT, i) = "kcal/m3":       FieldNamesCromatografia(THE_WIDH, i) = "1500"
    
    i = 1:       FieldNamesVeribox(THE_NAME, i) = "DateTime":                 FieldNamesVeribox(THE_DESC, i) = "Hora Entrada":          FieldNamesVeribox(THE_UNIT, i) = "":                FieldNamesVeribox(THE_WIDH, i) = "2400"
    i = 2:       FieldNamesVeribox(THE_NAME, i) = "StationName":              FieldNamesVeribox(THE_DESC, i) = "Nombre Estacion":       FieldNamesVeribox(THE_UNIT, i) = "":                FieldNamesVeribox(THE_WIDH, i) = "3000"
    i = 3:       FieldNamesVeribox(THE_NAME, i) = "pca":                   FieldNamesVeribox(THE_DESC, i) = "pca":                FieldNamesVeribox(THE_UNIT, i) = "":              FieldNamesVeribox(THE_WIDH, i) = "800"
    i = 4:       FieldNamesVeribox(THE_NAME, i) = "vol_c":                    FieldNamesVeribox(THE_DESC, i) = "vol_c":                 FieldNamesVeribox(THE_UNIT, i) = "m3":              FieldNamesVeribox(THE_WIDH, i) = "1500"
    i = 5:       FieldNamesVeribox(THE_NAME, i) = "vol_c_an":                   FieldNamesVeribox(THE_DESC, i) = "vol_c_an":                FieldNamesVeribox(THE_UNIT, i) = "m3":              FieldNamesVeribox(THE_WIDH, i) = "1000"
    i = 6:       FieldNamesVeribox(THE_NAME, i) = "vol_nc":                    FieldNamesVeribox(THE_DESC, i) = "vol_nc":                 FieldNamesVeribox(THE_UNIT, i) = "m3":              FieldNamesVeribox(THE_WIDH, i) = "1500"
    i = 7:       FieldNamesVeribox(THE_NAME, i) = "vol_nc_an":                   FieldNamesVeribox(THE_DESC, i) = "vol_nc_an":                FieldNamesVeribox(THE_UNIT, i) = "m3":              FieldNamesVeribox(THE_WIDH, i) = "1000"
    i = 8:       FieldNamesVeribox(THE_NAME, i) = "presion":                   FieldNamesVeribox(THE_DESC, i) = "presion":                FieldNamesVeribox(THE_UNIT, i) = "bar":              FieldNamesVeribox(THE_WIDH, i) = "1000"
    i = 9:       FieldNamesVeribox(THE_NAME, i) = "temperatura":                   FieldNamesVeribox(THE_DESC, i) = "temperatura":                FieldNamesVeribox(THE_UNIT, i) = "°C":              FieldNamesVeribox(THE_WIDH, i) = "1000"
    i = 10:       FieldNamesVeribox(THE_NAME, i) = "batt_vc":                   FieldNamesVeribox(THE_DESC, i) = "batt_vc":                FieldNamesVeribox(THE_UNIT, i) = "V":              FieldNamesVeribox(THE_WIDH, i) = "1000"
    i = 11:       FieldNamesVeribox(THE_NAME, i) = "batt_vb":                   FieldNamesVeribox(THE_DESC, i) = "batt_vb":                FieldNamesVeribox(THE_UNIT, i) = "mV":              FieldNamesVeribox(THE_WIDH, i) = "1000"
    i = 12:       FieldNamesVeribox(THE_NAME, i) = "alarma":                   FieldNamesVeribox(THE_DESC, i) = "Alarma":                FieldNamesVeribox(THE_UNIT, i) = "":              FieldNamesVeribox(THE_WIDH, i) = "1000"
    i = 13:       FieldNamesVeribox(THE_NAME, i) = "gsmq":                   FieldNamesVeribox(THE_DESC, i) = "nivel_señal":                FieldNamesVeribox(THE_UNIT, i) = "":              FieldNamesVeribox(THE_WIDH, i) = "1000"
   
    
    
    i = 1:       SortCrit(THE_NAME, i) = "DateTime":                SortCrit(THE_DESC, i) = "HoraEntrada"
    i = 2:       SortCrit(THE_NAME, i) = "StationName":             SortCrit(THE_DESC, i) = "Nombre de la Estacion"
    i = 3:       SortCrit(THE_NAME, i) = "rtu_id":                  SortCrit(THE_DESC, i) = "ID RTU"
    
    
    
'//////////////ACA ARRANCAN LOS DE GPRS///////////////////////////////////////////////////////////////////////
i = 1:  StationNamesGU(THE_NAME, i) = "Acabio Etanol-VillaMaría":         StationNamesGU(THE_DESC, i) = "Acabio Etanol-VillaMaría"
i = 2:  StationNamesGU(THE_NAME, i) = "ACCESO SUR SRL 851821":         StationNamesGU(THE_DESC, i) = "ACCESO SUR SRL 851821"
i = 3:  StationNamesGU(THE_NAME, i) = "AGAF S.A.                808968":         StationNamesGU(THE_DESC, i) = "AGAF S.A.                808968"
i = 4:  StationNamesGU(THE_NAME, i) = "Agroalimentos V. María":         StationNamesGU(THE_DESC, i) = "Agroalimentos V. María"
i = 5:  StationNamesGU(THE_NAME, i) = "AgroAndina - La Rioja":         StationNamesGU(THE_DESC, i) = "AgroAndina - La Rioja"
i = 6:  StationNamesGU(THE_NAME, i) = "ALBERDI  Carlos Paz 786172":         StationNamesGU(THE_DESC, i) = "ALBERDI  Carlos Paz 786172"
i = 7:  StationNamesGU(THE_NAME, i) = "Alberdi 2 Rio4°Ponce  779210":         StationNamesGU(THE_DESC, i) = "Alberdi 2 Rio4°Ponce  779210"
i = 8:  StationNamesGU(THE_NAME, i) = "ALFA SUR II  SRL      781660":         StationNamesGU(THE_DESC, i) = "ALFA SUR II  SRL      781660"
i = 9:  StationNamesGU(THE_NAME, i) = "Alican S. A. - Alcira Gigena":         StationNamesGU(THE_DESC, i) = "Alican S. A. - Alcira Gigena"
i = 10:  StationNamesGU(THE_NAME, i) = "ALT Colonia Caroya    810956":         StationNamesGU(THE_DESC, i) = "ALT Colonia Caroya    810956"
i = 11:  StationNamesGU(THE_NAME, i) = "ALT Villa General Belgrano":         StationNamesGU(THE_DESC, i) = "ALT Villa General Belgrano"
i = 12:  StationNamesGU(THE_NAME, i) = "ALTOS del VALLE     406044":         StationNamesGU(THE_DESC, i) = "ALTOS del VALLE     406044"
i = 13:  StationNamesGU(THE_NAME, i) = "AM Emprendimientos  815922":         StationNamesGU(THE_DESC, i) = "AM Emprendimientos  815922"
i = 14:  StationNamesGU(THE_NAME, i) = "Ausonia Conecar":         StationNamesGU(THE_DESC, i) = "Ausonia Conecar"
i = 15:  StationNamesGU(THE_NAME, i) = "AUTOGAS III Cat.       840301":         StationNamesGU(THE_DESC, i) = "AUTOGAS III Cat.       840301"
i = 16:  StationNamesGU(THE_NAME, i) = "Barale y GUIO II      897838":         StationNamesGU(THE_DESC, i) = "Barale y GUIO II      897838"
i = 17:  StationNamesGU(THE_NAME, i) = "BARALEyGUIO + d Eje   822163":         StationNamesGU(THE_DESC, i) = "BARALEyGUIO + d Eje   822163"
i = 18:  StationNamesGU(THE_NAME, i) = "Benzin SRL            943414":         StationNamesGU(THE_DESC, i) = "Benzin SRL            943414"
i = 19:  StationNamesGU(THE_NAME, i) = "BIFUEL(Av Richeri)    742895":         StationNamesGU(THE_DESC, i) = "BIFUEL(Av Richeri)    742895"
i = 20:  StationNamesGU(THE_NAME, i) = "Bio 4":         StationNamesGU(THE_DESC, i) = "Bio 4"
i = 21:  StationNamesGU(THE_NAME, i) = "BioFarma - Río IV":         StationNamesGU(THE_DESC, i) = "BioFarma - Río IV"
i = 22:  StationNamesGU(THE_NAME, i) = "Bornoromi               857610":         StationNamesGU(THE_DESC, i) = "Bornoromi               857610"
i = 23:  StationNamesGU(THE_NAME, i) = "CENTRO BEBIDAS    805774":         StationNamesGU(THE_DESC, i) = "CENTRO BEBIDAS    805774"
i = 24:  StationNamesGU(THE_NAME, i) = "Cerámica Nanini - C. Caroya":         StationNamesGU(THE_DESC, i) = "Cerámica Nanini - C. Caroya"
i = 25:  StationNamesGU(THE_NAME, i) = "Cervecería Quilmes  887071":         StationNamesGU(THE_DESC, i) = "Cervecería Quilmes  887071"
i = 26:  StationNamesGU(THE_NAME, i) = "CHICOM S.A.            778409":         StationNamesGU(THE_DESC, i) = "CHICOM S.A.            778409"
i = 27:  StationNamesGU(THE_NAME, i) = "Coca Cola":         StationNamesGU(THE_DESC, i) = "Coca Cola"
i = 28:  StationNamesGU(THE_NAME, i) = "COFICO S.R.L.          802153":         StationNamesGU(THE_DESC, i) = "COFICO S.R.L.          802153"
i = 29:  StationNamesGU(THE_NAME, i) = "Combustibles Serranos S. A. - La Falda":         StationNamesGU(THE_DESC, i) = "Combustibles Serranos S. A. - La Falda"
i = 30:  StationNamesGU(THE_NAME, i) = "Coop Gral. Paz - Los Surgentes":         StationNamesGU(THE_DESC, i) = "Coop Gral. Paz - Los Surgentes"
i = 31:  StationNamesGU(THE_NAME, i) = "Coop P.del MOLLE   757320":         StationNamesGU(THE_DESC, i) = "Coop P.del MOLLE   757320"
i = 32:  StationNamesGU(THE_NAME, i) = "Coop. Los Condores 892552":         StationNamesGU(THE_DESC, i) = "Coop. Los Condores 892552"
i = 33:  StationNamesGU(THE_NAME, i) = "Corgas SA":         StationNamesGU(THE_DESC, i) = "Corgas SA"
i = 34:  StationNamesGU(THE_NAME, i) = "Cotagro CAL            923636":         StationNamesGU(THE_DESC, i) = "Cotagro CAL            923636"
i = 35:  StationNamesGU(THE_NAME, i) = "DEL VALLE COMUNIC.  797880":         StationNamesGU(THE_DESC, i) = "DEL VALLE COMUNIC.  797880"
i = 36:  StationNamesGU(THE_NAME, i) = "Denso M. Argentina":         StationNamesGU(THE_DESC, i) = "Denso M. Argentina"
i = 37:  StationNamesGU(THE_NAME, i) = "Don Bosco SRL       879634":         StationNamesGU(THE_DESC, i) = "Don Bosco SRL       879634"
i = 38:  StationNamesGU(THE_NAME, i) = "EL CORRALITO         751828":         StationNamesGU(THE_DESC, i) = "EL CORRALITO         751828"
i = 39:  StationNamesGU(THE_NAME, i) = "El Portal                693165":         StationNamesGU(THE_DESC, i) = "El Portal                693165"
i = 40:  StationNamesGU(THE_NAME, i) = "EL PRACTICO           805786":         StationNamesGU(THE_DESC, i) = "EL PRACTICO           805786"
i = 41:  StationNamesGU(THE_NAME, i) = "ESER S.A.               758890":         StationNamesGU(THE_DESC, i) = "ESER S.A.               758890"
i = 42:  StationNamesGU(THE_NAME, i) = "ESSO El Arco           765437":         StationNamesGU(THE_DESC, i) = "ESSO El Arco           765437"
i = 43:  StationNamesGU(THE_NAME, i) = "ESSO Petrole Bulnes 774027":         StationNamesGU(THE_DESC, i) = "ESSO Petrole Bulnes 774027"
i = 44:  StationNamesGU(THE_NAME, i) = "ESSO PETROLERA   798123":         StationNamesGU(THE_DESC, i) = "ESSO PETROLERA   798123"
i = 45:  StationNamesGU(THE_NAME, i) = "ESSO Petrolera Arg.  522800":         StationNamesGU(THE_DESC, i) = "ESSO Petrolera Arg.  522800"
i = 46:  StationNamesGU(THE_NAME, i) = "Estación Ferreyra SRL":         StationNamesGU(THE_DESC, i) = "Estación Ferreyra SRL"
i = 47:  StationNamesGU(THE_NAME, i) = "Ex-GU     MONTICH              135":         StationNamesGU(THE_DESC, i) = "Ex-GU     MONTICH              135"
i = 48:  StationNamesGU(THE_NAME, i) = "Ex-OMEGA          430415":         StationNamesGU(THE_DESC, i) = "Ex-OMEGA          430415"
i = 49:  StationNamesGU(THE_NAME, i) = "FATIMA S.R.L.          760223":         StationNamesGU(THE_DESC, i) = "FATIMA S.R.L.          760223"
i = 50:  StationNamesGU(THE_NAME, i) = "Firenze":         StationNamesGU(THE_DESC, i) = "Firenze"
i = 51:  StationNamesGU(THE_NAME, i) = "FluidConex GNC - La Rioja":         StationNamesGU(THE_DESC, i) = "FluidConex GNC - La Rioja"
i = 52:  StationNamesGU(THE_NAME, i) = "Franco Fabril - Las Arias":         StationNamesGU(THE_DESC, i) = "Franco Fabril - Las Arias"
i = 53:  StationNamesGU(THE_NAME, i) = "Frigorífico Transur":         StationNamesGU(THE_DESC, i) = "Frigorífico Transur"
i = 54:  StationNamesGU(THE_NAME, i) = "Frucor":         StationNamesGU(THE_DESC, i) = "Frucor"
i = 55:  StationNamesGU(THE_NAME, i) = "GALACTIC S.A.       846648":         StationNamesGU(THE_DESC, i) = "GALACTIC S.A.       846648"
i = 56:  StationNamesGU(THE_NAME, i) = "Gas A 1°                   745097":         StationNamesGU(THE_DESC, i) = "Gas A 1°                   745097"
i = 57:  StationNamesGU(THE_NAME, i) = "Gas A 2°                   783050":         StationNamesGU(THE_DESC, i) = "Gas A 2°                   783050"
i = 58:  StationNamesGU(THE_NAME, i) = "Gas A° 3":         StationNamesGU(THE_DESC, i) = "Gas A° 3"
i = 59:  StationNamesGU(THE_NAME, i) = "Gecor Rio III            927445":         StationNamesGU(THE_DESC, i) = "Gecor Rio III            927445"
i = 60:  StationNamesGU(THE_NAME, i) = "Generadora Cba. Río III":         StationNamesGU(THE_DESC, i) = "Generadora Cba. Río III"
i = 61:  StationNamesGU(THE_NAME, i) = "GENERAL ROCA     847439":         StationNamesGU(THE_DESC, i) = "GENERAL ROCA     847439"
i = 62:  StationNamesGU(THE_NAME, i) = "General Roca GNC SRL":         StationNamesGU(THE_DESC, i) = "General Roca GNC SRL"
i = 63:  StationNamesGU(THE_NAME, i) = "GNC - Orbilex S.A.":         StationNamesGU(THE_DESC, i) = "GNC - Orbilex S.A."
i = 64:  StationNamesGU(THE_NAME, i) = "GNC - Vesinm Capdevila":         StationNamesGU(THE_DESC, i) = "GNC - Vesinm Capdevila"
i = 65:  StationNamesGU(THE_NAME, i) = "GNC Alcano S. R. L.":         StationNamesGU(THE_DESC, i) = "GNC Alcano S. R. L."
i = 66:  StationNamesGU(THE_NAME, i) = "GNC ALT Carlos Paz":         StationNamesGU(THE_DESC, i) = "GNC ALT Carlos Paz"
i = 67:  StationNamesGU(THE_NAME, i) = "GNC Alt S.A.- Alta Gracia":         StationNamesGU(THE_DESC, i) = "GNC Alt S.A.- Alta Gracia"
i = 68:  StationNamesGU(THE_NAME, i) = "GNC Altos Del Valle - Catamarca":         StationNamesGU(THE_DESC, i) = "GNC Altos Del Valle - Catamarca"
i = 69:  StationNamesGU(THE_NAME, i) = "GNC AM Emprendimmientos S.A.- Valparaíso":         StationNamesGU(THE_DESC, i) = "GNC AM Emprendimmientos S.A.- Valparaíso"
i = 70:  StationNamesGU(THE_NAME, i) = "GNC Anjor S. A. Martinolli":         StationNamesGU(THE_DESC, i) = "GNC Anjor S. A. Martinolli"
i = 71:  StationNamesGU(THE_NAME, i) = "GNC Anjor S. A. Rancagua":         StationNamesGU(THE_DESC, i) = "GNC Anjor S. A. Rancagua"
i = 72:  StationNamesGU(THE_NAME, i) = "GNC Armar SRL - A. Argentina 1860":         StationNamesGU(THE_DESC, i) = "GNC Armar SRL - A. Argentina 1860"
i = 73:  StationNamesGU(THE_NAME, i) = "GNC Arroyo de los Patos":         StationNamesGU(THE_DESC, i) = "GNC Arroyo de los Patos"
i = 74:  StationNamesGU(THE_NAME, i) = "GNC Autopista Villa Maria":         StationNamesGU(THE_DESC, i) = "GNC Autopista Villa Maria"
i = 75:  StationNamesGU(THE_NAME, i) = "GNC Barale S. A. Laguna Larga":         StationNamesGU(THE_DESC, i) = "GNC Barale S. A. Laguna Larga"
i = 76:  StationNamesGU(THE_NAME, i) = "GNC Barale SA - Villa del Lago":         StationNamesGU(THE_DESC, i) = "GNC Barale SA - Villa del Lago"
i = 77:  StationNamesGU(THE_NAME, i) = "GNC Barale y Ghio-Serrezuela 6001859890":         StationNamesGU(THE_DESC, i) = "GNC Barale y Ghio-Serrezuela 6001859890"
i = 78:  StationNamesGU(THE_NAME, i) = "GNC Benzini SRL- OHiggins":         StationNamesGU(THE_DESC, i) = "GNC Benzini SRL- OHiggins"
i = 79:  StationNamesGU(THE_NAME, i) = "GNC Bulcros SA - Saldán":         StationNamesGU(THE_DESC, i) = "GNC Bulcros SA - Saldán"
i = 80:  StationNamesGU(THE_NAME, i) = "GNC Carbonetti - C. Caroya":         StationNamesGU(THE_DESC, i) = "GNC Carbonetti - C. Caroya"
i = 81:  StationNamesGU(THE_NAME, i) = "GNC Carrizo y Cia SRL - Dean Funes":         StationNamesGU(THE_DESC, i) = "GNC Carrizo y Cia SRL - Dean Funes"
i = 82:  StationNamesGU(THE_NAME, i) = "GNC Centro de Bebidas SRL-Tanti":         StationNamesGU(THE_DESC, i) = "GNC Centro de Bebidas SRL-Tanti"
i = 83:  StationNamesGU(THE_NAME, i) = "GNC Coop. Alcira Gigena 6001833588":         StationNamesGU(THE_DESC, i) = "GNC Coop. Alcira Gigena 6001833588"
i = 84:  StationNamesGU(THE_NAME, i) = "GNC Dinkeldein SRL - RÍo IV":         StationNamesGU(THE_DESC, i) = "GNC Dinkeldein SRL - RÍo IV"
i = 85:  StationNamesGU(THE_NAME, i) = "GNC Don Antonio - Dean Funes":         StationNamesGU(THE_DESC, i) = "GNC Don Antonio - Dean Funes"
i = 86:  StationNamesGU(THE_NAME, i) = "GNC Enertrac S.A.":         StationNamesGU(THE_DESC, i) = "GNC Enertrac S.A."
i = 87:  StationNamesGU(THE_NAME, i) = "GNC Estación del Portal S. A.":         StationNamesGU(THE_DESC, i) = "GNC Estación del Portal S. A."
i = 88:  StationNamesGU(THE_NAME, i) = "GNC Estación Ferreyra II":         StationNamesGU(THE_DESC, i) = "GNC Estación Ferreyra II"
i = 89:  StationNamesGU(THE_NAME, i) = "GNC Fantini - La Puerta":         StationNamesGU(THE_DESC, i) = "GNC Fantini - La Puerta"
i = 90:  StationNamesGU(THE_NAME, i) = "GNC Fantini Balnearia 6001867425":         StationNamesGU(THE_DESC, i) = "GNC Fantini Balnearia 6001867425"
i = 91:  StationNamesGU(THE_NAME, i) = "GNC Fiorani e Hijos - Cruz del Eje":         StationNamesGU(THE_DESC, i) = "GNC Fiorani e Hijos - Cruz del Eje"
i = 92:  StationNamesGU(THE_NAME, i) = "GNC Gemaco S. A. - Río Cuarto":         StationNamesGU(THE_DESC, i) = "GNC Gemaco S. A. - Río Cuarto"
i = 93:  StationNamesGU(THE_NAME, i) = "GNC Gerena José OIL":         StationNamesGU(THE_DESC, i) = "GNC Gerena José OIL"
i = 94:  StationNamesGU(THE_NAME, i) = "GNC Gold SRL-Villa de Soto":         StationNamesGU(THE_DESC, i) = "GNC Gold SRL-Villa de Soto"
i = 95:  StationNamesGU(THE_NAME, i) = "GNC Imhoff SRL - Bialet Massé":         StationNamesGU(THE_DESC, i) = "GNC Imhoff SRL - Bialet Massé"
i = 96:  StationNamesGU(THE_NAME, i) = "GNC José Luis Jurado":         StationNamesGU(THE_DESC, i) = "GNC José Luis Jurado"
i = 97:  StationNamesGU(THE_NAME, i) = "GNC Juan Mondino e Hijos - Argüello":         StationNamesGU(THE_DESC, i) = "GNC Juan Mondino e Hijos - Argüello"
i = 98:  StationNamesGU(THE_NAME, i) = "GNC La Barranca SRL-Río IV":         StationNamesGU(THE_DESC, i) = "GNC La Barranca SRL-Río IV"
i = 99:  StationNamesGU(THE_NAME, i) = "GNC La Rotonda SAS - V. C. de América":         StationNamesGU(THE_DESC, i) = "GNC La Rotonda SAS - V. C. de América"
i = 100:  StationNamesGU(THE_NAME, i) = "GNC Las Marías I - La Rioja":         StationNamesGU(THE_DESC, i) = "GNC Las Marías I - La Rioja"
i = 101:  StationNamesGU(THE_NAME, i) = "GNC Las Marías II - La Rioja":         StationNamesGU(THE_DESC, i) = "GNC Las Marías II - La Rioja"
i = 102:  StationNamesGU(THE_NAME, i) = "GNC Las Marías III - La Rioja":         StationNamesGU(THE_DESC, i) = "GNC Las Marías III - La Rioja"
i = 103:  StationNamesGU(THE_NAME, i) = "GNC Las Marías IV - La Rioja":         StationNamesGU(THE_DESC, i) = "GNC Las Marías IV - La Rioja"
i = 104:  StationNamesGU(THE_NAME, i) = "GNC Luis Angel Pavone":         StationNamesGU(THE_DESC, i) = "GNC Luis Angel Pavone"
i = 105:  StationNamesGU(THE_NAME, i) = "GNC M. de Soldano - Los Cóndores":         StationNamesGU(THE_DESC, i) = "GNC M. de Soldano - Los Cóndores"
i = 106:  StationNamesGU(THE_NAME, i) = "GNC Marijo SA RIV":         StationNamesGU(THE_DESC, i) = "GNC Marijo SA RIV"
i = 107:  StationNamesGU(THE_NAME, i) = "GNC Monbian S. A.Oncativo":         StationNamesGU(THE_DESC, i) = "GNC Monbian S. A.Oncativo"
i = 108:  StationNamesGU(THE_NAME, i) = "GNC Muchay SRL - Catamarca 6002873526":         StationNamesGU(THE_DESC, i) = "GNC Muchay SRL - Catamarca 6002873526"
i = 109:  StationNamesGU(THE_NAME, i) = "GNC OIL Salsipuedes SRL-Río Ceballos":         StationNamesGU(THE_DESC, i) = "GNC OIL Salsipuedes SRL-Río Ceballos"
i = 110:  StationNamesGU(THE_NAME, i) = "GNC Omega":         StationNamesGU(THE_DESC, i) = "GNC Omega"
i = 111:  StationNamesGU(THE_NAME, i) = "GNC Patria S. A.":         StationNamesGU(THE_DESC, i) = "GNC Patria S. A."
i = 112:  StationNamesGU(THE_NAME, i) = "GNC Pavón - Luque":         StationNamesGU(THE_DESC, i) = "GNC Pavón - Luque"
i = 113:  StationNamesGU(THE_NAME, i) = "GNC Petrovalle - Mina Clavero":         StationNamesGU(THE_DESC, i) = "GNC Petrovalle - Mina Clavero"
i = 114:  StationNamesGU(THE_NAME, i) = "GNC Quimbaletes C Lago Tanti":         StationNamesGU(THE_DESC, i) = "GNC Quimbaletes C Lago Tanti"
i = 115:  StationNamesGU(THE_NAME, i) = "GNC Reparker - Ricardo Rojas 7405":         StationNamesGU(THE_DESC, i) = "GNC Reparker - Ricardo Rojas 7405"
i = 116:  StationNamesGU(THE_NAME, i) = "GNC Saby Hnos RIV":         StationNamesGU(THE_DESC, i) = "GNC Saby Hnos RIV"
i = 117:  StationNamesGU(THE_NAME, i) = "GNC San Cayetano-Río Ceballos":         StationNamesGU(THE_DESC, i) = "GNC San Cayetano-Río Ceballos"
i = 118:  StationNamesGU(THE_NAME, i) = "GNC Santa Ana GyG SRL":         StationNamesGU(THE_DESC, i) = "GNC Santa Ana GyG SRL"
i = 119:  StationNamesGU(THE_NAME, i) = "GNC Servicios-Falda del Carmen":         StationNamesGU(THE_DESC, i) = "GNC Servicios-Falda del Carmen"
i = 120:  StationNamesGU(THE_NAME, i) = "GNC Soluciones Ecológicas":         StationNamesGU(THE_DESC, i) = "GNC Soluciones Ecológicas"
i = 121:  StationNamesGU(THE_NAME, i) = "GNC Umpy - Catamarca":         StationNamesGU(THE_DESC, i) = "GNC Umpy - Catamarca"
i = 122:  StationNamesGU(THE_NAME, i) = "GNC Vieja Estación S. A.":         StationNamesGU(THE_DESC, i) = "GNC Vieja Estación S. A."
i = 123:  StationNamesGU(THE_NAME, i) = "GNC Yaguareté - La Falda":         StationNamesGU(THE_DESC, i) = "GNC Yaguareté - La Falda"
i = 124:  StationNamesGU(THE_NAME, i) = "GNC Yaguareté II - Cosquín":         StationNamesGU(THE_DESC, i) = "GNC Yaguareté II - Cosquín"
i = 125:  StationNamesGU(THE_NAME, i) = "GNC-Umpy 2-Catamarca":         StationNamesGU(THE_DESC, i) = "GNC-Umpy 2-Catamarca"
i = 126:  StationNamesGU(THE_NAME, i) = "Golden Peanut - Chucul":         StationNamesGU(THE_DESC, i) = "Golden Peanut - Chucul"
i = 127:  StationNamesGU(THE_NAME, i) = "Grupo Bongiovanni  Guatimozín 6001818606":         StationNamesGU(THE_DESC, i) = "Grupo Bongiovanni  Guatimozín 6001818606"
i = 128:  StationNamesGU(THE_NAME, i) = "Grupo Cavigliasso - Gral. Cabrera":         StationNamesGU(THE_DESC, i) = "Grupo Cavigliasso - Gral. Cabrera"
i = 129:  StationNamesGU(THE_NAME, i) = "GUISI SA                   787762":         StationNamesGU(THE_DESC, i) = "GUISI SA                   787762"
i = 130:  StationNamesGU(THE_NAME, i) = "Helacor S.A.           772994":         StationNamesGU(THE_DESC, i) = "Helacor S.A.           772994"
i = 131:  StationNamesGU(THE_NAME, i) = "HUGO LOPEZ VM     746962":         StationNamesGU(THE_DESC, i) = "HUGO LOPEZ VM     746962"
i = 132:  StationNamesGU(THE_NAME, i) = "Indacor - Montecristo":         StationNamesGU(THE_DESC, i) = "Indacor - Montecristo"
i = 133:  StationNamesGU(THE_NAME, i) = "Indelma SA             923637":         StationNamesGU(THE_DESC, i) = "Indelma SA             923637"
i = 134:  StationNamesGU(THE_NAME, i) = "Josa S. A. - General Cabrera":         StationNamesGU(THE_DESC, i) = "Josa S. A. - General Cabrera"
i = 135:  StationNamesGU(THE_NAME, i) = "Juan y Leonardo Sio SH":         StationNamesGU(THE_DESC, i) = "Juan y Leonardo Sio SH"
i = 136:  StationNamesGU(THE_NAME, i) = "La Arboleda de Mendiolaza S.A.":         StationNamesGU(THE_DESC, i) = "La Arboleda de Mendiolaza S.A."
i = 137:  StationNamesGU(THE_NAME, i) = "LA COMERCIAL GAS 822254":         StationNamesGU(THE_DESC, i) = "LA COMERCIAL GAS 822254"
i = 138:  StationNamesGU(THE_NAME, i) = "LA META S.R.L.        787765":         StationNamesGU(THE_DESC, i) = "LA META S.R.L.        787765"
i = 139:  StationNamesGU(THE_NAME, i) = "La Troja - Oncativo":         StationNamesGU(THE_DESC, i) = "La Troja - Oncativo"
i = 140:  StationNamesGU(THE_NAME, i) = "Libre 3051":         StationNamesGU(THE_DESC, i) = "Libre 3051"
i = 141:  StationNamesGU(THE_NAME, i) = "Luis A. Pavone          849025":         StationNamesGU(THE_DESC, i) = "Luis A. Pavone          849025"
i = 142:  StationNamesGU(THE_NAME, i) = "M.M.A.G. SRL         860891":         StationNamesGU(THE_DESC, i) = "M.M.A.G. SRL         860891"
i = 143:  StationNamesGU(THE_NAME, i) = "Maglione - Las Junturas":         StationNamesGU(THE_DESC, i) = "Maglione - Las Junturas"
i = 144:  StationNamesGU(THE_NAME, i) = "Manisel - Pasco":         StationNamesGU(THE_DESC, i) = "Manisel - Pasco"
i = 145:  StationNamesGU(THE_NAME, i) = "Marcos EBLAGON  965309":         StationNamesGU(THE_DESC, i) = "Marcos EBLAGON  965309"
i = 146:  StationNamesGU(THE_NAME, i) = "Marcuzzi Arnaldo  969325":         StationNamesGU(THE_DESC, i) = "Marcuzzi Arnaldo  969325"
i = 147:  StationNamesGU(THE_NAME, i) = "MARTOGLIO COM.  882125":         StationNamesGU(THE_DESC, i) = "MARTOGLIO COM.  882125"
i = 148:  StationNamesGU(THE_NAME, i) = "Maxion Montich S.A. 135":         StationNamesGU(THE_DESC, i) = "Maxion Montich S.A. 135"
i = 149:  StationNamesGU(THE_NAME, i) = "ME.CIRO Del Valle   889226":         StationNamesGU(THE_DESC, i) = "ME.CIRO Del Valle   889226"
i = 150:  StationNamesGU(THE_NAME, i) = "Melipal 3 - La Rioja":         StationNamesGU(THE_DESC, i) = "Melipal 3 - La Rioja"
i = 151:  StationNamesGU(THE_NAME, i) = "MELIPAL II S.A.":         StationNamesGU(THE_DESC, i) = "MELIPAL II S.A."
i = 152:  StationNamesGU(THE_NAME, i) = "MELIPAL S.A.           743001":         StationNamesGU(THE_DESC, i) = "MELIPAL S.A.           743001"
i = 153:  StationNamesGU(THE_NAME, i) = "Metalúrgica RAR     911762":         StationNamesGU(THE_DESC, i) = "Metalúrgica RAR     911762"
i = 154:  StationNamesGU(THE_NAME, i) = "Molinos Rio II":         StationNamesGU(THE_DESC, i) = "Molinos Rio II"
i = 155:  StationNamesGU(THE_NAME, i) = "NELSO II                 851873":         StationNamesGU(THE_DESC, i) = "NELSO II                 851873"
i = 156:  StationNamesGU(THE_NAME, i) = "NILDO LOPEZ III        772993":         StationNamesGU(THE_DESC, i) = "NILDO LOPEZ III        772993"
i = 157:  StationNamesGU(THE_NAME, i) = "Noal - Villa María":         StationNamesGU(THE_DESC, i) = "Noal - Villa María"
i = 158:  StationNamesGU(THE_NAME, i) = "Noal 2 - Villa María":         StationNamesGU(THE_DESC, i) = "Noal 2 - Villa María"
i = 159:  StationNamesGU(THE_NAME, i) = "ONOFRIO                 458274":         StationNamesGU(THE_DESC, i) = "ONOFRIO                 458274"
i = 160:  StationNamesGU(THE_NAME, i) = "Palmar S. A.":         StationNamesGU(THE_DESC, i) = "Palmar S. A."
i = 161:  StationNamesGU(THE_NAME, i) = "PAVONE Luis A       744273":         StationNamesGU(THE_DESC, i) = "PAVONE Luis A       744273"
i = 162:  StationNamesGU(THE_NAME, i) = "PETRO GRUTA  Cat.     767899":         StationNamesGU(THE_DESC, i) = "PETRO GRUTA  Cat.     767899"
i = 163:  StationNamesGU(THE_NAME, i) = "PETROSAM Samap.   811710":         StationNamesGU(THE_DESC, i) = "PETROSAM Samap.   811710"
i = 164:  StationNamesGU(THE_NAME, i) = "Porta":         StationNamesGU(THE_DESC, i) = "Porta"
i = 165:  StationNamesGU(THE_NAME, i) = "Promaiz Bioetanol - Alej. Roca":         StationNamesGU(THE_DESC, i) = "Promaiz Bioetanol - Alej. Roca"
i = 166:  StationNamesGU(THE_NAME, i) = "Prueba AyD":         StationNamesGU(THE_DESC, i) = "Prueba AyD"
i = 167:  StationNamesGU(THE_NAME, i) = "Prueba Gencontrol":         StationNamesGU(THE_DESC, i) = "Prueba Gencontrol"
i = 168:  StationNamesGU(THE_NAME, i) = "Prueba Taller":         StationNamesGU(THE_DESC, i) = "Prueba Taller"
i = 169:  StationNamesGU(THE_NAME, i) = "Punta del Agua":         StationNamesGU(THE_DESC, i) = "Punta del Agua"
i = 170:  StationNamesGU(THE_NAME, i) = "Punto Panorámico":         StationNamesGU(THE_DESC, i) = "Punto Panorámico"
i = 171:  StationNamesGU(THE_NAME, i) = "RAJORSI S.R.L         753654":         StationNamesGU(THE_DESC, i) = "RAJORSI S.R.L         753654"
i = 172:  StationNamesGU(THE_NAME, i) = "RAMONDA VillaMaria 806087":         StationNamesGU(THE_DESC, i) = "RAMONDA VillaMaria 806087"
i = 173:  StationNamesGU(THE_NAME, i) = "Red Petro SA         872211":         StationNamesGU(THE_DESC, i) = "Red Petro SA         872211"
i = 174:  StationNamesGU(THE_NAME, i) = "RSD Emprendimientos SA":         StationNamesGU(THE_DESC, i) = "RSD Emprendimientos SA"
i = 175:  StationNamesGU(THE_NAME, i) = "SAN MIGUEL S.R.L.  860375":         StationNamesGU(THE_DESC, i) = "SAN MIGUEL S.R.L.  860375"
i = 176:  StationNamesGU(THE_NAME, i) = "SAN URBANO S.A.    815052":         StationNamesGU(THE_DESC, i) = "SAN URBANO S.A.    815052"
i = 177:  StationNamesGU(THE_NAME, i) = "SAN VICENTE SRL  879381":         StationNamesGU(THE_DESC, i) = "SAN VICENTE SRL  879381"
i = 178:  StationNamesGU(THE_NAME, i) = "SANTA LUCIA SRL    823305":         StationNamesGU(THE_DESC, i) = "SANTA LUCIA SRL    823305"
i = 179:  StationNamesGU(THE_NAME, i) = "SANTA MARIA S.A.  864007":         StationNamesGU(THE_DESC, i) = "SANTA MARIA S.A.  864007"
i = 180:  StationNamesGU(THE_NAME, i) = "SERCAT S.A.           755985":         StationNamesGU(THE_DESC, i) = "SERCAT S.A.           755985"
i = 181:  StationNamesGU(THE_NAME, i) = "SERVI SUD S.A.       743993":         StationNamesGU(THE_DESC, i) = "SERVI SUD S.A.       743993"
i = 182:  StationNamesGU(THE_NAME, i) = "Servicios Agropecuarios":         StationNamesGU(THE_DESC, i) = "Servicios Agropecuarios"
i = 183:  StationNamesGU(THE_NAME, i) = "Servicios Industriales":         StationNamesGU(THE_DESC, i) = "Servicios Industriales"
i = 184:  StationNamesGU(THE_NAME, i) = "SERVICIOS S.A.       747919":         StationNamesGU(THE_DESC, i) = "SERVICIOS S.A.       747919"
i = 185:  StationNamesGU(THE_NAME, i) = "SGG ENOD La Rioja    194":         StationNamesGU(THE_DESC, i) = "SGG ENOD La Rioja    194"
i = 186:  StationNamesGU(THE_NAME, i) = "SGG SERVIO S.A.  842981":         StationNamesGU(THE_DESC, i) = "SGG SERVIO S.A.  842981"
i = 187:  StationNamesGU(THE_NAME, i) = "Sindicato del Seguro - La Falda":         StationNamesGU(THE_DESC, i) = "Sindicato del Seguro - La Falda"
i = 188:  StationNamesGU(THE_NAME, i) = "SUMI SA VillaDolores 798105":         StationNamesGU(THE_DESC, i) = "SUMI SA VillaDolores 798105"
i = 189:  StationNamesGU(THE_NAME, i) = "Vasquetto Rio IV":         StationNamesGU(THE_DESC, i) = "Vasquetto Rio IV"
i = 190:  StationNamesGU(THE_NAME, i) = "VESINM S.A.            729647":         StationNamesGU(THE_DESC, i) = "VESINM S.A.            729647"
i = 191:  StationNamesGU(THE_NAME, i) = "VILLAFUELCarlosPaz 791892":         StationNamesGU(THE_DESC, i) = "VILLAFUELCarlosPaz 791892"
i = 192:  StationNamesGU(THE_NAME, i) = "GNC ESSFE SAS 6001862649":         StationNamesGU(THE_DESC, i) = "GNC ESSFE SAS 6001862649"
i = 193:  StationNamesGU(THE_NAME, i) = "GNC Meteña Hnos S.A. Río III":         StationNamesGU(THE_DESC, i) = "GNC Meteña Hnos S.A. Río III"



i = 1:  StationNamesCamaras(THE_NAME, i) = "Aceitera Gral. Deheza":         StationNamesCamaras(THE_DESC, i) = "Aceitera Gral. Deheza"
i = 2:  StationNamesCamaras(THE_NAME, i) = "Adella Maria":         StationNamesCamaras(THE_DESC, i) = "Adella Maria"
i = 3:  StationNamesCamaras(THE_NAME, i) = "Alta Gracia":         StationNamesCamaras(THE_DESC, i) = "Alta Gracia"
i = 4:  StationNamesCamaras(THE_NAME, i) = "Anillo Córdoba":         StationNamesCamaras(THE_DESC, i) = "Anillo Córdoba"
i = 5:  StationNamesCamaras(THE_NAME, i) = "Arroyito Cespal":         StationNamesCamaras(THE_DESC, i) = "Arroyito Cespal"
i = 6:  StationNamesCamaras(THE_NAME, i) = "Arroyito Cespal 2":         StationNamesCamaras(THE_DESC, i) = "Arroyito Cespal 2"
i = 7:  StationNamesCamaras(THE_NAME, i) = "Atanor":         StationNamesCamaras(THE_DESC, i) = "Atanor"
i = 8:  StationNamesCamaras(THE_NAME, i) = "Autóctono":         StationNamesCamaras(THE_DESC, i) = "Autóctono"
i = 9:  StationNamesCamaras(THE_NAME, i) = "Balnearia":         StationNamesCamaras(THE_DESC, i) = "Balnearia"
i = 10:  StationNamesCamaras(THE_NAME, i) = "Banda Norte Rio 4°":         StationNamesCamaras(THE_DESC, i) = "Banda Norte Rio 4°"
i = 11:  StationNamesCamaras(THE_NAME, i) = "Bell Ville Baja":         StationNamesCamaras(THE_DESC, i) = "Bell Ville Baja"
i = 12:  StationNamesCamaras(THE_NAME, i) = "Bower":         StationNamesCamaras(THE_DESC, i) = "Bower"
i = 13:  StationNamesCamaras(THE_NAME, i) = "Bruzzone Alta-En Prueba":         StationNamesCamaras(THE_DESC, i) = "Bruzzone Alta-En Prueba"
i = 14:  StationNamesCamaras(THE_NAME, i) = "Bruzzone Odorización":         StationNamesCamaras(THE_DESC, i) = "Bruzzone Odorización"
i = 15:  StationNamesCamaras(THE_NAME, i) = "Buchardo":         StationNamesCamaras(THE_DESC, i) = "Buchardo"
i = 16:  StationNamesCamaras(THE_NAME, i) = "Bunge Cebal":         StationNamesCamaras(THE_DESC, i) = "Bunge Cebal"
i = 17:  StationNamesCamaras(THE_NAME, i) = "Cámara Carrara":         StationNamesCamaras(THE_DESC, i) = "Cámara Carrara"
i = 18:  StationNamesCamaras(THE_NAME, i) = "Cámara La Rioja":         StationNamesCamaras(THE_DESC, i) = "Cámara La Rioja"
i = 19:  StationNamesCamaras(THE_NAME, i) = "Cámara La Rioja Nueva":         StationNamesCamaras(THE_DESC, i) = "Cámara La Rioja Nueva"
i = 20:  StationNamesCamaras(THE_NAME, i) = "Canals":         StationNamesCamaras(THE_DESC, i) = "Canals"
i = 21:  StationNamesCamaras(THE_NAME, i) = "Central Térmica Arroyito":         StationNamesCamaras(THE_DESC, i) = "Central Térmica Arroyito"
i = 22:  StationNamesCamaras(THE_NAME, i) = "Central Térmica Arroyito 2":         StationNamesCamaras(THE_DESC, i) = "Central Térmica Arroyito 2"
i = 23:  StationNamesCamaras(THE_NAME, i) = "Charras":         StationNamesCamaras(THE_DESC, i) = "Charras"
i = 24:  StationNamesCamaras(THE_NAME, i) = "Cnel Baigorria":         StationNamesCamaras(THE_DESC, i) = "Cnel Baigorria"
i = 25:  StationNamesCamaras(THE_NAME, i) = "Cruz del Eje":         StationNamesCamaras(THE_DESC, i) = "Cruz del Eje"
i = 26:  StationNamesCamaras(THE_NAME, i) = "Dalmacio Vélez Sarsfield Baja":         StationNamesCamaras(THE_DESC, i) = "Dalmacio Vélez Sarsfield Baja"
i = 27:  StationNamesCamaras(THE_NAME, i) = "El Fortin":         StationNamesCamaras(THE_DESC, i) = "El Fortin"
i = 28:  StationNamesCamaras(THE_NAME, i) = "El Pantanillo Catamarca":         StationNamesCamaras(THE_DESC, i) = "El Pantanillo Catamarca"
i = 29:  StationNamesCamaras(THE_NAME, i) = "EPEC Dean Funes":         StationNamesCamaras(THE_DESC, i) = "EPEC Dean Funes"
i = 30:  StationNamesCamaras(THE_NAME, i) = "Epec Las Playas":         StationNamesCamaras(THE_DESC, i) = "Epec Las Playas"
i = 31:  StationNamesCamaras(THE_NAME, i) = "Epec Levalle":         StationNamesCamaras(THE_DESC, i) = "Epec Levalle"
i = 32:  StationNamesCamaras(THE_NAME, i) = "Epec Sudoeste":         StationNamesCamaras(THE_DESC, i) = "Epec Sudoeste"
i = 33:  StationNamesCamaras(THE_NAME, i) = "Equipo de prueba Taller":         StationNamesCamaras(THE_DESC, i) = "Equipo de prueba Taller"
i = 34:  StationNamesCamaras(THE_NAME, i) = "Equipo de prueba Taller ID 23":         StationNamesCamaras(THE_DESC, i) = "Equipo de prueba Taller ID 23"
i = 35:  StationNamesCamaras(THE_NAME, i) = "Equipo de prueba Taller ID 82":         StationNamesCamaras(THE_DESC, i) = "Equipo de prueba Taller ID 82"
i = 36:  StationNamesCamaras(THE_NAME, i) = "Ferreyra":         StationNamesCamaras(THE_DESC, i) = "Ferreyra"
i = 37:  StationNamesCamaras(THE_NAME, i) = "Fiat Auto":         StationNamesCamaras(THE_DESC, i) = "Fiat Auto"
i = 38:  StationNamesCamaras(THE_NAME, i) = "Gen Mediterránea RIV":         StationNamesCamaras(THE_DESC, i) = "Gen Mediterránea RIV"
i = 39:  StationNamesCamaras(THE_NAME, i) = "Gen. Mediterránea La Rioja":         StationNamesCamaras(THE_DESC, i) = "Gen. Mediterránea La Rioja"
i = 40:  StationNamesCamaras(THE_NAME, i) = "Generadora Riojana":         StationNamesCamaras(THE_DESC, i) = "Generadora Riojana"
i = 41:  StationNamesCamaras(THE_NAME, i) = "GNC Yaguareté - La Falda":         StationNamesCamaras(THE_DESC, i) = "GNC Yaguareté - La Falda"
i = 42:  StationNamesCamaras(THE_NAME, i) = "Gral Cabrera":         StationNamesCamaras(THE_DESC, i) = "Gral Cabrera"
i = 43:  StationNamesCamaras(THE_NAME, i) = "Gral. Cabrera Baja":         StationNamesCamaras(THE_DESC, i) = "Gral. Cabrera Baja"
i = 44:  StationNamesCamaras(THE_NAME, i) = "Gral. Cabrera Cromatógrafo":         StationNamesCamaras(THE_DESC, i) = "Gral. Cabrera Cromatógrafo"
i = 45:  StationNamesCamaras(THE_NAME, i) = "Gral. Deheza Alta":         StationNamesCamaras(THE_DESC, i) = "Gral. Deheza Alta"
i = 46:  StationNamesCamaras(THE_NAME, i) = "Gral. Deheza Baja":         StationNamesCamaras(THE_DESC, i) = "Gral. Deheza Baja"
i = 47:  StationNamesCamaras(THE_NAME, i) = "Holcim Malagueño":         StationNamesCamaras(THE_DESC, i) = "Holcim Malagueño"
i = 48:  StationNamesCamaras(THE_NAME, i) = "Huinca Renancó":         StationNamesCamaras(THE_DESC, i) = "Huinca Renancó"
i = 49:  StationNamesCamaras(THE_NAME, i) = "Inriville":         StationNamesCamaras(THE_DESC, i) = "Inriville"
i = 50:  StationNamesCamaras(THE_NAME, i) = "Intendencia":         StationNamesCamaras(THE_DESC, i) = "Intendencia"
i = 51:  StationNamesCamaras(THE_NAME, i) = "Iveco":         StationNamesCamaras(THE_DESC, i) = "Iveco"
i = 52:  StationNamesCamaras(THE_NAME, i) = "Jesús María":         StationNamesCamaras(THE_DESC, i) = "Jesús María"
i = 53:  StationNamesCamaras(THE_NAME, i) = "La Calera":         StationNamesCamaras(THE_DESC, i) = "La Calera"
i = 54:  StationNamesCamaras(THE_NAME, i) = "La Carlota":         StationNamesCamaras(THE_DESC, i) = "La Carlota"
i = 55:  StationNamesCamaras(THE_NAME, i) = "La Cumbrecita":         StationNamesCamaras(THE_DESC, i) = "La Cumbrecita"
i = 56:  StationNamesCamaras(THE_NAME, i) = "La Granja":         StationNamesCamaras(THE_DESC, i) = "La Granja"
i = 57:  StationNamesCamaras(THE_NAME, i) = "La Laguna":         StationNamesCamaras(THE_DESC, i) = "La Laguna"
i = 58:  StationNamesCamaras(THE_NAME, i) = "La Paz":         StationNamesCamaras(THE_DESC, i) = "La Paz"
i = 59:  StationNamesCamaras(THE_NAME, i) = "La Población":         StationNamesCamaras(THE_DESC, i) = "La Población"
i = 60:  StationNamesCamaras(THE_NAME, i) = "La Quinta Carlos Paz":         StationNamesCamaras(THE_DESC, i) = "La Quinta Carlos Paz"
i = 61:  StationNamesCamaras(THE_NAME, i) = "Laborde":         StationNamesCamaras(THE_DESC, i) = "Laborde"
i = 62:  StationNamesCamaras(THE_NAME, i) = "Laboulaye":         StationNamesCamaras(THE_DESC, i) = "Laboulaye"
i = 63:  StationNamesCamaras(THE_NAME, i) = "Las Américas":         StationNamesCamaras(THE_DESC, i) = "Las Américas"
i = 64:  StationNamesCamaras(THE_NAME, i) = "Las Junturas":         StationNamesCamaras(THE_DESC, i) = "Las Junturas"
i = 65:  StationNamesCamaras(THE_NAME, i) = "Las Varillas":         StationNamesCamaras(THE_DESC, i) = "Las Varillas"
i = 66:  StationNamesCamaras(THE_NAME, i) = "Las Varillas Baja":         StationNamesCamaras(THE_DESC, i) = "Las Varillas Baja"
i = 67:  StationNamesCamaras(THE_NAME, i) = "Loma Negra":         StationNamesCamaras(THE_DESC, i) = "Loma Negra"
i = 68:  StationNamesCamaras(THE_NAME, i) = "Luque":         StationNamesCamaras(THE_DESC, i) = "Luque"
i = 69:  StationNamesCamaras(THE_NAME, i) = "Mendiolaza":         StationNamesCamaras(THE_DESC, i) = "Mendiolaza"
i = 70:  StationNamesCamaras(THE_NAME, i) = "Mina Clavero":         StationNamesCamaras(THE_DESC, i) = "Mina Clavero"
i = 71:  StationNamesCamaras(THE_NAME, i) = "Morteros":         StationNamesCamaras(THE_DESC, i) = "Morteros"
i = 72:  StationNamesCamaras(THE_NAME, i) = "Nicolás Bruzzone":         StationNamesCamaras(THE_DESC, i) = "Nicolás Bruzzone"
i = 73:  StationNamesCamaras(THE_NAME, i) = "Noetinger":         StationNamesCamaras(THE_DESC, i) = "Noetinger"
i = 74:  StationNamesCamaras(THE_NAME, i) = "Odorización Chumbicha LR":         StationNamesCamaras(THE_DESC, i) = "Odorización Chumbicha LR"
i = 75:  StationNamesCamaras(THE_NAME, i) = "Odorización Ferreyra":         StationNamesCamaras(THE_DESC, i) = "Odorización Ferreyra"
i = 76:  StationNamesCamaras(THE_NAME, i) = "Odorización Valle de Conlara":         StationNamesCamaras(THE_DESC, i) = "Odorización Valle de Conlara"
i = 77:  StationNamesCamaras(THE_NAME, i) = "Odorización Yofre":         StationNamesCamaras(THE_DESC, i) = "Odorización Yofre"
i = 78:  StationNamesCamaras(THE_NAME, i) = "Paraísos":         StationNamesCamaras(THE_DESC, i) = "Paraísos"
i = 79:  StationNamesCamaras(THE_NAME, i) = "Pascanas":         StationNamesCamaras(THE_DESC, i) = "Pascanas"
i = 80:  StationNamesCamaras(THE_NAME, i) = "Pasco":         StationNamesCamaras(THE_DESC, i) = "Pasco"
i = 81:  StationNamesCamaras(THE_NAME, i) = "Paso del Durazno":         StationNamesCamaras(THE_DESC, i) = "Paso del Durazno"
i = 82:  StationNamesCamaras(THE_NAME, i) = "Patricios":         StationNamesCamaras(THE_DESC, i) = "Patricios"
i = 83:  StationNamesCamaras(THE_NAME, i) = "Pedro Funes":         StationNamesCamaras(THE_DESC, i) = "Pedro Funes"
i = 84:  StationNamesCamaras(THE_NAME, i) = "Pedro Viña":         StationNamesCamaras(THE_DESC, i) = "Pedro Viña"
i = 85:  StationNamesCamaras(THE_NAME, i) = "Petroq. RIII":         StationNamesCamaras(THE_DESC, i) = "Petroq. RIII"
i = 86:  StationNamesCamaras(THE_NAME, i) = "Posse":         StationNamesCamaras(THE_DESC, i) = "Posse"
i = 87:  StationNamesCamaras(THE_NAME, i) = "Pque. Industrial V. Dolores":         StationNamesCamaras(THE_DESC, i) = "Pque. Industrial V. Dolores"
i = 88:  StationNamesCamaras(THE_NAME, i) = "PRF B. Don Bosco":         StationNamesCamaras(THE_DESC, i) = "PRF B. Don Bosco"
i = 89:  StationNamesCamaras(THE_NAME, i) = "PRF B. Los Carolinos":         StationNamesCamaras(THE_DESC, i) = "PRF B. Los Carolinos"
i = 90:  StationNamesCamaras(THE_NAME, i) = "PRF B° Cerino Río III":         StationNamesCamaras(THE_DESC, i) = "PRF B° Cerino Río III"
i = 91:  StationNamesCamaras(THE_NAME, i) = "PRF Cementerio RIV":         StationNamesCamaras(THE_DESC, i) = "PRF Cementerio RIV"
i = 92:  StationNamesCamaras(THE_NAME, i) = "PRF Fortín del Pozo":         StationNamesCamaras(THE_DESC, i) = "PRF Fortín del Pozo"
i = 93:  StationNamesCamaras(THE_NAME, i) = "PRF Los Surgentes":         StationNamesCamaras(THE_DESC, i) = "PRF Los Surgentes"
i = 94:  StationNamesCamaras(THE_NAME, i) = "PRF Oliva":         StationNamesCamaras(THE_DESC, i) = "PRF Oliva"
i = 95:  StationNamesCamaras(THE_NAME, i) = "PRF Pza. Centenario V. M.":         StationNamesCamaras(THE_DESC, i) = "PRF Pza. Centenario V. M."
i = 96:  StationNamesCamaras(THE_NAME, i) = "PRF Rotonda Ruta 20":         StationNamesCamaras(THE_DESC, i) = "PRF Rotonda Ruta 20"
i = 97:  StationNamesCamaras(THE_NAME, i) = "PRI Alejandro Roca":         StationNamesCamaras(THE_DESC, i) = "PRI Alejandro Roca"
i = 98:  StationNamesCamaras(THE_NAME, i) = "PRI Hernando":         StationNamesCamaras(THE_DESC, i) = "PRI Hernando"
i = 99:  StationNamesCamaras(THE_NAME, i) = "PRI Juárez Celman":         StationNamesCamaras(THE_DESC, i) = "PRI Juárez Celman"
i = 100:  StationNamesCamaras(THE_NAME, i) = "PRI La Falda-Valle Hermoso":        StationNamesCamaras(THE_DESC, i) = "PRI La Falda-Valle Hermoso"
i = 101:  StationNamesCamaras(THE_NAME, i) = "Pza. Alberdi":        StationNamesCamaras(THE_DESC, i) = "Pza. Alberdi"
i = 102:  StationNamesCamaras(THE_NAME, i) = "Pza. Carballo":        StationNamesCamaras(THE_DESC, i) = "Pza. Carballo"
i = 103:  StationNamesCamaras(THE_NAME, i) = "Pza. F. Quiroga La Rioja":        StationNamesCamaras(THE_DESC, i) = "Pza. F. Quiroga La Rioja"
i = 104:  StationNamesCamaras(THE_NAME, i) = "Pza. Gral Paz":        StationNamesCamaras(THE_DESC, i) = "Pza. Gral Paz"
i = 105:  StationNamesCamaras(THE_NAME, i) = "Pza. Navarro":        StationNamesCamaras(THE_DESC, i) = "Pza. Navarro"
i = 106:  StationNamesCamaras(THE_NAME, i) = "Pza. Perón Carlos Paz":        StationNamesCamaras(THE_DESC, i) = "Pza. Perón Carlos Paz"
i = 107:  StationNamesCamaras(THE_NAME, i) = "Q. de Arguello":        StationNamesCamaras(THE_DESC, i) = "Q. de Arguello"
i = 108:  StationNamesCamaras(THE_NAME, i) = "Renault":        StationNamesCamaras(THE_DESC, i) = "Renault"
i = 109:  StationNamesCamaras(THE_NAME, i) = "Río III Fuerza Aerea":        StationNamesCamaras(THE_DESC, i) = "Río III Fuerza Aerea"
i = 110:  StationNamesCamaras(THE_NAME, i) = "Río III Sur":        StationNamesCamaras(THE_DESC, i) = "Río III Sur"
i = 111:  StationNamesCamaras(THE_NAME, i) = "RIV Ciudad":        StationNamesCamaras(THE_DESC, i) = "RIV Ciudad"
i = 112:  StationNamesCamaras(THE_NAME, i) = "Sampacho":        StationNamesCamaras(THE_DESC, i) = "Sampacho"
i = 113:  StationNamesCamaras(THE_NAME, i) = "San Francisco":        StationNamesCamaras(THE_DESC, i) = "San Francisco"
i = 114:  StationNamesCamaras(THE_NAME, i) = "San José de la Dormida":        StationNamesCamaras(THE_DESC, i) = "San José de la Dormida"
i = 115:  StationNamesCamaras(THE_NAME, i) = "San Vicente":        StationNamesCamaras(THE_DESC, i) = "San Vicente"
i = 116:  StationNamesCamaras(THE_NAME, i) = "Santa María de Punilla":        StationNamesCamaras(THE_DESC, i) = "Santa María de Punilla"
i = 117:  StationNamesCamaras(THE_NAME, i) = "Sarmiento":        StationNamesCamaras(THE_DESC, i) = "Sarmiento"
i = 118:  StationNamesCamaras(THE_NAME, i) = "Serrezuela":        StationNamesCamaras(THE_DESC, i) = "Serrezuela"
i = 119:  StationNamesCamaras(THE_NAME, i) = "Soenergy RIII":        StationNamesCamaras(THE_DESC, i) = "Soenergy RIII"
i = 120:  StationNamesCamaras(THE_NAME, i) = "Sudoeste":        StationNamesCamaras(THE_DESC, i) = "Sudoeste"
i = 121:  StationNamesCamaras(THE_NAME, i) = "Tancacha":        StationNamesCamaras(THE_DESC, i) = "Tancacha"
i = 122:  StationNamesCamaras(THE_NAME, i) = "TGN Acequias":        StationNamesCamaras(THE_DESC, i) = "TGN Acequias"
i = 123:  StationNamesCamaras(THE_NAME, i) = "TGN Bell Ville":        StationNamesCamaras(THE_DESC, i) = "TGN Bell Ville"
i = 124:  StationNamesCamaras(THE_NAME, i) = "TGN C.-Hernando":        StationNamesCamaras(THE_DESC, i) = "TGN C.-Hernando"
i = 125:  StationNamesCamaras(THE_NAME, i) = "TGN Cat. La Rioja":        StationNamesCamaras(THE_DESC, i) = "TGN Cat. La Rioja"
i = 126:  StationNamesCamaras(THE_NAME, i) = "TGN Centro Este II":        StationNamesCamaras(THE_DESC, i) = "TGN Centro Este II"
i = 127:  StationNamesCamaras(THE_NAME, i) = "TGN Colonia Caroya":        StationNamesCamaras(THE_DESC, i) = "TGN Colonia Caroya"
i = 128:  StationNamesCamaras(THE_NAME, i) = "TGN Córdoba Sur 1":        StationNamesCamaras(THE_DESC, i) = "TGN Córdoba Sur 1"
i = 129:  StationNamesCamaras(THE_NAME, i) = "TGN Córdoba Sur 2":        StationNamesCamaras(THE_DESC, i) = "TGN Córdoba Sur 2"
i = 130:  StationNamesCamaras(THE_NAME, i) = "TGN Este Córdoba":        StationNamesCamaras(THE_DESC, i) = "TGN Este Córdoba"
i = 131:  StationNamesCamaras(THE_NAME, i) = "TGN Ferreyra":        StationNamesCamaras(THE_DESC, i) = "TGN Ferreyra"
i = 132:  StationNamesCamaras(THE_NAME, i) = "TGN G. Pilar":        StationNamesCamaras(THE_DESC, i) = "TGN G. Pilar"
i = 133:  StationNamesCamaras(THE_NAME, i) = "TGN Las Playas":        StationNamesCamaras(THE_DESC, i) = "TGN Las Playas"
i = 134:  StationNamesCamaras(THE_NAME, i) = "TGN Leones":        StationNamesCamaras(THE_DESC, i) = "TGN Leones"
i = 135:  StationNamesCamaras(THE_NAME, i) = "TGN Loma Negra":        StationNamesCamaras(THE_DESC, i) = "TGN Loma Negra"
i = 136:  StationNamesCamaras(THE_NAME, i) = "TGN M. Juarez":        StationNamesCamaras(THE_DESC, i) = "TGN M. Juarez"
i = 137:  StationNamesCamaras(THE_NAME, i) = "TGN Malena RIV":        StationNamesCamaras(THE_DESC, i) = "TGN Malena RIV"
i = 138:  StationNamesCamaras(THE_NAME, i) = "TGN Manisero":        StationNamesCamaras(THE_DESC, i) = "TGN Manisero"
i = 139:  StationNamesCamaras(THE_NAME, i) = "TGN Maranzana":        StationNamesCamaras(THE_DESC, i) = "TGN Maranzana"
i = 140:  StationNamesCamaras(THE_NAME, i) = "TGN Oliva":        StationNamesCamaras(THE_DESC, i) = "TGN Oliva"
i = 141:  StationNamesCamaras(THE_NAME, i) = "TGN Pilar Arroyito":        StationNamesCamaras(THE_DESC, i) = "TGN Pilar Arroyito"
i = 142:  StationNamesCamaras(THE_NAME, i) = "TGN Rio IV 60-25 Bar":        StationNamesCamaras(THE_DESC, i) = "TGN Rio IV 60-25 Bar"
i = 143:  StationNamesCamaras(THE_NAME, i) = "TGN Toledo":        StationNamesCamaras(THE_DESC, i) = "TGN Toledo"
i = 144:  StationNamesCamaras(THE_NAME, i) = "TGN Va. María":        StationNamesCamaras(THE_DESC, i) = "TGN Va. María"
i = 145:  StationNamesCamaras(THE_NAME, i) = "TGN Valle de Calamuchita":        StationNamesCamaras(THE_DESC, i) = "TGN Valle de Calamuchita"
i = 146:  StationNamesCamaras(THE_NAME, i) = "TGN Villa Totoral":        StationNamesCamaras(THE_DESC, i) = "TGN Villa Totoral"
i = 147:  StationNamesCamaras(THE_NAME, i) = "TGN Yofre":        StationNamesCamaras(THE_DESC, i) = "TGN Yofre"
i = 148:  StationNamesCamaras(THE_NAME, i) = "Toledo":        StationNamesCamaras(THE_DESC, i) = "Toledo"
i = 149:  StationNamesCamaras(THE_NAME, i) = "Tránsito":        StationNamesCamaras(THE_DESC, i) = "Tránsito"
i = 150:  StationNamesCamaras(THE_NAME, i) = "Turismo Carlos Paz":        StationNamesCamaras(THE_DESC, i) = "Turismo Carlos Paz"
i = 151:  StationNamesCamaras(THE_NAME, i) = "Unquillo":        StationNamesCamaras(THE_DESC, i) = "Unquillo"
i = 152:  StationNamesCamaras(THE_NAME, i) = "Vál. Aut. Adelia María":        StationNamesCamaras(THE_DESC, i) = "Vál. Aut. Adelia María"
i = 153:  StationNamesCamaras(THE_NAME, i) = "Vál. Aut. Bower":        StationNamesCamaras(THE_DESC, i) = "Vál. Aut. Bower"
i = 154:  StationNamesCamaras(THE_NAME, i) = "Vál. Aut. Chuña":        StationNamesCamaras(THE_DESC, i) = "Vál. Aut. Chuña"
i = 155:  StationNamesCamaras(THE_NAME, i) = "Vál. Aut. Colonia Almada":        StationNamesCamaras(THE_DESC, i) = "Vál. Aut. Colonia Almada"
i = 156:  StationNamesCamaras(THE_NAME, i) = "Vál. Aut. Conlara":        StationNamesCamaras(THE_DESC, i) = "Vál. Aut. Conlara"
i = 157:  StationNamesCamaras(THE_NAME, i) = "Vál. Aut. Dalmacio Vélez":        StationNamesCamaras(THE_DESC, i) = "Vál. Aut. Dalmacio Vélez"
i = 158:  StationNamesCamaras(THE_NAME, i) = "Vál. Aut. Gen. Mediterránea RIV":        StationNamesCamaras(THE_DESC, i) = "Vál. Aut. Gen. Mediterránea RIV"
i = 159:  StationNamesCamaras(THE_NAME, i) = "Vál. Aut. Las Junturas":        StationNamesCamaras(THE_DESC, i) = "Vál. Aut. Las Junturas"
i = 160:  StationNamesCamaras(THE_NAME, i) = "Vál. Aut. Los Hornillos":        StationNamesCamaras(THE_DESC, i) = "Vál. Aut. Los Hornillos"
i = 161:  StationNamesCamaras(THE_NAME, i) = "Vál. Aut. Luque":        StationNamesCamaras(THE_DESC, i) = "Vál. Aut. Luque"
i = 162:  StationNamesCamaras(THE_NAME, i) = "Vál. Aut. N°1 La Falda":        StationNamesCamaras(THE_DESC, i) = "Vál. Aut. N°1 La Falda"
i = 163:  StationNamesCamaras(THE_NAME, i) = "Vál. Aut. N°1 Recreo":        StationNamesCamaras(THE_DESC, i) = "Vál. Aut. N°1 Recreo"
i = 164:  StationNamesCamaras(THE_NAME, i) = "Vál. Aut. N°2 Bruzzone":        StationNamesCamaras(THE_DESC, i) = "Vál. Aut. N°2 Bruzzone"
i = 165:  StationNamesCamaras(THE_NAME, i) = "Vál. Aut. N°2 Chumbicha-Catamarca":        StationNamesCamaras(THE_DESC, i) = "Vál. Aut. N°2 Chumbicha-Catamarca"
i = 166:  StationNamesCamaras(THE_NAME, i) = "Vál. Aut. N°2 Chumbicha-La Rioja":        StationNamesCamaras(THE_DESC, i) = "Vál. Aut. N°2 Chumbicha-La Rioja"
i = 167:  StationNamesCamaras(THE_NAME, i) = "Vál. Aut. N°2 Esquiú  CAT":        StationNamesCamaras(THE_DESC, i) = "Vál. Aut. N°2 Esquiú  CAT"
i = 168:  StationNamesCamaras(THE_NAME, i) = "Vál. Aut. N°2 Juárez Celman":        StationNamesCamaras(THE_DESC, i) = "Vál. Aut. N°2 Juárez Celman"
i = 169:  StationNamesCamaras(THE_NAME, i) = "Vál. Aut. N°2 Río Primero":        StationNamesCamaras(THE_DESC, i) = "Vál. Aut. N°2 Río Primero"
i = 170:  StationNamesCamaras(THE_NAME, i) = "Vál. Aut. N°2  Luca":        StationNamesCamaras(THE_DESC, i) = "Vál. Aut. N°2  Luca"
i = 171:  StationNamesCamaras(THE_NAME, i) = "Vál. Aut. N°3 Chumbicha-La Rioja":        StationNamesCamaras(THE_DESC, i) = "Vál. Aut. N°3 Chumbicha-La Rioja"
i = 172:  StationNamesCamaras(THE_NAME, i) = "Vál. Aut. N°3 La Guardia":        StationNamesCamaras(THE_DESC, i) = "Vál. Aut. N°3 La Guardia"
i = 173:  StationNamesCamaras(THE_NAME, i) = "Vál. Aut. N°4 Tránsito":        StationNamesCamaras(THE_DESC, i) = "Vál. Aut. N°4 Tránsito"
i = 174:  StationNamesCamaras(THE_NAME, i) = "Vál. Aut. N°5 San Martín":        StationNamesCamaras(THE_DESC, i) = "Vál. Aut. N°5 San Martín"
i = 175:  StationNamesCamaras(THE_NAME, i) = "Vál. Aut. N°6 Carranza":        StationNamesCamaras(THE_DESC, i) = "Vál. Aut. N°6 Carranza"
i = 176:  StationNamesCamaras(THE_NAME, i) = "Vál. Aut. N°6 La Francia":        StationNamesCamaras(THE_DESC, i) = "Vál. Aut. N°6 La Francia"
i = 177:  StationNamesCamaras(THE_NAME, i) = "Vál. Aut. N°7 Colonia Marina":        StationNamesCamaras(THE_DESC, i) = "Vál. Aut. N°7 Colonia Marina"
i = 178:  StationNamesCamaras(THE_NAME, i) = "Vál. Aut. Pozo del Molle":        StationNamesCamaras(THE_DESC, i) = "Vál. Aut. Pozo del Molle"
i = 179:  StationNamesCamaras(THE_NAME, i) = "Villa Allende":        StationNamesCamaras(THE_DESC, i) = "Villa Allende"
i = 180:  StationNamesCamaras(THE_NAME, i) = "Villa Allende Centro":        StationNamesCamaras(THE_DESC, i) = "Villa Allende Centro"
i = 181:  StationNamesCamaras(THE_NAME, i) = "Villa Carlos Paz - Keops":        StationNamesCamaras(THE_DESC, i) = "Villa Carlos Paz - Keops"
i = 182:  StationNamesCamaras(THE_NAME, i) = "Villa Carlos Paz - Keops 1.5bar":        StationNamesCamaras(THE_DESC, i) = "Villa Carlos Paz - Keops 1.5bar"
i = 183:  StationNamesCamaras(THE_NAME, i) = "Villa del Rosario Alta":        StationNamesCamaras(THE_DESC, i) = "Villa del Rosario Alta"
i = 184:  StationNamesCamaras(THE_NAME, i) = "Villa del Rosario Baja":        StationNamesCamaras(THE_DESC, i) = "Villa del Rosario Baja"
i = 185:  StationNamesCamaras(THE_NAME, i) = "Villa del Totoral":        StationNamesCamaras(THE_DESC, i) = "Villa del Totoral"
i = 186:  StationNamesCamaras(THE_NAME, i) = "Villa Dolores":        StationNamesCamaras(THE_DESC, i) = "Villa Dolores"
i = 187:  StationNamesCamaras(THE_NAME, i) = "Villa General Belgrano":        StationNamesCamaras(THE_DESC, i) = "Villa General Belgrano"
i = 188:  StationNamesCamaras(THE_NAME, i) = "Villa María":        StationNamesCamaras(THE_DESC, i) = "Villa María"
i = 189:  StationNamesCamaras(THE_NAME, i) = "Yofre":        StationNamesCamaras(THE_DESC, i) = "Yofre"
i = 190:  StationNamesCamaras(THE_NAME, i) = "Arroyo Cabral Baja":        StationNamesCamaras(THE_DESC, i) = "Arroyo Cabral Baja"
i = 191:  StationNamesCamaras(THE_NAME, i) = "James Craik":        StationNamesCamaras(THE_DESC, i) = "James Craik"


  

i = 1: StationNamesVeribox(THE_NAME, i) = "Acabio C. L.":         StationNamesVeribox(THE_DESC, i) = "Acabio C. L."
i = 2: StationNamesVeribox(THE_NAME, i) = "Allevard Rejna Argentina":         StationNamesVeribox(THE_DESC, i) = "Allevard Rejna Argentina"
i = 3: StationNamesVeribox(THE_NAME, i) = "Arcor - Caroya":         StationNamesVeribox(THE_DESC, i) = "Arcor - Caroya"
i = 4: StationNamesVeribox(THE_NAME, i) = "Arcor - Recreo":         StationNamesVeribox(THE_DESC, i) = "Arcor - Recreo"
i = 5: StationNamesVeribox(THE_NAME, i) = "Bagley Argentina S.A.":         StationNamesVeribox(THE_DESC, i) = "Bagley Argentina S.A."
i = 6: StationNamesVeribox(THE_NAME, i) = "Canteras Cerro Negro S.A.":         StationNamesVeribox(THE_DESC, i) = "Canteras Cerro Negro S.A."
i = 9: StationNamesVeribox(THE_NAME, i) = "Cartocor S.A.":         StationNamesVeribox(THE_DESC, i) = "Cartocor S.A."
i = 7: StationNamesVeribox(THE_NAME, i) = "Cerámica Catamarca S.R.L.":         StationNamesVeribox(THE_DESC, i) = "Cerámica Catamarca S.R.L."
i = 8: StationNamesVeribox(THE_NAME, i) = "Colortex":         StationNamesVeribox(THE_DESC, i) = "Colortex"
i = 10: StationNamesVeribox(THE_NAME, i) = "D. G. F. Militares RIII DPM":         StationNamesVeribox(THE_DESC, i) = "D. G. F. Militares RIII DPM"
i = 11: StationNamesVeribox(THE_NAME, i) = "D. G. F. Militares RIII DPQ":         StationNamesVeribox(THE_DESC, i) = "D. G. F. Militares RIII DPQ"
i = 12: StationNamesVeribox(THE_NAME, i) = "D. G. F. Militares VM":         StationNamesVeribox(THE_DESC, i) = "D. G. F. Militares VM"
i = 13: StationNamesVeribox(THE_NAME, i) = "DPAMA S.A.":         StationNamesVeribox(THE_DESC, i) = "DPAMA S.A."
i = 14: StationNamesVeribox(THE_NAME, i) = "Gas Carbónico Chiantore":         StationNamesVeribox(THE_DESC, i) = "Gas Carbónico Chiantore"
i = 15: StationNamesVeribox(THE_NAME, i) = "José Guma S.A.":         StationNamesVeribox(THE_DESC, i) = "José Guma S.A."
i = 16: StationNamesVeribox(THE_NAME, i) = "Lácteos La Cristina S.A.":         StationNamesVeribox(THE_DESC, i) = "Lácteos La Cristina S.A."
i = 17: StationNamesVeribox(THE_NAME, i) = "Metal Veneta S.A.":         StationNamesVeribox(THE_DESC, i) = "Metal Veneta S.A."
i = 18: StationNamesVeribox(THE_NAME, i) = "Molfino Hnos. S.A.":         StationNamesVeribox(THE_DESC, i) = "Molfino Hnos. S.A."
i = 19: StationNamesVeribox(THE_NAME, i) = "Olca S.A.I.C.":         StationNamesVeribox(THE_DESC, i) = "Olca S.A.I.C."
i = 20: StationNamesVeribox(THE_NAME, i) = "Porta Hnos. S. A.":         StationNamesVeribox(THE_DESC, i) = "Porta Hnos. S. A."
i = 21: StationNamesVeribox(THE_NAME, i) = "Refinería del Centro":         StationNamesVeribox(THE_DESC, i) = "Refinería del Centro"
i = 22: StationNamesVeribox(THE_NAME, i) = "Ricoltex":         StationNamesVeribox(THE_DESC, i) = "Ricoltex"
i = 23: StationNamesVeribox(THE_NAME, i) = "Tecotex":         StationNamesVeribox(THE_DESC, i) = "Tecotex"
i = 24: StationNamesVeribox(THE_NAME, i) = "Vitopel S.A.":         StationNamesVeribox(THE_DESC, i) = "Vitopel S.A."
i = 25: StationNamesVeribox(THE_NAME, i) = "Volkswagen Argentina S.A.":         StationNamesVeribox(THE_DESC, i) = "Volkswagen Argentina S.A."
i = 26: StationNamesVeribox(THE_NAME, i) = "Libre28664":         StationNamesVeribox(THE_DESC, i) = "Libre28664"
i = 27: StationNamesVeribox(THE_NAME, i) = "Libre28665":         StationNamesVeribox(THE_DESC, i) = "Libre28664-exRecreo"
    
    
i = 1:           StationNamesCromas(THE_NAME, i) = "Cromat. Norte":                    StationNamesCromas(THE_DESC, i) = "Cromat. Norte"
i = 2:           StationNamesCromas(THE_NAME, i) = "Cromat. Centro-Oeste":                   StationNamesCromas(THE_DESC, i) = "Cromat. Centro-Oeste"
i = 3:           StationNamesCromas(THE_NAME, i) = "Cromat. Mendoza Sur":                   StationNamesCromas(THE_DESC, i) = "Cromat. Cromat. Mendoza Sur"
    

End Sub



Private Function GetCurrentSQLString() As String


Dim i As Integer
Dim sListaCampos, sListaEstaciones As String
Dim strDateFrom, strDateTo, strStation, strTable, strSQL As String
Dim sWhere As String
Dim sOrderBy As String
Dim selectString As String

    strDateFrom = Format(dtFechaInicial.Value, "MM-dd-YYYY HH:mm")
    strDateTo = Format(dtFechaFinal.Value, "MM-dd-YYYY HH:mm")
    selectString = "*"
    Select Case cmbSortList1.Text
        Case "Cromatografia"
            strTable = "cromatografia"
        Case "Medicion"
            strTable = "medicion"
        Case "Grandes Usuarios"
            strTable = "GncStationsHisto"
         Case "Veribox"
            strTable = "ucv.ucv_data"
            'formateo la fecha para consulta a bd Veribox
            strDateFrom = Format(dtFechaInicial.Value, "YYYY-MM-dd HH:mm")
            strDateTo = Format(dtFechaFinal.Value, "YYYY-MM-dd HH:mm")
            'Inicializo la consulta  adecuada para  Veribox
            selectString = " DateTime,StationName,pca,vol_c,vol_c_an,vol_nc,vol_nc_an,presion,temperatura,batt_vc,batt_vb,alarma, dst "
        Case Else
            strTable = "GncStationsHisto"
    End Select
    
    
    sListaCampos = ""
    With lbCampos
        For i = 0 To .ListCount - 1
            If .Selected(i) Then
                sListaCampos = sListaCampos & """" & .List(i, 1) & """, "
            End If
        Next i
        If sListaCampos = "(" Or sListaCampos = "" Or Len(sListaCampos) < 3 Then
            sListaCampos = selectString
            
        Else
            sListaCampos = Left(sListaCampos, Len(sListaCampos) - 2) & ""
        End If
    End With

    sListaEstaciones = "AND ("
    With lbEstaciones
        For i = 0 To .ListCount - 1
            If .Selected(i) Then
                sListaEstaciones = sListaEstaciones & """StationName"" = '" & .List(i, 1) & "' OR "
            End If
        Next i
        If sListaEstaciones = "AND (" Or sListaEstaciones = "" Or sListaEstaciones = "(" Or Len(sListaEstaciones) < 5 Then
            sListaEstaciones = ""
        Else
            sListaEstaciones = Left(sListaEstaciones, Len(sListaEstaciones) - 4) & ")"
        End If
    End With
    
    
    If Trim(cbOrden1) <> "" Then sOrderBy = "ORDER BY """ & cbOrden1.Column(1) & IIf(bOrden1Asc, """ ASC", """ DESC")
    
    
    
    strSQL = "Select " & sListaCampos & " From """ + strTable + """ Where " & _
             "(""DateTime"" Between '" & strDateFrom & "' and '" & strDateTo & "') " & sListaEstaciones & sOrderBy & ";"
    
                          
    GetCurrentSQLString = strSQL
End Function


Private Function exportGridData()
Dim oExcel As Object
Dim xlApp As Excel.Application
Dim xlLibro As Excel.Workbook
Dim oBook As Object
Dim oSheet As Object
Dim Dia
Dim hora
Dim Dia_a
Dim hora_a
Dim Ruta_archivo

Dim i, j, k, maxStations As Integer



Dim sUserID As String
Dim sUserName As String
Dim sGroupName As String
Dim sType As String


  Dia = Format(Date, "YYYY-MM-dd")
  hora = Format(Time, "HH:mm")
  Dia_a = Format(Date, "_YYYY-MM-dd")
  hora_a = Format(Time, "_HH-mm")



  'Start a new workbook in Excel
  
  Set xlApp = New Excel.Application
  
  
  

   Select Case cmbSortList1.Text
        Case "Cromatografia"
            sType = "croma"
            Set xlLibro = xlApp.Workbooks.Open("C:\ECOGAS\REPORTE\Plantillas\datosMC.xlsx", True, True, , "")
            Set oBook = xlApp.Worksheets("CROMATOGRAFIA")
            Set oSheet = xlApp.Worksheets("CROMATOGRAFIA")
        Case "Medicion"
            sType = "med"
            Set xlLibro = xlApp.Workbooks.Open("C:\ECOGAS\REPORTE\Plantillas\datosMC.xlsx", True, True, , "")
            Set oBook = xlApp.Worksheets("MEDICION")
            Set oSheet = xlApp.Worksheets("MEDICION")
        Case "Grandes Usuarios"
            sType = "GU"
            Set xlLibro = xlApp.Workbooks.Open("C:\ECOGAS\REPORTE\Plantillas\datosGU.xlsx", True, True, , "")
            Set oBook = xlApp.Worksheets("GU")
            Set oSheet = xlApp.Worksheets("GU")
        Case Else
            Set xlLibro = xlApp.Workbooks.Open("C:\ECOGAS\REPORTE\Plantillas\datosGU.xlsx", True, True, , "")
            Set oBook = xlApp.Worksheets("GU")
            Set oSheet = xlApp.Worksheets("GU")
    End Select
    

  


  j = 5
  k = 5
  
  
            
  maxStations = 113
  For i = 0 To MSFlexGrid1.Rows - 1
    For j = 0 To MSFlexGrid1.Cols - 1

            oSheet.Cells(i + 3, j + 1).Value = MSFlexGrid1.TextMatrix(i, j)
            If (i > 1) Then
               If (j = 0) Then
                ' si es fecha lo formateo
               ' If IsDate(MSFlexGrid1.TextMatrix(i, j)) Then
               '     oSheet.Range("A:A").NumberFormat = "yyyy/MM/dd HH:mm:ss"
                    oSheet.Cells(i + 3, j + 1).Value = CDate(MSFlexGrid1.TextMatrix(i, j))
                  '  oSheet.Cells(i + 3, j + 1).NumberFormat = "YYYY-MM-dd HH:mm"
                End If
            End If
  
                
    Next j
  Next i

    
    oSheet.Range("A:A").NumberFormat = "yyyy/MM/dd HH:mm:ss"
    xlLibro.SaveAs ("C:\ECOGAS\REPORTE\postgres\postgres_" & sType & Dia_a & hora_a & ".xlsx")
    Ruta_archivo = System.MyNodeName
    
    Select Case System.MyNodeName
        Case "SRV01"
            MsgBox "Se generó el siguiente archivo" + Chr(13) + " \\ifix-srv01\Postgres\postgres_" & sType & Dia_a & hora_a & ".xlsx"
        Case "SRV02"
            MsgBox "Se generó el siguiente archivo" + Chr(13) + " \\ifix-srv02\Postgres\postgres_" & sType & Dia_a & hora_a & ".xlsx"
        Case "SRV03"
            MsgBox "Se generó el siguiente archivo" + Chr(13) + " \\ifix-srv03\Postgres\postgres_" & sType & Dia_a & hora_a & ".xlsx"
        Case "PC02"
            MsgBox "Se generó el siguiente archivo" + Chr(13) + " \\cpl-scada04a\Postgres\postgres_" & sType & Dia_a & hora_a & ".xlsx"
        Case "PC01"
            MsgBox "Se generó el siguiente archivo" + Chr(13) + " \\cpl-scada03a\Postgres\postgres_" & sType & Dia_a & hora_a & ".xlsx"
        Case Else
            MsgBox "Se generó el siguiente archivo" + Chr(13) + " \\ifix-srv03\postgres\postgres_" & sType & Dia_a & hora_a & ".xlsx"
    End Select
    

    
    
    'MsgBox "Se generó el siguiente archivo" + Chr(13) + " C:\ECOGAS\REPORTE\postgres\postgres_" & sType & Dia_a & hora_a & ".xlsx"
    
    
    
   xlApp.Quit
End Function

Private Sub ToggleSortOrder(i As Integer)
    If i = 1 Then bOrden1Asc = Not bOrden1Asc
    Orden1Asc.Visible = bOrden1Asc
    Orden1Desc.Visible = Not bOrden1Asc
End Sub
