Public Function cargarDatosPostgres(strSQL As String)

    Dim cnn As New ADODB.Connection
    Dim rst As New ADODB.Recordset
    Dim i, j, answer As Integer

    Dim xDate, xYear, xMonth, xDay, yDate, xResto As String

    
    
    cnn.ConnectionString = "Driver=PostgreSQL;" & _
                           "Server=" & "143.4.12.52" & ";" & _
                           "Port=" & "5434" & ";" & _
                           "User Id=" & "postgres" & ";" & _
                           "Password=" & "xxxxxxxxx" & ";" & _
                           "Database=" & "Gas_GNC" & ";"
                           
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
