Option Explicit

Const THE_NAME = 1
Const THE_DESC = 2
Const THE_COMMA = ", "
Const THE_COMMA_CRLF = ", " & vbCrLf

Const AREAS_ALARMA = 192

Const DT_FORMAT = "d/MMM/yy H:mm:ss"
Const NO_QUERY = "¡¡NO QUERY!!"
Const FieldQty As Integer = 50
Const SortCritQty As Integer = 3
Const TABU = "     "
Dim AlarmAreas(1 To 2, 1 To 200) As String

Dim bOrden1Asc As Boolean





Private Sub AlarmSummaryOCX1_AlarmListChanged()



End Sub

Private Sub AlarmSummaryOCX1_NewAlarm(strNode As String, strTag As String)
    SendAlarmMsg strNode, strTag
    
    
    
End Sub

Private Sub SendAlarmMsg(strNode As String, strTag As String)


'On Error GoTo HandError
On Error Resume Next
    'MsgBox (strNode & " " & strTag)
    Dim n, i As Integer
    Dim Tags() As String
    Dim Nodes() As String
    Dim AlmTime As String
    Dim AlmStat As String
    Dim AlmDesc As String
    Dim AlmValue As String
    Dim AlmUnit As String

    Dim S As String
    Dim D As Object
    
    ' Para envío de Telegram
    Dim objRequest As Object
    Dim strChatId As String
    Dim strMessage As String
    Dim strPostData As String
    Dim strResponse As String
    
    n = Len(strTag) / 30
    'MsgBox "Number of Alarms triggering NewAlarm event is " & n
    ReDim Tags(1 To n)
    ReDim Nodes(1 To n)
    
   
    For i = 1 To n
        S = Mid(strNode, ((i - 1) * 8) + 1, 8)
        Nodes(i) = Trim(S)
        S = Mid(strTag, ((i - 1) * 30) + 1, 30)
        Tags(i) = Trim(S)
        
        
        AlmStat = "Fix32." & Nodes(i) & "." & Tags(i) & ".A_CUALM"
        Set D = System.FindObject(AlmStat)
        AlmStat = D

        If Not (InStr(1, AlmStat, "COMM") > 0) Then 'excluyo los COM
            
            AlmTime = "Fix32." & Nodes(i) & "." & Tags(i) & ".A_ALMINTIME"
            Set D = System.FindObject(AlmTime)
            AlmTime = D
    

    
            AlmDesc = "Fix32." & Nodes(i) & "." & Tags(i) & ".A_DESC"
            Set D = System.FindObject(AlmDesc)
            AlmDesc = D
    

   
            AlmUnit = "Fix32." & Nodes(i) & "." & Tags(i) & ".A_EGUDESC"
    
            Set D = System.FindObject(AlmUnit)
    
            AlmUnit = D
            

       
            AlmValue = "Fix32." & Nodes(i) & "." & Tags(i) & ".F_CV"
       
            Set D = System.FindObject(AlmValue)
           
            AlmValue = D

    
     
                'Envío por telegram
                strChatId = -328253192
                strMessage = " " & AlmTime & " " & Tags(i) & "-" & AlmDesc & ": " & AlmValue & " " & AlmUnit & "-Estado: " & AlmStat
    
            


            'MsgBox Nodes(i) & "--" & Tags(i) & "--" & AlmStat & "--" & AlmTime & "--" & AlmDesc
           
        
            strPostData = "{""alarm"": """ & strMessage & """}"
                

            Set objRequest = CreateObject("Msxml2.XMLHTTP.6.0")
      

            With objRequest
         

                .Open "POST", "http://10.2.9.17:5000/mgateway/v1.0/tasks", False
                'ok   .Open "POST", "https://api.telegram.org/bot996303276:AAGzG4KOztClAl2jbjG99BPSh3QzSANW9Lw/sendMessage?", False
                ' .Open "POST", "https://api.telegram.org/bot996303276:AAGzG4KOztClAl2jbjG99BPSh3QzSANW9Lw/getMe", False
                '{"ok":true,"result":{"id":996303276,"is_bot":true,"first_name":"scadatest","username":"scadatest1_bot", _
                '"can_join_groups":true,"can_read_all_group_messages":false,"supports_inline_queries":false}}
                .SetRequestHeader "Content-Type", "application/json"
                .Send (strPostData)
                'GetSessionId = .responseText
                'MsgBox GetSessionId
                'MsgBox .ResponseText
            End With
        End If
    Next i
HandError:
    'HandleError
   ' MsgBox ("Se ha producido un error. Tipo de error = " & Err.Number & ". Descripción: " & Err.description)

End Sub



Private Sub AlarmSummaryOCX1_SeverityIncreased(strNode As String, strTag As String)

    SendAlarmMsg strNode, strTag
End Sub

Private Sub CFixPicture_Initialize()

 On Error GoTo HandError
    
    AlarmSummaryOCX1.CheckForNewAlarms = True
    AlarmSummaryOCX1.CheckForSeverityIncrease = True
    
    InitializeArrays
    'Agregar las opciones de ordenamiento al combo list
    cmbSortList.Clear
    cmbSortList.AddItem "Time In"
    cmbSortList.AddItem "Time Last"
    cmbSortList.AddItem "Block Type"
    cmbSortList.AddItem "Tagname"
    cmbSortList.AddItem "Priority"
    cmbSortList.AddItem "Node"
    cmbSortList.AddItem "Ack/Time"
    cmbSortList.AddItem "Ack/Priority"
    'Seleccionar la 1ra. opcion
    cmbSortList.ListIndex = 1

    CargarListaCampos lbEstaciones, AlarmAreas, AREAS_ALARMA
    
    
    lbEstaciones.Selected(0) = True
    lbEstaciones.Selected(2) = True
    
    AlarmSummaryOCX1.FilterString = "Area In ""CAMARAS"" AND Priority >= ""LOLO"""
   ' AlarmSummaryOCX1.FilterString = "Area In ""CAMARAS"""
    'AlarmSummaryOCX1.FilterString = "Area In ""CAMARAS"""
    
    Exit Sub
    
HandError:
    HandleError
    
    
End Sub



Private Sub btnConsultar_Borrar_Todas_Alarmas_Click()

    If System.LoginUserName = "USER" Or System.LoginUserName = "OREC" Or System.LoginUserName = "ADMIN" Or System.LoginUserName = "ADMLOC" Then
        AlarmSummaryOCX1.AckAllAlarms
        AlarmSummaryOCX1.DeleteAllAlarms
    Else
    MsgBox ("No tiene los Permisos para realizar la operación.")
    End If

End Sub




Private Sub btnConsultar_Borrar_Alarmas_Click()
    Dim sNode As String
    Dim sTag As String

    If System.LoginUserName = "USER" Or System.LoginUserName = "OREC" Or System.LoginUserName = "ADMIN" Or System.LoginUserName = "ADMLOC" Then
        AlarmSummaryOCX1.GetSelectedNodeTag sNode, sTag
        If (sNode <> "" And sTag <> "") Then
            AlarmSummaryOCX1.DelAlarm sNode, sTag
        End If
    Else
    MsgBox ("No tiene los Permisos para realizar la operación.")
    End If

End Sub




Private Sub btnConsultar2_Click()

'Reconocer la alarma seleccionada
 Dim sNode As String
 Dim sTag As String
 Dim boolTagSelec As Boolean
    
    If System.LoginUserName = "USER" Or System.LoginUserName = "OREC" Or System.LoginUserName = "ADMIN" Or System.LoginUserName = "ADMLOC" Then
        boolTagSelec = AlarmSummaryOCX1.GetSelectedNodeTag(sNode, sTag)
        If boolTagSelec Then
            AlarmSummaryOCX1.AckAlarm sNode, sTag
        End If
    Else
    MsgBox ("No tiene los Permisos para realizar la operación.")
    End If
    
    
    
    
End Sub


Private Sub btnConsultar1_Click()
    'Reconocer Todas las alarmas filtradas
    If System.LoginUserName = "USER" Or System.LoginUserName = "OREC" Or System.LoginUserName = "ADMIN" Or System.LoginUserName = "ADMLOC" Then
         AlarmSummaryOCX1.AckAllAlarms
    Else
         MsgBox ("No tiene los Permisos para realizar la operación.")
    End If
End Sub




Private Sub CheckBox1_Click()

End Sub

Private Sub Group_return_Click()

    ClosePicture
End Sub







Private Sub btnConsultar6Click()

Dim oExcel As Object
Dim xlApp As Excel.Application
Dim xlLibro As Excel.Workbook
Dim oBook As Object
Dim oSheet As Object

Dim i, iCount As Integer
Dim area As String
Dim dateIn As String
Dim timeIn As String
Dim dateLast As String
Dim timeLast As String
Dim status As String
Dim description As String
Dim value As String
Dim nada As String
Dim ack As Boolean

Dim Dia_a
Dim hora_a
Dim DIA
Dim hora


Dim sUserID As String
Dim sUserName As String
Dim sGroupName As String


On Error Resume Next

DIA = Format(Date, "YYYY-MM-dd")
hora = Format(Time, "HH:mm")
Dia_a = Format(Date, "_YYYY-MM-dd")
hora_a = Format(Time, "_HH-mm")



  'Start a new workbook in Excel
  Set xlApp = New Excel.Application
  Set xlLibro = xlApp.Workbooks.Open("C:\ECOGAS\REPORTE\Plantillas\Alarmas.xlsx", True, True, , "")

  Set oSheet = xlLibro.Worksheets("Alarmas")

    iCount = AlarmSummaryOCX1.TotalFilteredAlarms
    For i = 1 To iCount
        AlarmSummaryOCX1.SelectAlarmRow i, True
        AlarmSummaryOCX1.GetSelectedRow ack, nada, area, dateIn, dateLast, timeIn, timeLast, nada, nada, nada, status, description, value, nada, nada, nada, nada
        oSheet.Range("A" + CStr(i + 4)).value = IIf(ack, "ACK", "NACK")
        oSheet.Range("B" + CStr(i + 4)).value = area
        oSheet.Range("C" + CStr(i + 4)).value = timeLast
        oSheet.Range("D" + CStr(i + 4)).value = Format(dateLast, "YYYY-MM-dd")
        oSheet.Range("E" + CStr(i + 4)).value = timeIn
        oSheet.Range("F" + CStr(i + 4)).value = Format(dateIn, "YYYY-MM-dd")
        oSheet.Range("G" + CStr(i + 4)).value = description
        oSheet.Range("H" + CStr(i + 4)).value = value
        oSheet.Range("I" + CStr(i + 4)).value = status
        
    Next
   AlarmSummaryOCX1.SelectAlarmRow i - 1, False
   
   System.FixGetUserInfo sUserID, sUserName, sGroupName
   oSheet.Range("A2").value = "Usuario: " & sUserName
   oSheet.Range("A3").value = "Fecha: " & DIA & " " & hora
   
   'oSheet.Range("C:C").Select
   'oSheet.Range("D4").Activate
   'oSheet.Selection.NumberFormat = "hh:mm"
    
    
   oSheet.PageSetup.printArea = "=A1:J" & CStr(iCount + 4)
   
   If CheckBox1.value Then
    oSheet.PrintOut
   End If
   'SALVO,CIERRO E IMPRIMO ARCHIVO EXCEL TODO EN 1
   xlLibro.SaveAs ("C:\ECOGAS\REPORTE\Alarmas\Alarmas" & Dia_a & hora_a & ".xlsx")
    Select Case System.MyNodeName
        Case "SRV01"
            MsgBox "Se generó el siguiente archivo" + Chr(13) + " \\ifix-srv01\Alarmas\Alarmas" & Dia_a & hora_a & ".xlsx"
        Case "SRV02"
            MsgBox "Se generó el siguiente archivo" + Chr(13) + " \\ifix-srv02\Alarmas\Alarmas" & Dia_a & hora_a & ".xlsx"
        Case "SRV03"
            MsgBox "Se generó el siguiente archivo" + Chr(13) + " \\ifix-srv03\Alarmas\Alarmas" & Dia_a & hora_a & ".xlsx"
        Case "PC02"
            MsgBox "Se generó el siguiente archivo" + Chr(13) + " \\cpl-scada04\Alarmas\Alarmas" & Dia_a & hora_a & ".xlsx"
        Case "PC01"
            MsgBox "Se generó el siguiente archivo" + Chr(13) + " \\cpl-scada03\Alarmas\Alarmas" & Dia_a & hora_a & ".xlsx"
        Case Else
            MsgBox "Se generó el siguiente archivo" + Chr(13) + " C:\ECOGAS\REPORTE\Alarmas\Alarmas" & Dia_a & hora_a & ".xlsx"
    End Select
   'MsgBox "Se generó el siguiente archivo" + Chr(13) + " C:\ECOGAS\REPORTE\Alarmas\Alarmas" & Dia_a & hora_a & ".xlsx"
   
   xlApp.Quit
End Sub






Private Sub btnConsultar_Click()
    Dim sListaAreas As String
    Dim i As Integer
    sListaAreas = ""
    With lbEstaciones
        For i = 0 To .ListCount - 1
            If .Selected(i) Then
                sListaAreas = sListaAreas & "Area In " + Chr(34) & .List(i, 1) & Chr(34) & " OR "
            End If
        Next i
    End With
    sListaAreas = Left(sListaAreas, Len(sListaAreas) - 4)
    
    AlarmSummaryOCX1.FilterString = sListaAreas
    
   
End Sub

Private Sub cmbSortList_Change()
    'Reordenar la lista
    If cmbSortList.Text <> "" Then
        AlarmSummaryOCX1.SortColumnName = cmbSortList.Text
    End If
End Sub


Private Sub Fondo6_Click()

End Sub

Private Sub Fondo7_Click()

End Sub

Private Sub Group5_Click()

End Sub

Private Sub grpOrdenamiento_Click()

End Sub

Private Sub optSortAscending_Click()
    AlarmSummaryOCX1.SortOrderAscending = True
    optSortDescending.value = False
End Sub


Private Sub optSortDescending_Click()
    
    AlarmSummaryOCX1.SortOrderAscending = False
    optSortAscending.value = False

End Sub


Private Sub gListaAreas_Click()

End Sub


Private Sub lbEstaciones_Change()

End Sub

Private Sub lbEstaciones_Click()

End Sub




Private Sub InitializeArrays()
Dim i As Integer

i = 1:           AlarmAreas(THE_NAME, i) = "CAMARAS":     AlarmAreas(THE_DESC, i) = "CAMARAS"
i = 2:           AlarmAreas(THE_NAME, i) = "GU":          AlarmAreas(THE_DESC, i) = "GU"
i = 3:           AlarmAreas(THE_NAME, i) = "SCADA":        AlarmAreas(THE_DESC, i) = "SCADA"
i = 4:           AlarmAreas(THE_NAME, i) = "R002":        AlarmAreas(THE_DESC, i) = "R002"
i = 5:           AlarmAreas(THE_NAME, i) = "R003":        AlarmAreas(THE_DESC, i) = "R003"
i = 6:           AlarmAreas(THE_NAME, i) = "R004":        AlarmAreas(THE_DESC, i) = "R004"
i = 7:           AlarmAreas(THE_NAME, i) = "R005":        AlarmAreas(THE_DESC, i) = "R005"
i = 8:           AlarmAreas(THE_NAME, i) = "R006":        AlarmAreas(THE_DESC, i) = "R006"
i = 9:           AlarmAreas(THE_NAME, i) = "R007":        AlarmAreas(THE_DESC, i) = "R007"
i = 10:          AlarmAreas(THE_NAME, i) = "R008":        AlarmAreas(THE_DESC, i) = "R008"
i = 11:          AlarmAreas(THE_NAME, i) = "R009":        AlarmAreas(THE_DESC, i) = "R009"
i = 12:          AlarmAreas(THE_NAME, i) = "R010":        AlarmAreas(THE_DESC, i) = "R010"
i = 13:          AlarmAreas(THE_NAME, i) = "R011":        AlarmAreas(THE_DESC, i) = "R011"
i = 14:          AlarmAreas(THE_NAME, i) = "R012":        AlarmAreas(THE_DESC, i) = "R012"
i = 15:          AlarmAreas(THE_NAME, i) = "R013":        AlarmAreas(THE_DESC, i) = "R013"
i = 16:          AlarmAreas(THE_NAME, i) = "R014":        AlarmAreas(THE_DESC, i) = "R014"
i = 17:          AlarmAreas(THE_NAME, i) = "R015":        AlarmAreas(THE_DESC, i) = "R015"
i = 18:          AlarmAreas(THE_NAME, i) = "R016":        AlarmAreas(THE_DESC, i) = "R016"
i = 19:          AlarmAreas(THE_NAME, i) = "R017":        AlarmAreas(THE_DESC, i) = "R017"
i = 20:          AlarmAreas(THE_NAME, i) = "R018":        AlarmAreas(THE_DESC, i) = "R018"
i = 21:          AlarmAreas(THE_NAME, i) = "R019":        AlarmAreas(THE_DESC, i) = "R019"
i = 22:          AlarmAreas(THE_NAME, i) = "R020":        AlarmAreas(THE_DESC, i) = "R020"
i = 23:          AlarmAreas(THE_NAME, i) = "R021":        AlarmAreas(THE_DESC, i) = "R021"
i = 24:          AlarmAreas(THE_NAME, i) = "R022":        AlarmAreas(THE_DESC, i) = "R022"
i = 25:          AlarmAreas(THE_NAME, i) = "R023":        AlarmAreas(THE_DESC, i) = "R023"
i = 26:          AlarmAreas(THE_NAME, i) = "R024":        AlarmAreas(THE_DESC, i) = "R024"
i = 27:          AlarmAreas(THE_NAME, i) = "R025":        AlarmAreas(THE_DESC, i) = "R025"
i = 28:          AlarmAreas(THE_NAME, i) = "R026":        AlarmAreas(THE_DESC, i) = "R026"
i = 29:          AlarmAreas(THE_NAME, i) = "R027":        AlarmAreas(THE_DESC, i) = "R027"
i = 30:          AlarmAreas(THE_NAME, i) = "R028":        AlarmAreas(THE_DESC, i) = "R028"
i = 31:          AlarmAreas(THE_NAME, i) = "R029":        AlarmAreas(THE_DESC, i) = "R029"
i = 32:          AlarmAreas(THE_NAME, i) = "R030":        AlarmAreas(THE_DESC, i) = "R030"
i = 33:          AlarmAreas(THE_NAME, i) = "R031":        AlarmAreas(THE_DESC, i) = "R031"
i = 34:          AlarmAreas(THE_NAME, i) = "R032":        AlarmAreas(THE_DESC, i) = "R032"
i = 35:          AlarmAreas(THE_NAME, i) = "R033":        AlarmAreas(THE_DESC, i) = "R033"
i = 36:          AlarmAreas(THE_NAME, i) = "R034":        AlarmAreas(THE_DESC, i) = "R034"
i = 37:          AlarmAreas(THE_NAME, i) = "R035":        AlarmAreas(THE_DESC, i) = "R035"
i = 38:          AlarmAreas(THE_NAME, i) = "R036":        AlarmAreas(THE_DESC, i) = "R036"
i = 39:          AlarmAreas(THE_NAME, i) = "R037":        AlarmAreas(THE_DESC, i) = "R037"
i = 40:          AlarmAreas(THE_NAME, i) = "R038":        AlarmAreas(THE_DESC, i) = "R038"
i = 41:          AlarmAreas(THE_NAME, i) = "R039":        AlarmAreas(THE_DESC, i) = "R039"
i = 42:          AlarmAreas(THE_NAME, i) = "R040":        AlarmAreas(THE_DESC, i) = "R040"
i = 43:          AlarmAreas(THE_NAME, i) = "R041":        AlarmAreas(THE_DESC, i) = "R041"
i = 44:          AlarmAreas(THE_NAME, i) = "R042":        AlarmAreas(THE_DESC, i) = "R042"
i = 45:          AlarmAreas(THE_NAME, i) = "R043":        AlarmAreas(THE_DESC, i) = "R043"
i = 46:          AlarmAreas(THE_NAME, i) = "R044":        AlarmAreas(THE_DESC, i) = "R044"
i = 47:          AlarmAreas(THE_NAME, i) = "R045":        AlarmAreas(THE_DESC, i) = "R045"
i = 48:          AlarmAreas(THE_NAME, i) = "R046":        AlarmAreas(THE_DESC, i) = "R046"
i = 49:          AlarmAreas(THE_NAME, i) = "R047":        AlarmAreas(THE_DESC, i) = "R047"
i = 50:          AlarmAreas(THE_NAME, i) = "R048":        AlarmAreas(THE_DESC, i) = "R048"
i = 51:          AlarmAreas(THE_NAME, i) = "R049":        AlarmAreas(THE_DESC, i) = "R049"
i = 52:          AlarmAreas(THE_NAME, i) = "R050":        AlarmAreas(THE_DESC, i) = "R050"
i = 53:          AlarmAreas(THE_NAME, i) = "R051":        AlarmAreas(THE_DESC, i) = "R051"
i = 54:          AlarmAreas(THE_NAME, i) = "R052":        AlarmAreas(THE_DESC, i) = "R052"
i = 55:          AlarmAreas(THE_NAME, i) = "R053":        AlarmAreas(THE_DESC, i) = "R053"
i = 56:          AlarmAreas(THE_NAME, i) = "R054":        AlarmAreas(THE_DESC, i) = "R054"
i = 57:          AlarmAreas(THE_NAME, i) = "R055":        AlarmAreas(THE_DESC, i) = "R055"
i = 58:          AlarmAreas(THE_NAME, i) = "R056":        AlarmAreas(THE_DESC, i) = "R056"
i = 59:          AlarmAreas(THE_NAME, i) = "R057":        AlarmAreas(THE_DESC, i) = "R057"
i = 60:          AlarmAreas(THE_NAME, i) = "R058":        AlarmAreas(THE_DESC, i) = "R058"
i = 61:          AlarmAreas(THE_NAME, i) = "R059":        AlarmAreas(THE_DESC, i) = "R059"
i = 62:          AlarmAreas(THE_NAME, i) = "R060":        AlarmAreas(THE_DESC, i) = "R060"
i = 63:          AlarmAreas(THE_NAME, i) = "R061":        AlarmAreas(THE_DESC, i) = "R061"
i = 64:          AlarmAreas(THE_NAME, i) = "R062":        AlarmAreas(THE_DESC, i) = "R062"
i = 65:          AlarmAreas(THE_NAME, i) = "R063":        AlarmAreas(THE_DESC, i) = "R063"
i = 66:          AlarmAreas(THE_NAME, i) = "R064":        AlarmAreas(THE_DESC, i) = "R064"
i = 67:          AlarmAreas(THE_NAME, i) = "R065":        AlarmAreas(THE_DESC, i) = "R065"
i = 68:          AlarmAreas(THE_NAME, i) = "R066":        AlarmAreas(THE_DESC, i) = "R066"
i = 69:          AlarmAreas(THE_NAME, i) = "R067":        AlarmAreas(THE_DESC, i) = "R067"
i = 70:          AlarmAreas(THE_NAME, i) = "R068":        AlarmAreas(THE_DESC, i) = "R068"
i = 71:          AlarmAreas(THE_NAME, i) = "R069":        AlarmAreas(THE_DESC, i) = "R069"
i = 72:          AlarmAreas(THE_NAME, i) = "R070":        AlarmAreas(THE_DESC, i) = "R070"
i = 73:          AlarmAreas(THE_NAME, i) = "R071":        AlarmAreas(THE_DESC, i) = "R071"
i = 74:          AlarmAreas(THE_NAME, i) = "R072":        AlarmAreas(THE_DESC, i) = "R072"
i = 75:          AlarmAreas(THE_NAME, i) = "R073":        AlarmAreas(THE_DESC, i) = "R073"
i = 76:          AlarmAreas(THE_NAME, i) = "R074":        AlarmAreas(THE_DESC, i) = "R074"
i = 77:          AlarmAreas(THE_NAME, i) = "R075":        AlarmAreas(THE_DESC, i) = "R075"
i = 78:          AlarmAreas(THE_NAME, i) = "R076":        AlarmAreas(THE_DESC, i) = "R076"
i = 79:          AlarmAreas(THE_NAME, i) = "R077":        AlarmAreas(THE_DESC, i) = "R077"
i = 80:          AlarmAreas(THE_NAME, i) = "R078":        AlarmAreas(THE_DESC, i) = "R078"
i = 81:          AlarmAreas(THE_NAME, i) = "R079":        AlarmAreas(THE_DESC, i) = "R079"
i = 82:          AlarmAreas(THE_NAME, i) = "R080":        AlarmAreas(THE_DESC, i) = "R080"
i = 83:          AlarmAreas(THE_NAME, i) = "R081":        AlarmAreas(THE_DESC, i) = "R081"
i = 84:          AlarmAreas(THE_NAME, i) = "R082":        AlarmAreas(THE_DESC, i) = "R082"
i = 85:          AlarmAreas(THE_NAME, i) = "R083":        AlarmAreas(THE_DESC, i) = "R083"
i = 86:          AlarmAreas(THE_NAME, i) = "R084":        AlarmAreas(THE_DESC, i) = "R084"
i = 87:          AlarmAreas(THE_NAME, i) = "R085":        AlarmAreas(THE_DESC, i) = "R085"
i = 88:          AlarmAreas(THE_NAME, i) = "R086":        AlarmAreas(THE_DESC, i) = "R086"
i = 89:          AlarmAreas(THE_NAME, i) = "R087":        AlarmAreas(THE_DESC, i) = "R087"
i = 90:          AlarmAreas(THE_NAME, i) = "R088":        AlarmAreas(THE_DESC, i) = "R088"
i = 91:          AlarmAreas(THE_NAME, i) = "R089":        AlarmAreas(THE_DESC, i) = "R089"
i = 92:          AlarmAreas(THE_NAME, i) = "R090":        AlarmAreas(THE_DESC, i) = "R090"
i = 93:          AlarmAreas(THE_NAME, i) = "E001":        AlarmAreas(THE_DESC, i) = "E001"
i = 94:          AlarmAreas(THE_NAME, i) = "E002":        AlarmAreas(THE_DESC, i) = "E002"
i = 95:          AlarmAreas(THE_NAME, i) = "E003":        AlarmAreas(THE_DESC, i) = "E003"
i = 96:          AlarmAreas(THE_NAME, i) = "E004":        AlarmAreas(THE_DESC, i) = "E004"
i = 97:          AlarmAreas(THE_NAME, i) = "E005":        AlarmAreas(THE_DESC, i) = "E005"
i = 98:          AlarmAreas(THE_NAME, i) = "E006":        AlarmAreas(THE_DESC, i) = "E006"
i = 99:          AlarmAreas(THE_NAME, i) = "E007":        AlarmAreas(THE_DESC, i) = "E007"
i = 100:          AlarmAreas(THE_NAME, i) = "E008":        AlarmAreas(THE_DESC, i) = "E008"
i = 101:          AlarmAreas(THE_NAME, i) = "E009":        AlarmAreas(THE_DESC, i) = "E009"
i = 102:          AlarmAreas(THE_NAME, i) = "E010":        AlarmAreas(THE_DESC, i) = "E010"
i = 103:          AlarmAreas(THE_NAME, i) = "E011":        AlarmAreas(THE_DESC, i) = "E011"
i = 104:          AlarmAreas(THE_NAME, i) = "E012":        AlarmAreas(THE_DESC, i) = "E012"
i = 105:          AlarmAreas(THE_NAME, i) = "E013":        AlarmAreas(THE_DESC, i) = "E013"
i = 106:          AlarmAreas(THE_NAME, i) = "E014":        AlarmAreas(THE_DESC, i) = "E014"
i = 107:          AlarmAreas(THE_NAME, i) = "E015":        AlarmAreas(THE_DESC, i) = "E015"
i = 108:          AlarmAreas(THE_NAME, i) = "E016":        AlarmAreas(THE_DESC, i) = "E016"
i = 109:          AlarmAreas(THE_NAME, i) = "E017":        AlarmAreas(THE_DESC, i) = "E017"
i = 110:          AlarmAreas(THE_NAME, i) = "E018":        AlarmAreas(THE_DESC, i) = "E018"
i = 111:          AlarmAreas(THE_NAME, i) = "E019":        AlarmAreas(THE_DESC, i) = "E019"
i = 112:          AlarmAreas(THE_NAME, i) = "E020":        AlarmAreas(THE_DESC, i) = "E020"
i = 113:          AlarmAreas(THE_NAME, i) = "E021":        AlarmAreas(THE_DESC, i) = "E021"
i = 114:          AlarmAreas(THE_NAME, i) = "E022":        AlarmAreas(THE_DESC, i) = "E022"
i = 115:          AlarmAreas(THE_NAME, i) = "E023":        AlarmAreas(THE_DESC, i) = "E023"
i = 116:          AlarmAreas(THE_NAME, i) = "E024":        AlarmAreas(THE_DESC, i) = "E024"
i = 117:          AlarmAreas(THE_NAME, i) = "E025":        AlarmAreas(THE_DESC, i) = "E025"
i = 118:          AlarmAreas(THE_NAME, i) = "E026":        AlarmAreas(THE_DESC, i) = "E026"
i = 119:          AlarmAreas(THE_NAME, i) = "E027":        AlarmAreas(THE_DESC, i) = "E027"
i = 120:          AlarmAreas(THE_NAME, i) = "E028":        AlarmAreas(THE_DESC, i) = "E028"
i = 121:          AlarmAreas(THE_NAME, i) = "E029":        AlarmAreas(THE_DESC, i) = "E029"
i = 122:          AlarmAreas(THE_NAME, i) = "E030":        AlarmAreas(THE_DESC, i) = "E030"
i = 123:          AlarmAreas(THE_NAME, i) = "E031":        AlarmAreas(THE_DESC, i) = "E031"
i = 124:          AlarmAreas(THE_NAME, i) = "E032":        AlarmAreas(THE_DESC, i) = "E032"
i = 125:          AlarmAreas(THE_NAME, i) = "E033":        AlarmAreas(THE_DESC, i) = "E033"
i = 126:          AlarmAreas(THE_NAME, i) = "E034":        AlarmAreas(THE_DESC, i) = "E034"
i = 127:          AlarmAreas(THE_NAME, i) = "E035":        AlarmAreas(THE_DESC, i) = "E035"
i = 128:          AlarmAreas(THE_NAME, i) = "E036":        AlarmAreas(THE_DESC, i) = "E036"
i = 129:          AlarmAreas(THE_NAME, i) = "E037":        AlarmAreas(THE_DESC, i) = "E037"
i = 130:          AlarmAreas(THE_NAME, i) = "E038":        AlarmAreas(THE_DESC, i) = "E038"
i = 131:          AlarmAreas(THE_NAME, i) = "E039":        AlarmAreas(THE_DESC, i) = "E039"
i = 132:          AlarmAreas(THE_NAME, i) = "E040":        AlarmAreas(THE_DESC, i) = "E040"
i = 133:          AlarmAreas(THE_NAME, i) = "E041":        AlarmAreas(THE_DESC, i) = "E041"
i = 134:          AlarmAreas(THE_NAME, i) = "E042":        AlarmAreas(THE_DESC, i) = "E042"
i = 135:          AlarmAreas(THE_NAME, i) = "E043":        AlarmAreas(THE_DESC, i) = "E043"
i = 136:          AlarmAreas(THE_NAME, i) = "E044":        AlarmAreas(THE_DESC, i) = "E044"
i = 137:          AlarmAreas(THE_NAME, i) = "E045":        AlarmAreas(THE_DESC, i) = "E045"
i = 138:          AlarmAreas(THE_NAME, i) = "E046":        AlarmAreas(THE_DESC, i) = "E046"
i = 139:          AlarmAreas(THE_NAME, i) = "E047":        AlarmAreas(THE_DESC, i) = "E047"
i = 140:          AlarmAreas(THE_NAME, i) = "E048":        AlarmAreas(THE_DESC, i) = "E048"
i = 141:          AlarmAreas(THE_NAME, i) = "E049":        AlarmAreas(THE_DESC, i) = "E049"
i = 142:          AlarmAreas(THE_NAME, i) = "E050":        AlarmAreas(THE_DESC, i) = "E050"
i = 143:          AlarmAreas(THE_NAME, i) = "E051":        AlarmAreas(THE_DESC, i) = "E051"
i = 144:          AlarmAreas(THE_NAME, i) = "E052":        AlarmAreas(THE_DESC, i) = "E052"
i = 145:          AlarmAreas(THE_NAME, i) = "E053":        AlarmAreas(THE_DESC, i) = "E053"
i = 146:          AlarmAreas(THE_NAME, i) = "E054":        AlarmAreas(THE_DESC, i) = "E054"
i = 147:          AlarmAreas(THE_NAME, i) = "E055":        AlarmAreas(THE_DESC, i) = "E055"
i = 148:          AlarmAreas(THE_NAME, i) = "E056":        AlarmAreas(THE_DESC, i) = "E056"
i = 149:          AlarmAreas(THE_NAME, i) = "E057":        AlarmAreas(THE_DESC, i) = "E057"
i = 150:          AlarmAreas(THE_NAME, i) = "E058":        AlarmAreas(THE_DESC, i) = "E058"
i = 151:          AlarmAreas(THE_NAME, i) = "E059":        AlarmAreas(THE_DESC, i) = "E059"
i = 152:          AlarmAreas(THE_NAME, i) = "E060":        AlarmAreas(THE_DESC, i) = "E060"
i = 153:          AlarmAreas(THE_NAME, i) = "E061":        AlarmAreas(THE_DESC, i) = "E061"
i = 154:          AlarmAreas(THE_NAME, i) = "E062":        AlarmAreas(THE_DESC, i) = "E062"
i = 155:          AlarmAreas(THE_NAME, i) = "E063":        AlarmAreas(THE_DESC, i) = "E063"
i = 156:          AlarmAreas(THE_NAME, i) = "E064":        AlarmAreas(THE_DESC, i) = "E064"
i = 157:          AlarmAreas(THE_NAME, i) = "E065":        AlarmAreas(THE_DESC, i) = "E065"
i = 158:          AlarmAreas(THE_NAME, i) = "E066":        AlarmAreas(THE_DESC, i) = "E066"
i = 159:          AlarmAreas(THE_NAME, i) = "E067":        AlarmAreas(THE_DESC, i) = "E067"
i = 160:          AlarmAreas(THE_NAME, i) = "E068":        AlarmAreas(THE_DESC, i) = "E068"
i = 161:          AlarmAreas(THE_NAME, i) = "E069":        AlarmAreas(THE_DESC, i) = "E069"
i = 162:          AlarmAreas(THE_NAME, i) = "E070":        AlarmAreas(THE_DESC, i) = "E070"
i = 163:          AlarmAreas(THE_NAME, i) = "E071":        AlarmAreas(THE_DESC, i) = "E071"
i = 164:          AlarmAreas(THE_NAME, i) = "E072":        AlarmAreas(THE_DESC, i) = "E072"
i = 165:          AlarmAreas(THE_NAME, i) = "E073":        AlarmAreas(THE_DESC, i) = "E073"
i = 166:          AlarmAreas(THE_NAME, i) = "E074":        AlarmAreas(THE_DESC, i) = "E074"
i = 167:          AlarmAreas(THE_NAME, i) = "E075":        AlarmAreas(THE_DESC, i) = "E075"
i = 168:          AlarmAreas(THE_NAME, i) = "E076":        AlarmAreas(THE_DESC, i) = "E076"
i = 169:          AlarmAreas(THE_NAME, i) = "E077":        AlarmAreas(THE_DESC, i) = "E077"
i = 170:          AlarmAreas(THE_NAME, i) = "E078":        AlarmAreas(THE_DESC, i) = "E078"
i = 171:          AlarmAreas(THE_NAME, i) = "E079":        AlarmAreas(THE_DESC, i) = "E079"
i = 172:          AlarmAreas(THE_NAME, i) = "E080":        AlarmAreas(THE_DESC, i) = "E080"
i = 173:          AlarmAreas(THE_NAME, i) = "E081":        AlarmAreas(THE_DESC, i) = "E081"
i = 174:          AlarmAreas(THE_NAME, i) = "E082":        AlarmAreas(THE_DESC, i) = "E082"
i = 175:          AlarmAreas(THE_NAME, i) = "E083":        AlarmAreas(THE_DESC, i) = "E083"
i = 176:          AlarmAreas(THE_NAME, i) = "E084":        AlarmAreas(THE_DESC, i) = "E084"
i = 177:          AlarmAreas(THE_NAME, i) = "E085":        AlarmAreas(THE_DESC, i) = "E085"
i = 178:          AlarmAreas(THE_NAME, i) = "E086":        AlarmAreas(THE_DESC, i) = "E086"
i = 179:          AlarmAreas(THE_NAME, i) = "E087":        AlarmAreas(THE_DESC, i) = "E087"
i = 180:          AlarmAreas(THE_NAME, i) = "E088":        AlarmAreas(THE_DESC, i) = "E088"
i = 181:          AlarmAreas(THE_NAME, i) = "E089":        AlarmAreas(THE_DESC, i) = "E089"
i = 182:          AlarmAreas(THE_NAME, i) = "E090":        AlarmAreas(THE_DESC, i) = "E090"
i = 183:          AlarmAreas(THE_NAME, i) = "E091":        AlarmAreas(THE_DESC, i) = "E091"
i = 184:          AlarmAreas(THE_NAME, i) = "E092":        AlarmAreas(THE_DESC, i) = "E092"
i = 185:          AlarmAreas(THE_NAME, i) = "E093":        AlarmAreas(THE_DESC, i) = "E093"
i = 186:          AlarmAreas(THE_NAME, i) = "E094":        AlarmAreas(THE_DESC, i) = "E094"
i = 187:          AlarmAreas(THE_NAME, i) = "E095":        AlarmAreas(THE_DESC, i) = "E095"
i = 188:          AlarmAreas(THE_NAME, i) = "E096":        AlarmAreas(THE_DESC, i) = "E096"
i = 189:          AlarmAreas(THE_NAME, i) = "E097":        AlarmAreas(THE_DESC, i) = "E097"
i = 190:          AlarmAreas(THE_NAME, i) = "E098":        AlarmAreas(THE_DESC, i) = "E098"
i = 191:          AlarmAreas(THE_NAME, i) = "E099":        AlarmAreas(THE_DESC, i) = "E099"
i = 192:          AlarmAreas(THE_NAME, i) = "PRUEBA":        AlarmAreas(THE_DESC, i) = "PRUEBA"

    

    
  
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

Private Sub btnConsultar6_Click()
    Dim oExcel As Object
    Dim oBook As Object
    Dim oSheet As Object
    Dim i, iCount As Integer

Dim area As String
Dim dateIn As String
Dim timeIn As String
Dim dateLast As String
Dim timeLast As String
Dim status As String
Dim description As String
Dim value As String
Dim nada As String
Dim ack As Boolean

Dim sUserID As String
Dim sUserName As String
Dim sGroupName As String
On Error Resume Next
   'Start a new workbook in Excel
    Set oExcel = CreateObject("Excel.Application")
    Set oBook = oExcel.Workbooks.Add
    Set oSheet = oBook.Worksheets.Add
   
    Dim Dia_a
    Dim hora_a
    Dim DIA
    Dim hora
    DIA = Format(Date, "YYYY-MM-dd")
    hora = Format(Time, "HH:mm")
    Dia_a = Format(Date, "_YYYY-MM-dd")
    hora_a = Format(Time, "_HH-mm")
   

        oSheet.Range("A" + CStr(4)).value = "ACK"
        oSheet.Range("B" + CStr(4)).value = "AREA"
        oSheet.Range("C" + CStr(4)).value = "TIMELAST"
        oSheet.Range("D" + CStr(4)).value = "FECHA"
        oSheet.Range("E" + CStr(4)).value = "TIME IN"
        oSheet.Range("F" + CStr(4)).value = "FECHA"
        oSheet.Range("G" + CStr(4)).value = "DESCRIPCION"
        oSheet.Range("H" + CStr(4)).value = "VALOR"
        oSheet.Range("I" + CStr(4)).value = "ESTADO"
        oSheet.Range("A4:I4").Font.Bold = True
    iCount = AlarmSummaryOCX1.TotalFilteredAlarms
    For i = 1 To iCount
        AlarmSummaryOCX1.SelectAlarmRow i, True
        AlarmSummaryOCX1.GetSelectedRow ack, nada, area, dateIn, dateLast, timeIn, timeLast, nada, nada, nada, status, description, value, nada, nada, nada, nada
        oSheet.Range("A" + CStr(i + 4)).value = IIf(ack, "ACK", "NACK")
        oSheet.Range("B" + CStr(i + 4)).value = area
        oSheet.Range("C" + CStr(i + 4)).value = timeLast
        oSheet.Range("D" + CStr(i + 4)).value = Format(dateLast, "YYYY-MM-dd")
        oSheet.Range("E" + CStr(i + 4)).value = timeIn
        oSheet.Range("F" + CStr(i + 4)).value = Format(dateIn, "YYYY-MM-dd")
        oSheet.Range("G" + CStr(i + 4)).value = description
        oSheet.Range("H" + CStr(i + 4)).value = value
        oSheet.Range("I" + CStr(i + 4)).value = status
        
    Next
   AlarmSummaryOCX1.SelectAlarmRow i - 1, False
   oSheet.Cells.EntireColumn.AutoFit
   
   System.FixGetUserInfo sUserID, sUserName, sGroupName
   oSheet.Range("A2").value = "Usuario: " & sUserName
   oSheet.Range("A3").value = "Generado: " & DIA & " " & hora
   oSheet.Range("A1").value = "DISTRIBUIDORA DE GAS DEL CENTRO S. A. - Resúmen de Alarmas"
   oSheet.Range("A1").Font.Bold = True
   
   'oSheet.Range("C:C").Select
   'oSheet.Range("D4").Activate
   'oSheet.Selection.NumberFormat = "hh:mm"
    
    
   oSheet.PageSetup.printArea = "=A1:J" & CStr(iCount + 4)
   
If CheckBox1.value Then
    oSheet.PrintOut
   End If
   'SALVO,CIERRO E IMPRIMO ARCHIVO EXCEL TODO EN 1
   oBook.SaveAs ("C:\ECOGAS\REPORTE\Alarmas\Alarmas" & Dia_a & hora_a & ".xlsx")
    Select Case System.MyNodeName
        Case "SRV01"
            MsgBox "Se generó el siguiente archivo" + Chr(13) + " \\ifix-srv01\Alarmas\Alarmas" & Dia_a & hora_a & ".xlsx"
        Case "SRV02"
            MsgBox "Se generó el siguiente archivo" + Chr(13) + " \\ifix-srv02\Alarmas\Alarmas" & Dia_a & hora_a & ".xlsx"
        Case "SRV03"
            MsgBox "Se generó el siguiente archivo" + Chr(13) + " \\ifix-srv03\Alarmas\Alarmas" & Dia_a & hora_a & ".xlsx"
        Case "PC02"
            MsgBox "Se generó el siguiente archivo" + Chr(13) + " \\cpl-scada04\Alarmas\Alarmas" & Dia_a & hora_a & ".xlsx"
        Case "PC01"
            MsgBox "Se generó el siguiente archivo" + Chr(13) + " \\cpl-scada03\Alarmas\Alarmas" & Dia_a & hora_a & ".xlsx"
        Case Else
            MsgBox "Se generó el siguiente archivo" + Chr(13) + " C:\ECOGAS\REPORTE\Alarmas\Alarmas" & Dia_a & hora_a & ".xlsx"
    End Select
   'MsgBox "Se generó el siguiente archivo" + Chr(13) + " C:\ECOGAS\REPORTE\Alarmas\Alarmas" & Dia_a & hora_a & ".xlsx"
  
   oExcel.Quit
End Sub





Private Sub Boton_Prueba1_Click()

User.PicAnterior.CurrentValue = User.PicActual.CurrentValue
User.PicActual.CurrentValue = "Eco_sum_Alarmas_chart.grf"
replacepicture "Eco_sum_Alarmas_chart.grf"

End Sub

Private Sub PolyLine70_Click()
ClosePicture
End Sub
