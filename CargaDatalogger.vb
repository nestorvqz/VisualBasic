Public Sub CargarDatosDatalogger()

   Dim cnn As New ADODB.Connection
    Dim rst As New ADODB.Recordset
    Dim i, j, answer As Long
    Dim k As Integer
    Dim paso As Integer
    
    Dim Value(800000) As Double
    Dim Battery(800000) As Double
    Dim Times(800000) As Date
    Dim Quality(800000) As Long
    Dim DayTime1 As String
    
    Dim auxValue(8000) As Double
    Dim auxBattery(8000) As Double
    Dim auxTimes(8000) As Date
    Dim auxQuality(8000) As Long
    
    Dim strSQL As String
    
    Dim vtVal As Variant
    Dim vtBattery As Variant
    Dim vtDate As Variant
    Dim vtQual As Variant
    
    Dim sTable As String
    Dim sDateFrom As String
    Dim sDateTo As String
    Dim sOrderBy As String
    Dim sListaCampos As String
    Dim nvariable As Variant
    
    
    
    
    sTable = "datalogger.veribox_data"
    
    'formateo la fecha para consulta a bd Veribox
    sDateFrom = Format(Now() - 2, "YYYY-MM-dd HH:mm")
    sDateTo = Format(Now(), "YYYY-MM-dd HH:mm")
    
    
    sOrderBy = " ORDER BY datetime"
    sOrderBy = " GROUP BY datetime1 ORDER BY datetime1 "
  '  sOrderBy = " ORDER BY datetime"
    sListaCampos = " datetime, presion  "
sListaCampos = " *  "

    strSQL = "Select " & sListaCampos & " From """ + sTable + """ Where " & _
             "(""datetime"" Between '" & sDateFrom & "' and '" & sDateTo & "') " & sOrderBy & ";"
             
             
    'strSQL = "select datetime,stationname,vol_c,vol_nc from ucv.ucv_data order by datetime desc ""
    
    'strSQL = "select t.datetime, t.stationname,t.vol_c, t.vol_nc from ucv.ucv_data as t where date_part('hour',t.datetime)= 6 and date_part('day',now()-t.datetime)<1;"
  
    strSQL = "select * from datalogger.veribox_data ORDER BY datetime LIMIT 800000"
    
    strSQL = "Select " & sListaCampos & " From " + sTable + " Where " & _
             "(""datetime"" Between '" & sDateFrom & "' and '" & sDateTo & "') " & sOrderBy & " LIMIT 12000;"
    
      strSQL = "Select date_trunc('minute', datetime) as datetime1, avg(presion) AS presion1 " & " From " + sTable + " Where " & _
             "(""datetime"" Between '" & sDateFrom & "' and '" & sDateTo & "') " & sOrderBy & " LIMIT 12000;"
             
    strSQL = "Select " & sListaCampos & " From " + sTable + " Where " & _
             "(""datetime"" Between '" & sDateFrom & "' and '" & sDateTo & "') " & sOrderBy & " LIMIT 12000;"
    
      strSQL = "Select date_trunc('minute', datetime) as datetime1, avg(presion) AS presion1, min(batt_vb) AS battery1 " & " From " + sTable + " Where " & _
             "(""datetime"" Between '" & sDateFrom & "' and '" & sDateTo & "') " & sOrderBy & " LIMIT 12000;"
    'SELECT date_trunc('day', datetime) as datetime1, MIN(presion) AS presion1 FROM datalogger.veribox_data WHERE ("datetime" between '2021-11-01 06:00' AND '2021-11-29 06:00') GROUP BY datetime1 ORDER BY datetime1;


     
     'MsgBox (strSQL)
     
    cnn.ConnectionString = "Driver=PostgreSQL UNICODE;" & _
                           "Server=" & "10.2.8.30" & ";" & _
                           "Port=" & "5432" & ";" & _
                           "User Id=" & "postgres" & ";" & _
                           "Password=" & "ecopass" & ";" & _
                           "Database=" & "gas" & ";"
                           
    cnn.Open
    rst.Open strSQL, cnn
    i = 0
    

    If (cnn.State) Then
        If Not rst.EOF Then
            While (Not rst.EOF)
        
                   Value(i) = rst.Fields("presion1").Value ' Format(dtFechaInicial.Value, "YYYY-MM-dd HH:mm")
                   Battery(i) = rst.Fields("battery1").Value ' Format(dtFechaInicial.Value, "YYYY-MM-dd HH:mm")

                   DayTime1 = rst.Fields("datetime1").Value
                   DayTime1 = Format$(DayTime1, "dd/MM/yyyy HH:mm:ss")
                   
                   Times(i) = DayTime1
                   
        
                   Quality(i) = 192
                   i = i + 1
                   rst.MoveNext
        
            Wend
       End If
              
       
    End If

    
    
    If (i > 0) Then
    
     paso = Int(i / 1000)
     
     paso = 0
     k = 0
     For j = 0 To (i - 1)
     
         auxValue(k) = Value(j)
         auxTimes(k) = Times(j)
         auxBattery(k) = Battery(j) / 1000
         auxQuality(k) = 192
         k = k + 1
         j = j + paso
         
     Next j



     
     vtVal = auxValue
     vtBattery = auxBattery
     
     vtDate = auxTimes
     
     vtQual = auxQuality
     
     Pen13.SetSource "Presión", True
     Pen12.SetSource "TensiónBaterías", True
     Pen13.SetPenDataArray k, vtVal, vtDate, vtQual
     Pen12.SetPenDataArray k, vtBattery, vtDate, vtQual
     Pen12.StartTime = sDateFrom
     'Pen3.SetPenDataArray count - 1, vtVal, vtDate, vtQual
     Pen13.StartTime = sDateFrom
     '´'Pen3.HiLimit = 2
    ' Pen3.LoLimit = 0
     Chart4.Duration = DateDiff("s", sDateFrom, sDateTo)
     Chart4.RefreshChartData
    
   Else
        MsgBox ("SIN DATOS")
   End If
   rst.Close
   cnn.Close
End Sub
