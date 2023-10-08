Option Explicit

Sub UpdateMarket_Q1()

    Dim r As Long
    Dim ttr As Long
    Dim key As String
    Dim vl As String
    
    ttr = CountRows(ws_query, 1)
    
    For r = 2 To ttr
    
        If GetWorksheetColumnValue(ws_query, r, "Mercado") = "" Then
    
            key = GetWorksheetColumnValue(ws_query, r, "TipFt")
            vl = "BR"
            
            If key = "ZVEX" Or key = "ZEXT" Then
                vl = "EXPO"
            End If
            
            Call SetWorksheetColumnValue(ws_query, r, "Mercado", vl)
            
        End If
    
    Next

End Sub

Sub GetUserInvoiceROL_Q1()

    Dim r As Long
    Dim ttr As Long
    Dim key As String
    Dim i As Long

    Call ClearWorksheet(ws_temp, 1)
    Call ClearWorksheet(ws_vbrk, 2)
    Call EnvironmentVariables_SAP(True)
    
    ttr = CountRows(ws_query, 1)
    i = 0
    
    For r = 2 To ttr
    
        If GetWorksheetColumnValue(ws_query, r, "Usuário Fatura 2") = "" Then
        
            key = GetWorksheetColumnValue(ws_query, r, "Fatura 2")
                        
            i = i + 1
            ws_temp.Cells(i, 1) = key
            
            If i > 10000 Then
                Call UpdateVBRK_Q1
                Call UpdateUserInvoiceROL_Q1
                i = 0
            End If
            
        End If
    
    Next r
    
    If i > 1 Then
        Call UpdateVBRK_Q1
        Call UpdateUserInvoiceROL_Q1
        i = 0
    End If
        
    Call ClearWorksheet(ws_temp, 1)
    Call EnvironmentVariables_SAP(False)

End Sub

 Sub UpdateVBRK_Q1()

    Call ClearWorksheet(ws_vbrk, 2)
    
    ws_temp.Activate
    ws_temp.Range("A:A").Select
    Selection.Copy
    
    Call RunScriptSAP_ROL_F2_CREATOR_Q1
    Call SetTempData
    Call LoadTempData_ROL_F2_CREATOR_Q1
    Call ClearWorksheet(ws_temp, 1)

End Sub

Private Sub LoadTempData_ROL_F2_CREATOR_Q1()
    
    Dim i As Long
    Dim r As Long
    Dim ttr As Long

    Dim lineTemp As String
    Dim cls() As String
    Dim hdr() As String
    Dim c As Long
    Dim h As String
    Dim vl As String
                
    ttr = CountRows(ws_temp, 1)
    i = CountRows(ws_vbrk, 1)
    
    For r = 2 To ttr
    
        lineTemp = ws_temp.Cells(r, 1)
    
        If lineTemp Like "|*" Then
        
            cls = Split(lineTemp, "|")
        
            If IsArrayEmpty(hdr) Then
            
                hdr = cls
                
                For c = 0 To UBound(hdr)
                    hdr(c) = Trim(hdr(c))
                Next c
            
            Else
                
                For c = 0 To UBound(cls)
                                   
                    h = hdr(c)
                    vl = Trim(cls(c))
                                   
                    If h <> "" Then
                                                         
                        If h <> vl Then

                            Select Case h

                                Case "DocFat."
                                    i = i + 1
                                    Call SetWorksheetColumnValue(ws_vbrk, i, "Invoice", vl)
                                    
                                Case "Criado por"
                                    Call SetWorksheetColumnValue(ws_vbrk, i, "User", vl)
                                    
                                Case "Dt.criação"
                                    Call SetWorksheetColumnValue(ws_vbrk, i, "CreationDate", ConvertionOutputDate(vl))
                                    
                                Case "Data do faturamento"
                                    Call SetWorksheetColumnValue(ws_vbrk, i, "BillingDate", ConvertionOutputDate(vl))
                            
                            End Select
                            
                        End If
                        
                    End If
                    
                Next c
                
            End If
        
        End If
        
    Next r
    
End Sub

Sub UpdateUserInvoiceROL_Q1()

    Dim r As Long
    Dim ttr As Long
    Dim key As String
        
    Dim ttn As Long
    Dim items As Range
   ' Dim i As Variant
    
    'Dim c As Long
    Dim vl As String
    Dim dt_1 As Date
    Dim dt_2 As Date
    
    ttr = CountRows(ws_query, 1)
    ttn = CountRows(ws_vbrk, 1)
    
    Set items = ws_vbrk.Range("A1:D" & ttn)
    
    For r = 2 To ttr
    
        vl = ""
        dt_1 = 0
        dt_2 = 0
    
        '"Usuário Fatura 2"
        If GetWorksheetColumnValue(ws_query, r, "Usuário Fatura 2") = "" Then
        
            '"Fatura 2"
            key = GetWorksheetColumnValue(ws_query, r, "Fatura 2")
                                            
            With Application.WorksheetFunction
            
                On Error Resume Next
                vl = .VLookup(key * 1, items, 2, False)
                
                If vl <> "" Then
                
                    Call SetWorksheetColumnValue(ws_query, r, "Usuário Fatura 2", vl)
                    
                    On Error Resume Next
                    dt_1 = .VLookup(key * 1, items, 3, False)
                    If dt_1 > 0 Then
                        Call SetWorksheetColumnValue(ws_query, r, "Data Criação Fatura 2", dt_1)
                    End If
                    
                    On Error Resume Next
                    dt_2 = .VLookup(key * 1, items, 4, False)
                    If dt_2 > 0 Then
                        Call SetWorksheetColumnValue(ws_query, r, "Data Contabilização Fatura 2", dt_2)
                    End If
    
                End If
                
            End With
        
        End If
    
    Next r

End Sub

Sub UpdateAccountingType_Q1()

    Dim r As Long
    Dim ttr As Long
    Dim key As String
    Dim user As String
    Dim vl As String
    
    ttr = CountRows(ws_query, 1)
    
    For r = 2 To ttr
    
        If GetWorksheetColumnValue(ws_query, r, "Contabilização Fatura 2") = "" Then
    
            key = GetWorksheetColumnValue(ws_query, r, "Texto")
            user = GetWorksheetColumnValue(ws_query, r, "Usuário Fatura 2")
            vl = "INTEGRACAO"
            
            If key Like "Finalizado Manualmente (AANTUNES*" Or key Like "Finalizado Manualmente (PYKOSZ*" Then
                
                vl = "ZTSD401"
                
            ElseIf user = "AANTUNES" Or user = "PYKOSZ" Then
            
                vl = "VF01"
            
            End If
            
            Call SetWorksheetColumnValue(ws_query, r, "Contabilização Fatura 2", vl)
            
        End If
        
    Next

End Sub

Sub GetDocTranspInfo_Q1()

    Dim r As Long
    Dim ttr As Long
    Dim key As String
    Dim i As Long

    Call ClearWorksheet(ws_temp, 1)
    Call ClearWorksheet(ws_vttk, 2)
    Call EnvironmentVariables_SAP(True)
    
    ttr = CountRows(ws_query, 1)
    i = 0
    
    For r = 2 To ttr
    
        If GetWorksheetColumnValue(ws_query, r, "Cod Transp") = "" Then
        
            key = GetWorksheetColumnValue(ws_query, r, "Transporte")
            
            If key = "" Then
            
                Call SetNullDocTranspInfo_Q1(r)
    
            Else
                
                i = i + 1
                ws_temp.Cells(i, 1) = key
                
                If i > 10000 Then
                    Call UpdateVTTK_Q1
                    Call UpdateBaseByVTTK_Q1
                    i = 0
                End If
            
            End If
            
        End If
    
    Next r
    
    If i > 1 Then
        Call UpdateVTTK_Q1
        Call UpdateBaseByVTTK_Q1
        i = 0
    End If
        
    Call ClearWorksheet(ws_temp, 1)
    Call EnvironmentVariables_SAP(False)
    
End Sub

Sub SetNullDocTranspInfo_Q1(r As Long)

    Call SetWorksheetColumnValue(ws_query, r, "Cod Transp", "N/D")
    Call SetWorksheetColumnValue(ws_query, r, "Itinerário", "N/D")
    Call SetWorksheetColumnValue(ws_query, r, "Data Coleta", "N/D")
    Call SetWorksheetColumnValue(ws_query, r, "Nome Transportadora", "N/D")
    Call SetWorksheetColumnValue(ws_query, r, "Apelido", "N/D")
    Call SetWorksheetColumnValue(ws_query, r, "Tempo Itinerário", "N/D")

End Sub

Sub UpdateVTTK_Q1()

    Call ClearWorksheet(ws_vttk, 2)

    ws_temp.Activate
    ws_temp.Range("A:A").Select
    Selection.Copy
    
    Call RunScriptSAP_ROL_DT_TO_LFA1_Q1
    Call SetTempData
    Call LoadTempData_ROL_DT_TO_LFA1_Q1

    Call ClearWorksheet(ws_temp, 1)

End Sub

Sub LoadTempData_ROL_DT_TO_LFA1_Q1()
    
    Dim i As Long
    Dim r As Long
    Dim ttr As Long

    Dim lineTemp As String
    Dim cls() As String
    Dim hdr() As String
    Dim c As Long
    Dim h As String
    Dim vl As String
                
    ttr = CountRows(ws_temp, 1)
    i = CountRows(ws_vttk, 1)
    
    For r = 2 To ttr
    
        lineTemp = ws_temp.Cells(r, 1)
    
        If lineTemp Like "|*" Then
        
            cls = Split(lineTemp, "|")
        
            If IsArrayEmpty(hdr) Then
            
                hdr = cls
                
                For c = 0 To UBound(hdr)
                    hdr(c) = Trim(hdr(c))
                Next c
            
            Else
                
                For c = 0 To UBound(cls)
                                   
                    h = hdr(c)
                    vl = Trim(cls(c))
                                   
                    If h <> "" Then
                                                         
                        If h <> vl Then

                            Select Case h

                                Case "Transporte"
                                    i = i + 1
                                    Call SetWorksheetColumnValue(ws_vttk, i, "DT", vl)
                                    
                                Case "Itin."
                                    Call SetWorksheetColumnValue(ws_vttk, i, "Route", vl)
                                    
                                Case "ForncServ."
                                    Call SetWorksheetColumnValue(ws_vttk, i, "CarrierCod", vl)
                                    
                                Case "Nome 1"
                                    Call SetWorksheetColumnValue(ws_vttk, i, "CarrierName", vl)
                                    
                                Case "InAtTransp"
                                    Call SetWorksheetColumnValue(ws_vttk, i, "DepartureDate", ConvertionOutputDate(vl))
                                    
                                Case "DurGlobPlan."
                                    Call SetWorksheetColumnValue(ws_vttk, i, "TransitTime", vl)
                            
                            End Select
                            
                        End If
                        
                    End If
                    
                Next c
                
            End If
        
        End If
        
    Next r
    
End Sub

Sub UpdateBaseByVTTK_Q1()

    Dim r As Long
    Dim ttr As Long
    Dim key As String
        
    Dim ttn As Long
    Dim items As Range
    'Dim i As Variant
    
    'Dim c As Long
    Dim vl As String
    Dim cod As String
    Dim route As String
    Dim dt As Date
    Dim carrier As String
    Dim route_time As Long
    
    ttr = CountRows(ws_query, 1)
    ttn = CountRows(ws_vttk, 1)
    
    Set items = ws_vttk.Range("A1:F" & ttn)
    
    For r = 2 To ttr
    
        vl = ""
        cod = ""
        route = ""
        dt = 0
        carrier = ""
        route_time = 0
    
        If GetWorksheetColumnValue(ws_query, r, "Cod Transp") = "" Then
        
            key = GetWorksheetColumnValue(ws_query, r, "Transporte")
            
            With Application.WorksheetFunction
                
                On Error Resume Next
                vl = .VLookup(key * 1, items, 1, False)
                
                If vl <> "" Then
                
                    On Error Resume Next
                    cod = .VLookup(key * 1, items, 3, False)
                    If cod <> "" Then
                        Call SetWorksheetColumnValue(ws_query, r, "Cod Transp", cod)
                    End If
                
                    On Error Resume Next
                    route = .VLookup(key * 1, items, 2, False)
                    If route <> "" Then
                        Call SetWorksheetColumnValue(ws_query, r, "Itinerário", route)
                    End If
                    
                    On Error Resume Next
                    dt = .VLookup(key * 1, items, 5, False)
                    If dt > 0 Then
                        Call SetWorksheetColumnValue(ws_query, r, "Data Coleta", dt)
                    End If
                    
                    On Error Resume Next
                    carrier = .VLookup(key * 1, items, 4, False)
                    If carrier <> "" Then
                        Call SetWorksheetColumnValue(ws_query, r, "Nome Transportadora", carrier)
                    End If
                    
                    On Error Resume Next
                    route_time = .VLookup(key * 1, items, 6, False)
                    If route_time > 0 Then
                        Call SetWorksheetColumnValue(ws_query, r, "Tempo Itinerário", route_time)
                    End If
                
                End If
            
            End With
            
        End If
    
    Next
 
End Sub

Sub UpdateCarrierNickname()

    Dim r As Long
    Dim ttr As Long
    Dim key As String
        
    Dim ttn As Long
    Dim items As Range
    Dim vl As String
    
    Set wb_external = Workbooks.Open(ws_q1.Range("B3") & ws_q1.Range("C3"))
    Set ws_external = wb_external.Worksheets(CStr(ws_q1.Range("D3")))
    
    ttr = CountRows(ws_query, 1)
    ttn = CountRows(ws_external, 1)
    
    Set items = ws_external.Range("A1:F" & ttn)
    
    For r = 2 To ttr
    
        If GetWorksheetColumnValue(ws_query, r, "Apelido") = "" Then
    
            key = GetWorksheetColumnValue(ws_query, r, "Cod Transp")
            
            If IsNumeric(key) Then
                
                With Application.WorksheetFunction
                
                    On Error Resume Next
                    vl = .VLookup(key, items, 6, False)
                    
                    If vl <> "" Then
                        
                        Call SetWorksheetColumnValue(ws_query, r, "Apelido", vl)
                        
                    End If
                    
                End With
            
            End If
        
        End If
    
    Next
    
    Call ClearExternalParameters

End Sub
