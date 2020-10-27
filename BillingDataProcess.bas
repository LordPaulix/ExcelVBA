Attribute VB_Name = "BillingDataProcess"
Option Explicit
Sub CallBillingDataGeneration()
    
    Dim pCounter As Integer
    
    Dim mConsolidationWS As Worksheet
    Dim pListWS As Worksheet
    Dim tTrackerWS As Worksheet
    Dim bDataWS As Worksheet
    Dim rscDataWS As Worksheet
    
    Dim pName As String
    Dim haveProject As Boolean
    Dim rscName As Integer
    Dim pNameSrc As Integer
    Dim iLastEntry As Long
    Dim iLastRowTimeTracker As Long
    Dim dRowSrc As Integer
    Dim rowDst As Integer
    Dim tTrackerName As Integer
    
    Dim rowCounter As Integer
    
    Dim modulo As Double
    
    Initialize
    
    If pWSExist("Consolidated Master Creation") = True Then
    
        CallBillingDataGeneration
    
    Else
        pNameSrc = 2
        
        On Error Resume Next
        
        Set mConsolidationWS = ThisWorkbook.Sheets.Add(after:=ThisWorkbook.Sheets(6))
        mConsolidationWS.Name = "Consolidated Master Creation"

        Do While WS_PLCREATION.Cells(pNameSrc, 1) <> ""
            If WS_PLCREATION.Cells(pNameSrc, 2).Value = "Yes" Then
                rscName = 2
                Do While WS_RDEXTACT.Cells(rscName, 1).Value <> ""
                    tTrackerName = 2
                        Do While WS_TTEXTRACT.Cells(tTrackerName, 1).Value <> ""
                            If WS_TTEXTRACT.Cells(tTrackerName, 2).Value = WS_RDEXTACT.Cells(rscName, 1).Value Then
                                If WS_TTEXTRACT.Cells(tTrackerName, 8).Value = WS_PLCREATION.Cells(pNameSrc, 1) Then
                                    
                                    With mConsolidationWS
                                        iLastRowTimeTracker = .Range("A" & Rows.Count).End(xlUp).Row + 1
                                        
                                        .Cells(1, 1) = "Account"
                                        .Cells(1, 2) = "Account Code"
                                        .Cells(1, 3) = "Section"
                                        .Cells(1, 4) = "Working Time"
                                        .Cells(1, 5) = "Project Name"
                                        .Cells(1, 6) = "Work Group"
                                        .Cells(1, 7) = "Task Name"
                                        .Cells(1, 8) = "Task Code"
                                        .Cells(1, 9) = "Unit Price"
                                        .Cells(1, 10) = "Billing Ammount"
                                        'Account Name
                                        .Cells(iLastRowTimeTracker, 1) = WS_TTEXTRACT.Cells(tTrackerName, 1).Value
                                        'Account code
                                        .Cells(iLastRowTimeTracker, 2) = WS_TTEXTRACT.Cells(tTrackerName, 2).Value
                                        'Section
                                        .Cells(iLastRowTimeTracker, 3) = WS_TTEXTRACT.Cells(tTrackerName, 3).Value
                                        'Working Time
                                        .Cells(iLastRowTimeTracker, 4) = WS_TTEXTRACT.Cells(tTrackerName, 6).Value
                                        'Project Name
                                        .Cells(iLastRowTimeTracker, 5) = WS_TTEXTRACT.Cells(tTrackerName, 8).Value
                                        'Work Group
                                        .Cells(iLastRowTimeTracker, 6) = WS_TTEXTRACT.Cells(tTrackerName, 14).Value
                                        'Task Name
                                        .Cells(iLastRowTimeTracker, 7) = WS_TTEXTRACT.Cells(tTrackerName, 19).Value
                                        'Task Code
                                        .Cells(iLastRowTimeTracker, 8) = WS_TTEXTRACT.Cells(tTrackerName, 20).Value
                                        'Unit Price
                                        .Cells(iLastRowTimeTracker, 9) = "=IFERROR(VLOOKUP(RC[-3],'Unit Price Data Extract'!C1:C3,3,FALSE),0)"
                                        'Billing Amount
                                        .Cells(iLastRowTimeTracker, 10) = (.Cells(iLastRowTimeTracker, 4).Value) * (.Cells(iLastRowTimeTracker, 9).Value)
                                    End With
                                End If
                            End If
                        tTrackerName = tTrackerName + 1
                        Loop
                    rscName = rscName + 1
                    Loop
                If pWSAdded(WS_PLCREATION.Cells(pNameSrc, 1).Value) = False Then
                    ReportGenerate WS_PLCREATION.Cells(pNameSrc, 1).Value
                End If
            End If
        pNameSrc = pNameSrc + 1
        Loop
    End If
    
    With mConsolidationWS
        .Cells.WrapText = False
        .Cells.Font.Name = "Calibri"
        .Cells.Font.Size = 10
        .Rows(1).AutoFilter
    End With
    On Error GoTo 0
End Sub
Public Sub ReportGenerate(wSheetName As String)
    On Error GoTo ERR_HANDLER

    Dim wSheetCreate As Worksheet
    
    Dim pTable As PivotTable
    Dim pObjField As PivotField
    Dim pc As PivotCache
    Dim rSheetName As String
    Dim pivotItm As PivotItem
    
    Dim sectionName As String
    
    Dim counter As Integer
    

    If pWSExist(wSheetName) = True Then
        CallBillingDataGeneration
    Else
        Set wSheetCreate = ThisWorkbook.Sheets.Add(after:=ThisWorkbook.Sheets(7))
        
        rSheetName = wSheetName
        rSheetName = Replace(Replace(Replace(Replace(Replace(Replace(Replace(wSheetName, ":", ""), _
                        "/", ""), "\", ""), "*", ""), "[", ""), "]", ""), "?", "")
                        
        If Len(wSheetName) > 31 Then
            rSheetName = Left$(rSheetName, 31)
        End If
        
        wSheetCreate.Name = rSheetName
        
        Set pc = ThisWorkbook.PivotCaches.Create(xlDatabase, "Consolidated Master Creation!R1C1:R1048576C10", xlPivotTableVersion14)
        Set pTable = pc.CreatePivotTable(wSheetCreate.Range("A7"), wSheetName, , xlPivotTableVersion14)
        
        'On Error GoTo 0
        
        Set pObjField = wSheetCreate.PivotTables(wSheetName).PivotFields("Section")
            pObjField.Orientation = xlPageField
            pObjField.Position = 1
        
        Set pObjField = wSheetCreate.PivotTables(wSheetName).PivotFields("Project Name")
            pObjField.Orientation = xlPageField
            pObjField.Position = 2
                
        Set pObjField = wSheetCreate.PivotTables(wSheetName).PivotFields("Task Name")
            pObjField.Orientation = xlRowField
            pObjField.Position = 1
            pObjField.Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
            False, False)
       
        Set pObjField = wSheetCreate.PivotTables(wSheetName).PivotFields("Unit Price")
            pObjField.Orientation = xlRowField
            pObjField.Position = 2
        
        With ActiveSheet.PivotTables(wSheetName)
            .AddDataField ActiveSheet.PivotTables( _
                wSheetName).PivotFields("Working Time"), "合計 / Working Time", xlSum
            .AddDataField ActiveSheet.PivotTables( _
                wSheetName).PivotFields("Billing Ammount"), "合計 / Billing Ammount", _
                xlSum
            .PivotFields("Project Name"). _
                CurrentPage = "(All)"
            .PivotFields("Project Name"). _
                EnableMultiplePageItems = True
            .PivotSelect "'Task Name'[All]", _
                xlLabelOnly, False
            .InGridDropZones = True
            .RowAxisLayout xlTabularRow
            With .PivotFields("Project Name")
                '.PivotItems(wSheetName).Visible = True
                For Each pivotItm In .PivotItems
                    If pivotItm <> wSheetName Then
                        pivotItm.Visible = False
                    End If
                Next
                .EnableMultiplePageItems = True
            End With
            .PivotCache. _
                MissingItemsLimit = xlMissingItemsNone
        End With
                
        With wSheetCreate
            .Cells(1, 1) = "Project"
            .Cells(2, 1) = wSheetName
            .Cells.WrapText = True
            .Cells(1, 1).Font.Name = "Calibri"
            .Cells(1, 1).Font.Size = 18
            .Cells(1, 1).Font.Bold = True
        End With

        Range("A7").Select
           
    End If
     
    CountSection (wSheetName)

    If wSheetCreate.Cells(2, 1) = "" Then
        pWSExist (rSheetName)
    End If
     
     
     'On Error GoTo 0
Exit Sub

ERR_HANDLER:
        pWSExist (rSheetName)
End Sub
Public Function CountSection(wSheet As String)

    Dim pT As PivotTable
    Dim pi As PivotItem
    Dim pF As PivotField
    
    Dim pTable As PivotTable
    Dim pObjField As PivotField
    Dim pc As PivotCache
    Dim rSheetName As String
    
    Dim pivotItm As PivotItem
    Dim itmChecker As String
    Dim i As Integer
       
    Dim counter As Integer
    
    Dim iLoop As Long

    rSheetName = wSheet
    rSheetName = Replace(Replace(Replace(Replace(Replace(Replace(Replace(wSheet, ":", ""), _
                    "/", ""), "\", ""), "*", ""), "[", ""), "]", ""), "?", "")
                    
    If Len(wSheet) > 31 Then
        rSheetName = Left$(wSheet, 31)
    End If
    
    Set pT = Sheets(rSheetName).PivotTables(wSheet)
    Set pF = pT.PageFields("Section")
    
    counter = ActiveSheet.Range("A8:A" & Rows.Count).End(xlToRight).Column + 2
    
    For Each pi In pF.PivotItems
        
        pi.Value = pi.Value
        
        If pi.Value = "(blank)" Or pi.Value = "(空白)" Then
            iLoop = iLoop + 1
        
        Else
        
            'CountSection = pi.Value

            Set pc = ThisWorkbook.PivotCaches.Create(xlDatabase, "Consolidated Master Creation!R1C1:R1048576C10", xlPivotTableVersion14)
            Set pTable = pc.CreatePivotTable(Sheets(rSheetName).Cells(7, counter), wSheet & counter, , xlPivotTableVersion14)
            
            Set pObjField = ActiveSheet.PivotTables(wSheet & counter).PivotFields("Section")
                pObjField.Orientation = xlPageField
                pObjField.Position = 1
            
            Set pObjField = ActiveSheet.PivotTables(wSheet & counter).PivotFields("Project Name")
                pObjField.Orientation = xlPageField
                pObjField.Position = 2
        
            
            Set pObjField = ActiveSheet.PivotTables(wSheet & counter).PivotFields("Task Name")
                pObjField.Orientation = xlRowField
                pObjField.Position = 1
                pObjField.Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
                 False, False)

            Set pObjField = ActiveSheet.PivotTables(wSheet & counter).PivotFields("Unit Price")
                 pObjField.Orientation = xlRowField
                 pObjField.Position = 2

            ActiveSheet.PivotTables(wSheet & counter).AddDataField ActiveSheet.PivotTables( _
                wSheet & counter).PivotFields("Working Time"), "合計 / Working Time", xlSum
            ActiveSheet.PivotTables(wSheet & counter).AddDataField ActiveSheet.PivotTables( _
                 wSheet & counter).PivotFields("Billing Ammount"), "合計 / Billing Ammount", _
                 xlSum

            ActiveSheet.PivotTables(wSheet & counter).PivotSelect "'Task Name'[All]", _
                 xlLabelOnly, True

            With ActiveSheet.PivotTables(wSheet & counter)
                .InGridDropZones = True
                .RowAxisLayout xlTabularRow
                With .PivotFields("Task Name")
                    .PivotFilters.Add Type:=xlCaptionDoesNotEqual, Value1:="(空白)"
                    
                End With
                
                With .PivotFields("Project Name")
                    '.PivotItems(rSheetName).Visible = True
                    For Each pivotItm In .PivotItems
                        If pivotItm <> wSheet Then
                            pivotItm.Visible = False
                        End If
                    Next
                    .EnableMultiplePageItems = True
                End With
                                
                With .PivotFields("Section")
                    .CurrentPage = pi.Value
                End With
            End With
            
            With ActiveSheet
                .Cells.WrapText = True
            End With
                        
            ActiveSheet.PivotTables(wSheet & counter).PivotCache. _
            MissingItemsLimit = xlMissingItemsNone
                        
'            Dim converter As String
'            converter = Cells(7, counter).Address(RowAbsolute:=False, ColumnAbsolute:=False)

            If Application.WorksheetFunction.CountA(Columns(counter)) = 4 Then
                ActiveSheet.PivotTables(wSheet & counter).TableRange2.Clear
            End If
               
           counter = counter + (counter - 1)
        
        End If

    Next pi

    iLoop = 1
    
    Application.StatusBar = ""
    
    If Sheets(rSheetName).Cells(2, 1) = "" Then
        pWSExist (rSheetName)
    End If
    
End Function
Private Function CheckForValue(pvtPivotName As Object, strFieldName As String, strCheckValue As String) As Boolean

    Dim i As Integer

    With pvPivotName
        For i = i To .PivotFields(strFieldName).PivotItems.Count
            If strCheckValue = .PivotFields(strFieldName).PivotItems(i).Name Then
                CheckForValue = True
                Exit Function
            End If
        Next i
    End With
    
    CheckForValue = False
End Function
