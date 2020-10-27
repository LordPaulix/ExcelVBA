Attribute VB_Name = "InputSheetProcess"
Option Explicit
Dim nDoneCount As Long
Dim nFileCount As Long
Dim nPreP As Long
Dim strREC As String
Private Function SearchCount(strPath) As Long
    Dim objFs
    Dim objSub
    Dim objFld
    Dim objFl
    Dim j As Long
    Dim nFileCount
    
    Set objFs = CreateObject("Scripting.FileSystemObject")
    Set objFld = objFs.GetFolder(strPath)
    
    nFileCount = objFld.Files.Count
    For Each objSub In objFld.SubFolders
        nFileCount = nFileCount + SearchCount(objSub.Path)
    Next
    
    SearchCount = nFileCount
End Function
Private Function InputSheetCreation(strPath As String, wSheet As String) As Boolean
    
    Dim newWS As Worksheet
    Dim wB As Workbook
    Dim wS As Worksheets
    Dim fName As String
    Dim myPath As String
            
    Dim nRowSrc As Integer
    Dim nRowDst As Integer
    
    Dim counter As Integer
    Dim tTrackerColCount As Integer
       
'    toolWorkbookName = ActiveWorkbook.Name
'    Set wbFs = Workbooks(toolWorkbookName)
'
'    Set objFs = CreateObject("Scripting.FileSystemObject")

    If (strPath = "") Then
        MsgBox "Please provide the location", vbInformation
        Sheet1.timeTrackerPath = ""

    InputSheetCreation = False

    Else
          
            Application.DisplayAlerts = False
            Application.StatusBar = "Billing Data Creation In-progress"
                                        
            Select Case wSheet
            
                Case "TimeTracker Extract"
                    
                    Application.StatusBar = "TimeTracker Data Extraction In-progress"
                        If pWSExist(wSheet) = False Then
                        
                            Set wB = Workbooks.Open(strPath, 0, True)
                            Set newWS = ThisWorkbook.Sheets.Add(after:=ThisWorkbook.Sheets(2))
                            newWS.Name = wSheet
                            
                            With wB.Sheets(1)
                                tTrackerColCount = .Cells(1, .Columns.Count).End(xlToLeft).Column
                            End With
                            
                            With newWS
    
                                nRowSrc = 1
                                nRowDst = 1
    
                                Do While wB.Sheets(1).Cells(nRowSrc, 1).Value <> ""
                                    For counter = 1 To tTrackerColCount
                                        .Cells(nRowDst, counter).Value = wB.Sheets(1).Cells(nRowSrc, counter).Value
                                    Next counter
        
                                nRowSrc = nRowSrc + 1
                                nRowDst = nRowDst + 1
    
                                Loop
    
                            .Cells.WrapText = False
                            .Cells.Font.Name = "Calibri"
                            .Cells.Font.Size = 10
                            .Rows(1).AutoFilter
    
                            End With
                        wB.Close
                    End If
                    
                Case "Resource Data Extract"
                    Application.StatusBar = "Resource Data Extraction In-progress"
                    If pWSExist(wSheet) = False Then
                        Set wB = Workbooks.Open(strPath, 0, True)
                        Set newWS = ThisWorkbook.Sheets.Add(after:=ThisWorkbook.Sheets(3))
                        newWS.Name = wSheet
                        
                        With wB
                            .Sheets(1).Range("A1:B65536").Copy
                        End With
        
                        With newWS
                            .Range("A1").PasteSpecial xlPasteValues
                            .Cells.WrapText = False
                            .Cells.Font.Name = "Calibri"
                            .Cells.Font.Size = 10
                            .Rows(1).AutoFilter
                        End With
                        
                        wB.Close
                        
                    End If
    
                Case "Unit Price Data Extract"
                    Application.StatusBar = "Unit Price Data Extraction In-progress"
                    If pWSExist(wSheet) = False Then
                        
                        Set wB = Workbooks.Open(strPath, 0, True)
                        Set newWS = ThisWorkbook.Sheets.Add(after:=ThisWorkbook.Sheets(4))
                        newWS.Name = wSheet
                        
                        With wB
                            .Sheets(1).Range("A1:C65536").Copy
                        End With
        
                        With newWS
                            .Range("A1").PasteSpecial xlPasteValues
                            .Cells.WrapText = False
                            .Cells.Font.Name = "Calibri"
                            .Cells.Font.Size = 10
                            .Rows(1).AutoFilter
                        End With
                        
                        wB.Close
                
                    End If
                Case Else
            
            End Select
                            
            Application.CutCopyMode = False
    
    InputSheetCreation = True
    
    End If

End Function
Private Sub SearchFiles()

    Dim j As Long
    Dim sTargetPath As String
    
    Application.Calculation = xlCalculationManual
    
    With ThisWorkbook.Sheets("File List")
    
    Application.ScreenUpdating = False
    sTargetPath = .Cells(1, 1)
    
    .Cells(1, 2) = SearchCount(sTargetPath)
    nDoneCount = 0
    Call InputSheetCreation(sTargetPath)
    
    End With
    
    Application.Calculation = xlCalculationAutomatic
    
End Sub
Sub CallTimeTrackerPath()

    If InputSheetCreation(ThisWorkbook.Sheets("BillingData_Interface").timeTrackerPath.Text, _
            "TimeTracker Extract") = True Then
       Call CallResourceDataPath
    
    Else
        MsgBox "Importing failed."
    End If

End Sub
Sub CallResourceDataPath()
    If InputSheetCreation(ThisWorkbook.Sheets("BillingData_Interface").resourceDataPath.Text, _
            "Resource Data Extract") = True Then
        Call CallUnitPriceDataPath
    Else
        MsgBox "Importing failed."
    End If
End Sub
Sub CallUnitPriceDataPath()
    If InputSheetCreation(ThisWorkbook.Sheets("BillingData_Interface").unitPriceDataPath.Text, _
            "Unit Price Data Extract") = True Then
        PListCreation
    End If
End Sub
Private Function PListCreation()
    Dim pListCreationWS As Worksheet
    Dim wsSource As Worksheet
    Dim iLastRow As Integer
    Dim counter As Integer
    
    Dim obj As Object
    Dim i As Long
    Dim cellVal As Variant

        If pWSExist("Project List Creation") = False Then
            
            Application.StatusBar = "Creating Project List..."
            
            Set pListCreationWS = ThisWorkbook.Sheets.Add(after:=ThisWorkbook.Sheets(5))
            pListCreationWS.Name = "Project List Creation"
            
            Set wsSource = Sheets("TimeTracker Extract")
            
            Set obj = CreateObject("Scripting.Dictionary")
            
            iLastRow = wsSource.Range("A" & Rows.Count).End(xlUp).Row + 1
    
            cellVal = wsSource.Range("H2:H" & iLastRow)
    
            With pListCreationWS
                
                For i = 1 To UBound(cellVal, 1)
                    obj(cellVal(i, 1)) = 1
                Next i
                
                .Range("A2").Resize(obj.Count) = Application.Transpose(obj.keys)
                .Cells(1, 1) = "Project Name"
                .Cells(1, 2) = "Billing File Creation"
                
                counter = 1
                
                Do While .Cells(counter + 1, 1) <> ""
                
                    Select Case .Cells(counter + 1, 1)
                    
                        Case "COMMON", "Estabilshment SQA system" _
                             , "Improvement SQA system", "Meter process training" _
                             , "MM Tasks", "POS Sidecheck", "SQA Training"
                        
                            .Cells(counter + 1, 2) = ""
                        
                        Case Else
                        
                            .Cells(counter + 1, 2) = ""
                             
                    End Select
                
                counter = counter + 1
                
                Loop
                .Cells.WrapText = False
                .Cells.Font.Name = "Calibri"
                .Cells.Font.Size = 10
                .Rows(1).AutoFilter
                Columns("A:A").EntireColumn.AutoFit
                .Range("A2:A99999").Sort key1:=Range("A2:A99999"), order1:=xlAscending, Header:=xlNo
            End With
        End If
        Application.StatusBar = "Project List Creation Done"
End Function
