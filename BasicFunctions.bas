Attribute VB_Name = "BasicFunctions"
'-----------------------------------------
'Function: pWSExist
'Parameter: wSheetName
'Return: Boolean
'Description: Check if worksheet name already exist
'-----------------------------------------
Public Function pWSExist(wSheetName As String, Optional wB As Workbook) As Boolean
    
    Dim sht As Worksheet
     
    bBoolean = False


    For Each oSheet In ThisWorkbook.Sheets
        If oSheet.Name = wSheetName Then
            Application.DisplayAlerts = False
            ActiveWorkbook.Sheets(wSheetName).Delete
            bBoolean = True
            Exit For
        End If
    Next oSheet

    pWSExist = bBoolean
        
End Function
Public Function pWSAdded(wSheetName As String, Optional wB As Workbook) As Boolean
    
    Dim sht As Worksheet
     
    bBoolean = False


    For Each oSheet In ThisWorkbook.Sheets
        If oSheet.Name = wSheetName Then
            'ActiveWorkbook.Sheets(wSheetName).Delete
            bBoolean = True
            Exit For
        End If
    Next oSheet

    pWSAdded = bBoolean
        
End Function
'-------------------------------------------
'Function: browseFile
'Parameter: None
'Return: String, Select File information
'Description: File Selection dialog box
'-------------------------------------------
Public Function browseFile() As String
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "Select File"
        .AllowMultiSelect = False
        .Filters.Add "Excel Files", "*.xl??"
        .Show
        
        If (.SelectedItems.Count > 0) Then
            browseFile = .SelectedItems(1)
        Else
            browseFile = ""
            Exit Function
        End If
    End With
End Function
