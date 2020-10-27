VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Private Sub generateInputSheet_Click()

    CallTimeTrackerPath

End Sub
Private Sub generateBillingData_Click()

    CallBillingDataGeneration
    
End Sub
Private Sub btn_BrowseTimeTrackerData_Click()
Dim sFile As String
Dim sName As String

    sFile = browseFile()
    If (sFile <> "") Then
        sName = Split(Trim(sFile), "\")(UBound(Split(Trim(sFile), "\")))
        
        If (sFile <> "") Then
'        And (InStr(sName, "TimeSheetDB") > 0) Then
            timeTrackerPath = sFile
        Else
            MsgBox "Selected file was not a valid Database", vbOKOnly + vbInformation, "Invalid Database"
        End If
    Else
        MsgBox "No File selected", vbOKOnly + vbInformation, "Invalid Database"
    End If
End Sub
Private Sub btn_BrowseResourceData_Click()
Dim sFile As String
Dim sName As String

    sFile = browseFile()
    If (sFile <> "") Then
        sName = Split(Trim(sFile), "\")(UBound(Split(Trim(sFile), "\")))
        
        If (sFile <> "") Then
            resourceDataPath = sFile
        Else
            MsgBox "Selected file was not a valid", vbOKOnly + vbInformation, "Invalid Database"
        End If
    Else
        MsgBox "No File selected", vbOKOnly + vbInformation, "Invalid Database"
    End If
End Sub
Private Sub btn_BrowseUnitPriceData_Click()
Dim sFile As String
Dim sName As String

    sFile = browseFile()
    If (sFile <> "") Then
        sName = Split(Trim(sFile), "\")(UBound(Split(Trim(sFile), "\")))
        
        If (sFile <> "") Then
            unitPriceDataPath = sFile
        Else
            MsgBox "Selected file was not a valid ordersheet", vbOKOnly + vbInformation, "Invalid Database"
        End If
    Else
        MsgBox "No File selected", vbOKOnly + vbInformation, "Invalid Database"
    End If
End Sub
