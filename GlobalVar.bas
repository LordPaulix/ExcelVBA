Attribute VB_Name = "GlobalVar"
Option Explicit

'Workbook
Public WB_BILLING_DATA_TOOL As Workbook

'Worksheets
Public WS_TTEXTRACT As Worksheet
Public WS_RDEXTACT As Worksheet
Public WS_UPDEXTRACT As Worksheet
Public WS_PLCREATION As Worksheet
Public WS_CONSOLIDATION_MASTER As Worksheet

'Sheet Name
Public Const SHT_TTEXTRACT As String = "TimeTracker Extract"
Public Const SHT_RDEXTACT As String = "Resource Data Extract"
Public Const SHT_UPDEXTRACT As String = "Unit Price Data Extract"
Public Const SHT_PLCREATION As String = "Project List Creation"
Public Const SHT_CONSOLIDATION_MASTER As String = "Consolidated Master Creation"
Public Const SHT_VERSION As String = "Version"


