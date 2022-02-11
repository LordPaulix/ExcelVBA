Sub ReplaceStringInFile()

Dim sBuf As String
Dim sTemp As String
Dim iFileNum As Integer
Dim sFileName As String
Dim newFileName As String




' Edit as needed
sFileName = "D:\Bibi\Personal\Test.txt"

iFileNum = FreeFile
Open sFileName For Input As iFileNum

Do Until EOF(iFileNum)
    Line Input #iFileNum, sBuf
    sTemp = sTemp & sBuf & vbCrLf
Loop
Close iFileNum

sTemp = Replace(sTemp, "953", "954")

iFileNum = FreeFile
newFileName = "D:\Bibi\Personal\Test123.txt"
Open newFileName For Output As iFileNum
Print #iFileNum, sTemp
Close iFileNum

End Sub
