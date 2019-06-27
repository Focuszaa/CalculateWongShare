Attribute VB_Name = "Module2"
'
'Author Nirut Phengjaiwong
'Version 1.0.0
'Date: 21-Oct-2017 : 15:36
'

Function Check_wong_edit(name As String, dates As Date, Count As Long) As Double
    
    Dim checkindex As Boolean
    checkindex = False
    'Dim Last_index As Long
    Dim datess As Date
    Dim old_name As String
    
    For i = 2 To Count
        old_name = Sheets("ÃÇÁÇ§áªÃì").Cells(i, 2).Value
        If old_name = name Then
           Check_wong_edit = i
           GoTo exitLoop
        End If
    Next i
    Check_wong_edit = 0
exitLoop:
    
    
End Function
