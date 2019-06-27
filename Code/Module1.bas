Attribute VB_Name = "Module1"
'
'Author Nirut Phengjaiwong
'Version 1.0.0
'Date: 21-Oct-2017 : 15:36
'
Public Sum_handDead As Long


Sub calculate()

Dim Sum_forward As Double
Dim Sum_reverse As Double
Dim Sum_income As Double
Dim Round As Double
Dim EachPayment As Double
Dim OwnerPayment As Double
Dim Plus_interest As Double
Dim bottomline As Double
Dim count_forward As Integer
Dim FreeForOwner As Double
Dim Fee As Double


Round = ActiveSheet.Cells(5, 2).Value
EachPayment = ActiveSheet.Cells(6, 2).Value
OwnerPayment = ActiveSheet.Cells(9, 2).Value
Plus_interest = ActiveSheet.Cells(10, 2).Value
FreeForOwner = ActiveSheet.Cells(4, 2).Value
Fee = ActiveSheet.Cells(8, 2).Value


'initail sum income
Sum_imcome = 0
bottomline = (EachPayment + OwnerPayment) * Round
ActiveSheet.Cells(3, 7).Value = ActiveSheet.Cells(4, 2).Value

For j = 1 To Round + 1
    ActiveSheet.Cells(j + 2, 5).Value = j
Next j

'forward and reverse
For i = 1 To Round

  'income
  ActiveSheet.Cells(i + 3, 7).Value = bottomline + Sum_income
  Sum_income = Sum_income + 50
  
  'forward
  count_forward = Round - i
  Sum_forward = (Plus_interest + EachPayment) * count_forward
    
  'reverse
  Sum_reverse = (EachPayment) * i
  
  'Assign value
  ActiveSheet.Cells(i + 3, 8).Value = (Sum_forward + Sum_reverse + (FreeForOwner / Round)) * -1
  ActiveSheet.Cells(i + 3, 9).Value = (Fee) * -1
  ActiveSheet.Cells(i + 3, 10).Value = (Sum_forward + Sum_reverse + Fee + (FreeForOwner / Round)) * -1
  ActiveSheet.Cells(i + 3, 11).Value = ActiveSheet.Cells(i + 3, 10).Value + ActiveSheet.Cells(i + 3, 7).Value
  ActiveSheet.Cells(i + 3, 13).Value = (ActiveSheet.Cells(11, 2).Value) * -1
  
    'calculate summary
  'ActiveSheet.Cells(i + 3, 13).Value =
    
  'clear value all Sums
  Sum_forward = 0
  Sum_reverse = 0
  
Next i

    

End Sub

Sub Owner_calculation()

Dim Round As Double
Dim Round_date As Double
Dim OwnerPayment As Double
Dim Dead As Double
Dim EachPayment As Double


'
'Checking name on each row, if all names already inputed we will input it on member sheet calcuation
'
'

Round = ActiveSheet.Cells(5, 2).Value
Dim i As Integer
    For i = 1 To Round
        If ActiveSheet.Cells(i + 2, 6).Value = "" Then
            MsgBox "กรุณาใส่ข้อมูลสมาชิคของวงแชร์ให้ครบ"
            Exit Sub
        End If
    Next i
i = 0

Round = ActiveSheet.Cells(5, 2).Value
Round_date = ActiveSheet.Cells(2, 3).Value
OwnerPayment = ActiveSheet.Cells(9, 2).Value
EachPayment = ActiveSheet.Cells(6, 2).Value


'
'Calculate date and input to the sheets
'
'

ActiveSheet.Cells(3, 16).Value = ActiveSheet.Cells(2, 2).Value
ActiveSheet.Cells(28, 8).Value = ActiveSheet.Cells(2, 2).Value
For i = 0 To Round - 1
    ActiveSheet.Cells(i + 4, 16).Value = ActiveSheet.Cells(i + 3, 16).Value + Round_date
    ActiveSheet.Cells(28, 9 + i).Value = ActiveSheet.Cells(i + 3, 16).Value + Round_date
Next i

'
'calcualte payment for owner
'
'
For i = 1 To Round
    ActiveSheet.Cells(i + 3, 17).Value = (OwnerPayment * Round) * -1
Next i

'
'Assign handDead
'

ActiveSheet.Cells(100, 1).Value = "=SUM(L3:L23)"
Sum_handDead = ActiveSheet.Cells(100, 1).Value

'
'calculate payment for each times.
'calcualte summery for onwer
'

'รายรับ (มือฟรี)
'ActiveSheet.Cells(27, 17).Value = ActiveSheet.Cells(3, 7).Value
ActiveSheet.Cells(19, 2).Value = ActiveSheet.Cells(3, 7).Value
'ค่าดูแล
'ActiveSheet.Cells(28, 17).Value = "=SUM(I3:I24)*-1"
ActiveSheet.Cells(20, 2).Value = "=SUM(I3:I24)*-1"
'รายจ่าย
'ActiveSheet.Cells(29, 17).Value = "=SUM(Q3:Q24)"
ActiveSheet.Cells(21, 2).Value = "=SUM(Q3:Q24)"
'มือตาย
'ActiveSheet.Cells(100, 1).Value = ActiveSheet.Cells(100, 1).Value * EachPayment * -1
ActiveSheet.Cells(22, 2).Value = ActiveSheet.Cells(100, 1).Value * EachPayment * -1

'หักหนี้
'ActiveSheet.Cells(33, 17).Value = "=SUM(M3:M24)*-1"
ActiveSheet.Cells(25, 2).Value = "=SUM(M3:M24)*-1"

'รวมก่อนหักหนี้
'ActiveSheet.Cells(31, 17).Value = "=SUM(Q27:Q30)"
ActiveSheet.Cells(23, 2).Value = "=SUM(B19:B22)"

'ยอดสุทธิ
'ActiveSheet.Cells(34, 17).Value = "=SUM(Q31:Q33)"
ActiveSheet.Cells(26, 2).Value = "=SUM(B23:B25)"

'
'input all member into member calcutate sheets
'
'

For i = 1 To Round + 1
'name of X
   ActiveSheet.Cells(i + 28, 6).Value = ActiveSheet.Cells(i + 2, 6).Value
    'name of Y
   ActiveSheet.Cells(52, 7 + i).Value = ActiveSheet.Cells(i + 2, 6).Value
    ' Receiving of each person
   ActiveSheet.Cells(i + 28, 7).Value = ActiveSheet.Cells(i + 2, 7).Value
    'Fee and service charge
   ActiveSheet.Cells(49, 7 + i).Value = ActiveSheet.Cells(i + 2, 9).Value
    '
    
    If ActiveSheet.Cells(i + 2, 17).Value < 0 Then
        ActiveSheet.Cells(29, 7 + i).Value = ActiveSheet.Cells(i + 2, 17).Value * -1
    End If
    
   Free_price = ActiveSheet.Cells(4, 2).Value
   
   If Free_price > 0 Then
    If Round - i >= 0 Then
        ActiveSheet.Cells(29 + i, 8).Value = Free_price / Round
    End If
   End If
   
Next i

'
'Forward and revese calcualation
'
'

Dim Sum_forward As Double
Dim Sum_reverse As Double
Dim Sum_income As Double


Round = ActiveSheet.Cells(5, 2).Value
EachPayment = ActiveSheet.Cells(6, 2).Value
OwnerPayment = ActiveSheet.Cells(9, 2).Value
Plus_interest = ActiveSheet.Cells(10, 2).Value
FreeForOwner = ActiveSheet.Cells(4, 2).Value
Fee = ActiveSheet.Cells(8, 2).Value

Dim countx As Integer

For i = 0 To Round - 1
 countx = 1
  
  For y = Round - i To Round - 1
    'forward
    ActiveSheet.Cells(29 + countx, 9 + i).Value = Plus_interest + EachPayment
    countx = countx + 1
  Next y
  
  For j = i To Round - 1

    'reverse
    ActiveSheet.Cells(j + 30, 9 + i).Value = EachPayment
  Next j
 
Next i

Dim sh As Worksheet
Dim Count_handDead As Integer
Dim indexes()
Dim Count_round As Integer
Dim Minus_count As Integer

Set sh = ThisWorkbook.Sheets("คำนวน")

Count_handDead = sh.Cells(4, 23).Rows.Count


'indexes(1 To Count_handDead)

' hightlight handDead
If Count_handDead > 0 Then
    For i = 1 To Round
        If ActiveSheet.Cells(3 + i, 12).Value > 0 Then
            Count_round = ActiveSheet.Cells(3 + i, 12).Value
            Minus_count = Round - Count_round
            For j = 1 To Count_round
                ActiveSheet.Cells(29 + i, 8 + Minus_count + j).Select
                With Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .ThemeColor = xlThemeColorAccent2
                    .TintAndShade = -0.249977111117893
                    .PatternTintAndShade = 0
                End With
            Next j
        End If
    Next i
End If






End Sub


Sub Credit_dept()

Dim Round As Double
Round = ActiveSheet.Cells(5, 2).Value

For i = 1 To Round
    ActiveSheet.Cells(i + 3, 14).Value = ActiveSheet.Cells(i + 3, 13).Value + ActiveSheet.Cells(i + 3, 11).Value
Next i

End Sub

Sub clear()

    Range("E3:E23").Select
    Selection.ClearContents
    Range("G3:N23").Select
    Selection.ClearContents
    Range("P3:R23").Select
    Selection.ClearContents
    Range("F4:F23").Select
    Selection.ClearContents
    Range("B19:B26").Select
    Selection.ClearContents
    
    'ลูกแชร์
    Range("H28:R49").Select
    ActiveWindow.SmallScroll Down:=6
    Selection.ClearContents
    Range("F29:G48").Select
    Selection.ClearContents
    Range("H52:R52").Select
    Selection.ClearContents
    'Clear Color
    Range("H29:R48").Select
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
End Sub

Sub Update_data()
    Dim Round As Double
    'Dim Names As String
    
    Round = ActiveSheet.Cells(5, 2).Value
    'Names = ActiveSheet.Cells(3, 6).Value
    For i = 1 To Round
        If ActiveSheet.Cells(i + 2, 6).Value = "" Then
            MsgBox "กรุณาใส่ข้อมูลสมาชิคของวงแชร์ให้ครบ"
            Exit Sub
        End If
    Next i
    
    '
    'initail an array for collect data from calculation sheet and move data to fact data sheet
    '
    Dim sheet_dead As Worksheet
    Set sheet_dead = ThisWorkbook.Sheets("คำนวน")

    Dim Count_dead As Long
    Dim sum_price_dead As Double
    
    Count_dead = sheet_dead.Range("L3:L24").Rows.Count
    If Count_dead > 0 Then
        sum_price_dead = ActiveSheet.Cells(100, 1).Value
    End If
    
    '
    '
    '
    Dim EachPayment As Double
    EachPayment = ActiveSheet.Cells(6, 2).Value
    
    
    Dim Data()
    Round = Round + 1
    ReDim Data(1 To Round, 1 To 7)
    
    'K = sh.Range("A1", sh.Range("A1").End(xlDown)).Rows.Count
    
    For i = 1 To Round
        
            'date
            Data(i, 1) = Sheets("คำนวน").Cells(i + 2, 16).Value
            'Name_wong
            Data(i, 2) = Sheets("คำนวน").Cells(3, 2).Value
            'Name
            Data(i, 3) = Sheets("คำนวน").Cells(i + 2, 6).Value
            'Dead
            'If Sheets("คำนวน").Cells(i + 2, 12).Value <> 0 Then
            '    Data(i, 6) = Sheets("คำนวน").Cells(i + 2, 12).Value * EachPayment
            'End If
            
            'decrese dept
            If Sheets("คำนวน").Cells(i + 2, 13).Value <> 0 Then
                Data(i, 6) = Sheets("คำนวน").Cells(i + 2, 13).Value
            End If
            
            'paid or receive
                If Data(i, 3) = "ท้าว" Then
                    Data(i, 4) = Sheets("คำนวน").Cells(i + 2, 7).Value
                Else
                    If Sheets("คำนวน").Cells(i + 2, 17).Value > 0 Then
                        Data(i, 4) = Sheets("คำนวน").Cells(i + 2, 17).Value
                    End If
                End If
    Next i
    'Fee , Service charge
            If Sheets("คำนวน").Cells(8, 2).Value <> 0 Then
                Data(1, 7) = Sheets("คำนวน").Cells(8, 2).Value * Sheets("คำนวน").Cells(5, 2).Value
            End If
    
    
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Sheets("รวมวงแชร์")

    Dim k As Long

    k = sh.Range("A1", sh.Range("A1").End(xlDown)).Rows.Count
    If k = 1048576 Then
        k = 1
    End If
    'MsgBox (k)
    
    
    'part checking wongName
    Dim wongName As Double
    Dim dates As Date
    dates = Data(1, 1)
    
    If Sheets("รวมวงแชร์").Cells(2, 1).Value = "" Then
        wongName = 1
    Else
        wongName = Check_wong(Sheets("คำนวน").Cells(3, 2).Value, dates, k)
    End If
    
    If wongName <> 0 Then
        For i = 1 To Round
            'date
            Sheets("รวมวงแชร์").Cells(i + k, 1).Value = Data(i, 1)
            'Name_wong
            Sheets("รวมวงแชร์").Cells(i + k, 2).Value = Data(i, 2)
            'Name
            Sheets("รวมวงแชร์").Cells(i + k, 3).Value = Data(i, 3)
            
            'paid or receive
                If Data(i, 3) = "ท้าว" Then
                    Sheets("รวมวงแชร์").Cells(i + k, 5).Value = Data(i, 4)
                Else
                    Sheets("รวมวงแชร์").Cells(i + k, 4).Value = Data(i, 4)
                End If
            'Decrese dept
            Sheets("รวมวงแชร์").Cells(i + k, 7).Value = Data(i, 6)
            'Fee, service charge
            Sheets("รวมวงแชร์").Cells(i + k, 8).Value = Data(i, 7)
        Next i
        MsgBox "เพิ่มวงแชร์เสร็จเรียบร้อยแล้ว"
    Else
        MsgBox "ชออภัยคุณมีวงแชร์ชื่อนี้ในระบบเรียบร้อยแล้วและวงนี้ยังไม่ปิดวง"
        Exit Sub
    End If
    
    
End Sub


' this function is for checking wong name and when we need avoid dupplication

Function Check_wong(name As String, dates As Date, Count As Long) As Double
    
    Dim checkindex As Boolean
    checkindex = False
    'Dim Last_index As Long
    Dim datess As Date
    Dim old_name As String
    
    For i = 2 To Count
        old_name = Sheets("รวมวงแชร์").Cells(i, 2).Value
        If old_name = name Then
            'Check_wong = i
            'checkindex = True
            'Exit For
            
            For j = 1 To 50
                If Sheets("รวมวงแชร์").Cells(j + i, 2).Value <> name Or Sheets("รวมวงแชร์").Cells(j + i, 2).Value = "" Then
                  'First_index = i
                  'Last_index = i + j
                  
                  'old date
                    datess = Sheets("รวมวงแชร์").Cells(j + i - 1, 1).Value
                    
                    If datess < dates Then
                        Check_wong = i
                        checkindex = True
                        GoTo exitLoop
                        'Exit For
                    Else
                        Check_wong = 0
                        checkindex = False
                        GoTo exitLoop
                    End If
                  
                  'Exit For
                End If
            Next j
        End If
    Next i
    
    Check_wong = 1
    
exitLoop:
    
    
End Function

Sub edit_wong()

    Dim Round As Double
    'Dim Names As String
    
    Round = ActiveSheet.Cells(5, 2).Value
    'Names = ActiveSheet.Cells(3, 6).Value
    For i = 1 To Round
        If ActiveSheet.Cells(i + 2, 6).Value = "" Then
            MsgBox "กรุณาใส่ข้อมูลสมาชิคของวงแชร์ให้ครบ"
            Exit Sub
        End If
    Next i
    
    '
    'initail an array for collect data from calculation sheet and move data to fact data sheet
    '
    Dim sheet_dead As Worksheet
    Set sheet_dead = ThisWorkbook.Sheets("คำนวน")

    Dim Count_dead As Long
    Dim sum_price_dead As Double
    
    Count_dead = sheet_dead.Range("L3:L24").Rows.Count
    If Count_dead > 0 Then
        sum_price_dead = ActiveSheet.Cells(100, 1).Value
    End If
    
    '
    '
    '
    Dim EachPayment As Double
    EachPayment = ActiveSheet.Cells(6, 2).Value
    
    
    Dim Data()
    Round = Round + 1
    ReDim Data(1 To Round, 1 To 7)
    
    'K = sh.Range("A1", sh.Range("A1").End(xlDown)).Rows.Count
    
    For i = 1 To Round
        
            'date
            Data(i, 1) = Sheets("คำนวน").Cells(i + 2, 16).Value
            'Name_wong
            Data(i, 2) = Sheets("คำนวน").Cells(3, 2).Value
            'Name
            Data(i, 3) = Sheets("คำนวน").Cells(i + 2, 6).Value
            'Dead
            If Sheets("คำนวน").Cells(i + 2, 12).Value <> 0 Then
                Data(i, 5) = Sheets("คำนวน").Cells(i + 2, 12).Value * EachPayment
            End If
            
            'decrese dept
            If Sheets("คำนวน").Cells(i + 2, 13).Value <> 0 Then
                Data(i, 6) = Sheets("คำนวน").Cells(i + 2, 13).Value
            End If
            
            
            'paid or receive
                If Data(i, 3) = "ท้าว" Then
                    Data(i, 4) = Sheets("คำนวน").Cells(i + 2, 7).Value
                Else
                    Data(i, 4) = Sheets("คำนวน").Cells(i + 2, 17).Value
                End If
    Next i
    'Fee , Service charge
            If Sheets("คำนวน").Cells(8, 2).Value <> 0 Then
                Data(1, 7) = Sheets("คำนวน").Cells(8, 2).Value * Sheets("คำนวน").Cells(5, 2).Value
            End If
    
    
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Sheets("รวมวงแชร์")

    Dim k As Long

    k = sh.Range("A1", sh.Range("A1").End(xlDown)).Rows.Count
    If k = 1048576 Then
        k = 1
    End If
    'MsgBox (k)
    
    
    'part checking wongName
    Dim wongName As Double
    Dim dates As Date
    dates = Data(1, 1)
    
    If Sheets("รวมวงแชร์").Cells(2, 1).Value = "" Then
        wongName = 1
    Else
        wongName = Check_wong_edit(Sheets("คำนวน").Cells(3, 2).Value, dates, k)
    End If
    Dim replace_index As Double
    replace_index = wongName
    
    
    If wongName > 1 Then
        For i = 1 To Round
            'date
            Sheets("รวมวงแชร์").Cells(replace_index, 1).Value = Data(i, 1)
            'Name_wong
            Sheets("รวมวงแชร์").Cells(replace_index, 2).Value = Data(i, 2)
            'Name
            Sheets("รวมวงแชร์").Cells(replace_index, 3).Value = Data(i, 3)
            
            'paid or receive
                If Data(i, 3) = "ท้าว" Then
                    Sheets("รวมวงแชร์").Cells(replace_index, 5).Value = Data(i, 4)
                Else
                    Sheets("รวมวงแชร์").Cells(replace_index, 4).Value = Data(i, 4)
                End If
            'Dead
                If Data(i, 5) > 0 Then
                    Sheets("รวมวงแชร์").Cells(replace_index, 6).Value = Data(i, 5) * -1
                End If
            'Decrese dept
            Sheets("รวมวงแชร์").Cells(replace_index, 7).Value = Data(i, 6)
            'Fee, service charge
            Sheets("รวมวงแชร์").Cells(replace_index, 8).Value = Data(i, 7)
            
            replace_index = replace_index + 1
        Next i
        MsgBox "แก้ไขเสร็จเรียบร้อยแล้ว"
    Else
        MsgBox "ขออภัยไม่พบวงที่คุณต้องการเปลี่ยน"
        Exit Sub
    End If
    

End Sub

