Option Explicit

Sub Calculate_Target()

Dim wb As Workbook: Set wb = ThisWorkbook

Dim SKUList As Worksheet
Dim QualitySpec As Worksheet
Dim mainPage As Worksheet
Dim Target As Worksheet



Dim fromId As Long
Dim toId As Long
Dim LastRowQualitySpec As Long
Dim i As Long
Dim rng_from As Range
Dim rng_to As Range
Dim rownumber_from As Long
Dim rownumber_to As Long
Dim trigger As Long


'Sheet definitions
Set SKUList = wb.Sheets("SKU_List")
Set QualitySpec = wb.Sheets("Quality_Spec")
Set mainPage = wb.Sheets("Main")
Set Target = wb.Sheets("Target")

'Static Values

Dim DataCodeMethod As Integer
DataCodeMethod = 28

Dim Parameter1 As Integer
Parameter1 = 31

Dim Parameter2 As Integer
Parameter2 = 34

Dim Parameter3 As Integer
Parameter3 = 35

Dim Parameter4 As Integer
Parameter4 = 36
                              'COLUMN NUMBERS ON QualitySpec SHEED
Dim Parameter5 As Integer
Parameter5 = 41

Dim Parameter6 As Integer
Parameter6 = 42

Dim Parameter7 As Integer
Parameter7 = 49

Dim Parameter8 As Integer
Parameter8 = 50

Dim Parameter9 As Integer
Parameter9 = 56


LastRowQualitySpec = QualitySpec.Cells(QualitySpec.Rows.Count, "E").End(xlUp).Row


fromId = mainPage.Cells(3, 2)
toId = mainPage.Cells(4, 2)
trigger = 0

If fromId = 0 And toId = 0 Then
        MsgBox "Mevcut ve Yeni SKU Numaralarını Giriniz."
        Exit Sub
Else
    For i = 2 To LastRowQualitySpec
        If fromId = QualitySpec.Cells(i, 5) Then
            trigger = 1
            Exit For
        End If
    Next i

    If trigger = 0 Then
        MsgBox "Mevcut SKU Numarası Girilmedi. Lütfen Giriniz."
        Exit Sub
    End If
    trigger = 0

    For i = 2 To LastRowQualitySpec
        If toId = QualitySpec.Cells(i, 5) Then
            trigger = 1
            Exit For
        End If
    Next i
    
    If trigger = 0 Then
        MsgBox "Yeni SKU Numarası Girilmedi. Lütfen Giriniz."
        Exit Sub
    End If
End If

'market

Set rng_from = QualitySpec.Columns("E:E").Find(what:=fromId, LookIn:=xlFormulas, Lookat:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
If Not rng_from Is Nothing Then
    rownumber_from = rng_from.Row
    mainPage.Cells(10, 1).Value = Right(QualitySpec.Cells(rownumber_from, 6).Value, 3)
    
End If

Set rng_to = QualitySpec.Columns("E:E").Find(what:=toId, LookIn:=xlFormulas, Lookat:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
    
If Not rng_to Is Nothing Then
    rownumber_to = rng_to.Row
    mainPage.Cells(10, 4).Value = Right(QualitySpec.Cells(rownumber_to, 6).Value, 3)
       
End If

'comparing market
If mainPage.Cells(10, 1).Value <> mainPage.Cells(10, 4).Value Then

    mainPage.Cells(10, 8).Value = "Changed"
    
Else
    mainPage.Cells(10, 8).Value = "Not Changed"
    
End If


'Brand

If Not rng_from Is Nothing Then
    rownumber_from = rng_from.Row
    mainPage.Cells(13, 1).Value = Mid(QualitySpec.Cells(rownumber_from, 6).Value, 4, 3)
    
End If

Set rng_to = QualitySpec.Columns("E:E").Find(what:=toId, LookIn:=xlFormulas, Lookat:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
    
If Not rng_to Is Nothing Then
    rownumber_to = rng_to.Row
    mainPage.Cells(13, 4).Value = Mid(QualitySpec.Cells(rownumber_to, 6).Value, 4, 3)
       
End If

'comparing brand
If mainPage.Cells(13, 1).Value <> mainPage.Cells(13, 4).Value Then
    mainPage.Cells(13, 8).Value = "Changed"
    
Else
    mainPage.Cells(13, 8).Value = "Not Changed"
    
End If


'Data Code Method:

Set rng_from = QualitySpec.Columns("E:E").Find(what:=fromId, LookIn:=xlFormulas, Lookat:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
    
If Not rng_from Is Nothing Then
    rownumber_from = rng_from.Row
    mainPage.Cells(16, 1).Value = QualitySpec.Cells(rownumber_from, DataCodeMethod).Value
 
End If
    
Set rng_to = QualitySpec.Columns("E:E").Find(what:=toId, LookIn:=xlFormulas, Lookat:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
    
If Not rng_to Is Nothing Then
    rownumber_to = rng_to.Row
    mainPage.Cells(16, 4).Value = QualitySpec.Cells(rownumber_to, DataCodeMethod).Value
       
End If

Dim DataCode1 As String
Dim DataCode2 As String

'comparing of data code method
DataCode1 = mainPage.Cells(16, 1).Value
DataCode2 = mainPage.Cells(16, 4).Value

If DataCode1 = DataCode2 Then

    mainPage.Cells(16, 8).Value = "Not Changed"
    'Değişmiyorsa yeşil
Else

    mainPage.Cells(16, 8).Value = "Changed"
    'değişiyorsa kırmızı
    
End If


'Parameter1:

Set rng_from = QualitySpec.Columns("E:E").Find(what:=fromId, LookIn:=xlFormulas, Lookat:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
  
If Not rng_from Is Nothing Then

    rownumber_from = rng_from.Row
    
    mainPage.Cells(19, 1).Value = QualitySpec.Cells(rownumber_from, Parameter1).Value
        
End If
        
    
Set rng_to = QualitySpec.Columns("E:E").Find(what:=toId, LookIn:=xlFormulas, Lookat:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
  
If Not rng_to Is Nothing Then

    rownumber_to = rng_to.Row
    
    mainPage.Cells(19, 4).Value = QualitySpec.Cells(rownumber_to, Parameter1).Value

End If

'comparing of 'Parameter1:

If mainPage.Cells(19, 1).Value <> mainPage.Cells(19, 4).Value Then

    mainPage.Cells(19, 8).Value = "Changed"
Else
    mainPage.Cells(19, 8).Value = "Not Changed"
    
End If

     
   
' Parameter2

Set rng_from = QualitySpec.Columns("E:E").Find(what:=fromId, LookIn:=xlFormulas, Lookat:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
  
If Not rng_from Is Nothing Then

    rownumber_from = rng_from.Row
    
    mainPage.Cells(22, 1).Value = QualitySpec.Cells(rownumber_from, Parameter2).Value

End If
        
    
Set rng_to = QualitySpec.Columns("E:E").Find(what:=toId, LookIn:=xlFormulas, Lookat:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
  
If Not rng_to Is Nothing Then

    rownumber_to = rng_to.Row
    
    mainPage.Cells(22, 4).Value = QualitySpec.Cells(rownumber_to, Parameter2).Value
        
End If

'comparing Parameter2

If mainPage.Cells(22, 1) <> mainPage.Cells(22, 4) Then

    mainPage.Cells(22, 8).Value = "Changed"
Else
    
    mainPage.Cells(22, 8).Value = "Not Changed"
    
End If

'checking if IA or not

If mainPage.Cells(22, 1).Value = 0 Then

    mainPage.Cells(22, 1).Value = "Ment"
    
End If

If mainPage.Cells(22, 4).Value = 0 Then

    mainPage.Cells(22, 4).Value = "Ment"
    
End If

'Parameter3

Set rng_from = QualitySpec.Columns("E:E").Find(what:=fromId, LookIn:=xlFormulas, Lookat:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
  
If Not rng_from Is Nothing Then

    rownumber_from = rng_from.Row
    
    mainPage.Cells(25, 1).Value = QualitySpec.Cells(rownumber_from, Parameter3).Value

End If
        
    
Set rng_to = QualitySpec.Columns("E:E").Find(what:=toId, LookIn:=xlFormulas, Lookat:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
  
If Not rng_to Is Nothing Then

    rownumber_to = rng_to.Row
    
    mainPage.Cells(25, 4).Value = QualitySpec.Cells(rownumber_to, Parameter3).Value

        
End If

'comparing Parameter3



If mainPage.Cells(25, 1).Value <> mainPage.Cells(25, 4).Value Then

    mainPage.Cells(25, 8) = "Changed"
Else
     mainPage.Cells(25, 8) = "Not Changed"
     
End If

'checking if IA or not

If mainPage.Cells(25, 1).Value = 0 Then

    mainPage.Cells(25, 1).Value = "Ment"
    
End If

If mainPage.Cells(25, 4).Value = 0 Then

    mainPage.Cells(25, 4).Value = "Ment                                                            "
    
End If
                                                      

'Parameter4

Set rng_from = QualitySpec.Columns("E:E").Find(what:=fromId, LookIn:=xlFormulas, Lookat:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
  
If Not rng_from Is Nothing Then

    rownumber_from = rng_from.Row
    
    mainPage.Cells(28, 1).Value = QualitySpec.Cells(rownumber_from, Parameter4).Value

End If
        
    
Set rng_to = QualitySpec.Columns("E:E").Find(what:=toId, LookIn:=xlFormulas, Lookat:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
  
If Not rng_to Is Nothing Then

    rownumber_to = rng_to.Row
    
    mainPage.Cells(28, 4).Value = QualitySpec.Cells(rownumber_to, Parameter4).Value

        
End If

' comparing Parameter4

If mainPage.Cells(28, 1).Value <> mainPage.Cells(28, 4).Value Then

    mainPage.Cells(28, 8).Value = "Changed"
Else
    mainPage.Cells(28, 8).Value = "Not Changed"
    
End If

'checking if IA or not
If mainPage.Cells(28, 1).Value = 0 Then
    mainPage.Cells(28, 1).Value = "Mentollü"
End If

If mainPage.Cells(28, 4).Value = 0 Then
    mainPage.Cells(28, 4).Value = "Mentollü"
End If


'Parameter5

Set rng_from = QualitySpec.Columns("E:E").Find(what:=fromId, LookIn:=xlFormulas, Lookat:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
  
If Not rng_from Is Nothing Then

    rownumber_from = rng_from.Row
    
    mainPage.Cells(31, 1).Value = QualitySpec.Cells(rownumber_from, Parameter5).Value

End If
        
    
Set rng_to = QualitySpec.Columns("E:E").Find(what:=toId, LookIn:=xlFormulas, Lookat:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
  
If Not rng_to Is Nothing Then

    rownumber_to = rng_to.Row
    
    mainPage.Cells(31, 4).Value = QualitySpec.Cells(rownumber_to, Parameter5).Value
    
End If

'Parameter5
If mainPage.Cells(31, 1).Value <> mainPage.Cells(31, 4).Value Then

    mainPage.Cells(31, 8) = "Changed"
    
Else

    mainPage.Cells(31, 8) = "Not Changed"

End If



'Parameter6

Set rng_from = QualitySpec.Columns("E:E").Find(what:=fromId, LookIn:=xlFormulas, Lookat:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
  
If Not rng_from Is Nothing Then

    rownumber_from = rng_from.Row
    
    mainPage.Cells(34, 1).Value = QualitySpec.Cells(rownumber_from, Parameter6).Value

End If
        
    
Set rng_to = QualitySpec.Columns("E:E").Find(what:=toId, LookIn:=xlFormulas, Lookat:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
  
If Not rng_to Is Nothing Then

    rownumber_to = rng_to.Row
    
    mainPage.Cells(34, 4).Value = QualitySpec.Cells(rownumber_to, Parameter6).Value
    
End If

'comparing Parameter6

If mainPage.Cells(34, 1).Value <> mainPage.Cells(34, 4).Value Then

    mainPage.Cells(34, 8).Value = "Changed"
Else

    mainPage.Cells(34, 8).Value = "Not Changed"
    
End If
   
'FCP/TX Width

Set rng_from = QualitySpec.Columns("E:E").Find(what:=fromId, LookIn:=xlFormulas, Lookat:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
  
If Not rng_from Is Nothing Then

    rownumber_from = rng_from.Row
    
    mainPage.Cells(37, 1).Value = QualitySpec.Cells(rownumber_from, Parameter8).Value

End If
        
    
Set rng_to = QualitySpec.Columns("E:E").Find(what:=toId, LookIn:=xlFormulas, Lookat:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
  
If Not rng_to Is Nothing Then

    rownumber_to = rng_to.Row
    
    mainPage.Cells(37, 4).Value = QualitySpec.Cells(rownumber_to, Parameter8).Value
    
End If

'comparing Parameter8

If mainPage.Cells(37, 1).Value <> mainPage.Cells(37, 4).Value Then

    mainPage.Cells(37, 8).Value = "Changed"
    
Else
    mainPage.Cells(37, 8).Value = "Not Changed"

End If

'Parameter7

Set rng_from = QualitySpec.Columns("E:E").Find(what:=fromId, LookIn:=xlFormulas, Lookat:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
  
If Not rng_from Is Nothing Then

    rownumber_from = rng_from.Row
    
    mainPage.Cells(40, 1).Value = QualitySpec.Cells(rownumber_from, Parameter7).Value

End If
        
    
Set rng_to = QualitySpec.Columns("E:E").Find(what:=toId, LookIn:=xlFormulas, Lookat:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
  
If Not rng_to Is Nothing Then

    rownumber_to = rng_to.Row
    
    mainPage.Cells(40, 4).Value = QualitySpec.Cells(rownumber_to, Parameter7).Value
    
End If

If mainPage.Cells(40, 1).Value <> mainPage.Cells(40, 4).Value Then

    mainPage.Cells(40, 8) = "Changed"
    
Else

    mainPage.Cells(40, 8) = "Not Changed"
    
End If

'X size change

mainPage.Cells(43, 1).Value = QualitySpec.Cells(rownumber_from, Parameter8).Value & "x" & QualitySpec.Cells(rownumber_from, Parameter7).Value

mainPage.Cells(43, 4).Value = QualitySpec.Cells(rownumber_to, Parameter8).Value & "x" & QualitySpec.Cells(rownumber_to, Parameter7).Value

'comparing X size change
If mainPage.Cells(43, 1).Value <> mainPage.Cells(43, 4).Value Then

    mainPage.Cells(43, 8).Value = "Changed"
    
Else
    mainPage.Cells(43, 8).Value = "Not Changed"

End If


'Parameter9

Set rng_from = QualitySpec.Columns("E:E").Find(what:=fromId, LookIn:=xlFormulas, Lookat:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
  
If Not rng_from Is Nothing Then

    rownumber_from = rng_from.Row
    
    mainPage.Cells(46, 1).Value = QualitySpec.Cells(rownumber_from, Parameter9).Value

End If
        
    
Set rng_to = QualitySpec.Columns("E:E").Find(what:=toId, LookIn:=xlFormulas, Lookat:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
  
If Not rng_to Is Nothing Then
    rownumber_to = rng_to.Row
    
    mainPage.Cells(46, 4).Value = QualitySpec.Cells(rownumber_to, Parameter9).Value
    
End If

    'comparision of Parameter9
    
    If mainPage.Cells(46, 1).Value <> mainPage.Cells(46, 4).Value Then
    
        mainPage.Cells(46, 8).Value = "Changed"
        
    Else
        mainPage.Cells(46, 8).Value = "Not Changed"
        
    
  
End If

'coloring the changed cells
Dim C As Long

For C = 8 To 50

If mainPage.Cells(C, 8).Value = "Changed" Then
    mainPage.Cells(C, 8).Interior.color = 255 '(red color)
    

ElseIf mainPage.Cells(C, 8).Value = "Not Changed" Then
    mainPage.Cells(C, 8).Interior.color = 5287936 '(green color)

Else
    mainPage.Cells(C, 8).Interior.color = xlNone

End If

Next




'Target Setting


Dim Target1 As Integer
Target1 = 0

'Brand Change

If mainPage.Cells(13, 8).Value = "Changed" Then
    Target1 = Target1 + Target.Cells(11, 3).Value
End If
                
                
'Embosser Change

If mainPage.Cells(16, 8).Value = "Changed" Then
    Target1 = Target1 + Target.Cells(7, 3).Value
End If


'NX Size Change

'20x44 ---> 17x43
If mainPage.Cells(43, 1).Value = "20,00 mmx44,00 mm" And mainPage.Cells(43, 4).Value = "17,00 mmx43,00 mm" Then
    Target1 = Target1 + Target.Cells(3, 3).Value
End If

If mainPage.Cells(43, 1).Value = "17,00 mmx43,00 mm" And mainPage.Cells(43, 4).Value = "20,00 mmx44,00 mm" Then
    Target1 = Target1 + Target.Cells(3, 3).Value
End If
'20x44 ---> 12x33
If mainPage.Cells(43, 1).Value = "20,00 mmx44,00 mm" And mainPage.Cells(43, 4).Value = "12,00 mmx33,00 mm" Then
    Target = Target1 + Target.Cells(4, 3).Value
End If

If mainPage.Cells(43, 1).Value = "12,00 mmx33,00 mm" And mainPage.Cells(43, 4).Value = "20,00 mmx44,00 mm" Then
    Target1 = Target1 + Target.Cells(4, 3).Value
End If

'12x33 ---> 17x43
If mainPage.Cells(43, 1).Value = "12,00 mmx33,00 mm" And mainPage.Cells(43, 4).Value = "17,00 mmx43,00 mm" Then
    Target1 = Target1 + Target.Cells(5, 3).Value
End If

If mainPage.Cells(43, 1).Value = "17,00 mmx43,00 mm" And mainPage.Cells(43, 4).Value = "12,00 mmx33,00 mm" Then
    Target1 = Target1 + Target.Cells(5, 3).Value
End If

'18x40 ---> 20x44
If mainPage.Cells(43, 1).Value = "18,00 mmx40,00 mm" And mainPage.Cells(43, 4).Value = "20,00 mmx44,00 mm" Then
    Target1 = Target1 + Target.Cells(2, 3).Value
End If

If mainPage.Cells(43, 1).Value = "20,00 mmx44,00 mm" And mainPage.Cells(43, 4).Value = "18,00 mmx40,00 mm" Then
    Target1 = Target1 + Target.Cells(2, 3).Value
End If

'IF Change
If mainPage.Cells(31, 8).Value = "Changed" Or mainPage.Cells(34, 8).Value = "Changed" Then
    Target1 = Target1 + Target.Cells(9, 3).Value
End If

mainPage.Cells(5, 2).Value = Target1

'Maker Size Change
Dim Target2 As Integer

If mainPage.Cells(46, 8).Value = "Changed" Then
    Target2 = Target.Cells(13, 3)
End If

    
'Comparing Target1 and Target2
If Target1 > Target2 Then
    mainPage.Cells(5, 2).Value = Target1
End If
If Target2 > Target1 Then
    mainPage.Cells(5, 2).Value = Target2
End If


End Sub

Sub Clear()

Dim wb As Workbook: Set wb = ThisWorkbook

Dim mainPage As Worksheet

Set mainPage = wb.Sheets("Main")

Sheets("Main").Range("A10", "D10").Value = ""
Sheets("Main").Range("A13", "D13").Value = ""
Sheets("Main").Range("A16", "D16").Value = ""
Sheets("Main").Range("A19", "D19").Value = ""
Sheets("Main").Range("A22", "D22").Value = ""
Sheets("Main").Range("A25", "D25").Value = ""
Sheets("Main").Range("A28", "D28").Value = ""
Sheets("Main").Range("A31", "D31").Value = ""
Sheets("Main").Range("A34", "D34").Value = ""
Sheets("Main").Range("A37", "D37").Value = ""
Sheets("Main").Range("A40", "D40").Value = ""
Sheets("Main").Range("A43", "D43").Value = ""
Sheets("Main").Range("A46", "D46").Value = ""

Sheets("Main").Range("H10").Value = ""
Sheets("Main").Range("H13").Value = ""
Sheets("Main").Range("H16").Value = ""
Sheets("Main").Range("H19").Value = ""
Sheets("Main").Range("H22").Value = ""
Sheets("Main").Range("H25").Value = ""
Sheets("Main").Range("H28").Value = ""
Sheets("Main").Range("H31").Value = ""
Sheets("Main").Range("H34").Value = ""
Sheets("Main").Range("H37").Value = ""
Sheets("Main").Range("H40").Value = ""
Sheets("Main").Range("H43").Value = ""
Sheets("Main").Range("H46").Value = ""

Sheets("Main").Range("B5").Value = ""

Dim C As Long

For C = 10 To 50
mainPage.Cells(C, 8).Interior.color = xlNone 'removing color

Next

End Sub
