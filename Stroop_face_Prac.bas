Attribute VB_Name = "Module3"
Sub 表格整理_stroop_face_prac()
'
' 表格整理_stroop_face_prac 巨集
'
'
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    sheetname = ActiveSheet.Name
    
    SheetReplace ("工作表2")
    SheetReplace ("表格整理")
    
    NewSheetName = ActiveSheet.Name
    Sheets(NewSheetName).Select
    Range("A1").Value = "cond"
    Range("B1").Value = "Sex"
    Range("C1").Value = "Emotion"
    Range("D1").Value = "Color"
    Range("E1").Value = "ACC"
    Range("F1").Value = "RT"
    
    Sheets(sheetname).Select
    m = ActiveSheet.UsedRange.Rows.Count
    n = 1
    
    If Range("B2").Value = "" Then
        For i = 1 To m
            S = Range("A" & i).Value & Range("B" & i).Value & Range("C" & i).Value & Range("D" & i).Value & Range("E" & i).Value
            If InStr(1, S, "*** LogFrame Start ***", vbTextCompare) > 0 Then
                n = n + 1
            ElseIf InStr(1, S, "*** LogFrame End ***", vbTextCompare) > 0 Then
                n = n + 0
            ElseIf InStr(1, S, "cond:", vbTextCompare) > 0 Then
                Sheets("工作表2").Select
                Range("A1").Value = S
                Call splitstr(":")
                Range("B1").Select
                Selection.Copy
                Sheets("表格整理").Select
                Range("A" & CStr(n)).Select
                ActiveSheet.Paste
                Sheets(sheetname).Select
            ElseIf InStr(1, S, "Face:", vbTextCompare) > 0 Then
                Sheets("工作表2").Select
                Range("A1").Value = S
                Call splitstr(":")
                Range("A1").Value = Range("B1").Value
                Call splitstr("_")
                
                Range("A1").Select
                Selection.Copy
                Sheets("表格整理").Select
                Range("B" & CStr(n)).Select
                ActiveSheet.Paste
                
                Sheets("工作表2").Select
                Range("B1").Select
                Selection.Copy
                Sheets("表格整理").Select
                Range("C" & CStr(n)).Select
                ActiveSheet.Paste
                Sheets(sheetname).Select
            ElseIf InStr(1, S, "PicName:", vbTextCompare) > 0 Then
                Sheets("工作表2").Select
                Range("A1").Value = S
                Call splitstr(":")
                Range("B1").Value = Left(Trim(Range("B1").Value), 1)
                Range("B1").Select
                Selection.Copy
                Sheets("表格整理").Select
                Range("D" & CStr(n)).Select
                ActiveSheet.Paste
                Sheets(sheetname).Select
            ElseIf InStr(1, S, "Target.ACC:", vbTextCompare) > 0 Then
                Sheets("工作表2").Select
                Range("A1").Value = S
                Call splitstr(":")
                Range("B1").Select
                Selection.Copy
                Sheets("表格整理").Select
                Range("E" & CStr(n)).Select
                ActiveSheet.Paste
                Sheets(sheetname).Select
            ElseIf InStr(1, S, "Target.RT:", vbTextCompare) > 0 Then
                Sheets("工作表2").Select
                Range("A1").Value = S
                Call splitstr(":")
                Range("B1").Select
                Selection.Copy
                Sheets("表格整理").Select
                Range("F" & CStr(n)).Select
                ActiveSheet.Paste
                Sheets(sheetname).Select
            End If
        Next i
    Else
        For i = 1 To m
            S = Range("A" & i).Value '& Range("B" & i).Value & Range("C" & i).Value & Range("D" & i).Value & Range("E" & i).Value
            If InStr(1, S, "*** LogFrame Start ***", vbTextCompare) > 0 Then
                n = n + 1
            ElseIf InStr(1, S, "*** LogFrame End ***", vbTextCompare) > 0 Then
                n = n + 0
            ElseIf InStr(1, S, "cond", vbTextCompare) > 0 Then
                'Sheets("工作表2").Select
                'Range("A1").Value = S
                'Call splitstr(":")
                Range("B" & i).Select
                Selection.Copy
                Sheets("表格整理").Select
                Range("A" & CStr(n)).Select
                ActiveSheet.Paste
                Sheets(sheetname).Select
            ElseIf InStr(1, S, "Face", vbTextCompare) > 0 Then
                If InStr(1, Left(Trim(S), 4), "Face", vbTextCompare) > 0 Then
                    Sheets(sheetname).Select
                    S = Range("B" & i).Value
                    Sheets("工作表2").Select
                    Range("A1").Value = S
                    'Call splitstr(":")
                    'Range("A1").Value = Range("B1").Value
                    Call splitstr("_")
                    
                    Range("A1").Select
                    Selection.Copy
                    Sheets("表格整理").Select
                    Range("B" & CStr(n)).Select
                    ActiveSheet.Paste
                    
                    Sheets("工作表2").Select
                    Range("B1").Select
                    Selection.Copy
                    Sheets("表格整理").Select
                    Range("C" & CStr(n)).Select
                    ActiveSheet.Paste
                    Sheets(sheetname).Select
                End If
            ElseIf InStr(1, S, "PicName", vbTextCompare) > 0 Then
                'Sheets("工作表2").Select
                'Range("A1").Value = S
                'Call splitstr(":")
                Range("B" & i).Value = Left(Trim(Range("B" & i).Value), 1)
                Range("B" & i).Select
                Selection.Copy
                Sheets("表格整理").Select
                Range("D" & CStr(n)).Select
                ActiveSheet.Paste
                Sheets(sheetname).Select
            ElseIf InStr(1, S, "Target.ACC", vbTextCompare) > 0 Then
                'Sheets("工作表2").Select
                'Range("A1").Value = S
                'Call splitstr(":")
                Range("B" & i).Select
                Selection.Copy
                Sheets("表格整理").Select
                Range("E" & CStr(n)).Select
                ActiveSheet.Paste
                Sheets(sheetname).Select
            ElseIf InStr(1, S, "Target.RT", vbTextCompare) > 0 Then
                'Sheets("工作表2").Select
                'Range("A1").Value = S
                'Call splitstr(":")
                Range("B" & i).Select
                Selection.Copy
                Sheets("表格整理").Select
                Range("F" & CStr(n)).Select
                ActiveSheet.Paste
                Sheets(sheetname).Select
            End If
        Next i
    End If
    
    Sheets("表格整理").Select
    Range("I3").Select
    ActiveCell.FormulaR1C1 = "快樂"
    Range("J3").Select
    ActiveCell.FormulaR1C1 = "悲傷"
    Range("K3").Select
    ActiveCell.FormulaR1C1 = "驚訝"
    Range("L3").Select
    ActiveCell.FormulaR1C1 = "中性"
    Range("H4").Select
    ActiveCell.FormulaR1C1 = "總trial數"
    Range("H5").Select
    ActiveCell.FormulaR1C1 = "ACC(正確的trial數)"
    Range("H6").Select
    ActiveCell.FormulaR1C1 = "平均RT(正確的trial數)"
    Columns("C:C").Select
    Selection.AutoFilter
    
    ActiveSheet.Range("$C$1:$C$" & ActiveSheet.UsedRange.Rows.Count).AutoFilter Field:=1, Criteria1:="H"
    trialsN = Application.WorksheetFunction.Subtotal(3, Range("C:C")) - 1 'Happy trials Number
    Selection.AutoFilter
    Range("I4").Select
    ActiveCell.FormulaR1C1 = trialsN
    
    ActiveSheet.Range("$C$1:$C$" & ActiveSheet.UsedRange.Rows.Count).AutoFilter Field:=1, Criteria1:="S"
    trialsN = Application.WorksheetFunction.Subtotal(3, Range("C:C")) - 1 'Sad trials Number
    Selection.AutoFilter
    Range("J4").Select
    ActiveCell.FormulaR1C1 = trialsN
    
    ActiveSheet.Range("$C$1:$C$" & ActiveSheet.UsedRange.Rows.Count).AutoFilter Field:=1, Criteria1:="SC"
    trialsN = Application.WorksheetFunction.Subtotal(3, Range("C:C")) - 1 'Scare trials Number
    Selection.AutoFilter
    Range("K4").Select
    ActiveCell.FormulaR1C1 = trialsN
     
    ActiveSheet.Range("$C$1:$C$" & ActiveSheet.UsedRange.Rows.Count).AutoFilter Field:=1, Criteria1:="N"
    trialsN = Application.WorksheetFunction.Subtotal(3, Range("C:C")) - 1 'Neutral trials Number
    Selection.AutoFilter
    Range("L4").Select
    ActiveCell.FormulaR1C1 = trialsN
    
    Rows("1:1").Select
    Selection.AutoFilter
    
    ActiveSheet.Range("$A$1:$F$" & ActiveSheet.UsedRange.Rows.Count).AutoFilter Field:=3, Criteria1:="H"
    ActiveSheet.Range("$A$1:$F$" & ActiveSheet.UsedRange.Rows.Count).AutoFilter Field:=5, Criteria1:="1"
    trialsN = Application.WorksheetFunction.Subtotal(3, Range("C:C")) - 1 'Happy trials Number
    Selection.AutoFilter
    Range("I5").Select
    ActiveCell.FormulaR1C1 = trialsN
    
    ActiveSheet.Range("$A$1:$F$" & ActiveSheet.UsedRange.Rows.Count).AutoFilter Field:=3, Criteria1:="S"
    ActiveSheet.Range("$A$1:$F$" & ActiveSheet.UsedRange.Rows.Count).AutoFilter Field:=5, Criteria1:="1"
    trialsN = Application.WorksheetFunction.Subtotal(3, Range("C:C")) - 1 'Sad trials Number
    Selection.AutoFilter
    Range("J5").Select
    ActiveCell.FormulaR1C1 = trialsN
    
    ActiveSheet.Range("$A$1:$F$" & ActiveSheet.UsedRange.Rows.Count).AutoFilter Field:=3, Criteria1:="SC"
    ActiveSheet.Range("$A$1:$F$" & ActiveSheet.UsedRange.Rows.Count).AutoFilter Field:=5, Criteria1:="1"
    trialsN = Application.WorksheetFunction.Subtotal(3, Range("C:C")) - 1 'Scare trials Number
    Selection.AutoFilter
    Range("K5").Select
    ActiveCell.FormulaR1C1 = trialsN
    
    ActiveSheet.Range("$A$1:$F$" & ActiveSheet.UsedRange.Rows.Count).AutoFilter Field:=3, Criteria1:="N"
    ActiveSheet.Range("$A$1:$F$" & ActiveSheet.UsedRange.Rows.Count).AutoFilter Field:=5, Criteria1:="1"
    trialsN = Application.WorksheetFunction.Subtotal(3, Range("C:C")) - 1 'Neutral trials Number
    Selection.AutoFilter
    Range("L5").Select
    ActiveCell.FormulaR1C1 = trialsN
    
    ActiveSheet.Range("$A$1:$F$" & ActiveSheet.UsedRange.Rows.Count).AutoFilter Field:=3, Criteria1:="H"
    ActiveSheet.Range("$A$1:$F$" & ActiveSheet.UsedRange.Rows.Count).AutoFilter Field:=5, Criteria1:="1"
    If Range("A1").End(xlDown).Row <> Rows.Count Then  '條件成立: 篩選後有資料
        'End(xlDown);往下到最後有資料的儲存格,Row:儲存格的列號
        'Rows.Count:工作表的總列數
        meanRT = Application.WorksheetFunction.Subtotal(1, Range("F:F"))  'Happy trials RT
    Else
        meanRT = 0
    End If
    Selection.AutoFilter
    Range("I6").Select
    ActiveCell.FormulaR1C1 = meanRT
    
    ActiveSheet.Range("$A$1:$F$" & ActiveSheet.UsedRange.Rows.Count).AutoFilter Field:=3, Criteria1:="S"
    ActiveSheet.Range("$A$1:$F$" & ActiveSheet.UsedRange.Rows.Count).AutoFilter Field:=5, Criteria1:="1"
    If Range("A1").End(xlDown).Row <> Rows.Count Then  '條件成立: 篩選後有資料
        'End(xlDown);往下到最後有資料的儲存格,Row:儲存格的列號
        'Rows.Count:工作表的總列數
        meanRT = Application.WorksheetFunction.Subtotal(1, Range("F:F"))  'Sad trials RT
    Else
        meanRT = 0
    End If
    Selection.AutoFilter
    Range("J6").Select
    ActiveCell.FormulaR1C1 = meanRT
    
    ActiveSheet.Range("$A$1:$F$" & ActiveSheet.UsedRange.Rows.Count).AutoFilter Field:=3, Criteria1:="SC"
    ActiveSheet.Range("$A$1:$F$" & ActiveSheet.UsedRange.Rows.Count).AutoFilter Field:=5, Criteria1:="1"
    If Range("A1").End(xlDown).Row <> Rows.Count Then  '條件成立: 篩選後有資料
        'End(xlDown);往下到最後有資料的儲存格,Row:儲存格的列號
        'Rows.Count:工作表的總列數
        meanRT = Application.WorksheetFunction.Subtotal(1, Range("F:F"))  'Scare trials RT
    Else
        meanRT = 0
    End If
    Selection.AutoFilter
    Range("K6").Select
    ActiveCell.FormulaR1C1 = meanRT
    
    ActiveSheet.Range("$A$1:$F$" & ActiveSheet.UsedRange.Rows.Count).AutoFilter Field:=3, Criteria1:="N"
    ActiveSheet.Range("$A$1:$F$" & ActiveSheet.UsedRange.Rows.Count).AutoFilter Field:=5, Criteria1:="1"
    If Range("A1").End(xlDown).Row <> Rows.Count Then  '條件成立: 篩選後有資料
        'End(xlDown);往下到最後有資料的儲存格,Row:儲存格的列號
        'Rows.Count:工作表的總列數
        meanRT = Application.WorksheetFunction.Subtotal(1, Range("F:F"))  'Neural trials RT
    Else
        meanRT = 0
    End If
    Selection.AutoFilter
    Range("L6").Select
    ActiveCell.FormulaR1C1 = meanRT
    
    Range("H3:L6").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ThemeColor = 6
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ThemeColor = 6
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ThemeColor = 6
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ThemeColor = 6
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ThemeColor = 6
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ThemeColor = 6
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Range("H3:L3").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent2
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("H3:L6").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ThemeColor = 6
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ThemeColor = 6
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Columns("H:H").EntireColumn.AutoFit
    Columns("I:I").EntireColumn.AutoFit
    Columns("J:J").EntireColumn.AutoFit
    Columns("K:K").EntireColumn.AutoFit
    Columns("L:L").EntireColumn.AutoFit
    
    Sheets("工作表2").Delete
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
End Sub

Sub splitstr(sParam1 As String)
'
' splitstr 巨集
'
'
    Range("A1").TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :=sParam1, FieldInfo:=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True

End Sub

Sub SheetReplace(mySheetName As String)
    'Application.DisplayAlerts = False
    On Error Resume Next
    Worksheets(mySheetName).Delete
    Err.Clear
    'Application.DisplayAlerts = True
    Worksheets.Add.Name = mySheetName
    'MsgBox "The sheet named ''" & mySheetName & "'' has been replaced."
End Sub


