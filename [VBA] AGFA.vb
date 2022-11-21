'***********************************
'** Author: Marco Cot DAS:A669714 **
'***********************************
'
' ACCOUNT: AGFA
' Importing data from PBI and returning SLA
' per language and per agent
'
'
Sub AGFA(Optional HideMe As String)
'
' AGFA Macro
'

'
    Application.ScreenUpdating = False
    ActiveSheet.Unprotect "NeverEdit"
    'CHECK DATA AVAIL
    Sheets("DATA").Select
    Range("S:W").Copy
    Range("R1").Select
    ActiveSheet.Paste
    Range("W:W").ClearContents
        If Range("A2") = "" Then
            Sheets("TABLE").Select
            MsgBox ("Data not pasted")
        Else
        'AVOID OVERWRITING
        Sheets("DATA").Select
        If Range("Q1") <> 1998 Then
            Sheets("DATA CHAT").Select
            If Range("O1") <> 1998 Then
                Range("A1").Select
                'CONVERSION TALKING TIME
                Sheets("DATA").Select
                Columns("L:L").Select
                Application.CutCopyMode = False
               Selection.TextToColumns Destination:=Range("L1"), DataType:=xlDelimited, _
                TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
                Semicolon:=False, Comma:=True, Space:=False, Other:=False, FieldInfo _
                :=Array(1, 1), TrailingMinusNumbers:=True
                Sheets("DATA CHAT").Select
                Columns("J:J").Select
                Application.CutCopyMode = False
                Selection.TextToColumns Destination:=Range("J1"), DataType:=xlDelimited, _
                TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
                Semicolon:=False, Comma:=True, Space:=False, Other:=False, FieldInfo _
                :=Array(1, 1), TrailingMinusNumbers:=True
                'CONVERSION WAITING TIME
                Sheets("DATA").Select
                Columns("I:I").Select
                Application.CutCopyMode = False
                Selection.TextToColumns Destination:=Range("I1"), DataType:=xlDelimited, _
                TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
                Semicolon:=False, Comma:=True, Space:=False, Other:=False, FieldInfo _
                :=Array(1, 1), TrailingMinusNumbers:=True
                Sheets("DATA CHAT").Select
                Columns("H:H").Select
                Application.CutCopyMode = False
                Selection.TextToColumns Destination:=Range("H1"), DataType:=xlDelimited, _
                TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
                Semicolon:=False, Comma:=True, Space:=False, Other:=False, FieldInfo _
                :=Array(1, 1), TrailingMinusNumbers:=True
                'REPLACE LANGUAGE
                Sheets("DATA").Select
                Range("T:T").Select
                Selection.Replace What:="FR", Replacement:="French", LookAt:=xlPart, _
                SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
                Selection.Replace What:="DE", Replacement:="German", LookAt:=xlPart, _
                SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
                Selection.Replace What:="NL", Replacement:="Dutch", LookAt:=xlPart, _
                SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
                Selection.Replace What:="ES", Replacement:="Spanish", LookAt:=xlPart, _
                SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
                Selection.Replace What:="IT", Replacement:="Italian", LookAt:=xlPart, _
                SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
                Selection.Replace What:="PT-BR", Replacement:="Portuguese", LookAt:=xlPart, _
                SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
                Sheets("DATA CHAT").Select
                Range("Q:Q").Select
                Selection.Replace What:="FR", Replacement:="French", LookAt:=xlPart, _
                SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
                Selection.Replace What:="DE", Replacement:="German", LookAt:=xlPart, _
                SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
                Selection.Replace What:="NL", Replacement:="Dutch", LookAt:=xlPart, _
                SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
                Selection.Replace What:="ES", Replacement:="Spanish", LookAt:=xlPart, _
                SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
                Selection.Replace What:="IT", Replacement:="Italian", LookAt:=xlPart, _
                SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
                Selection.Replace What:="PT-BR", Replacement:="Portuguese", LookAt:=xlPart, _
                SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
                'CHECK ASA SLA
                Sheets("DATA").Select
                Range("Q1").Select
                ActiveCell.FormulaR1C1 = "1998"
                Columns("V:V").Select
                Selection.Copy
                Columns("Y:Y").Select
                Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
                Range("Y1").Select
                ActiveCell.FormulaR1C1 = "ASA%"
                Range("Y2").Select
                ActiveCell.FormulaR1C1 = "=IF(RC[-14]<TABLE!R1C[-12],""OK"",""OUT SLA"")"
                Range("Y2").Select
                Selection.Copy
                Range(Selection, Selection.End(xlDown)).Select
                ActiveSheet.Paste
                Columns("Y:Y").Select
                Application.CutCopyMode = False
                Selection.Copy
                Columns("W:W").Select
                ActiveSheet.Paste
                Columns("Y:Y").Select
                Selection.ClearContents
                Sheets("By Agent").Select
                ThisWorkbook.RefreshAll
                Range("B:B").NumberFormat = "h:mm:ss;@"
                Sheets("DATA CHAT").Select
                Range("O1").Select
                ActiveCell.FormulaR1C1 = "1998"
                Columns("P:P").Select
                Selection.Copy
                Columns("Y:Y").Select
                Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
                Range("Y1").Select
                ActiveCell.FormulaR1C1 = "ASA%"
                Range("Y2").Select
                ActiveCell.FormulaR1C1 = "=IF(RC[-11]<TABLE!R1C[-7],""OK"",""OUT SLA"")"
                Range("Y2").Select
                Selection.Copy
                Range(Selection, Selection.End(xlDown)).Select
                ActiveSheet.Paste
                Columns("Y:Y").Select
                Application.CutCopyMode = False
                Selection.Copy
                Columns("S:S").Select
                ActiveSheet.Paste
                Columns("Y:Y").Select
                Selection.ClearContents
                Sheets("TABLE").Select
                MsgBox ("Process has been completed")
            Else
            Sheets("TABLE").Select
            MsgBox ("Data already processed")
            End If
        Else
        Sheets("TABLE").Select
        MsgBox ("Data already processed")
        End If
    End If
        ActiveSheet.Protect "NeverEdit"
    End Sub
