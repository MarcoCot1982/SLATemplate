'***********************************
'** Author: Marco Cot DAS:A669714 **
'***********************************
'*
'* ACCOUNT: DENTSPLY
'* Importing data from PBI and returning SLA
'* with manual check false attempt and ASA threshold
'*
Sub Dentsply(Optional HideMe As String)
'
' Dentsply Macro
'

'
    Application.ScreenUpdating = False
    Sheets("DATA").Select
    Range("A1").Select
'CONVERSION TALKING TIME
    Columns("L:L").Select
    Application.CutCopyMode = False
    Selection.TextToColumns Destination:=Range("L1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=True, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
'CONVERSION WAITING TIME
    Columns("I:I").Select
    Application.CutCopyMode = False
    Selection.TextToColumns Destination:=Range("I1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=True, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
'REPLACE LANGUAGE
    Range("T:T").Select
    Selection.Replace What:="AT", Replacement:="Austria", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="DE", Replacement:="German", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="EN", Replacement:="English", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="ES", Replacement:="Spanish", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="IT", Replacement:="Italian", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="PT", Replacement:="Portuguese", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="FR", Replacement:="French", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
'CHECK ASA SLA
    Columns("V:V").Select
    Selection.Copy
    Columns("Y:Z").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("Y1").Select
    ActiveCell.FormulaR1C1 = "ASA%"
    Range("Y2").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-14]<TABLE!R2C[-14],""OK"",""OUT SLA"")"
    Selection.Copy
    Range(Selection, Selection.End(xlDown)).Select
    ActiveSheet.Paste
'CHECK WAITING vs FALSE ATTEMPT
    Range("Z1").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "WAITING FA"
    Range("Z2").Select
    ActiveCell.FormulaR1C1 = "=IF(DATA!RC9<TABLE!R2C9,""WITHIN"",""OUT"")"
    Range("Z2").Select
    Selection.Copy
    Range(Selection, Selection.End(xlDown)).Select
    ActiveSheet.Paste
    Columns("Y:Z").Select
    Application.CutCopyMode = False
    Selection.Copy
    Columns("W:X").Select
    ActiveSheet.Paste
    Columns("Y:Z").Select
    Selection.ClearContents
'END
    Sheets("TABLE").Select
    Range("B15").Select
    Application.ScreenUpdating = True
    MsgBox ("Process has been completed")
End Sub