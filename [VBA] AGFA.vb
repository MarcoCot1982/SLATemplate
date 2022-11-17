'***********************************
'** Author: Marco Cot DAS:A669714 **
'***********************************
'*
'* Importing data from PBI and returning SLA
'* AGFA
'*
'*
Sub AGFA()
'
' AGFA Macro
'

'
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
    Selection.Replace What:="FR", Replacement:="French", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Range("T:T").Select
    Selection.Replace What:="DE", Replacement:="German", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Range("T:T").Select
    Selection.Replace What:="NL", Replacement:="Dutch", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Range("T:T").Select
    Selection.Replace What:="ES", Replacement:="Spanish", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Range("T:T").Select
    Selection.Replace What:="IT", Replacement:="Italian", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Range("T:T").Select
    Selection.Replace What:="PT-BR", Replacement:="Portuguese", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
'CHECK ASA SLA
    Columns("V:V").Select
    Selection.Copy
    Columns("Y:Y").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("Y12").Select
    Selection.End(xlDown).Select
    Selection.End(xlUp).Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "ASA%"
    Range("Y2").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-16]<TABLE!R1C[-14],""OK"",""OUT SLA"")"
    Range("Y2").Select
    Selection.Copy
    Range(Selection, Selection.End(xlDown)).Select
    ActiveSheet.Paste
    Columns("Y:Y").Select
    Application.CutCopyMode = False
    Selection.Cut
    Columns("W:W").Select
    ActiveSheet.Paste
    Sheets("TABLE").Select
    Range("B14").Select    
    MsgBox ("Process has been completed")
End Sub