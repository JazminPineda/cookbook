' Este ejm es para actualizar las conexiones e insertar fechas en la celda de la hoja "Resumen" C4
Sub m_ActAll()
    Worksheets("Resumen").Activate
    Range("G3").Select
    ActiveWorkbook.RefreshAll
    Range("C4").Select
    ActiveCell.FormulaR1C1 = "=TODAY()-1"
    Range("C4").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Range("K3:P3").Select
End Sub