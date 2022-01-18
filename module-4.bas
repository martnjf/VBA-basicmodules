Attribute VB_Name = "Módulo3"
Sub PegarFiltro()
    'Indica una celda de la base filtrada:
    UnaCelda = "B4"
    
    'Copiado de celdas visibles
    Range(UnaCelda).CurrentRegion.SpecialCells(xlCellTypeVisible).Copy
    
    'Ir a celda de destino
    Sheets.Add Before:=Sheets(1)
    Worksheets(1).Select
    Range("B1").Select
    
    'pegado de valores
    'Sheets.Add(After:=Sheets(Sheets.Count)).Name = x.Value
    Selection.PasteSpecial Paste:=xlValues
    
    'Pegado de formatos
    Selection.PasteSpecial Paste:=xlFormats
End Sub
