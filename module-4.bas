Attribute VB_Name = "Módulo4"
Sub PegarFiltro()
   '***************************
    'This current file copies only the visible cells from a Sheet, creating a new one and pasting it on it.
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
