Attribute VB_Name = "Módulo3"
Sub PegarFiltro()
    'Indica una celda de la base filtrada:
    UnaCelda = "B4"
    
    'Copiado de celdas visibles
    Range(UnaCelda).CurrentRegion.SpecialCells(xlCellTypeVisible).Copy
    
    'Ir a celda de destino
    Sheets("Filtradov2").Activate
    Range("B4").Select
    
    'pegado de valores
    Selection.PasteSpecial Paste:=xlValues
    
    'Pegado de formatos
    Selection.PasteSpecial Paste:=xlFormats
End Sub
