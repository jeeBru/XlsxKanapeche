Sub reset()


Dim LastRow As Long 'last row in the worksheet
    
'LastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row 'JB : la colonne B contenant les identifiants est utilisée pour juger de la fin du tableau car elle est en théorie 100% remplie
    
Worksheets("Sources").Range(Cells(2, 2), Cells(100, 5)).ClearContents
Worksheets("Compilation").Range(Cells(3, 1), Cells(300, 23)).ClearContents
Worksheets("Debug").ClearContents

End Sub
