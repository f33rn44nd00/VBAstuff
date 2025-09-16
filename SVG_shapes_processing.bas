Option Explicit 

Sub GetShapeNames()
'
Dim shp As Shape 
Dim i As Long 

 i = 1 
 For Each shp In ActiveSheet.Shapes 
 ActiveSheet.Range("M1").Offset(i, 0).Value = _
 ActiveSheet.Shapes(i).Name 
 i = i + 1 
 Next shp 

End Sub 

Sub SetShapeNames() 
Dim shp As Shape 
Dim i As Long 

 i = 1 
 For Each shp In ActiveSheet.Shapes 
 ActiveSheet.Shapes(i).Name = _ 
 ActiveSheet.Range("N1").Offset(i, 0).Value 
 i = i + 1 
 Next shp 

End Sub

