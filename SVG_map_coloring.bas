Attribute VB_Name = "Module2"


Sub SelectFunctionColumn()
' Deprecated: chooses macro based on the existence of conditional formatting in the selected cell
' Replaced by "Worksheet change" Sub in "Mapas" sheet
Dim ColNm As String

ColNm = ActiveCell.Offset(2 - ActiveCell.Row).Value
If ColNm = "" Then
    MsgBox ("No se encontró la columna!")
    Exit Sub
ElseIf ActiveCell.FormatConditions.Count <> 0 Then
    Color_Format (ColNm)
Else
    Color_New (ColNm)
End If
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
'Hacky method: identifies cursor placement in worksheet, supposedly triggers on update of the dropdown in B3
Dim ColNm As String
    If Target.Address = "$B$3" Then
        ColNm = Range("$B$3").Value
        Color_New (ColNm)
        Color_Format (ColNm)
    End If
End Sub

Sub Color_Format(ColNm)
' Just change the name of the column in Set ColorColumn
' Does not recognize colors from conditional formatting. Colors must be put manually.

    Dim ws As Worksheet
    Dim ws2 As Worksheet
    Dim tbl As ListObject
    Dim stateColumn As Range
    Dim colorColumn As Range
    Dim cell As Range
    Dim shape As shape
    Dim colorIndex As Long
    
    ' Set the worksheet containing the map
    Set ws = ThisWorkbook.Sheets("Mapas")
    
    ' Set the worksheet containing the table
    Set ws2 = ThisWorkbook.Sheets("Tablero Municipios")
    
    ' Set the Excel table
    Set tbl = ws2.ListObjects("T_Municipios")
    
    ' Set the range for the municipio column and color column
    Set municipiosColumn = tbl.ListColumns("Municipios").DataBodyRange
    Set colorColumn = tbl.ListColumns(ColNm).DataBodyRange
    
    ' Loop through each municipio in the column
    For Each cell In municipiosColumn
        On Error Resume Next
        ' Get the shape corresponding to the municipio name
        nameshape = cell.Value
        Set shape = ws.Shapes(nameshape)
        If Err.Number = 0 Then
        
        ' Check if the shape exists
        If Not shape Is Nothing Then
            ' Get the color of the corresponding cell in the color column
            colorIndex = xlNone ' Default color if not found
            
            ' Find the corresponding municipio in the color column and get its color
            Dim municipiosIndex As Long
            municipiosIndex = cell.Row - municipiosColumn.Row + 1 ' Adjust index for data body range
            If municipiosIndex >= 1 And municipiosIndex <= colorColumn.Rows.Count Then
                If Not IsEmpty(colorColumn.Cells(municipiosIndex, 1)) Then
                    colorIndex = colorColumn.Cells(municipiosIndex, 1).DisplayFormat.Interior.Color
                End If
            End If
            
            ' Set the color of the shape based on the color index
            If colorIndex <> xlNone Then
                shape.Fill.ForeColor.RGB = colorColumn.Cells(municipiosIndex, 1).DisplayFormat.Interior.Color
            Else
                ' Handle case when color is not found
                ' You can set a default color or take other actions
                'MsgBox "Name: " & nameshape
                shape.Fill.ForeColor.RGB = RGB(255, 255, 255)
            End If
        Else
            ' Handle case when shape is not found
            ' You can display a message or take other actions
            MsgBox "No shape"
        End If
        End If
    Next cell
    ws.Shapes("TextboxMap").TextFrame.Characters.Text = ColNm
End Sub


Sub Color_New(ColNm)
' Just change the name of the column in Set ColorColumn
' Does not recognize colors from conditional formatting. Colors must be put manually.

    Dim ws As Worksheet
    Dim ws2 As Worksheet
    Dim tbl As ListObject
    Dim stateColumn As Range
    Dim colorColumn As Range
    Dim cell As Range
    Dim shape As shape
    Dim colorIndex As Long
    
    ' Set the worksheet containing the map
    Set ws = ThisWorkbook.Sheets("Mapas")
    
    ' Set the worksheet containing the table
    Set ws2 = ThisWorkbook.Sheets("Tablero Municipios")
    
    ' Set the Excel table
    Set tbl = ws2.ListObjects("T_Municipios")
    
    ' Set the range for the state column and color column
    Set municipiosColumn = tbl.ListColumns("Municipios").DataBodyRange
    Set colorColumn = tbl.ListColumns(ColNm).DataBodyRange
    
    ' Loop through each state in the state column
    For Each cell In municipiosColumn
        On Error Resume Next
        ' Get the shape corresponding to the municipios
        nameshape = cell.Value
        Set shape = ws.Shapes(nameshape)
        If Err.Number = 0 Then
        
        ' Check if the shape exists
        If Not shape Is Nothing Then
            ' Get the color of the corresponding cell in the color column
            colorIndex = xlNone ' Default color if not found
            
            ' Find the corresponding municipios in the color column and get its color
            Dim municipiosIndex As Long
            municipiosIndex = cell.Row - municipiosColumn.Row + 1 ' Adjust index for data body range
            If municipiosIndex >= 1 And municipiosIndex <= colorColumn.Rows.Count Then
                If Not IsEmpty(colorColumn.Cells(municipiosIndex, 1)) Then
                    colorIndex = colorColumn.Cells(municipiosIndex, 1).Interior.colorIndex
                End If
            End If
            
            ' Set the color of the shape based on the color index
            If colorIndex <> xlNone Then
                shape.Fill.ForeColor.RGB = colorColumn.Cells(municipiosIndex, 1).Interior.Color
            Else
                ' Handle case when color is not found
                ' You can set a default color or take other actions
                'MsgBox "Name: " & nameshape
                shape.Fill.ForeColor.RGB = RGB(255, 255, 255)
            End If
        Else
            ' Handle case when shape is not found
            ' You can display a message or take other actions
            MsgBox "No shape"
        End If
        End If
    Next cell
    ws.Shapes("TextboxMap").TextFrame.Characters.Text = ColNm
End Sub
