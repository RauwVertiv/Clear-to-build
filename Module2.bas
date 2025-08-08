Attribute VB_Name = "Module2"
' ============================================
' MÓDULO 2: Lógica principal de resaltado
' ============================================

Public prevCell As Range
Public prevFormat As Object ' Objeto para guardar todo el formato
Public tempCell As Range ' Celda temporal para guardar el formato original

Public Sub ResaltarCeldaActiva(ByVal Target As Range)
    ' Evita errores con selecciones múltiples
    If Target.Cells.CountLarge > 1 Then Exit Sub
    
    ' Restaura el formato original de la celda previa
    If Not prevCell Is Nothing Then
        On Error Resume Next
        RestaurarFormatoOriginal prevCell
        On Error GoTo 0
    End If
    
    ' Guarda el formato original de la nueva celda activa
    GuardarFormatoOriginal Target
    
    ' Aplica el color de resaltado manteniendo el resto del formato
    Target.Interior.Color = RGB(255, 255, 153)  ' Amarillo claro
    
    ' Guarda referencia a la nueva celda activa
    Set prevCell = Target
End Sub

Private Sub GuardarFormatoOriginal(ByVal cell As Range)
    ' Crear una celda temporal invisible para guardar el formato exacto
    On Error Resume Next
    
    ' Buscar una celda vacía muy lejos para usar como temporal
    Dim ws As Worksheet
    Set ws = cell.Worksheet
    Set tempCell = ws.Cells(1048576, 16384) ' Última celda posible en Excel
    
    ' Limpiar la celda temporal
    tempCell.Clear
    
    ' Copiar el formato exacto de la celda original a la temporal
    cell.Copy
    tempCell.PasteSpecial xlPasteFormats
    
    ' Limpiar el portapapeles
    Application.CutCopyMode = False
    
    ' También guardar en Dictionary como respaldo
    Set prevFormat = CreateObject("Scripting.Dictionary")
    
    With cell
        ' Guardar propiedades del interior
        ' Verificar si tiene relleno (xlNone = -4142 significa sin relleno)
        If .Interior.ColorIndex = xlNone Then
            prevFormat.Add "InteriorColorIndex", xlNone
            prevFormat.Add "HasFill", False
        Else
            prevFormat.Add "InteriorColor", .Interior.Color
            prevFormat.Add "InteriorColorIndex", .Interior.ColorIndex
            prevFormat.Add "InteriorPattern", .Interior.Pattern
            prevFormat.Add "InteriorPatternColor", .Interior.PatternColor
            prevFormat.Add "InteriorPatternColorIndex", .Interior.PatternColorIndex
            prevFormat.Add "InteriorTintAndShade", .Interior.TintAndShade
            prevFormat.Add "HasFill", True
        End If
        
        ' Guardar propiedades de bordes solo si existen
        Dim hasBorders As Boolean
        hasBorders = (.Borders(xlEdgeLeft).LineStyle <> xlNone) Or _
                    (.Borders(xlEdgeTop).LineStyle <> xlNone) Or _
                    (.Borders(xlEdgeBottom).LineStyle <> xlNone) Or _
                    (.Borders(xlEdgeRight).LineStyle <> xlNone) Or _
                    (.Borders(xlInsideVertical).LineStyle <> xlNone) Or _
                    (.Borders(xlInsideHorizontal).LineStyle <> xlNone)
        
        prevFormat.Add "HasBorders", hasBorders
        
        If hasBorders Then
            ' Guardar todos los tipos de bordes
            prevFormat.Add "BordersLineStyle", Array(.Borders(xlEdgeLeft).LineStyle, _
                                                    .Borders(xlEdgeTop).LineStyle, _
                                                    .Borders(xlEdgeBottom).LineStyle, _
                                                    .Borders(xlEdgeRight).LineStyle, _
                                                    .Borders(xlInsideVertical).LineStyle, _
                                                    .Borders(xlInsideHorizontal).LineStyle)
            prevFormat.Add "BordersWeight", Array(.Borders(xlEdgeLeft).Weight, _
                                                .Borders(xlEdgeTop).Weight, _
                                                .Borders(xlEdgeBottom).Weight, _
                                                .Borders(xlEdgeRight).Weight, _
                                                .Borders(xlInsideVertical).Weight, _
                                                .Borders(xlInsideHorizontal).Weight)
            prevFormat.Add "BordersColor", Array(.Borders(xlEdgeLeft).Color, _
                                               .Borders(xlEdgeTop).Color, _
                                               .Borders(xlEdgeBottom).Color, _
                                               .Borders(xlEdgeRight).Color, _
                                               .Borders(xlInsideVertical).Color, _
                                               .Borders(xlInsideHorizontal).Color)
        End If
        
        ' Guardar propiedades de fuente
        prevFormat.Add "FontName", .Font.Name
        prevFormat.Add "FontSize", .Font.Size
        prevFormat.Add "FontBold", .Font.Bold
        prevFormat.Add "FontItalic", .Font.Italic
        prevFormat.Add "FontColor", .Font.Color
        prevFormat.Add "FontUnderline", .Font.Underline
        
        ' Guardar propiedades de alineación
        prevFormat.Add "HorizontalAlignment", .HorizontalAlignment
        prevFormat.Add "VerticalAlignment", .VerticalAlignment
        prevFormat.Add "WrapText", .WrapText
        prevFormat.Add "Orientation", .Orientation
        prevFormat.Add "IndentLevel", .IndentLevel
        
        ' Guardar formato de número
        prevFormat.Add "NumberFormat", .NumberFormat
    End With
    
    On Error GoTo 0
End Sub

Private Sub RestaurarFormatoOriginal(ByVal cell As Range)
    If tempCell Is Nothing And prevFormat Is Nothing Then Exit Sub
    
    On Error Resume Next
    
    ' Método 1: Usar la celda temporal (más preciso para colores)
    If Not tempCell Is Nothing Then
        ' Guardar el valor de la celda antes de restaurar el formato
        Dim cellValue As Variant
        cellValue = cell.Value
        
        ' Copiar el formato exacto desde la celda temporal
        tempCell.Copy
        cell.PasteSpecial xlPasteFormats
        
        ' Restaurar el valor original
        cell.Value = cellValue
        
        ' Limpiar portapapeles
        Application.CutCopyMode = False
        
        ' Limpiar la celda temporal
        tempCell.Clear
        Set tempCell = Nothing
    Else
        ' Método 2: Usar Dictionary como respaldo
        With cell
            ' Restaurar propiedades del interior
            If prevFormat("HasFill") Then
                ' La celda tenía relleno, restaurar en el orden correcto
                ' Primero restaurar el patrón y colores del patrón
                .Interior.Pattern = prevFormat("InteriorPattern")
                .Interior.PatternColor = prevFormat("InteriorPatternColor")
                .Interior.PatternColorIndex = prevFormat("InteriorPatternColorIndex")
                
                ' Luego restaurar el ColorIndex (esto es crucial para colores temáticos)
                .Interior.ColorIndex = prevFormat("InteriorColorIndex")
                
                ' Finalmente TintAndShade (esto puede cambiar la apariencia del color)
                .Interior.TintAndShade = prevFormat("InteriorTintAndShade")
                
                ' Solo restaurar Color si ColorIndex es xlColorIndexAutomatic o xlColorIndexNone
                ' De lo contrario, el ColorIndex ya define el color correctamente
                If prevFormat("InteriorColorIndex") = xlColorIndexAutomatic Or _
                   prevFormat("InteriorColorIndex") = xlColorIndexNone Then
                    .Interior.Color = prevFormat("InteriorColor")
                End If
            Else
                ' La celda no tenía relleno, restaurar a "No Fill"
                .Interior.ColorIndex = xlNone
            End If
            
            ' Restaurar bordes solo si los tenía
            If prevFormat("HasBorders") Then
                Dim bordersLineStyle As Variant
                Dim bordersWeight As Variant
                Dim bordersColor As Variant
                
                bordersLineStyle = prevFormat("BordersLineStyle")
                bordersWeight = prevFormat("BordersWeight")
                bordersColor = prevFormat("BordersColor")
                
                ' Restaurar todos los tipos de bordes
                .Borders(xlEdgeLeft).LineStyle = bordersLineStyle(0)
                .Borders(xlEdgeTop).LineStyle = bordersLineStyle(1)
                .Borders(xlEdgeBottom).LineStyle = bordersLineStyle(2)
                .Borders(xlEdgeRight).LineStyle = bordersLineStyle(3)
                .Borders(xlInsideVertical).LineStyle = bordersLineStyle(4)
                .Borders(xlInsideHorizontal).LineStyle = bordersLineStyle(5)
                
                ' Solo aplicar weight y color si el borde existe
                If bordersLineStyle(0) <> xlNone Then
                    .Borders(xlEdgeLeft).Weight = bordersWeight(0)
                    .Borders(xlEdgeLeft).Color = bordersColor(0)
                End If
                If bordersLineStyle(1) <> xlNone Then
                    .Borders(xlEdgeTop).Weight = bordersWeight(1)
                    .Borders(xlEdgeTop).Color = bordersColor(1)
                End If
                If bordersLineStyle(2) <> xlNone Then
                    .Borders(xlEdgeBottom).Weight = bordersWeight(2)
                    .Borders(xlEdgeBottom).Color = bordersColor(2)
                End If
                If bordersLineStyle(3) <> xlNone Then
                    .Borders(xlEdgeRight).Weight = bordersWeight(3)
                    .Borders(xlEdgeRight).Color = bordersColor(3)
                End If
                If bordersLineStyle(4) <> xlNone Then
                    .Borders(xlInsideVertical).Weight = bordersWeight(4)
                    .Borders(xlInsideVertical).Color = bordersColor(4)
                End If
                If bordersLineStyle(5) <> xlNone Then
                    .Borders(xlInsideHorizontal).Weight = bordersWeight(5)
                    .Borders(xlInsideHorizontal).Color = bordersColor(5)
                End If
            Else
                ' La celda no tenía bordes, eliminar cualquier borde
                .Borders(xlEdgeLeft).LineStyle = xlNone
                .Borders(xlEdgeTop).LineStyle = xlNone
                .Borders(xlEdgeBottom).LineStyle = xlNone
                .Borders(xlEdgeRight).LineStyle = xlNone
                .Borders(xlInsideVertical).LineStyle = xlNone
                .Borders(xlInsideHorizontal).LineStyle = xlNone
            End If
            
            ' Restaurar propiedades de fuente
            .Font.Name = prevFormat("FontName")
            .Font.Size = prevFormat("FontSize")
            .Font.Bold = prevFormat("FontBold")
            .Font.Italic = prevFormat("FontItalic")
            .Font.Color = prevFormat("FontColor")
            .Font.Underline = prevFormat("FontUnderline")
            
            ' Restaurar propiedades de alineación
            .HorizontalAlignment = prevFormat("HorizontalAlignment")
            .VerticalAlignment = prevFormat("VerticalAlignment")
            .WrapText = prevFormat("WrapText")
            .Orientation = prevFormat("Orientation")
            .IndentLevel = prevFormat("IndentLevel")
            
            ' Restaurar formato de número
            .NumberFormat = prevFormat("NumberFormat")
        End With
    End If
    
    On Error GoTo 0
    
    ' Limpiar referencias
    Set prevFormat = Nothing
End Sub
