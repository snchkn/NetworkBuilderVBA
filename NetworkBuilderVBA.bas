Attribute VB_Name = "Module1"
Sub DrawNetworkDiagram()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Задание 4")
    
    ' Clear previous shapes
    ws.Shapes.SelectAll
    On Error Resume Next
    Selection.Delete
    On Error GoTo 0
    
    ' Define node properties
    Dim nodeWidth As Double, nodeHeight As Double
    nodeWidth = 50
    nodeHeight = 30
    Dim startX As Double, startY As Double
    startX = 100
    startY = 300  ' Расположение под таблицей
    Dim nodeSpacingX As Double, nodeSpacingY As Double
    nodeSpacingX = 100
    nodeSpacingY = 80
    
    ' Node positions (arranged in a custom pattern)
    Dim nodePositions As Variant
    nodePositions = Array( _
        Array(startX, startY), _
        Array(startX, startY + nodeSpacingY), _
        Array(startX, startY + 2 * nodeSpacingY), _
        Array(startX + nodeSpacingX, startY + 3 * nodeSpacingY), _
        Array(startX + nodeSpacingX, startY), _
        Array(startX + 2 * nodeSpacingX, startY + 2 * nodeSpacingY), _
        Array(startX + 2 * nodeSpacingX, startY), _
        Array(startX + 3 * nodeSpacingX, startY + nodeSpacingY), _
        Array(startX + 3 * nodeSpacingX, startY + 3 * nodeSpacingY), _
        Array(startX + 4 * nodeSpacingX, startY), _
        Array(startX + 4 * nodeSpacingX, startY + 2 * nodeSpacingY), _
        Array(startX + 5 * nodeSpacingX, startY + nodeSpacingY), _
        Array(startX + 5 * nodeSpacingX, startY + 3 * nodeSpacingY), _
        Array(startX + 6 * nodeSpacingX, startY + 2 * nodeSpacingY), _
        Array(startX + 6 * nodeSpacingX, startY + 4 * nodeSpacingY), _
        Array(startX + 7 * nodeSpacingX, startY + 3 * nodeSpacingY) _
    )
    
    ' Draw nodes
    Dim i As Integer
    For i = LBound(nodePositions) To UBound(nodePositions)
        Dim node As shape
        Set node = ws.Shapes.AddShape(msoShapeOval, nodePositions(i)(0) - (nodeWidth / 2), nodePositions(i)(1) - (nodeHeight / 2), nodeWidth, nodeHeight)
        node.Fill.ForeColor.RGB = RGB(218, 112, 214) ' Сиреневый цвет
        node.Line.Weight = 1.5
        node.Name = "a" & (i + 1)
        node.TextFrame.Characters.Text = "a" & (i + 1)
        node.TextFrame.HorizontalAlignment = xlHAlignCenter
        node.TextFrame.VerticalAlignment = xlVAlignCenter
        node.TextFrame.Characters.Font.Size = 8
        node.TextFrame.Characters.Font.Color = RGB(0, 0, 0) ' Черный цвет текста
    Next i
    
    ' Draw connections
    Dim connections As Variant
    connections = Array( _
        Array(1, 4), Array(1, 5), Array(1, 6), _
        Array(2, 7), Array(2, 8), _
        Array(3, 9), Array(3, 10), Array(3, 11), _
        Array(6, 12), Array(6, 13), _
        Array(10, 14), Array(10, 15), _
        Array(11, 16), Array(13, 16), Array(14, 16) _
    )
    
    Dim startNode As shape, endNode As shape
    Dim sX As Double, sY As Double, eX As Double, eY As Double
    Dim deltaX As Double, deltaY As Double, distance As Double
    Dim unitX As Double, unitY As Double

    For i = LBound(connections) To UBound(connections)
        Set startNode = ws.Shapes("a" & connections(i)(0))
        Set endNode = ws.Shapes("a" & connections(i)(1))
        
        ' Calculate center coordinates for nodes
        sX = startNode.Left + (startNode.Width / 2)
        sY = startNode.Top + (startNode.Height / 2)
        eX = endNode.Left + (endNode.Width / 2)
        eY = endNode.Top + (endNode.Height / 2)
        
        ' Calculate direction from start to end
        deltaX = eX - sX
        deltaY = eY - sY
        distance = Sqr(deltaX ^ 2 + deltaY ^ 2)
        unitX = deltaX / distance
        unitY = deltaY / distance
        
        ' Adjust start and end to ensure arrowheads touch the oval edges
        sX = sX + unitX * (nodeWidth / 2)
        sY = sY + unitY * (nodeHeight / 2)
        eX = eX - unitX * (nodeWidth / 2)
        eY = eY - unitY * (nodeHeight / 2)
        
        ' Draw the connection
        Dim conn As shape
        Set conn = ws.Shapes.AddConnector(msoConnectorStraight, sX, sY, eX, eY)
        conn.Line.Weight = 2
        conn.Line.ForeColor.RGB = Choose(((i Mod 5) + 1), RGB(255, 0, 0), RGB(0, 128, 0), RGB(0, 0, 255), RGB(255, 165, 0), RGB(128, 0, 128))
        conn.Line.EndArrowheadStyle = msoArrowheadTriangle
    Next i
End Sub


