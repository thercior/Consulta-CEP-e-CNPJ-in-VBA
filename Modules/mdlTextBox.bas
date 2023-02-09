Attribute VB_Name = "mdlTextBox"
Public FontSize As Integer
Public FontName As String
Public fColor, eColor, tColor, bColor

Public Sub TxtColor(fColorValue As String, eColorValue As String, tColorValue As String, Optional bColorValue As String)
    fColor = fColorValue 'TextBox ForeColor
    eColor = eColorValue 'When TextBox Enter
    tColor = tColorValue 'Title and bottom line Color
    bColor = bColorValue 'Background Color
End Sub
