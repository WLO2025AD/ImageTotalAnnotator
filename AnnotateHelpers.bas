Attribute VB_Name = "AnnotateHelpers"
Option Explicit

Sub AnnotateSelectedImage()
    On Error GoTo EH
    Dim shp As Shape
    Dim rng As Range

    If Selection.InlineShapes.Count > 0 Then
        Selection.InlineShapes(1).ConvertToShape
    End If

    If Selection.ShapeRange.Count > 0 Then
        Set shp = Selection.ShapeRange(1)
        Set rng = shp.Anchor.Duplicate
    Else
        Set rng = Selection.Range
    End If

    Dim cvs As Shape
    Set cvs = ActiveDocument.Shapes.AddCanvas(rng)
    cvs.Line.Visible = msoFalse
    cvs.Fill.Visible = msoFalse
    cvs.WrapFormat.Type = wdWrapTight

    Dim ann As Shape
    Set ann = cvs.CanvasItems.AddShape(Type:=msoShapeRectangle, _
        Left:=10, Top:=10, Width:=120, Height:=60)
    With ann
        .Fill.Visible = msoFalse
        .Line.ForeColor.RGB = RGB(220, 20, 60)
        .Line.Weight = 2.25
    End With

    Dim tb As Shape
    Set tb = cvs.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=10, Top:=80, Width:=220, Height:=28)
    With tb.TextFrame2
        .TextRange.Text = "주석 입력 (더블클릭하여 편집)"
        .TextRange.Font.Size = 11
    End With

    MsgBox "Annotate canvas inserted. Use Insert > Shapes to add arrows, lines, etc.", vbInformation
    Exit Sub
EH:
    MsgBox "Failed to insert annotation canvas: " & Err.Description, vbExclamation
End Sub
