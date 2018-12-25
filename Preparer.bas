Sub Preparer()
    ' Recorded Mn, 22.10.2018
    '
    ' Description:
    '     Prepare file for cutter plotter

    '---------------------------------------------------------------------------
    'Correct color for marks
    Dim trueMakrsColor As New Color
    trueMakrsColor.RGBAssign 30, 24, 21

    'Set correct color for marks
    Dim marks As Shape
    Dim newMarks as New ShapeRange
    For Each marks In ActivePage.Shapes.FindShapes()
      If Not marks.Fill.Type = cdrNoOutline Then
        newMarks.Add marks
      End If
    Next marks
    newMarks.CreateSelection
    newMarks.SetOutlineProperties cdrNoOutline
    newMarks.ApplyUniformFill trueMakrsColor

    'Delete all guidelines
    Dim gl As New ShapeRange
    Dim sgl As Shape
    For Each sgl In ActivePage.FindShapes(Type:=cdrGuidelineShape)
    gl.Add sgl
    Next sgl
    gl.Delete

    Dim s As Shape
    'Pick every object in active page and put it in "array" sr...
    '...if the object does not have a fill
    Dim sr As new ShapeRange
    For Each s In ActivePage.Shapes.FindShapes()
      If s.Fill.Type = cdrNoFill Then
        sr.Add s
      End If
    Next s

    sr.CreateSelection 'make selection of every object in sr

    sr.SetOutlineProperties 0.003 'Change abris thickness to hairline

    sr.Group 'Make a group of objects

    'Move objects 0,3mm to the left
    ActiveSelection.Move -0.011811, 0#
    'Move objects 0,5mm to the top
    ActiveSelection.Move 0#, 0.019685
End Sub
