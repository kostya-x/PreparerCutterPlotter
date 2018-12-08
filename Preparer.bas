Sub Preparer()
    ' Recorded Mn, 22.10.2018
    '
    ' Description:
    '     Prepare file for cutter plotter
    MsgBox "IT works!"

    '---------------------------------------------------------------------------
    'Move objects 0,3mm to the left
    ActiveSelection.Move -0.011811, 0#
    'Move objects 0,5mm to the top
    ActiveSelection.Move 0#, 0.019685
End Sub
