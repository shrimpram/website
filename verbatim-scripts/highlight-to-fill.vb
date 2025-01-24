Sub ConvertHighlightsToFills()
    Dim rng As Range
    Dim wordRange As Range
    Dim highlightColor As Long

    ' Exit if there is no selection or if the selection is empty
    If Selection Is Nothing Or Selection.Range.Text = "" Then
        MsgBox "Please select text to apply the fill change.", vbExclamation
        Exit Sub
    End If

    ' Speed up the macro by disabling screen updates and other processes
    Application.ScreenUpdating = False

    ' Set the range to the current selection
    Set rng = Selection.Range

    ' Loop through each word (or character range) within the selection
    For Each wordRange In rng.Words
        If wordRange.HighlightColorIndex <> wdNoHighlight Then
            ' Map the HighlightColorIndex to an RGB color value
            highlightColor = MapHighlightToRGB(wordRange.HighlightColorIndex)

            ' Apply the highlight color as fill
            wordRange.Shading.BackgroundPatternColor = highlightColor

            ' Remove the highlight
            wordRange.HighlightColorIndex = wdNoHighlight
        End If
    Next wordRange

    ' Restore screen updates
    Application.ScreenUpdating = True
End Sub

Function MapHighlightToRGB(highlightIndex As WdColorIndex) As Long
    Select Case highlightIndex
        Case wdYellow: MapHighlightToRGB = RGB(255, 255, 0)
        Case wdBrightGreen: MapHighlightToRGB = RGB(0, 255, 0)
        Case wdTurquoise: MapHighlightToRGB = RGB(0, 255, 255)
        Case wdPink: MapHighlightToRGB = RGB(255, 0, 255)
        Case wdBlue: MapHighlightToRGB = RGB(0, 0, 255)
        Case wdRed: MapHighlightToRGB = RGB(255, 0, 0)
        Case wdDarkBlue: MapHighlightToRGB = RGB(0, 0, 139)
        Case wdTeal: MapHighlightToRGB = RGB(0, 128, 128)
        Case wdGreen: MapHighlightToRGB = RGB(0, 128, 0)
        Case wdViolet: MapHighlightToRGB = RGB(238, 130, 238)
        Case wdDarkRed: MapHighlightToRGB = RGB(139, 0, 0)
        Case wdDarkYellow: MapHighlightToRGB = RGB(128, 128, 0)
        Case wdGray25: MapHighlightToRGB = RGB(210, 210, 210)
        Case wdGray50: MapHighlightToRGB = RGB(128, 128, 128)
        Case Else: MapHighlightToRGB = RGB(255, 255, 255) ' Default to white for unknown colors
    End Select
End Function
