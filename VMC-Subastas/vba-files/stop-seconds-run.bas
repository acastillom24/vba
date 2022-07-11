Private Sub stopRun(time As Integer)

    'pausa la ejecución en un tiempo establecido
    '
    'Args:
    '   time (Integer): tiempo a pausa la ejecución
    '

    Dim timeFormat As String
    
    If (time < 10) Then
        timeFormat = "0:00:0" & time
        Else:
            timeFormat = "0:00:" & time
    End If

    Application.Wait (Now + TimeValue(timeFormat))

End Sub