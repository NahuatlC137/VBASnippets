'Determines if a given value exists in a range.
'Default returns are True and False, with the option to return given true and false values.

Function EXISTSIN(ByVal CELL As Range, ByVal LOOK_IN_RANGE As Range, Optional ByVal TRUE_MATCH As String, Optional ByVal FALSE_MATCH As String)

Dim Match As Boolean

    Match = IsNumeric(Application.Match(CELL.Value, LOOK_IN_RANGE, 0))
    
    Select Case Match
    
        Case True
            If TRUE_MATCH = "" Then
                EXISTSIN = Match
            Else
                EXISTSIN = TRUE_MATCH
            End If
            
        Case False
        
            If FALSE_MATCH = "" Then
                EXISTSIN = Match
            Else
                EXISTSIN = FALSE_MATCH
            End If
    
    End Select

End Function
