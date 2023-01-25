Option Explicit


Public Source As ListObject

Function GetObjPropertiesByUID(ByVal uid As String) As Collection

On Error GoTo Catch:
    Dim UIDIDX As Long
    Dim col As ListColumn
    Dim Row As ListRow
    Dim TempObj As SDxObject
    Dim Prop As SDxProperty
    Dim I As Integer
    
    ' To Access Object Functional
    Set TempObj = New SDxObject
    TempObj.Init "Temp", New Collection
    
    For Each col In Source.ListColumns
        Set Prop = New SDxProperty
        Prop.Name = GetMappingToByFrom(col.Name)
        TempObj.Properties.Add Prop
    
        If col.Name = "UID" Then
            UIDIDX = col.Index
        End If
    Next col
    
    For Each Row In Source.ListRows
        If Row.Range(UIDIDX).Cells.Value = uid Then
            For I = 1 To Row.Range.Count
                TempObj.Properties(I).Value = Row.Range(I).Cells.Value
            Next I
        End If
        
    Next Row
    
    If TempObj.GetProperty("UID").Value = "" Then
        Set TempObj.Properties = Nothing
    End If
    
    Set GetObjPropertiesByUID = TempObj.Properties

    Exit Function

Catch:
    Call Err.Raise(10001, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)

End Function
