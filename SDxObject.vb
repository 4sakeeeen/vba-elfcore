Public DefUID As String
Public Properties As Collection


Sub Init(ByVal ObjDefUID As String, ByVal ObjProperties As Collection)

DefUID = ObjDefUID
Set Properties = ObjProperties

End Sub


Public Function GetProperty(Name As String) As SDxProperty

On Error GoTo Catch
    For Each Property In Properties
        If Property.Name = "UID" Then
            Set GetProperty = Property
        End If
    Next Property
    Exit Function

Catch:
    Call Err.Raise(10001, "SDxObject.GetProperty", "Property not found for '" & Name & "'.")
End Function
