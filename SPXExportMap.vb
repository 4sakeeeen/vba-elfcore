Option Explicit

Private Const mapFromNameColumn = "A"
Private Const mapToNameColumn = "B"
Private Const mapTypeColumn = "C"
Private Const mapAdditionalColumn = "D"
Private Const mapFunctionColumn = "E"
Private Const mapArg1Column = "F"
Private Const mapArg2Column = "G"
Private Const mapArg3Column = "H"

Private m_mapCollection As Collection

Private Sub Class_Initialize()
    Set m_mapCollection = New Collection
End Sub

Public Sub ReadMap(mapWorksheet As Excel.Worksheet)
    Dim currentRow As Integer
    currentRow = 4
    Dim isMoreData As Boolean
    isMoreData = True
    
    While isMoreData
        Dim mapFromName As String
        mapFromName = mapWorksheet.Range(mapFromNameColumn & currentRow).Value
        
        Dim MapToName As String
        MapToName = mapWorksheet.Range(mapToNameColumn & currentRow).Value
        
        Dim MapType As String
        MapType = mapWorksheet.Range(mapTypeColumn & currentRow).Value
        
        Dim mapAdditional As String
        mapAdditional = mapWorksheet.Range(mapAdditionalColumn & currentRow).Value
        
        Dim MapFunction As String
        MapFunction = mapWorksheet.Range(mapFunctionColumn & currentRow).Value
        
        Dim mapArg1 As String
        mapArg1 = mapWorksheet.Range(mapArg1Column & currentRow).Value
        
        Dim mapArg2 As String
        mapArg2 = mapWorksheet.Range(mapArg2Column & currentRow).Value
        
        Dim mapArg3 As String
        mapArg3 = mapWorksheet.Range(mapArg3Column & currentRow).Value
        
        If mapFromName <> "" And MapToName <> "" Then
            '// If we have from/to names then consider this a valid entry
            Dim MapEntry As SPXExportMapEntry
            Set MapEntry = New SPXExportMapEntry
            Call MapEntry.Init(mapFromName, MapToName, MapType, mapAdditional, MapFunction, mapArg1, mapArg2, mapArg3)
            Call m_mapCollection.Add(MapEntry, mapFromName)
        Else
            '// Invalid map entry. Consider this the end of the map.
            isMoreData = False
        End If
        
        currentRow = currentRow + 1
    Wend
End Sub

Public Function GetMapEntry(mapFromName As String) As SPXExportMapEntry
    On Error GoTo Catch
    Set GetMapEntry = m_mapCollection.Item(mapFromName)
    Exit Function
Catch:
    Call Err.Raise(10001, "SPXExportMap.GetMap", "Map entry not found for '" & mapFromName & "'.")
End Function

