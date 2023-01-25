Option Explicit


Sub RegenerateTableNames()

Dim Sheet As Worksheet
Dim Table As ListObject

For Each Sheet In ThisWorkbook.Worksheets
    If Sheet.Name <> "README" And Sheet.Name <> "Mapping" Then
        For Each Table In Sheet.ListObjects
            Table.Name = GetMappingToByFrom(Table.ListColumns(1).Name)
        Next Table
    End If
Next Sheet

End Sub


Function GetMappingToByFrom(MapFrom As String) As String

Dim Sheet As Worksheet
Dim Result As String
Dim Row As ListRow

On Error GoTo Catch
    For Each Sheet In ThisWorkbook.Worksheets
        If Sheet.Name = "Mapping" Then
            For Each Row In Sheet.ListObjects(1).ListRows
                If Row.Range(1).Cells.Value = MapFrom Then
                    Result = Row.Range(2).Cells.Value
                End If
            Next Row
            
            GetMappingToByFrom = Result
        End If
    Next Sheet
    Exit Function

Catch:
    Call Err.Raise(10001, ".GetMappingToByFrom", Err.Description)

End Function


Sub MergeXML()

On Error GoTo Catch
    Dim SDxObj As SDxObject
    Dim SDxObjs As Collection
    Dim XMLFilePath As String
    Dim ExistsTable As ELFTable
    Dim ExistsUID As String
    Dim ExistsProperties As Collection
    Dim Mapping As SPXExportMap
    Dim MappingItem As SPXExportMapEntry
    
    XMLFilePath = InputBox("Enter XML source file path", "Choose source...")
    Set SDxObjs = ParseXML(XMLFilePath)
    
    Set Mapping = New SPXExportMap
    Call Mapping.ReadMap(ThisWorkbook.Worksheets("Mapping"))
    
    For Each SDxObj In SDxObjs
        ExistsUID = SDxObj.GetProperty("UID").Value
        Set ExistsTable = New ELFTable
        
        If SDxObj.DefUID = "Rel" Then
            Set MappingItem = Mapping.GetMapEntry(SDxObj.Properties.Item("DefUID").Value)
            
            If MappingItem.MapType = "Rel" Then
                Set ExistsTable.Source = GetTableByName(MappingItem.MapToName)
                Set ExistsProperties = ExistsTable.GetObjPropertiesByUID(ExistsUID)
            End If
        Else
            Set ExistsTable.Source = GetTableByName(SDxObj.DefUID)
            Set ExistsProperties = ExistsTable.GetObjPropertiesByUID(ExistsUID)
        End If
    Next SDxObj

    Exit Sub

Catch:
    Call Err.Raise(10001, "MergeXML", Err.Description, Err.HelpFile, Err.HelpContext)

End Sub


Function GetTableByName(Name As String) As ListObject

Dim Sheet As Worksheet
Dim Table As ListObject
Dim Result As ListObject

On Error GoTo Catch
    For Each Sheet In ThisWorkbook.Sheets
        For Each Table In Sheet.ListObjects
            If Table.Name = Name Then
                Set Result = Table
            End If
        Next Table
    Next Sheet
    
    If Result Is Nothing Then
        GoTo Catch
    End If
    
    Set GetTableByName = Result
    Exit Function

Catch:
    Call Err.Raise(10001, "GetTableByName", "No table found '" & Name & "'", Err.HelpFile, Err.HelpContext)
End Function


Function ParseXML(filePath As String) As Collection
'Returns dictionary of objects: key - UID, value - SDxObject

On Error GoTo Catch
    Dim XDoc As Object, Root As Object
    Dim container As IXMLDOMElement
    Dim obj As IXMLDOMElement
    Dim Intf As IXMLDOMElement
    Dim Attrib As IXMLDOMAttribute
    Dim Prop As SDxProperty

    Set XDoc = CreateObject("MSXML2.DOMDocument")
    XDoc.async = False: XDoc.validateOnParse = False
    XDoc.Load ("D:\SDx\repos\tmp\exp.xml")
    Set container = XDoc.DocumentElement
    
    Dim Objects As New Collection
    
    For Each obj In container.ChildNodes
        Dim SDxObj As SDxObject
        Set SDxObj = New SDxObject
        Call SDxObj.Init(obj.BaseName, New Collection)
        
        For Each Intf In obj.ChildNodes
            For Each Attrib In Intf.Attributes
                Set Prop = New SDxProperty
                Call Prop.Init(Intf.BaseName, Attrib.BaseName, Attrib.Text)
                Call SDxObj.Properties.Add(Prop, Attrib.BaseName)
            Next Attrib
        Next Intf
            
        Call Objects.Add(SDxObj, SDxObj.Properties.Item("UID").Value)
    Next obj
    
    Set ParseXML = Objects

    Exit Function

Catch:
    Call Err.Raise(10001, "ParseXML", Err.Description, Err.HelpFile, Err.HelpContext)

End Function
