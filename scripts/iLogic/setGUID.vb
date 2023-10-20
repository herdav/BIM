Sub Main()
    Call AssemblySample()
End Sub

Public Sub AssemblySample()
    ' Get the active assembly.
    Dim oAsmDoc As AssemblyDocument
    oAsmDoc = ThisApplication.ActiveDocument

    ' Call the function that does the recursion.
    Call Assembly(oAsmDoc.ComponentDefinition.Occurrences, 1)
End Sub

Private Sub Assembly(Occurrences As ComponentOccurrences, _
                             Level As Integer)
    ' Iterate through all of the occurrence in this collection.  This
    ' represents the occurrences at the top level of an assembly.
    Dim oOcc As ComponentOccurrence
	
    For Each oOcc In Occurrences
		 ' Print the name of the current occurrence.
'	     Logger.Info(oOcc.Name)
'        Logger.Info("InternalName:" & oOcc.Definition.Document.InternalName)
'        Logger.Info("RevisionID:" & oOcc.Definition.Document.RevisionId)
'        Logger.Info("DatabaseRevisionID:" & oOcc.Definition.Document.DatabaseRevisionId)
'        Logger.Info("ModelGeometryVersion:" & oOcc.Definition.ModelGeometryVersion)

        ' Check to see if this occurrence represents a subassembly
        ' and recursively call this function to traverse through it.
        If oOcc.DefinitionDocumentType = kAssemblyDocumentObject Then
            Call Assembly(oOcc.SubOccurrences, Level + 1)
        End If
    Next
	
    Dim GUID As String
    GUID = oOcc.Definition.Document.InternalName
    GUID = Replace(GUID, "{", "")
    GUID = Replace(GUID, "}", "")
	
	iProperties.Value("Custom", "GUID") = GUID
	InventorVb.DocumentUpdate()
End Sub
