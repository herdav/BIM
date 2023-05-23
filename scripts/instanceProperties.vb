Sub Main()
	getInstanceProperties()
	setInstanceProperties()
End Sub

Sub getInstanceProperties()
	Dim instanceName = "Farbe"
    ' Get the active assembly document
    Dim asmDoc As AssemblyDocument
    asmDoc = ThisApplication.ActiveDocument
    
    ' Iterate through all occurrences in the assembly
    Dim occ As ComponentOccurrence
    For Each occ In asmDoc.ComponentDefinition.Occurrences
        ' Get the name of the component
        Dim componentName As String
        componentName = occ.Name
		
		' Print the component name
		Logger.Info(componentName & ": " & iProperties.InstanceValue(componentName, instanceName))
		iProperties.InstanceValue(componentName, "iName") = componentName
    Next
End Sub

Sub setInstanceProperties()
	Dim instanceName = "iName"
    ' Get the active assembly document
    Dim asmDoc As AssemblyDocument
    asmDoc = ThisApplication.ActiveDocument
    
    ' Iterate through all occurrences in the assembly
    Dim occ As ComponentOccurrence
    For Each occ In asmDoc.ComponentDefinition.Occurrences
        ' Get the name of the component
        Dim componentName As String
        componentName = occ.Name

		iProperties.InstanceValue(componentName, instanceName) = componentName
    Next
End Sub
