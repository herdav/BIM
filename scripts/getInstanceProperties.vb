Sub Main()
	getInstanceProperties()
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
    Next
End Sub
