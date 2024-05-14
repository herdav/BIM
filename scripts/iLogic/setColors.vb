'Set Colors
'Version 14.05.2024 by David Herren @ WiVi AG

Dim oApp As Inventor.Application
oApp = ThisApplication

Dim oAsmDoc As AssemblyDocument
oAsmDoc = oApp.ActiveDocument

If oAsmDoc.DocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then
    MessageBox.Show("Das aktive Dokument ist keine Baugruppe.", "Fehler")
    Exit Sub
End If

For Each oComponent In oAsmDoc.ComponentDefinition.Occurrences
    If oComponent.Name.StartsWith("A-") Then
		Component.Color(oComponent.Name) = "_Abbruch"
		Logger.Info("Komponentenname: " & oComponent.Name & ">" & Component.Color(oComponent.Name))
    End If
	If oComponent.Name.StartsWith("B-") Then
		Component.Color(oComponent.Name) = "_Bestand"
		Logger.Info("Komponentenname: " & oComponent.Name & ">" & Component.Color(oComponent.Name))
    End If
Next

MessageBox.Show("Ausgabe der Komponentennamen abgeschlossen.", "iLogic")
