'Set Instance Values from Excel
'Version 13.05.2024 by David Herren @ WiVi AG

Public Class EigenschaftInfo
    Public Property Datentyp As String
    Public Property ID As String

    Public Sub New(Datentyp As String, ID As String)
        Me.Datentyp = Datentyp
        Me.ID = ID
    End Sub
End Class

Sub Main()
    LoadFDKAndSetEigenschaften()
End Sub

Sub LoadFDKAndSetEigenschaften()
    Dim filterValue As String = InputBox("Bitte geben Sie den FDK Objekttyp exakt an.", "FDK Objekttyp")

    If String.IsNullOrEmpty(filterValue) Then
        MessageBox.Show("Kein Wert eingegeben. Das Skript wird beendet.")
        Exit Sub
    End If

    Dim result = LoadFDKFromExcel(filterValue)
    Dim FDK = result.Item1
    Dim count = result.Item2

    If count > 0 Then
        MessageBox.Show(String.Format("Der Objekttyp '{0}' wurde {1} Mal genau gefunden.", filterValue, count))
    Else
        MessageBox.Show(String.Format("Der Objekttyp '{0}' wurde nicht genau gefunden.", filterValue))
    End If

    Dim componentsModified As Integer = SetEigenschaften(FDK, filterValue)

    MessageBox.Show(String.Format("{0} Komponenten wurden geändert.", componentsModified))

    Dim restartResult As DialogResult = MessageBox.Show("Möchten Sie weitere Komponenten ändern?", "Neustart", MessageBoxButtons.YesNo)
    If restartResult = DialogResult.Yes Then
        LoadFDKAndSetEigenschaften()
    Else
        Exit Sub ' Skript wird beendet, wenn "Nein" oder eine andere Antwort ausgewählt wird
    End If
End Sub

Function LoadFDKFromExcel(ByVal filterValue As String) As Tuple(Of Dictionary(Of String, EigenschaftInfo), Integer)
    Dim localFDK As New Dictionary(Of String, EigenschaftInfo)()
    Dim matchCount As Integer = 0

    Dim row As Integer = 2
    While Not String.IsNullOrEmpty(GoExcel.CellValue("3rd Party:data", "FDK", "A" & row)) 'ID Obejktgruppe
        ' Überprüfe, ob in Spalte K 'Nein' steht, und überspringe die Zeile falls ja
        Dim skipRow As String = GoExcel.CellValue("3rd Party:data", "FDK", "K" & row)
        If skipRow = "Nein" Then
            row += 1
            Continue While
        End If

        Dim objekttypNameDE As String = GoExcel.CellValue("3rd Party:data", "FDK", "D" & row) 'ObjekttypNameDE

        If objekttypNameDE = filterValue Then 
            Dim eigenschaft As String = GoExcel.CellValue("3rd Party:data", "FDK", "F" & row) 'Eigenschaft
            Dim datentyp As String = GoExcel.CellValue("3rd Party:data", "FDK", "H" & row)   'Format
            Dim id As String = GoExcel.CellValue("3rd Party:data", "FDK", "E" & row)         'ID Eigenschaft

            If Not localFDK.ContainsKey(eigenschaft) Then
                localFDK.Add(eigenschaft, New EigenschaftInfo(datentyp, id))
                Logger.Info(id + "_" + eigenschaft + "_" + datentyp)
                matchCount += 1
            End If
        End If

        row += 1
    End While

    Return New Tuple(Of Dictionary(Of String, EigenschaftInfo), Integer)(localFDK, matchCount)
End Function

Function SetEigenschaften(ByVal eigenschaftenDict As Dictionary(Of String, EigenschaftInfo), ByVal filter As String) As Integer
    Dim componentsModified As Integer = 0 ' Zählvariable

    Dim oDoc As Document = ThisDoc.Document
    Dim oAsm As AssemblyDocument
    Dim oCompOccs As ComponentOccurrences

    If oDoc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
        oAsm = TryCast(oDoc, AssemblyDocument)
        oCompOccs = oAsm.ComponentDefinition.Occurrences
    Else
        MessageBox.Show("Dieses Dokument ist keine Baugruppe!")
        Exit Function
    End If

    Dim filterPattern As String = "-" & filter & ":"
    For Each oCompOcc In oCompOccs
        If oCompOcc.Name.Contains(filterPattern) Then
            componentsModified += 1

            For Each kvp In eigenschaftenDict
                Dim eigenschaft As String = kvp.Key
                Dim datentyp As String = kvp.Value.Datentyp
                Dim id As String = kvp.Value.ID

                Dim wert As Object = "nicht definiert"

                Select Case datentyp
                    Case "String"
                        wert = "nicht definiert"
                    Case "Double", "Real"
                        wert = 0.0
                    Case "Date", "Datum"
                        wert = Date.Now
                    Case "Boolean"
                        wert = False
                    Case "Integer"
                        wert = 0
                    Case Else
                        wert = "nicht definiert"
                End Select

                iProperties.InstanceValue(oCompOcc.Name, id & " " & eigenschaft) = wert
            Next
        End If
    Next

    Return componentsModified
End Function
