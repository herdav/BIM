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
    Dim filterValue As String = InputBox("Bitte geben Sie den FDK Objekttyp an.", "FDK Objekttyp")

    If String.IsNullOrEmpty(filterValue) Then
        MessageBox.Show("Kein Wert eingegeben. Das Skript wird beendet.")
        Exit Sub
    End If

    Dim result = LoadFDKFromExcel(filterValue)
    Dim FDK = result.Item1
    Dim count = result.Item2

    If count > 0 Then
        MessageBox.Show(String.Format("Der Objekttyp '{0}' wurde {1} Mal gefunden.", filterValue, count))
    Else
        MessageBox.Show(String.Format("Der Objekttyp '{0}' wurde nicht gefunden.", filterValue))
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
    While Not String.IsNullOrEmpty(GoExcel.CellValue("3rd Party:data", "FDK", "A" & row))
        Dim objekttypNameDE As String = GoExcel.CellValue("3rd Party:data", "FDK", "D" & row)

		If objekttypNameDE.Contains(filterValue) Then 
		    Dim eigenschaft As String = GoExcel.CellValue("3rd Party:data", "FDK", "F" & row)
		    Dim datentyp As String = GoExcel.CellValue("3rd Party:data", "FDK", "H" & row)
		    Dim id As String = GoExcel.CellValue("3rd Party:data", "FDK", "E" & row)
		
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

    For Each oCompOcc In oCompOccs
        If oCompOcc.Name.Contains(filter) Then
            componentsModified += 1
            iProperties.InstanceValue(oCompOcc.Name, "WIV_1 InstanzName") = oCompOcc.Name

            ' Extrahiere den Kilometrierungswert, falls vorhanden, und füge ihn als Instanzeigenschaft hinzu
            Dim kmRegex As New System.Text.RegularExpressions.Regex("KM(\d+\.\d+)")
            Dim match As System.Text.RegularExpressions.Match = kmRegex.Match(oCompOcc.Name)
            If match.Success Then
                Dim kmValue As Double = Convert.ToDouble(match.Groups(1).Value)
                iProperties.InstanceValue(oCompOcc.Name, "Kilometrierung") = kmValue.ToString("F3")
            End If

            ' Extrahiere die Mastbezeichnung, falls vorhanden
            Dim mastRegex As New System.Text.RegularExpressions.Regex("Mast-(.*?)-")
            Dim mastMatch As System.Text.RegularExpressions.Match = mastRegex.Match(oCompOcc.Name)
            If mastMatch.Success Then
                Dim mastValue As String = mastMatch.Groups(1).Value
                iProperties.InstanceValue(oCompOcc.Name, "Mastbezeichnung") = mastValue
            End If

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
