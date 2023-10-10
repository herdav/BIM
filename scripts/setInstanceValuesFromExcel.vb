Public Class EigenschaftInfo
    Public Property Datentyp As String
    Public Property ID As String

    Public Sub New(datentyp As String, id As String)
        Me.Datentyp = datentyp
        Me.ID = id
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
        MessageBox.Show(String.Format("Der Objekttyp '{0}' wurde in der Excel-Liste {1} Mal gefunden.", filterValue, count))
    Else
        MessageBox.Show(String.Format("Der Objekttyp '{0}' wurde in der Excel-Liste nicht gefunden.", filterValue))
    End If

    SetEigenschaften(FDK, filterValue)
End Sub

Function LoadFDKFromExcel(ByVal filterValue As String) As Tuple(Of Dictionary(Of String, EigenschaftInfo), Integer)
    Dim localFDK As New Dictionary(Of String, EigenschaftInfo)()
    Dim matchCount As Integer = 0

    Dim row As Integer = 3
    While Not String.IsNullOrEmpty(GoExcel.CellValue("3rd Party:data", "fdk", "A" & row))
        Dim objectType As String = GoExcel.CellValue("3rd Party:data", "fdk", "D" & row)

        If objectType.Contains(filterValue) Then 
            Dim eigenschaft As String = GoExcel.CellValue("3rd Party:data", "fdk", "F" & row)
            Dim datentyp As String = GoExcel.CellValue("3rd Party:data", "fdk", "H" & row)
            Dim id As String = GoExcel.CellValue("3rd Party:data", "fdk", "E" & row)

            localFDK.Add(eigenschaft, New EigenschaftInfo(datentyp, id))
            Logger.Info(id + "_" + eigenschaft + "_" + datentyp)

            matchCount += 1
        End If

        row += 1
    End While

    Return New Tuple(Of Dictionary(Of String, EigenschaftInfo), Integer)(localFDK, matchCount)
End Function

Sub SetEigenschaften(ByVal eigenschaftenDict As Dictionary(Of String, EigenschaftInfo), ByVal filter As String)

    Dim oDoc As Document = ThisDoc.Document
    Dim oAsm As AssemblyDocument
    Dim oCompOccs As ComponentOccurrences

    If oDoc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
        oAsm = TryCast(oDoc, AssemblyDocument)
        oCompOccs = oAsm.ComponentDefinition.Occurrences
    Else
        MessageBox.Show("Dieses Dokument ist keine Baugruppe!")
        Exit Sub
    End If

    For Each oCompOcc In oCompOccs
        If oCompOcc.Name.Contains(filter) Then
            iProperties.InstanceValue(oCompOcc.Name, "Exemplarname") = oCompOcc.Name

            For Each kvp In eigenschaftenDict
                Dim eigenschaft As String = kvp.Key
                Dim datentyp As String = kvp.Value.Datentyp
                Dim id As String = kvp.Value.ID

                Dim wert As Object = "zu definieren"

                Select Case datentyp
                    Case "String"
                        wert = "zu definieren"
                    Case "Double", "Real"
                        wert = 0.0
                    Case "Date", "Datum"
                        wert = Date.Now
                    Case "Boolean"
                        wert = False
                    Case "Integer"
                        wert = 0
                    Case Else
                        wert = "zu definieren"
                End Select

                iProperties.InstanceValue(oCompOcc.Name, id & " " & eigenschaft) = wert
            Next
        End If
    Next
End Sub
