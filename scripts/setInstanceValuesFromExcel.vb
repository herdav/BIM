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
    Dim FDK As Dictionary(Of String, EigenschaftInfo) = LoadFDKFromExcel()
    SetEigenschaften(FDK)
End Sub

Function LoadFDKFromExcel() As Dictionary(Of String, EigenschaftInfo)
    Dim localFDK As New Dictionary(Of String, EigenschaftInfo)()

    Dim row As Integer = 1
    While Not String.IsNullOrEmpty(GoExcel.CellValue("3rd Party:data", "fdk", "A" & row))
        Dim eigenschaft As String = GoExcel.CellValue("3rd Party:data", "fdk", "B" & row)
        Dim datentyp As String = GoExcel.CellValue("3rd Party:data", "fdk", "C" & row)
        Dim id As String = GoExcel.CellValue("3rd Party:data", "fdk", "A" & row)

        localFDK.Add(eigenschaft, New EigenschaftInfo(datentyp, id))
        row += 1
        Logger.Info(id + "_" + eigenschaft + "_" + datentyp)
    End While

    Return localFDK
End Function

Sub SetEigenschaften(ByVal eigenschaftenDict As Dictionary(Of String, EigenschaftInfo))

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
        If oCompOcc.Name.Contains("Mast") Then
            iProperties.InstanceValue(oCompOcc.Name, "Exemplarname") = oCompOcc.Name

            For Each kvp In eigenschaftenDict
                Dim eigenschaft As String = kvp.Key
                Dim datentyp As String = kvp.Value.Datentyp
                Dim id As String = kvp.Value.ID

                Dim wert As Object = "zu definieren" ' Standardwert

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

            Logger.Info(oCompOcc.Name)
        End If
    Next
End Sub
