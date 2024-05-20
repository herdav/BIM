'Set Instance Values from Excel
'Version 20.05.2024 by David Herren @ WiVi AG

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
  While True
    Dim filterValue As String = SelectFDKType()

    If String.IsNullOrEmpty(filterValue) OrElse filterValue = "Bitte eine Auswahl treffen" Then
      If filterValue Is Nothing Then
        'MessageBox.Show("Skript wurde abgebrochen.")
        Exit Sub
      End If
      MessageBox.Show("Kein gültiger Wert ausgewählt.")
      Continue While
    End If

    Dim excelFilterValue As String = If(filterValue = "Ausleger", "Ausleger (FS)", filterValue)
    Dim result = LoadFDKFromExcel(excelFilterValue)
    Dim FDK = result.Item1
    Dim count = result.Item2

    If count > 0 Then
      MessageBox.Show(String.Format("Der Objekttyp '{0}' wurde {1} Mal gefunden.", excelFilterValue, count))
    Else
      MessageBox.Show(String.Format("Der Objekttyp '{0}' wurde nicht gefunden.", excelFilterValue))
    End If

    Dim modelFilterValue As String = "Ausleger"
    Dim componentsModified As Integer = SetEigenschaften(FDK, modelFilterValue)

    MessageBox.Show(String.Format("{0} Komponenten wurden geändert.", componentsModified))

    Dim restartResult As DialogResult = MessageBox.Show("Möchten Sie weitere Komponenten ändern?", "Neustart", MessageBoxButtons.YesNo)
    If restartResult = DialogResult.No Then
      Exit While
    End If
  End While
End Sub

Function SelectFDKType() As String
  Dim validFilterValues As String() = {"Bitte eine Auswahl treffen", "Spurhalterabzug", "Spurhalter", "Fahrdraht", "Ausleger"}
  Dim form As New System.Windows.Forms.Form
  Dim comboBox As New System.Windows.Forms.ComboBox
  Dim buttonOK As New System.Windows.Forms.Button
  Dim buttonCancel As New System.Windows.Forms.Button

  form.Text = "Set Instance Value from Excel"
  form.Width = 300
  form.Height = 180
  form.StartPosition = FormStartPosition.CenterScreen

  comboBox.DropDownStyle = ComboBoxStyle.DropDownList
  comboBox.Items.AddRange(validFilterValues)
  comboBox.SelectedIndex = 0
  comboBox.Left = 50
  comboBox.Top = 30
  comboBox.Width = 200
  form.Controls.Add(comboBox)

  buttonOK.Text = "OK"
  buttonOK.Left = 50
  buttonOK.Top = 70
  buttonOK.Width = 80
  AddHandler buttonOK.Click, Sub(sender, e) form.DialogResult = System.Windows.Forms.DialogResult.OK
  form.Controls.Add(buttonOK)

  buttonCancel.Text = "Abbrechen"
  buttonCancel.Left = 150
  buttonCancel.Top = 70
  buttonCancel.Width = 80
  AddHandler buttonCancel.Click, Sub(sender, e) form.DialogResult = System.Windows.Forms.DialogResult.Cancel
  form.Controls.Add(buttonCancel)

  If form.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
    If comboBox.SelectedItem IsNot Nothing Then
      Return comboBox.SelectedItem.ToString()
    End If
  End If

  Return Nothing
End Function

Function LoadFDKFromExcel(ByVal filterValue As String) As Tuple(Of Dictionary(Of String, EigenschaftInfo), Integer)
  Dim localFDK As New Dictionary(Of String, EigenschaftInfo)()
  Dim matchCount As Integer = 0

  Dim row As Integer = 2
  While Not String.IsNullOrEmpty(GoExcel.CellValue("3rd Party:data", "FDK", "A" & row)) 'ID Obejktgruppe
    Dim skipRow As String = GoExcel.CellValue("3rd Party:data", "FDK", "K" & row)
    If skipRow = "Nein" Then
      row += 1
      Continue While
    End If

    Dim objekttypNameDE As String = GoExcel.CellValue("3rd Party:data", "FDK", "D" & row) 'ObjekttypNameDE

    If objekttypNameDE = filterValue Then
      Dim eigenschaft As String = GoExcel.CellValue("3rd Party:data", "FDK", "F" & row) 'Eigenschaft
      Dim datentyp As String = GoExcel.CellValue("3rd Party:data", "FDK", "H" & row) 'Format
      Dim id As String = GoExcel.CellValue("3rd Party:data", "FDK", "E" & row) 'ID Eigenschaft

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
  Dim componentsModified As Integer = 0

  Dim oDoc As Document = ThisDoc.Document
  Dim oAsm As AssemblyDocument
  Dim oCompOccs As ComponentOccurrences

  If oDoc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
    oAsm = TryCast(oDoc, AssemblyDocument)
    oCompOccs = oAsm.ComponentDefinition.Occurrences
  Else
    MessageBox.Show("Dieses Dokument ist keine Baugruppe!")
    Return componentsModified
  End If

  For Each oCompOcc In oCompOccs
    ' Suche nach dem Pattern im Namen der Komponenten
    If oCompOcc.Name.Contains(filter) Then
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
        Logger.Info(oCompOcc.Name & " " & eigenschaft & ": " & wert)
      Next
    End If
  Next

  Return componentsModified
End Function
