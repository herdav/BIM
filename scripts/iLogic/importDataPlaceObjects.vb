'Import Data from Excel and Place Obejcts on SKLT
'Version 14.05.2024 by David Herren @ WiVi AG

Imports System.Windows.Forms

Sub Main()
    Dim result = SelectFDKType()
    If result IsNot Nothing Then
        Dim selection As String = result.Item1
        Dim n_max As Integer = result.Item2
        Select Case selection
            Case "Ausleger", "Spurhalter", "Spurhalterabzug"
                Generate(selection, n_max)
            Case "Importiere Daten von eingebettetem Excel"
                ImportData(n_max)
            Case "Ausleger & Spurhalter & Spurhalterabzug"
                ImportData(n_max)
                Generate("Ausleger", n_max)
                Generate("Spurhalter", n_max)
                Generate("Spurhalterabzug", n_max)
            Case Else
                MessageBox.Show("Kein gültiger Wert ausgewählt.")
        End Select
    Else
        MessageBox.Show("Skript wurde abgebrochen.")
    End If
End Sub

Function SelectFDKType() As Tuple(Of String, Integer)
    Dim validFilterValues As String() = {"Bitte eine Auswahl treffen", "Ausleger", "Spurhalter", "Spurhalterabzug", "Importiere Daten von eingebettetem Excel", "Ausleger & Spurhalter & Spurhalterabzug"}
    Dim form As New System.Windows.Forms.Form
    Dim comboBox As New System.Windows.Forms.ComboBox
    Dim textBox As New System.Windows.Forms.TextBox
    Dim label As New System.Windows.Forms.Label
    Dim buttonOK As New System.Windows.Forms.Button
    Dim buttonCancel As New System.Windows.Forms.Button

    form.Text = "Aktion ausführen"
    form.Width = 400
    form.Height = 200
    form.StartPosition = FormStartPosition.CenterScreen

    comboBox.DropDownStyle = ComboBoxStyle.DropDownList
    comboBox.Items.AddRange(validFilterValues)
    comboBox.SelectedIndex = 0
    comboBox.Left = 50
    comboBox.Top = 30
    comboBox.Width = 250
    form.Controls.Add(comboBox)

    label.Text = "n_max:"
    label.Left = 50
    label.Top = 70
    form.Controls.Add(label)

    textBox.Text = "100"
    textBox.Left = 150
    textBox.Top = 70
    textBox.Width = 100
    form.Controls.Add(textBox)

    buttonOK.Text = "OK"
    buttonOK.Left = 50
    buttonOK.Top = 110
    buttonOK.Width = 80
    AddHandler buttonOK.Click, Sub(sender, e) form.DialogResult = System.Windows.Forms.DialogResult.OK
    form.Controls.Add(buttonOK)

    buttonCancel.Text = "Abbrechen"
    buttonCancel.Left = 150
    buttonCancel.Top = 110
    buttonCancel.Width = 80
    AddHandler buttonCancel.Click, Sub(sender, e) form.DialogResult = System.Windows.Forms.DialogResult.Cancel
    form.Controls.Add(buttonCancel)

    If form.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
        If comboBox.SelectedItem IsNot Nothing AndAlso IsNumeric(textBox.Text) Then
            Dim selection As String = comboBox.SelectedItem.ToString()
            Dim n_max As Integer = Integer.Parse(textBox.Text)
            Return New Tuple(Of String, Integer)(selection, n_max)
        End If
    End If

    Return Nothing
End Function

Sub Generate(selection As String, n_max As Integer)
    '[ Setup
    Dim n = 100          'number of parts / datasets
    Dim n_min = 1        'set min range
    Dim r = 5            'number of 1st row
    Dim c = 8            'number of columns !!
    Dim f = 1000         'convert data to mm for vectors
    Dim name_bks(n), name_dir(n), position_bks(n), position_dir(n), name_sklt(n), name_trwk(n), name_sptr(n), name_shaz(n)
    Dim x_red = GoExcel.CellValue("3rd Party:data", "data", "T" & 1) 'shift X !!
    Dim y_red = GoExcel.CellValue("3rd Party:data", "data", "U" & 1) 'shift Y !!
    Dim z_red = GoExcel.CellValue("3rd Party:data", "data", "V" & 1) 'shift Z !!
    Dim typ(n), xe(n), yn(n), zz(n), xd(n), yd(n), a(n), switch(n), shift(n)
    Dim tempX(n), tempY(n)
    Dim ht(n), hf(n), alr(n), δ(n), ε(n), lmr(n), hp(n), α(n)
    Dim ObjekttypID(n), ObjekttypNameDE(n), Status(n), stat(n), Objektcode(n), DIDOK(n)    'FDK / BIM Data
    Dim KM(n) As String, fBIMDeX = 1 '30.48 Factor for BIMDeX Export?
    Dim nr(n) As String
    Logger.Info("Datasets: " & n)
    ']

    '[ Import Data from excel
    For i = n_min - 1 To n_max - 1 'Load data from excel and save it in an array
        nr(i)     =  GoExcel.CellValue("3rd Party:data", "data", "C"  & i + r)              'Nr
        typ(i)    =  GoExcel.CellValue("3rd Party:data", "data", "E"  & i + r)              'Typ
        hf(i)     =  GoExcel.CellValue("3rd Party:data", "data", "F"  & i + r)              'Höhe zwischen Fahrdraht und Gleis
        hp(i)     =  GoExcel.CellValue("3rd Party:data", "data", "G"  & i + r)              'Höhe zwischen Tragseil und Gleis
        ht(i)     =  GoExcel.CellValue("3rd Party:data", "data", "H"  & i + r)              'Höhe zwischen Tragrohrachse und Gleis
        alr(i)    =  GoExcel.CellValue("3rd Party:data", "data", "J"  & i + r)              'Abstand zwischen Gleisachse und Befestigung
        α(i)      =  GoExcel.CellValue("3rd Party:data", "data", "M"  & i + r)              'Neigung Gleisachse
        δ(i)      =  GoExcel.CellValue("3rd Party:data", "data", "N"  & i + r)              'Versatz Gleisachse
        ε(i)      =  GoExcel.CellValue("3rd Party:data", "data", "O"  & i + r)              'Versatz Fahrzeugachse
        xe(i)     = (GoExcel.CellValue("3rd Party:data", "data", "T"  & i + r) - x_red) * f 'X(E)-Achse Fundament
        yn(i)     = (GoExcel.CellValue("3rd Party:data", "data", "U"  & i + r) - y_red) * f 'Y(N)-Achse Fundament
        zz(i)     = (GoExcel.CellValue("3rd Party:data", "data", "V"  & i + r) - z_red) * f 'Z(Z)-Achse Fundament
        xd(i)     = (GoExcel.CellValue("3rd Party:data", "data", "W"  & i + r) - x_red) * f 'X(E)-Achse Ausrichtung
        yd(i)     = (GoExcel.CellValue("3rd Party:data", "data", "X"  & i + r) - y_red) * f 'Y(N)-Achse Ausrichtung
        switch(i) =  GoExcel.CellValue("3rd Party:data", "data", "Y"  & i + r)              'Switch bks and dir
        lmr(i)    =  GoExcel.CellValue("3rd Party:data", "data", "Z"  & i + r)              'Befestigung Fahrdraht Links, Mitte oder Rechts der Gleisachse
        shift(i)  =  GoExcel.CellValue("3rd Party:data", "data", "AA" & i + r)              'Position Ausrichtung Tragwerk schieben (innen <> aussen)
        KM(i)     = GoExcel.CellValue("3rd Party:data", "data", "AG" & i + r)
        Status(i) = GoExcel.CellValue("3rd Party:data", "data", "AJ" & i + r)
        
        If (switch(i) = "true") 'Switch bks and dir
            tempX(i) = xd(i)
            tempY(i) = yd(i)
            xd(i) = xe(i)
            yd(i) = yn(i)
            xe(i) = tempX(i)
            yn(i) = tempY(i)
        End If
        
        If (Status(i) = "Bestand")
            stat(i) = "B"
        Else If (Status(i) = "Abbruch")
            stat(i) = "A"
        Else If (Status(i) = "Neu")
            stat(i) = "N"
        Else 
            stat(i) = "nicht definiert"
        End If
    Next
    Logger.Info("Done setup.")
    ']

    '[ Generate position oriented parts
    For i = n_min - 1 To n_max - 1
        ' Set a reference to the assembly component definition.
        Dim oAsmCompDef As AssemblyComponentDefinition
        oAsmCompDef = ThisApplication.ActiveDocument.ComponentDefinition

        ' Set a reference to the transient geometry object.
        Dim oTG As TransientGeometry
        oTG = ThisApplication.TransientGeometry

        ' Create a matrix.  A new matrix is initialized with an identity matrix.
        Dim oMatrix As Matrix
        oMatrix = oTG.CreateMatrix
        
        ' Set the rotation of the matrix about the Z axis.
        If ((xe(i) > xd(i) And yn(i) > yd(i)) Or yd(i) < yn(i)) 'Check direction of rotation
            Call oMatrix.SetToRotation(-a(i) * PI / 180, oTG.CreateVector(0, 0, 1), oTG.CreatePoint(0, 0, 0)) 'rotate left  (-)
        Else
            Call oMatrix.SetToRotation(a(i) * PI / 180, oTG.CreateVector(0, 0, 1), oTG.CreatePoint(0, 0, 0))  'rotate right (+)
        End If

        ' Set the translation portion of the matrix so the part will be positioned
        Call oMatrix.SetTranslation(oTG.CreateVector(xe(i)/10, yn(i)/10, zz(i)/10))  'distances in a matrix always defined in centimeters

        ' Add the occurrence depending on the selection and typ(i).
        Dim oOcc As ComponentOccurrence
        Dim name_suffix As String

        If selection = "Ausleger" Then
            name_suffix = "Ausleger"
            name_sklt(i) = stat(i) & "-QP-" & nr(i) & "-SKLT:" & i
            
            If (typ(i) = "Typ10A") Then
                oOcc = oAsmCompDef.Occurrences.Add(ThisDoc.Path & "..\..\Tragwerke\Typ10A-Tragwerk.iam", oMatrix)
                name_trwk(i) = stat(i) & "-QP-" & nr(i) & "-" & name_suffix & "-Typ10A:" & i
                oOcc.Name = name_trwk(i)
                Constraints.AddMate("ht" & i + 1, name_trwk(i).ToString, "ht-X-Achse", name_sklt(i).ToString, "XY-Ebene", offset := ht(i))
                Constraints.AddFlush("alr" & i + 1, name_trwk(i).ToString, "YZ-Ebene", name_sklt(i).ToString, "Gleisachse-YZ-Ebene", offset := alr(i))
                Constraints.AddFlush("XZ-Ebene" & i + 1, name_trwk(i).ToString, "XZ-Ebene", name_sklt(i).ToString, "XZ-Ebene", offset := 0)
                
            Else If (typ(i) = "Typ10R") Then
                oOcc = oAsmCompDef.Occurrences.Add(ThisDoc.Path & "..\..\Tragwerke\Typ10R-Tragwerk.iam", oMatrix)
                name_trwk(i) = stat(i) & "-QP-" & nr(i) & "-" & name_suffix & "-Typ10R:" & i
                oOcc.Name = name_trwk(i)
                Constraints.AddMate("ht" & i + 1, name_trwk(i).ToString, "ht-X-Achse", name_sklt(i).ToString, "XY-Ebene", offset := ht(i))
                Constraints.AddFlush("alr" & i + 1, name_trwk(i).ToString, "YZ-Ebene", name_sklt(i).ToString, "Gleisachse-YZ-Ebene", offset := alr(i))
                Constraints.AddFlush("XZ-Ebene" & i + 1, name_trwk(i).ToString, "XZ-Ebene", name_sklt(i).ToString, "XZ-Ebene", offset := 0)
                
            Else If (typ(i) = "Typ21A") Then
                oOcc = oAsmCompDef.Occurrences.Add(ThisDoc.Path & "..\..\Tragwerke\Typ21A-Tragwerk.iam", oMatrix)
                name_trwk(i) = stat(i) & "-QP-" & nr(i) & "-" & name_suffix & "-Typ21A:" & i
                oOcc.Name = name_trwk(i)
                Constraints.AddMate("ht" & i + 1, name_trwk(i).ToString, "ht-X-Achse", name_sklt(i).ToString, "XY-Ebene", offset := ht(i))
                If (shift(i) = "true") Then
                    Constraints.AddFlush("alr" & i + 1, name_trwk(i).ToString, "YZ-Ebene", name_sklt(i).ToString, "Gleisachse-YZ-Ebene", offset := alr(i) * -1)
                    Constraints.AddFlush("XZ-Ebene" & i + 1, name_trwk(i).ToString, "XZ-Ebene", name_sklt(i).ToString, "XZ-Ebene", offset := 0)
                Else
                    Constraints.AddMate("alr" & i + 1, name_trwk(i).ToString, "YZ-Ebene", name_sklt(i).ToString, "Gleisachse-YZ-Ebene", offset := alr(i) * -1)
                    Constraints.AddMate("XZ-Ebene" & i + 1, name_trwk(i).ToString, "XZ-Ebene", name_sklt(i).ToString, "XZ-Ebene", offset := 0)
                End If
            
            Else If (typ(i) = "Typ21R") Then
                oOcc = oAsmCompDef.Occurrences.Add(ThisDoc.Path & "..\..\Tragwerke\Typ21R-Tragwerk.iam", oMatrix)
                name_trwk(i) = stat(i) & "-QP-" & nr(i) & "-" & name_suffix & "-Typ21R:" & i
                oOcc.Name = name_trwk(i)
                Constraints.AddMate("ht" & i + 1, name_trwk(i).ToString, "ht-X-Achse", name_sklt(i).ToString, "XY-Ebene", offset := ht(i))
                If (shift(i) = "true") Then
                    Constraints.AddFlush("alr" & i + 1, name_trwk(i).ToString, "YZ-Ebene", name_sklt(i).ToString, "Gleisachse-YZ-Ebene", offset := alr(i) * -1)
                    Constraints.AddFlush("XZ-Ebene" & i + 1, name_trwk(i).ToString, "XZ-Ebene", name_sklt(i).ToString, "XZ-Ebene", offset := 0)
                Else
                    Constraints.AddMate("alr" & i + 1, name_trwk(i).ToString, "YZ-Ebene", name_sklt(i).ToString, "Gleisachse-YZ-Ebene", offset := alr(i) * -1)
                    Constraints.AddMate("XZ-Ebene" & i + 1, name_trwk(i).ToString, "XZ-Ebene", name_sklt(i).ToString, "XZ-Ebene", offset := 0)
                End If
            End If
            Logger.Info(name_trwk(i))
        
        Else If selection = "Spurhalter" Then
            name_suffix = "Spurhalter"
            name_sklt(i) = stat(i) & "-QP-" & nr(i) & "-SKLT:" & i    
            
            If (typ(i) = "Typ10A" Or typ(i) = "Typ10R" Or typ(i) = "Typ21A" Or typ(i) = "Typ21R") Then
                oOcc = oAsmCompDef.Occurrences.Add(ThisDoc.Path & "..\..\Tragwerke\Spurhalter.ipt", oMatrix)
                name_sptr(i) = stat(i) & "-QP-" & nr(i) & "-" & name_suffix & ":" & i
                oOcc.Name = name_sptr(i)
                Constraints.AddFlush("sptr-XZ-Ebene" & i + 1, name_sptr(i).ToString, "XZ-Ebene", name_sklt(i).ToString, "XZ-Ebene", offset := 0)
                Constraints.AddMate("sptr-ε-Y-Achse" & i + 1, name_sptr(i).ToString, "ε-Y-Achse", name_sklt(i).ToString, "ε-Y-Achse", offset := 0)
                Constraints.AddMate("sptr-alr-Z" & i + 1, name_sptr(i).ToString, "Spurhalterabzug-Y-Achse", name_sklt(i).ToString, "XY-Ebene", offset := ht(i) -225)
            End If
            Logger.Info(name_sptr(i))
        
        Else If selection = "Spurhalterabzug" Then
            name_suffix = "Spurhalterabzug"
            name_sklt(i) = stat(i) & "-QP-" & nr(i) & "-SKLT:" & i    

            If (typ(i) = "Typ10A" Or typ(i) = "Typ10R" Or typ(i) = "Typ21A" Or typ(i) = "Typ21R") Then
                oOcc = oAsmCompDef.Occurrences.Add(ThisDoc.Path & "..\..\Tragwerke\Spurhalterabzug.ipt", oMatrix)
                name_shaz(i) = stat(i) & "-QP-" & nr(i) & "-" & name_suffix & ":" & i
                name_sptr(i) = stat(i) & "-QP-" & nr(i) & "-Spurhalter:" & i
                oOcc.Name = name_shaz(i)
                Constraints.AddFlush("shaz-XZ-Ebene" & i + 1, name_shaz(i).ToString, "XZ-Ebene", name_sklt(i).ToString, "XZ-Ebene", offset := 0)
                Constraints.AddMate("shaz-Y-Achse" & i + 1, name_shaz(i).ToString, "Y-Achse", name_sptr(i).ToString, "Spurhalterabzug-Y-Achse", offset := 0)
                Constraints.AddAngle("shaz-Winkel" & i + 1, name_shaz(i).ToString, "Z-Achse", name_sklt(i).ToString, "Z-Achse", angle := 0.0)
            End If
            Logger.Info(name_shaz(i))
        End If
    Next
    ']

    '[ Set instance values
    For i = n_min - 1 To n_max - 1
        If (typ(i) = "Typ10A" Or typ(i) = "Typ10R" Or typ(i) = "Typ21A" Or typ(i) = "Typ21R") '& " xx" > makes sure iProperties is a text
            If selection = "Ausleger" Then
                iProperties.InstanceValue(name_trwk(i), "PTY_0 Status")          = Status(i)
                iProperties.InstanceValue(name_trwk(i), "PTY_0 Objektcode")      = "tl03"
                iProperties.InstanceValue(name_trwk(i), "PTY_0 DIDOK")           = "HIGT"
                iProperties.InstanceValue(name_trwk(i), "PTY_0 ObjekttypID")     = "OBJ_FS_18"
                iProperties.InstanceValue(name_trwk(i), "PTY_0 ObjekttypNameDE") = "Ausleger (FS)"
            
            Else If selection = "Spurhalter" Then
                iProperties.InstanceValue(name_sptr(i), "PTY_0 Kilometrierung")  = KM(i)
                iProperties.InstanceValue(name_sptr(i), "PTY_0 Status")          = Status(i)
                iProperties.InstanceValue(name_sptr(i), "PTY_0 Objektcode")      = "tl03"
                iProperties.InstanceValue(name_sptr(i), "PTY_0 DIDOK")           = "HIGT"
                iProperties.InstanceValue(name_sptr(i), "PTY_0 ObjekttypID")     = "OBJ_FS_16"
                iProperties.InstanceValue(name_sptr(i), "PTY_0 ObjekttypNameDE") = "Spurhalter"
            
            Else If selection = "Spurhalterabzug" Then
                iProperties.InstanceValue(name_shaz(i), "PTY_0 Kilometrierung")  = KM(i)
                iProperties.InstanceValue(name_shaz(i), "PTY_0 Status")          = Status(i)
                iProperties.InstanceValue(name_shaz(i), "PTY_0 Objektcode")      = "tl03"
                iProperties.InstanceValue(name_shaz(i), "PTY_0 DIDOK")           = "HIGT"
                iProperties.InstanceValue(name_shaz(i), "PTY_0 ObjekttypID")     = "OBJ_FS_106"
                iProperties.InstanceValue(name_shaz(i), "PTY_0 ObjekttypNameDE") = "Spurhalterabzug"
            End If
        End If
    Next
    ']

    '[ End of routine
    Logger.Info("Update Model.")
    InventorVb.DocumentUpdate() 'Update model if run script
    iLogicVb.UpdateWhenDone = True
    ThisApplication.CommandManager.ControlDefinitions.Item("AssemblyRebuildAllCmd").Execute
    Logger.Info("Finished.")
    ']
End Sub

Sub ImportData(n_max As Integer)
    '[ Setup
    Dim n = 100          'number of parts / datasets
    Dim n_min = 1        'set min range
    Dim r = 5            'number of 1st row
    Dim c = 8            'number of columns !!
    Dim f = 1000         'convert data to mm for vectors
    Dim name_bks(n), name_dir(n), position_bks(n), position_dir(n), name_sklt(n), name_trwk(n), name_sptr(n), name_shaz(n)
    Dim x_red = GoExcel.CellValue("3rd Party:data", "data", "T" & 1) 'shift X !!
    Dim y_red = GoExcel.CellValue("3rd Party:data", "data", "U" & 1) 'shift Y !!
    Dim z_red = GoExcel.CellValue("3rd Party:data", "data", "V" & 1) 'shift Z !!
    Dim typ(n), xe(n), yn(n), zz(n), xd(n), yd(n), a(n), switch(n), shift(n)
    Dim tempX(n), tempY(n)
    Dim ht(n), hf(n), alr(n), δ(n), ε(n), lmr(n), hp(n), α(n)
    Dim ObjekttypID(n), ObjekttypNameDE(n), Status(n), stat(n), Objektcode(n), DIDOK(n)    'FDK / BIM Data
    Dim KM(n) As String, fBIMDeX = 1 '30.48 Factor for BIMDeX Export?
    Dim nr(n) As String
    Logger.Info("Datasets: " & n)
    ']

    '[ Import Data from excel
    For i = n_min - 1 To n_max - 1 'Load data from excel and save it in an array
        nr(i)     =  GoExcel.CellValue("3rd Party:data", "data", "C"  & i + r)              'Nr
        typ(i)    =  GoExcel.CellValue("3rd Party:data", "data", "E"  & i + r)              'Typ
        hf(i)     =  GoExcel.CellValue("3rd Party:data", "data", "F"  & i + r)              'Höhe zwischen Fahrdraht und Gleis
        hp(i)     =  GoExcel.CellValue("3rd Party:data", "data", "G"  & i + r)              'Höhe zwischen Tragseil und Gleis
        ht(i)     =  GoExcel.CellValue("3rd Party:data", "data", "H"  & i + r)              'Höhe zwischen Tragrohrachse und Gleis
        alr(i)    =  GoExcel.CellValue("3rd Party:data", "data", "J"  & i + r)              'Abstand zwischen Gleisachse und Befestigung
        α(i)      =  GoExcel.CellValue("3rd Party:data", "data", "M"  & i + r)              'Neigung Gleisachse
        δ(i)      =  GoExcel.CellValue("3rd Party:data", "data", "N"  & i + r)              'Versatz Gleisachse
        ε(i)      =  GoExcel.CellValue("3rd Party:data", "data", "O"  & i + r)              'Versatz Fahrzeugachse
        xe(i)     = (GoExcel.CellValue("3rd Party:data", "data", "T"  & i + r) - x_red) * f 'X(E)-Achse Fundament
        yn(i)     = (GoExcel.CellValue("3rd Party:data", "data", "U"  & i + r) - y_red) * f 'Y(N)-Achse Fundament
        zz(i)     = (GoExcel.CellValue("3rd Party:data", "data", "V"  & i + r) - z_red) * f 'Z(Z)-Achse Fundament
        xd(i)     = (GoExcel.CellValue("3rd Party:data", "data", "W"  & i + r) - x_red) * f 'X(E)-Achse Ausrichtung
        yd(i)     = (GoExcel.CellValue("3rd Party:data", "data", "X"  & i + r) - y_red) * f 'Y(N)-Achse Ausrichtung
        switch(i) =  GoExcel.CellValue("3rd Party:data", "data", "Y"  & i + r)              'Switch bks and dir
        lmr(i)    =  GoExcel.CellValue("3rd Party:data", "data", "Z"  & i + r)              'Befestigung Fahrdraht Links, Mitte oder Rechts der Gleisachse
        shift(i)  =  GoExcel.CellValue("3rd Party:data", "data", "AA" & i + r)              'Position Ausrichtung Tragwerk schieben (innen <> aussen)
        KM(i)     = GoExcel.CellValue("3rd Party:data", "data", "AG" & i + r)
        Status(i) = GoExcel.CellValue("3rd Party:data", "data", "AJ" & i + r)
        
        If (switch(i) = "true") Then 'Switch bks and dir
            tempX(i) = xd(i)
            tempY(i) = yd(i)
            xd(i) = xe(i)
            yd(i) = yn(i)
            xe(i) = tempX(i)
            yn(i) = tempY(i)
        End If
        
        If (Status(i) = "Bestand") Then
            stat(i) = "B"
        Else If (Status(i) = "Abbruch") Then
            stat(i) = "A"
        Else If (Status(i) = "Neu") Then
            stat(i) = "N"
        Else 
            stat(i) = "nicht definiert"
        End If
        ']

        '[ Generate bks-parts in model / assembly
        name_bks(i) = stat(i) & "-QP-" & nr(i) & "-BKS:" & i
        position_bks(i) = ThisAssembly.Geometry.Point(xe(i), yn(i), zz(i))
        Components.Add(name_bks(i), ThisDoc.Path & "..\..\SKLT\bks.ipt", position_bks(i), True, False)
        ']

        '[ Generate direction-parts in model / assembly
        name_dir(i) = stat(i) & "-QP-" & nr(i) & "-DIR:" & i
        position_dir(i) = ThisAssembly.Geometry.Point(xd(i), yd(i), zz(i))
        Components.Add(name_dir(i), ThisDoc.Path & "..\..\SKLT\dir.ipt", position_dir(i), True, False)
        ']

        '[ Calculate direction between bks(i) and dir(i)
        a(i) = Measure.Angle(name_dir(i), "center", name_bks(i), "center", name_bks(i), "direction")
        Logger.Info(a(i).ToString)
        ']
    Next

    Logger.Info("Direction parts generated.")

    '[ Generate position oriented parts
    For i = n_min - 1 To n_max - 1
        ' Set a reference to the assembly component definition.
        Dim oAsmCompDef As AssemblyComponentDefinition
        oAsmCompDef = ThisApplication.ActiveDocument.ComponentDefinition

        ' Set a reference to the transient geometry object.
        Dim oTG As TransientGeometry
        oTG = ThisApplication.TransientGeometry

        ' Create a matrix.  A new matrix is initialized with an identity matrix.
        Dim oMatrix As Matrix
        oMatrix = oTG.CreateMatrix
        
        ' Set the rotation of the matrix about the Z axis.
        If ((xe(i) > xd(i) And yn(i) > yd(i)) Or yd(i) < yn(i)) 'Check direction of rotation
            Call oMatrix.SetToRotation(-a(i) * PI / 180, oTG.CreateVector(0, 0, 1), oTG.CreatePoint(0, 0, 0)) 'rotate left  (-)
        Else
            Call oMatrix.SetToRotation(a(i) * PI / 180, oTG.CreateVector(0, 0, 1), oTG.CreatePoint(0, 0, 0))  'rotate right (+)
        End If

        ' Set the translation portion of the matrix so the part will be positioned
        Call oMatrix.SetTranslation(oTG.CreateVector(xe(i)/10, yn(i)/10, zz(i)/10))  'distances in a matrix always defined in centimeters

        ' Add the occurrence depending on typ(i).
        Dim oOcc As ComponentOccurrence
        oOcc = oAsmCompDef.Occurrences.Add(ThisDoc.Path & "..\..\SKLT\sklt.iam", oMatrix)
        name_sklt(i) = stat(i) & "-QP-" & nr(i) & "-SKLT:" & i
        oOcc.Name = name_sklt(i)
        
        oOcc.Flexible = True
        oOcc.Grounded = True

        Logger.Info(name_sklt(i))
    Next
    ']

    '[ Set constraints
    For i = n_min - 1 To n_max - 1
        Constraints.AddMate("soll" & i + 1, name_dir(i).ToString, "Z-Achse", {name_sklt(i), "sklt-dir:1"}, "Z-Achse",
            offset := 0.0, e1InferredType := InferredTypeEnum.kNoInference, e2InferredType := InferredTypeEnum.kNoInference,
            solutionType := MateConstraintSolutionTypeEnum.kNoSolutionType)
        Constraints.AddFlush("hf" & i + 1, name_dir(i).ToString, "XY-Ebene", {name_sklt(i), "sklt-hf:1" }, "XY-Ebene", offset := hf(i))
        If (lmr(i) = "R") Then
            Constraints.AddMate("δ" & i + 1, name_dir(i).ToString, "Z-Achse", {name_sklt(i), "sklt-delta:1" }, "YZ-Ebene",
                offset := δ(i), e1InferredType := InferredTypeEnum.kNoInference, e2InferredType := InferredTypeEnum.kNoInference,
                solutionType := MateConstraintSolutionTypeEnum.kNoSolutionType)
            Constraints.AddAngle("α" & i + 1, name_bks(i).ToString, "Z-Achse", {name_sklt(i), "lichtraumprofil:1" }, "Z-Achse",
                angle := α(i) * 180 / PI * -1, solutionType := AngleConstraintSolutionTypeEnum.kDirectedSolution)
            If (α(i) < 0) Then 'R & α(i) < 0
                Constraints.AddMate("ε" & i + 1, name_dir(i).ToString, "Z-Achse", {name_sklt(i), "sklt-epsilon:1" }, "YZ-Ebene",
                    offset := δ(i) + ε(i), e1InferredType := InferredTypeEnum.kNoInference, e2InferredType := InferredTypeEnum.kNoInference,
                    solutionType := MateConstraintSolutionTypeEnum.kNoSolutionType)
            Else 'R & α(i) >= 0
                If (ε(i) >= 0) Then
                    Constraints.AddMate("ε" & i + 1, name_dir(i).ToString, "Z-Achse", {name_sklt(i), "sklt-epsilon:1" }, "YZ-Ebene",
                        offset := δ(i) + ε(i), e1InferredType := InferredTypeEnum.kNoInference, e2InferredType := InferredTypeEnum.kNoInference,
                        solutionType := MateConstraintSolutionTypeEnum.kNoSolutionType)
                Else
                    Constraints.AddMate("ε" & i + 1, name_dir(i).ToString, "Z-Achse", {name_sklt(i), "sklt-epsilon:1" }, "YZ-Ebene",
                        offset := δ(i) * -1 - ε(i), e1InferredType := InferredTypeEnum.kNoInference, e2InferredType := InferredTypeEnum.kNoInference,
                        solutionType := MateConstraintSolutionTypeEnum.kNoSolutionType)
                End If
            End If
        Else
            Constraints.AddMate("δ" & i + 1, name_dir(i).ToString, "Z-Achse", {name_sklt(i), "sklt-delta:1" }, "YZ-Ebene",
                offset := δ(i) * -1, e1InferredType := InferredTypeEnum.kNoInference, e2InferredType := InferredTypeEnum.kNoInference,
                solutionType := MateConstraintSolutionTypeEnum.kNoSolutionType)
            Constraints.AddAngle("α" & i + 1, name_bks(i).ToString, "Z-Achse", {name_sklt(i), "lichtraumprofil:1" }, "Z-Achse",
                angle := α(i) * 180 / PI, solutionType := AngleConstraintSolutionTypeEnum.kDirectedSolution)
            If (α(i) < 0) Then 'L & α(i) < 0
                Constraints.AddMate("ε" & i + 1, name_dir(i).ToString, "Z-Achse", {name_sklt(i), "sklt-epsilon:1" }, "YZ-Ebene",
                    offset := δ(i) * -1 - ε(i), e1InferredType := InferredTypeEnum.kNoInference, e2InferredType := InferredTypeEnum.kNoInference,
                    solutionType := MateConstraintSolutionTypeEnum.kNoSolutionType)
            Else 'L & α(i) >= 0
                Constraints.AddMate("ε" & i + 1, name_dir(i).ToString, "Z-Achse", {name_sklt(i), "sklt-epsilon:1" }, "YZ-Ebene",
                    offset := δ(i) * -1 - ε(i), e1InferredType := InferredTypeEnum.kNoInference, e2InferredType := InferredTypeEnum.kNoInference,
                    solutionType := MateConstraintSolutionTypeEnum.kNoSolutionType)
            End If    
        End If
        Constraints.AddFlush("hp" & i + 1, name_dir(i).ToString, "XY-Ebene", {name_sklt(i), "sklt-hp:1" }, "XY-Ebene", offset := hp(i))
    Next    
    ']

    '[ End of routine
    InventorVb.DocumentUpdate() 'Update model if run script
    Logger.Info("Finished.")
    ']
End Sub
