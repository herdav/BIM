'Place Parts on SKLT
'Version 20.05.2024 by David Herren @ WiVi AG

Imports System.Windows.Forms

Class placePartsOnSKLT
    Private n As Integer = 100          'number of parts / datasets
    Private n_min As Integer = 1        'set min range
    Private n_max As Integer = 100      'set max range
    Private r As Integer = 5            'number of 1st row
    Private c As Integer = 8            'number of columns !!
    Private f As Integer = 1000         'convert data to mm for vectors
    Private name_bks(n), name_dir(n), position_bks(n), position_dir(n), name_sklt(n), name_trwk(n), name_sptr(n), name_shaz(n)
    Private x_red As Double
    Private y_red As Double
    Private z_red As Double
    Private typ(n), xe(n), yn(n), zz(n), xd(n), yd(n), a(n), switch(n), shift(n)
    Private tempX(n), tempY(n)
    Private ht(n), hf(n), alr(n), δ(n), ε(n), lmr(n), hp(n), α(n)
    Private ObjekttypID(n), ObjekttypNameDE(n), Status(n), stat(n), Objektcode(n), DIDOK(n)    'FDK / BIM Data
    Private KM(n) As String
    Private fBIMDeX As Integer = 1 '30.48 Factor for BIMDeX Export?
    Private nr(n) As String

    Sub Main()
        Dim selection = SelectFDKType(n_max)
        If selection Is Nothing Then
            Logger.Info("Keine Auswahl getroffen oder ungültiger Wert für n_max.")
            Exit Sub
        End If
        
        Dim action As String = selection.Item1
        n_max = selection.Item2
        
        ImportData(n_max)
        
        If action = "Ausleger" Or action = "Alle" Then
            '[ Generate position oriented parts: Ausleger
            For i = n_min - 1 To n_max - 1
                ' Set a reference to the assembly component definition.
                Dim oAsmCompDef As AssemblyComponentDefinition
                oAsmCompDef = ThisApplication.ActiveDocument.ComponentDefinition
        
                ' Set a reference to the transient geometry object.
                Dim oTG As TransientGeometry
                oTG = ThisApplication.TransientGeometry
        
                ' Create a matrix.  A new matrix is initialized with an identity matrix.
                Dim oMatrix As Matrix = CreateMatrix(a(i), xe(i), xd(i), yn(i), yd(i), zz(i))
        
                ' Add the occurrence depending on typ(i).
                Dim oOcc As ComponentOccurrence
                
                name_sklt(i) = stat(i) & "-QP-" & nr(i) & "-SKLT:" & i
                
                If (typ(i) = "Typ10A") Then
                    oOcc = oAsmCompDef.Occurrences.Add(ThisDoc.Path & "..\..\Tragwerke\Typ10A-Tragwerk.iam", oMatrix)
                    name_trwk(i) = stat(i) & "-QP-" & nr(i) & "-Ausleger-Typ10A:" & i
                    oOcc.Name = name_trwk(i)
                    Constraints.AddMate("ht" & i + 1, name_trwk(i).ToString, "ht-X-Achse", name_sklt(i).ToString, "XY-Ebene", offset := ht(i))
                    Constraints.AddFlush("alr" & i + 1, name_trwk(i).ToString, "YZ-Ebene", name_sklt(i).ToString, "Gleisachse-YZ-Ebene", offset := alr(i))
                    Constraints.AddFlush("XZ-Ebene" & i + 1, name_trwk(i).ToString, "XZ-Ebene", name_sklt(i).ToString, "XZ-Ebene", offset := 0)
                    
                ElseIf (typ(i) = "Typ10R") Then
                    oOcc = oAsmCompDef.Occurrences.Add(ThisDoc.Path & "..\..\Tragwerke\Typ10R-Tragwerk.iam", oMatrix)
                    name_trwk(i) = stat(i) & "-QP-" & nr(i) & "-Ausleger-Typ10R:" & i
                    oOcc.Name = name_trwk(i)
                    Constraints.AddMate("ht" & i + 1, name_trwk(i).ToString, "ht-X-Achse", name_sklt(i).ToString, "XY-Ebene", offset := ht(i))
                    Constraints.AddFlush("alr" & i + 1, name_trwk(i).ToString, "YZ-Ebene", name_sklt(i).ToString, "Gleisachse-YZ-Ebene", offset := alr(i))
                    Constraints.AddFlush("XZ-Ebene" & i + 1, name_trwk(i).ToString, "XZ-Ebene", name_sklt(i).ToString, "XZ-Ebene", offset := 0)
                    
                ElseIf (typ(i) = "Typ21A") Then
                    oOcc = oAsmCompDef.Occurrences.Add(ThisDoc.Path & "..\..\Tragwerke\Typ21A-Tragwerk.iam", oMatrix)
                    name_trwk(i) = stat(i) & "-QP-" & nr(i) & "-Ausleger-Typ21A:" & i
                    oOcc.Name = name_trwk(i)
                    Constraints.AddMate("ht" & i + 1, name_trwk(i).ToString, "ht-X-Achse", name_sklt(i).ToString, "XY-Ebene", offset := ht(i))
                    If (shift(i) = "true") Then
                        Constraints.AddFlush("alr" & i + 1, name_trwk(i).ToString, "YZ-Ebene", name_sklt(i).ToString, "Gleisachse-YZ-Ebene", offset := alr(i) * -1)
                        Constraints.AddFlush("XZ-Ebene" & i + 1, name_trwk(i).ToString, "XZ-Ebene", name_sklt(i).ToString, "XZ-Ebene", offset := 0)
                    Else
                        Constraints.AddMate("alr" & i + 1, name_trwk(i).ToString, "YZ-Ebene", name_sklt(i).ToString, "Gleisachse-YZ-Ebene", offset := alr(i) * -1)
                        Constraints.AddMate("XZ-Ebene" & i + 1, name_trwk(i).ToString, "XZ-Ebene", name_sklt(i).ToString, "XZ-Ebene", offset := 0)  
                    End If
                
                ElseIf (typ(i) = "Typ21R") Then
                    oOcc = oAsmCompDef.Occurrences.Add(ThisDoc.Path & "..\..\Tragwerke\Typ21R-Tragwerk.iam", oMatrix)
                    name_trwk(i) = stat(i) & "-QP-" & nr(i) & "-Ausleger-Typ21R:" & i
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
            Next
        End If
        
        If action = "Spurhalter" Or action = "Alle" Then
            '[ Generate position oriented parts: Spurhalter
            For i = n_min - 1 To n_max - 1
                ' Set a reference to the assembly component definition.
                Dim oAsmCompDef As AssemblyComponentDefinition
                oAsmCompDef = ThisApplication.ActiveDocument.ComponentDefinition
        
                ' Set a reference to the transient geometry object.
                Dim oTG As TransientGeometry
                oTG = ThisApplication.TransientGeometry
        
                ' Create a matrix.  A new matrix is initialized with an identity matrix.
                Dim oMatrix As Matrix = CreateMatrix(a(i), xe(i), xd(i), yn(i), yd(i), zz(i))
        
                ' Add the occurrence depending on typ(i).
                Dim sptr As ComponentOccurrence
                
                name_sklt(i) = stat(i) & "-QP-" & nr(i) & "-SKLT:" & i  
                
                If (typ(i) = "Typ10A") Then
                    sptr = oAsmCompDef.Occurrences.Add(ThisDoc.Path & "..\..\Tragwerke\Spurhalter.ipt", oMatrix)
                    name_sptr(i) = stat(i) & "-QP-" & nr(i) & "-Spurhalter:" & i
                    sptr.Name = name_sptr(i)
                    Constraints.AddFlush("sptr-XZ-Ebene" & i + 1, name_sptr(i).ToString, "XZ-Ebene", name_sklt(i).ToString, "XZ-Ebene", offset := 0)
                    Constraints.AddMate("sptr-ε-Y-Achse" & i + 1, name_sptr(i).ToString, "ε-Y-Achse", name_sklt(i).ToString, "ε-Y-Achse", offset := 0)
                    Constraints.AddMate("sptr-alr-Z" & i + 1, name_sptr(i).ToString, "Spurhalterabzug-Y-Achse", name_sklt(i).ToString, "XY-Ebene", offset := ht(i) -225)
                    
                ElseIf (typ(i) = "Typ10R") Then
                    sptr = oAsmCompDef.Occurrences.Add(ThisDoc.Path & "..\..\Tragwerke\Spurhalter.ipt", oMatrix)
                    name_sptr(i) = stat(i) & "-QP-" & nr(i) & "-Spurhalter:" & i
                    sptr.Name = name_sptr(i)
                    Constraints.AddMate("sptr-XZ-Ebene" & i + 1, name_sptr(i).ToString, "XZ-Ebene", name_sklt(i).ToString, "XZ-Ebene", offset := 0)
                    Constraints.AddMate("sptr-ε-Y-Achse" & i + 1, name_sptr(i).ToString, "ε-Y-Achse", name_sklt(i).ToString, "ε-Y-Achse", offset := 0)
                    Constraints.AddMate("sptr-alr-Z" & i + 1, name_sptr(i).ToString, "Spurhalterabzug-Y-Achse", name_sklt(i).ToString, "XY-Ebene", offset := ht(i) - 225)
                    
                ElseIf (typ(i) = "Typ21A") Then
                    sptr = oAsmCompDef.Occurrences.Add(ThisDoc.Path & "..\..\Tragwerke\Spurhalter.ipt", oMatrix)
                    name_sptr(i) = stat(i) & "-QP-" & nr(i) & "-Spurhalter:" & i
                    sptr.Name = name_sptr(i)
                    Constraints.AddFlush("sptr-XZ-Ebene" & i + 1, name_sptr(i).ToString, "XZ-Ebene", name_sklt(i).ToString, "XZ-Ebene", offset := 0)
                    Constraints.AddMate("sptr-ε-Y-Achse" & i + 1, name_sptr(i).ToString, "ε-Y-Achse", name_sklt(i).ToString, "ε-Y-Achse", offset := 0)
                    Constraints.AddMate("sptr-alr-Z" & i + 1, name_sptr(i).ToString, "Spurhalterabzug-Y-Achse", name_sklt(i).ToString, "XY-Ebene", offset := ht(i) -225)
                    
                ElseIf (typ(i) = "Typ21R") Then
                    sptr = oAsmCompDef.Occurrences.Add(ThisDoc.Path & "..\..\Tragwerke\Spurhalter.ipt", oMatrix)
                    name_sptr(i) = stat(i) & "-QP-" & nr(i) & "-Spurhalter:" & i
                    sptr.Name = name_sptr(i)
                    Constraints.AddMate("sptr-XZ-Ebene" & i + 1, name_sptr(i).ToString, "XZ-Ebene", name_sklt(i).ToString, "XZ-Ebene", offset := 0)
                    Constraints.AddMate("sptr-ε-Y-Achse" & i + 1, name_sptr(i).ToString, "ε-Y-Achse", name_sklt(i).ToString, "ε-Y-Achse", offset := 0)
                    Constraints.AddMate("sptr-alr-Z" & i + 1, name_sptr(i).ToString, "Spurhalterabzug-Y-Achse", name_sklt(i).ToString, "XY-Ebene", offset := ht(i) - 225)
                End If
                Logger.Info(name_sptr(i))
            Next
        End If
        
        If action = "Spurhalterabzug" Or action = "Alle" Then
            '[ Generate position oriented parts: Spurhalterabzug
            For i = n_min - 1 To n_max - 1
                ' Set a reference to the assembly component definition.
                Dim oAsmCompDef As AssemblyComponentDefinition
                oAsmCompDef = ThisApplication.ActiveDocument.ComponentDefinition
        
                ' Set a reference to the transient geometry object.
                Dim oTG As TransientGeometry
                oTG = ThisApplication.TransientGeometry
        
                ' Create a matrix.  A new matrix is initialized with an identity matrix.
                Dim oMatrix As Matrix = CreateMatrix(a(i), xe(i), xd(i), yn(i), yd(i), zz(i))
        
                ' Add the occurrence depending on typ(i).
                Dim shaz As ComponentOccurrence
                
                name_sklt(i) = stat(i) & "-QP-" & nr(i) & "-SKLT:" & i  
                
                If (typ(i) = "Typ10A" Or typ(i) = "Typ10R" Or typ(i) = "Typ21A" Or typ(i) = "Typ21R") Then
                    shaz = oAsmCompDef.Occurrences.Add(ThisDoc.Path & "..\..\Tragwerke\Spurhalterabzug.ipt", oMatrix)
                    name_shaz(i) = stat(i) & "-QP-" & nr(i) & "-Spurhalterabzug:" & i
                    name_sptr(i) = stat(i) & "-QP-" & nr(i) & "-Spurhalter:" & i
                    shaz.Name = name_shaz(i)
                    Constraints.AddFlush("shaz-XZ-Ebene"  & i + 1, name_shaz(i).ToString, "XZ-Ebene", name_sklt(i).ToString, "XZ-Ebene", offset := 0)
                    Constraints.AddMate("shaz-Y-Achse"   & i + 1, name_shaz(i).ToString, "Y-Achse", name_sptr(i).ToString, "Spurhalterabzug-Y-Achse", offset := 0)
                    Constraints.AddAngle("shaz-Winkel"    & i + 1, name_shaz(i).ToString, "Z-Achse", name_sklt(i).ToString, "Z-Achse", angle := 0.0)
        
                    Logger.Info(name_shaz(i))
                End If
            Next
        End If
        
        '[ Set instance values
        For i = n_min - 1 To n_max - 1
            If action = "Ausleger" Or action = "Alle" Then
                If (typ(i) = "Typ10A" Or typ(i) = "Typ10R" Or typ(i) = "Typ21A" Or typ(i) = "Typ21R") Then
                    ' iProperties.InstanceValue(name_trwk(i), "WIV_1 InstanzName")     = name_trwk(i)
                    ' iProperties.InstanceValue(name_trwk(i), "WIV_2 Typ")             = typ(i)
                    ' iProperties.InstanceValue(name_trwk(i), "WIV_3 QP")              = nr(i)
                    ' iProperties.InstanceValue(name_trwk(i), "WIV_4 Epsilon")         = ε(i)
                    iProperties.InstanceValue(name_trwk(i), "PTY_0 Kilometrierung")  = KM(i)
                    iProperties.InstanceValue(name_trwk(i), "PTY_0 Status")          = Status(i)
                    iProperties.InstanceValue(name_trwk(i), "PTY_0 Objektcode")      = "tl03"
                    iProperties.InstanceValue(name_trwk(i), "PTY_0 DIDOK")           = "HIGT"
                    iProperties.InstanceValue(name_trwk(i), "PTY_0 ObjekttypID")     = "OBJ_FS_18"
                    iProperties.InstanceValue(name_trwk(i), "PTY_0 ObjekttypNameDE") = "Ausleger (FS)"
                End If
            End If
            
            If action = "Spurhalter" Or action = "Alle" Then
                If (typ(i) = "Typ10A" Or typ(i) = "Typ10R" Or typ(i) = "Typ21A" Or typ(i) = "Typ21R") Then
                    iProperties.InstanceValue(name_sptr(i), "PTY_0 Status")          = Status(i)
                    iProperties.InstanceValue(name_sptr(i), "PTY_0 Objektcode")      = "tl03"
                    iProperties.InstanceValue(name_sptr(i), "PTY_0 DIDOK")           = "HIGT"
                    iProperties.InstanceValue(name_sptr(i), "PTY_0 ObjekttypID")     = "OBJ_FS_16"
                    iProperties.InstanceValue(name_sptr(i), "PTY_0 ObjekttypNameDE") = "Spurhalter"
                End If
            End If
            
            If action = "Spurhalterabzug" Or action = "Alle" Then
                If (typ(i) = "Typ10A" Or typ(i) = "Typ10R" Or typ(i) = "Typ21A" Or typ(i) = "Typ21R") Then
                    iProperties.InstanceValue(name_shaz(i), "PTY_0 Status")          = Status(i)
                    iProperties.InstanceValue(name_shaz(i), "PTY_0 Objektcode")      = "tl03"
                    iProperties.InstanceValue(name_shaz(i), "PTY_0 DIDOK")           = "HIGT"
                    iProperties.InstanceValue(name_shaz(i), "PTY_0 ObjekttypID")     = "OBJ_FS_106"
                    iProperties.InstanceValue(name_shaz(i), "PTY_0 ObjekttypNameDE") = "Spurhalterabzug"
                End If
            End If
        Next
        
        '[ End of routine
        Logger.Info("Update Model.")
        InventorVb.DocumentUpdate() 'Update model if run script
        iLogicVb.UpdateWhenDone = True
        ThisApplication.CommandManager.ControlDefinitions.Item("AssemblyRebuildAllCmd").Execute
        Logger.Info("Finished.")
    End Sub
    
    ' Function for selecting the action to be executed and for entering n_max
    Function SelectFDKType(defaultMax As Integer) As Tuple(Of String, Integer)
        Dim validFilterValues As String() = {"Bitte eine Auswahl treffen", "Ausleger", "Spurhalter", "Spurhalterabzug", "Alle"}
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
    
        textBox.Text = defaultMax.ToString()
        textBox.Left = 150
        textBox.Top = 70
        textBox.Width = 100
        form.Controls.Add(textBox)
    
        buttonOK.Text = "OK"
        buttonOK.Left = 50
        buttonOK.Top = 110
        buttonOK.Width = 80
        AddHandler buttonOK.Click, Sub(sender, e) form.DialogResult = DialogResult.OK
        form.Controls.Add(buttonOK)
    
        buttonCancel.Text = "Abbrechen"
        buttonCancel.Left = 150
        buttonCancel.Top = 110
        buttonCancel.Width = 80
        AddHandler buttonCancel.Click, Sub(sender, e) form.DialogResult = DialogResult.Cancel
        form.Controls.Add(buttonCancel)
    
        If form.ShowDialog() = DialogResult.OK Then
            If comboBox.SelectedItem IsNot Nothing AndAlso IsNumeric(textBox.Text) Then
                Dim selection As String = comboBox.SelectedItem.ToString()
                Dim n_max As Integer = Integer.Parse(textBox.Text)
                Return New Tuple(Of String, Integer)(selection, n_max)
            End If
        End If
    
        Return Nothing
    End Function
    
    Sub ImportData(n_max As Integer)
        x_red = GoExcel.CellValue("3rd Party:data", "data", "T" & 1)
        y_red = GoExcel.CellValue("3rd Party:data", "data", "U" & 1)
        z_red = GoExcel.CellValue("3rd Party:data", "data", "V" & 1)
    
        For i = n_min - 1 To n_max - 1
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
            KM(i)     = GoExcel.CellValue("3rd Party:data", "data", "AG" & i + r)               'Kilometrierung
            Status(i) = GoExcel.CellValue("3rd Party:data", "data", "AJ" & i + r)               'Status: Neu, Abbruch, Bestand, Referenz (Bestand - Abbruch) 
    
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
            ElseIf (Status(i) = "Abbruch") Then
                stat(i) = "A"
            ElseIf (Status(i) = "Neu") Then
                stat(i) = "N"
            Else
                stat(i) = "nicht definiert"
            End If
        Next
        Logger.Info("Done setup.")
    End Sub

    ' Function to create matrix with rotation and translation
    Function CreateMatrix(angle As Double, x_e As Double, x_d As Double, y_n As Double, y_d As Double, z_z As Double) As Matrix
        ' Set a reference to the transient geometry object.
        Dim oTG As TransientGeometry
        oTG = ThisApplication.TransientGeometry

        ' Create a matrix.  A new matrix is initialized with an identity matrix.
        Dim oMatrix As Matrix
        oMatrix = oTG.CreateMatrix
        
        ' Set the rotation of the matrix about the Z axis.
        If ((x_e > x_d And y_n > y_d) Or y_d < y_n) Then
            ' Check direction of rotation
            Call oMatrix.SetToRotation(-angle * PI / 180, oTG.CreateVector(0, 0, 1), oTG.CreatePoint(0, 0, 0)) 'rotate left  (-)
        Else
            Call oMatrix.SetToRotation(angle * PI / 180, oTG.CreateVector(0, 0, 1), oTG.CreatePoint(0, 0, 0))  'rotate right (+)
        End If

        ' Set the translation portion of the matrix so the part will be positioned
        Call oMatrix.SetTranslation(oTG.CreateVector(x_e / 10, y_n / 10, z_z / 10))  'distances in a matrix always defined in centimeters

        Return oMatrix
    End Function

End Class
