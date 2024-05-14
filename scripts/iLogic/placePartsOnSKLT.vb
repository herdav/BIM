'Place Parts on SKLT
'Version 14.05.2024 by David Herren @ WiVi AG

'[ Setup
Dim n = 100          'number of parts / datasets
Dim n_min = 1        'set min range
Dim n_max = 100	     'set max range
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
For i = n_min - 1 To n_max - 1 'Load data from excel an safe it in an array
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
    ' Set a reference to the assembly component definintion.
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
		Call oMatrix.SetToRotation(a(i) * PI / 180, oTG.CreateVector(0, 0, 1), oTG.CreatePoint(0, 0, 0))  'rotete right (+)
	End If

    ' Set the translation portion of the matrix so the part will be positioned
    Call oMatrix.SetTranslation(oTG.CreateVector(xe(i)/10, yn(i)/10, zz(i)/10))  'distances in a matrix always defined in centimeters

    ' Add the occurrence depending on typ(i).
    Dim oOcc As ComponentOccurrence
	
	name_sklt(i) = stat(i) & "-QP-" & nr(i) & "-SKLT:" & i
	
	If (typ(i) = "Typ10A")
		oOcc = oAsmCompDef.Occurrences.Add(ThisDoc.Path & "..\..\Tragwerke\Typ10A-Tragwerk.iam", oMatrix)
		name_trwk(i) = stat(i) & "-QP-" & nr(i) & "-Ausleger-Typ10A:" & i
		oOcc.Name = name_trwk(i)
		Constraints.AddMate("ht" & i + 1, name_trwk(i).ToString, "ht-X-Achse", name_sklt(i).ToString, "XY-Ebene", offset := ht(i))
		Constraints.AddFlush("alr" & i + 1, name_trwk(i).ToString, "YZ-Ebene", name_sklt(i).ToString, "Gleisachse-YZ-Ebene", offset := alr(i))
		Constraints.AddFlush("XZ-Ebene" & i + 1, name_trwk(i).ToString, "XZ-Ebene", name_sklt(i).ToString, "XZ-Ebene", offset := 0)
		
	Else If (typ(i) = "Typ10R")
		oOcc = oAsmCompDef.Occurrences.Add(ThisDoc.Path & "..\..\Tragwerke\Typ10R-Tragwerk.iam", oMatrix)
		name_trwk(i) = stat(i) & "-QP-" & nr(i) & "-Ausleger-Typ10R:" & i
		oOcc.Name = name_trwk(i)
		Constraints.AddMate ("ht"         & i + 1, name_trwk(i).ToString, "ht-X-Achse", name_sklt(i).ToString, "XY-Ebene",            offset := ht(i))
		Constraints.AddFlush("alr"          & i + 1, name_trwk(i).ToString, "YZ-Ebene",   name_sklt(i).ToString, "Gleisachse-YZ-Ebene", offset := alr(i))
		Constraints.AddFlush("XZ-Ebene"   & i + 1, name_trwk(i).ToString, "XZ-Ebene",   name_sklt(i).ToString, "XZ-Ebene",            offset := 0)
		
	Else If (typ(i) = "Typ21A")
		oOcc = oAsmCompDef.Occurrences.Add(ThisDoc.Path & "..\..\Tragwerke\Typ21A-Tragwerk.iam", oMatrix)
		name_trwk(i) = stat(i) & "-QP-" & nr(i) & "-Ausleger-Typ21A:" & i
		oOcc.Name = name_trwk(i)
		Constraints.AddMate ("ht"           & i + 1, name_trwk(i).ToString, "ht-X-Achse", name_sklt(i).ToString, "XY-Ebene", offset := ht(i))
		If (shift(i) = "true")
			Constraints.AddFlush("alr"        & i + 1, name_trwk(i).ToString, "YZ-Ebene", name_sklt(i).ToString, "Gleisachse-YZ-Ebene", offset := alr(i) * -1)
			Constraints.AddFlush("XZ-Ebene" & i + 1, name_trwk(i).ToString, "XZ-Ebene", name_sklt(i).ToString, "XZ-Ebene", offset := 0)
		Else
            Constraints.AddMate("alr"         & i + 1, name_trwk(i).ToString, "YZ-Ebene", name_sklt(i).ToString, "Gleisachse-YZ-Ebene", offset := alr(i) * -1)
			Constraints.AddMate("XZ-Ebene"  & i + 1, name_trwk(i).ToString, "XZ-Ebene", name_sklt(i).ToString, "XZ-Ebene", offset := 0)	
		End If
	
	Else If (typ(i) = "Typ21R")
		oOcc = oAsmCompDef.Occurrences.Add(ThisDoc.Path & "..\..\Tragwerke\Typ21R-Tragwerk.iam", oMatrix)
		name_trwk(i) = stat(i) & "-QP-" & nr(i) & "-Ausleger-Typ21R:" & i
		oOcc.Name = name_trwk(i)
		Constraints.AddMate ("ht"           & i + 1, name_trwk(i).ToString, "ht-X-Achse", name_sklt(i).ToString, "XY-Ebene", offset := ht(i))
		If (shift(i) = "true")	
			Constraints.AddFlush("alr"        & i + 1, name_trwk(i).ToString, "YZ-Ebene", name_sklt(i).ToString,   "Gleisachse-YZ-Ebene", offset := alr(i) * -1)
			Constraints.AddFlush("XZ-Ebene" & i + 1, name_trwk(i).ToString, "XZ-Ebene", name_sklt(i).ToString,   "XZ-Ebene", offset := 0)
		Else
			Constraints.AddMate("alr"         & i + 1, name_trwk(i).ToString, "YZ-Ebene", name_sklt(i).ToString,   "Gleisachse-YZ-Ebene", offset := alr(i) * -1)
			Constraints.AddMate("XZ-Ebene"  & i + 1, name_trwk(i).ToString, "XZ-Ebene", name_sklt(i).ToString,   "XZ-Ebene", offset := 0)
		End If
	End If
	Logger.Info(name_trwk(i))
Next
']

'[ Set instance values
For i = n_min - 1 To n_max - 1
	If (typ(i) = "Typ10A" Or typ(i) = "Typ10R" Or typ(i) = "Typ21A" Or typ(i) = "Typ21R") '& " xx" > makes sure iProperties is a text
'		iProperties.InstanceValue(name_trwk(i), "WIV_1 InstanzName")     = name_trwk(i)
'		iProperties.InstanceValue(name_trwk(i), "WIV_2 Typ")             = typ(i)
'		iProperties.InstanceValue(name_trwk(i), "WIV_3 QP")              = nr(i)
'		iProperties.InstanceValue(name_trwk(i), "WIV_4 Epsilon")         = ε(i)
'		iProperties.InstanceValue(name_trwk(i), "PTY_0 Kilometrierung")  = KM(i)
		iProperties.InstanceValue(name_trwk(i), "PTY_0 Status")          = Status(i)
		iProperties.InstanceValue(name_trwk(i), "PTY_0 Objektcode")      = "tl03"
		iProperties.InstanceValue(name_trwk(i), "PTY_0 DIDOK")           = "HIGT"
		iProperties.InstanceValue(name_trwk(i), "PTY_0 ObjekttypID")     = "OBJ_FS_18"
		iProperties.InstanceValue(name_trwk(i), "PTY_0 ObjekttypNameDE") = "Ausleger (FS)"

	End If
Next
']

'[ End of rutine
Logger.Info("Update Model.")
InventorVb.DocumentUpdate() 'Update modell if run script
iLogicVb.UpdateWhenDone = True
ThisApplication.CommandManager.ControlDefinitions.Item("AssemblyRebuildAllCmd").Execute
Logger.Info("Finished.")
']
