'Import data from excel and generate oriented parts
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
	']

	'[ Generate bks-parts in modell / assembly
	name_bks(i) = stat(i) & "-QP-" & nr(i) & "-BKS:" & i
	
	position_bks(i) = ThisAssembly.Geometry.Point(xe(i), yn(i), zz(i))
	Components.Add(name_bks(i), ThisDoc.Path & "..\..\SKLT\bks.ipt", position_bks(i), True, False)
	']

	'[ Generate direction-parts in modell / assembly
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
	If (lmr(i) = "R")
		Constraints.AddMate("δ" & i + 1, name_dir(i).ToString, "Z-Achse", {name_sklt(i), "sklt-delta:1" }, "YZ-Ebene",
			offset := δ(i), e1InferredType := InferredTypeEnum.kNoInference, e2InferredType := InferredTypeEnum.kNoInference,
			solutionType := MateConstraintSolutionTypeEnum.kNoSolutionType)
		Constraints.AddAngle("α" & i + 1, name_bks(i).ToString, "Z-Achse", {name_sklt(i), "lichtraumprofil:1" }, "Z-Achse",
			angle := α(i) * 180 / PI * -1, solutionType := AngleConstraintSolutionTypeEnum.kDirectedSolution)
		If (α(i) < 0) 'R & α(i) < 0
			Constraints.AddMate("ε" & i + 1, name_dir(i).ToString, "Z-Achse", {name_sklt(i), "sklt-epsilon:1" }, "YZ-Ebene",
				offset := δ(i) + ε(i), e1InferredType := InferredTypeEnum.kNoInference, e2InferredType := InferredTypeEnum.kNoInference,
				solutionType := MateConstraintSolutionTypeEnum.kNoSolutionType)
		Else 'R & α(i) >= 0
			If (ε(i) >= 0)
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
		If (α(i) < 0) 'L & α(i) < 0
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

'[ End of rutine
InventorVb.DocumentUpdate() 'Update modell if run script
Logger.Info("Finished.")
']
