' -------------------------------------------------- '
' Import data from excel and generate oriented parts '
'   Version 23.05.2023 by David Herren @ WiVi AG
' -------------------------------------------------- '

'[ Setup
Dim n = 100          'number of parts / datasets
Dim n_min = 1
Dim n_max = 100
Dim r = 4            'number of 1st row
Dim c = 8            'number of columns !!
Dim f = 1000         'convert data to mm for vectors
Dim name_bks(n), name_dir(n), position_bks(n), position_dir(n), name_sklt(n)
Dim x_red = GoExcel.CellValue("3rd Party:data", "data", "R" & 1) 'reduced number space
Dim y_red = GoExcel.CellValue("3rd Party:data", "data", "S" & 1)
Dim z_red = GoExcel.CellValue("3rd Party:data", "data", "T" & 1)
Dim nr(n), typ(n), xe(n), yn(n), zz(n), xd(n), yd(n), a(n), switch(n)
Dim tempX(n), tempY(n)
Dim hp_sp(n), hf(n), δ(n), ε(n), lmr(n), hp(n), α(n)
Logger.Info("Datasets: " & n)
']

'[ Generate direction parts from excel
For i = n_min - 1 To n_max - 1 'Load data from excel an safe it in an array
	nr(i)     =  GoExcel.CellValue("3rd Party:data", "data", "C" & i + r)              'Nr
	typ(i)    =  GoExcel.CellValue("3rd Party:data", "data", "E" & i + r)              'Typ
	xe(i)     = (GoExcel.CellValue("3rd Party:data", "data", "R" & i + r) - x_red) * f 'X(E)-Achse Fundament
	yn(i)     = (GoExcel.CellValue("3rd Party:data", "data", "S" & i + r) - y_red) * f 'Y(N)-Achse Fundament
	zz(i)     = (GoExcel.CellValue("3rd Party:data", "data", "T" & i + r) - z_red) * f 'Z(Z)-Achse Fundament
	xd(i)     = (GoExcel.CellValue("3rd Party:data", "data", "U" & i + r) - x_red) * f 'X(E)-Achse Ausrichtung
	yd(i)     = (GoExcel.CellValue("3rd Party:data", "data", "V" & i + r) - y_red) * f 'Y(N)-Achse Ausrichtung
	switch(i) =  GoExcel.CellValue("3rd Party:data", "data", "W" & i + r)              'Switch bks and dir
	hp_sp(i)  =  GoExcel.CellValue("3rd Party:data", "data", "H" & i + r)              'Höhe zwischen Tragrohrachse und Gleis
	hf(i)     =  GoExcel.CellValue("3rd Party:data", "data", "F" & i + r)              'Höhe zwischen Fahrdraht und Gleis
	δ(i)      =  GoExcel.CellValue("3rd Party:data", "data", "L" & i + r)              'Versatz Gleisachse
	ε(i)      =  GoExcel.CellValue("3rd Party:data", "data", "M" & i + r)              'Versatz Fahrzeugachse
	lmr(i)    =  GoExcel.CellValue("3rd Party:data", "data", "X" & i + r)              'Aufhängung Links, Mitte oder Rechts des Gleises
	hp(i)     =  GoExcel.CellValue("3rd Party:data", "data", "G" & i + r)              'Höhe zwischen Tragseil und Gleis
	α(i)      =  GoExcel.CellValue("3rd Party:data", "data", "K" & i + r)              'Neigung Gleisachse
	
	If (switch(i) = "true") 'Switch bks and dir
		tempX(i) = xd(i)
		tempY(i) = yd(i)
		xd(i) = xe(i)
		yd(i) = yn(i)
		xe(i) = tempX(i)
		yn(i) = tempY(i)
	End If

	'[ Generate bks-parts in modell / assembly
	name_bks(i) = "bks:" & nr(i)
	position_bks(i) = ThisAssembly.Geometry.Point(xe(i), yn(i), zz(i))
	Components.Add(name_bks(i), "bks.ipt", position_bks(i), True, False)
	']

	'[ Generate direction-parts in modell / assembly
	name_dir(i) = "dir:" & nr(i)
	position_dir(i) = ThisAssembly.Geometry.Point(xd(i), yd(i), zz(i))
	Components.Add(name_dir(i), "dir.ipt", position_dir(i), True, False)
	']
	
	'[ Calculate direction between bks(i) and dir(i)
	a(i) = Measure.Angle(name_dir(i), "center", name_bks(i), "center", name_bks(i), "direction")
	'Logger.Info(a(i).ToString)
	']
Next
Logger.Info("Direction parts generated.")
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
	
	If (typ(i) = "Typ10")
		oOcc = oAsmCompDef.Occurrences.Add(ThisDoc.Path & "\sklt.iam", oMatrix)
		name_sklt(i) = i + 1 & "-sklt-Typ10:" & nr(i)
		oOcc.Name = name_sklt(i)
	Else If (typ(i) = "Typ21")
		oOcc = oAsmCompDef.Occurrences.Add(ThisDoc.Path & "\sklt.iam", oMatrix)
		name_sklt(i) = i + 1 & "-sklt-Typ21:" & nr(i)
		oOcc.Name = name_sklt(i)
	Else If (typ(i) = "Typ31")			
		oOcc = oAsmCompDef.Occurrences.Add(ThisDoc.Path & "\sklt.iam", oMatrix)
		name_sklt(i) = i + 1 & "-sklt-Typ31:" & nr(i)
		oOcc.Name = name_sklt(i)
	Else If (typ(i) = "Typ82")				
		oOcc = oAsmCompDef.Occurrences.Add(ThisDoc.Path & "\sklt.iam", oMatrix)
		name_sklt(i) = i + 1 & "-sklt-Typ82:" & nr(i)
		oOcc.Name = name_sklt(i)
	Else
		oOcc = oAsmCompDef.Occurrences.Add(ThisDoc.Path & "\sklt.iam", oMatrix)
		name_sklt(i) = i + 1 & "-sklt-Andere:" & nr(i)
		oOcc.Name = name_sklt(i)		
	End If
	
	oOcc.Flexible = True
	oOcc.Grounded = True

	Logger.Info(name_sklt(i))
Next
']

'[ Set constraints
For i = n_min - 1 To n_max - 1
	Constraints.AddMate("soll" & i + 1, name_dir(i).ToString, "Z-Achse", {name_sklt(i), "sklt-dir:1"}, "Z-Achse",
	    offset := 0.0, e1InferredType := InferredTypeEnum.kNoInference, e2InferredType := InferredTypeEnum.kNoInference,
	    solutionType := MateConstraintSolutionTypeEnum.kNoSolutionType,
	    biasPoint1 := Nothing, biasPoint2 := Nothing)
	Constraints.AddFlush("hf" & i + 1, name_dir(i).ToString, "XY-Ebene", {name_sklt(i), "sklt-hf:1" }, "XY-Ebene",
		offset := hf(i), biasPoint1 := Nothing, biasPoint2 := Nothing)
	If (lmr(i) = "R")
		Constraints.AddMate("δ" & i + 1, name_dir(i).ToString, "Z-Achse", {name_sklt(i), "sklt-delta:1" }, "YZ-Ebene",
			offset := δ(i), e1InferredType := InferredTypeEnum.kNoInference, e2InferredType := InferredTypeEnum.kNoInference,
			solutionType := MateConstraintSolutionTypeEnum.kNoSolutionType,
			biasPoint1 := Nothing, biasPoint2 := Nothing)
		Constraints.AddAngle("α" & i + 1, name_bks(i).ToString, "Z-Achse", {name_sklt(i), "lichtraumprofil:1" }, "Z-Achse",
			angle := α(i) * 180 / PI * -1, solutionType := AngleConstraintSolutionTypeEnum.kDirectedSolution, 
            refVecComponent := Nothing, refEntityName := Nothing, biasPoint1 := Nothing, biasPoint2 := Nothing)
		If (α(i) < 0) 'R & α(i) < 0
			Constraints.AddMate("ε" & i + 1, name_dir(i).ToString, "Z-Achse", {name_sklt(i), "sklt-epsilon:1" }, "YZ-Ebene",
				offset := δ(i) + ε(i), e1InferredType := InferredTypeEnum.kNoInference, e2InferredType := InferredTypeEnum.kNoInference,
				solutionType := MateConstraintSolutionTypeEnum.kNoSolutionType,
				biasPoint1 := Nothing, biasPoint2 := Nothing)
		Else 'R & α(i) >= 0
			If (ε(i) >= 0)
				Constraints.AddMate("ε" & i + 1, name_dir(i).ToString, "Z-Achse", {name_sklt(i), "sklt-epsilon:1" }, "YZ-Ebene",
					offset := δ(i) + ε(i), e1InferredType := InferredTypeEnum.kNoInference, e2InferredType := InferredTypeEnum.kNoInference,
					solutionType := MateConstraintSolutionTypeEnum.kNoSolutionType,
					biasPoint1 := Nothing, biasPoint2 := Nothing)
			Else
				Constraints.AddMate("ε" & i + 1, name_dir(i).ToString, "Z-Achse", {name_sklt(i), "sklt-epsilon:1" }, "YZ-Ebene",
					offset := δ(i) * -1 - ε(i), e1InferredType := InferredTypeEnum.kNoInference, e2InferredType := InferredTypeEnum.kNoInference,
					solutionType := MateConstraintSolutionTypeEnum.kNoSolutionType,
					biasPoint1 := Nothing, biasPoint2 := Nothing)
			End If
		End If
	Else
		Constraints.AddMate("δ" & i + 1, name_dir(i).ToString, "Z-Achse", {name_sklt(i), "sklt-delta:1" }, "YZ-Ebene",
			offset := δ(i) * -1, e1InferredType := InferredTypeEnum.kNoInference, e2InferredType := InferredTypeEnum.kNoInference,
			solutionType := MateConstraintSolutionTypeEnum.kNoSolutionType,
			biasPoint1 := Nothing, biasPoint2 := Nothing)
		Constraints.AddAngle("α" & i + 1, name_bks(i).ToString, "Z-Achse", {name_sklt(i), "lichtraumprofil:1" }, "Z-Achse",
			angle := α(i) * 180 / PI, solutionType := AngleConstraintSolutionTypeEnum.kDirectedSolution, 
            refVecComponent := Nothing, refEntityName := Nothing, biasPoint1 := Nothing, biasPoint2 := Nothing)
		If (α(i) < 0) 'L & α(i) < 0
			Constraints.AddMate("ε" & i + 1, name_dir(i).ToString, "Z-Achse", {name_sklt(i), "sklt-epsilon:1" }, "YZ-Ebene",
				offset := δ(i) * -1 - ε(i), e1InferredType := InferredTypeEnum.kNoInference, e2InferredType := InferredTypeEnum.kNoInference,
				solutionType := MateConstraintSolutionTypeEnum.kNoSolutionType,
				biasPoint1 := Nothing, biasPoint2 := Nothing)
		Else 'L & α(i) >= 0
			Constraints.AddMate("ε" & i + 1, name_dir(i).ToString, "Z-Achse", {name_sklt(i), "sklt-epsilon:1" }, "YZ-Ebene",
				offset := δ(i) * -1 - ε(i), e1InferredType := InferredTypeEnum.kNoInference, e2InferredType := InferredTypeEnum.kNoInference,
				solutionType := MateConstraintSolutionTypeEnum.kNoSolutionType,
				biasPoint1 := Nothing, biasPoint2 := Nothing)
		End If	
	End If
	Constraints.AddFlush("hp" & i + 1, name_dir(i).ToString, "XY-Ebene", {name_sklt(i), "sklt-hp:1" }, "XY-Ebene",
		offset := hp(i), biasPoint1 := Nothing, biasPoint2 := Nothing)
Next	
']

'[ Set instance values
For i = n_min - 1 To n_max - 1
	iProperties.InstanceValue(name_sklt(i), "name") = name_sklt(i)
	iProperties.InstanceValue(name_sklt(i), "typ")  = typ(i)
	iProperties.InstanceValue(name_sklt(i), "nr")   = nr(i)
	iProperties.InstanceValue(name_sklt(i), "ε")    = ε(i)
	iProperties.InstanceValue(name_sklt(i), "hf")   = hf(i)
Next
']

'[ End of rutine
InventorVb.DocumentUpdate() 'Update modell if run script
Logger.Info("Finished.")
']
