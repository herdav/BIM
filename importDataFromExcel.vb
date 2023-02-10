' -------------------------------------------------- '
' Import data from excel and generate oriented parts '
' -------------------------------------------------- '

Dim n = 79  'number of parts / datasets
Dim r = 5
Dim c = 8 'number of columns !!
Dim x_red, y_red, z_red 'reduce number space
Dim name_bks(n), name_dir(n), position_bks(n), position_dir(n)
Dim f = 1000 'convert data to mm

'n     = GoExcel.CellValue("3rd Party:data", "data", "B" & 1)
x_red = GoExcel.CellValue("3rd Party:data", "data", "Q" & 1)
y_red = GoExcel.CellValue("3rd Party:data", "data", "R" & 1)
z_red = GoExcel.CellValue("3rd Party:data", "data", "S" & 1)

Logger.Info("Datasets: " & n)

Dim nr(n), typ(n), xe(n), yn(n), zz(n), xd(n), yd(n), a(n), rot(n)
Dim tempX(n), tempY(n)

Dim generateParts = False

For i = 0 To n - 1
	'[ Load data from excel an safe it in array
	nr(i)  =  GoExcel.CellValue("3rd Party:data", "data", "C" & i + r) 				'Nr
	typ(i) =  GoExcel.CellValue("3rd Party:data", "data", "E" & i + r) 	 			'Typ
	xe(i)  = (GoExcel.CellValue("3rd Party:data", "data", "Q" & i + r) - x_red) * f 'X(E)-Achse Fundament
	yn(i)  = (GoExcel.CellValue("3rd Party:data", "data", "R" & i + r) - y_red) * f 'Y(N)-Achse Fundament
	zz(i)  = (GoExcel.CellValue("3rd Party:data", "data", "S" & i + r) - z_red) * f 'Z(Z)-Achse Fundament
	xd(i)  = (GoExcel.CellValue("3rd Party:data", "data", "T" & i + r) - x_red) * f 'X(E)-Achse Ausrichtung
	yd(i)  = (GoExcel.CellValue("3rd Party:data", "data", "U" & i + r) - y_red) * f 'Y(N)-Achse Ausrichtung
	rot(i) =  GoExcel.CellValue("3rd Party:data", "data", "V" & i + r) 	 			'Rotate 180grd
	']
	
	If (rot(i) = "true") 'Switch bks and dir
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
	Components.Add(name_bks(i), "bks.ipt", position_bks(i))
	']

	'[ Generate direction-parts in modell / assembly
	name_dir(i) = "dir:" & nr(i)
	position_dir(i) = ThisAssembly.Geometry.Point(xd(i), yd(i), zz(i))
	Components.Add(name_dir(i), "dir.ipt", position_dir(i))
	']
	
	'[ Calculate direction between bks(i) and dir(i)
	a(i) = Measure.Angle(name_dir(i), "center", name_bks(i), "center", name_bks(i), "direction")
	'Logger.Info(a(i).ToString)
	']
	
Next
Logger.Info("Direction parts generated.")

For i = 0 To n - 1
	'[ Position oriented parts.
	    ' Set a reference to the assembly component definintion.
	    ' This assumes an assembly document is open.
		
	    Dim oAsmCompDef As AssemblyComponentDefinition
	    oAsmCompDef = ThisApplication.ActiveDocument.ComponentDefinition
	
	    ' Set a reference to the transient geometry object.
	    Dim oTG As TransientGeometry
	    oTG = ThisApplication.TransientGeometry
	
	    ' Create a matrix.  A new matrix is initialized with an identity matrix.
	    Dim oMatrix As Matrix
	    oMatrix = oTG.CreateMatrix
		
	    ' Set the rotation of the matrix about the Z axis.
		If (xe(i) > xd(i) And yn(i) > yd(i)) 'Check direction of rotation
			Call oMatrix.SetToRotation(-a(i) * PI / 180, oTG.CreateVector(0, 0, 1), oTG.CreatePoint(0, 0, 0)) 'rotate left  (-)
		Else
			If (yd(i) < yn(i))
				Call oMatrix.SetToRotation(-a(i) * PI / 180, oTG.CreateVector(0, 0, 1), oTG.CreatePoint(0, 0, 0)) 'rotate left  (-)
			Else
				Call oMatrix.SetToRotation(a(i) * PI / 180, oTG.CreateVector(0, 0, 1), oTG.CreatePoint(0, 0, 0))  'rotete right (+)
			End If
		End If
	
	    ' Set the translation portion of the matrix so the part will be positioned
	    Call oMatrix.SetTranslation(oTG.CreateVector(xe(i)/10, yn(i)/10, zz(i)/10))  'distances in a matrix always defined in centimeters
	
	    ' Add the occurrence depending on typ(i).
	    Dim oOcc As ComponentOccurrence
		If (typ(i) = "A")
			oOcc = oAsmCompDef.Occurrences.Add("C:\Users\Public\VWS\Data\Mitarbeiter Ing Büro\Wi&Vi AG\Herren\Projekte\2201.006_SBB_GPL_Stadelhofen_und_Riesbachtunnel\Konzept\prt-A.ipt", oMatrix)
		ElseIf (typ(i) = "B")
			oOcc = oAsmCompDef.Occurrences.Add("C:\Users\Public\VWS\Data\Mitarbeiter Ing Büro\Wi&Vi AG\Herren\Projekte\2201.006_SBB_GPL_Stadelhofen_und_Riesbachtunnel\Konzept\prt-B.ipt", oMatrix)
		Else
			oOcc = oAsmCompDef.Occurrences.Add("C:\Users\Public\VWS\Data\Mitarbeiter Ing Büro\Wi&Vi AG\Herren\Projekte\2201.006_SBB_GPL_Stadelhofen_und_Riesbachtunnel\Hirschgrabentunnel\typ-A.iam", oMatrix)
		End If
		
		Logger.Info("Object " & i + 1 & " - Nr: " & nr(i) & " as " & typ(i) & " generated.")
	']
Next

InventorVb.DocumentUpdate() 'Update modell if run script

Logger.Info("Finished.")
