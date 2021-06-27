REM  *****  BASIC  *****

Sub OptVersMacro
Dim NBe as integer 'Nb Etudiants
Dim Nbs as integer 'Nb Terrains de stages
Dim Doc As Object
Dim Sheet As Object
Dim CellRangeAddress As New com.sun.star.table.CellRangeAddress
Dim CellAddress As New com.sun.star.table.CellAddress
Dim Cell As object
Dim oCurseur As Object 
dim feuille_dest as object
dim cell_dest as integer
dim nbl as integer
dim nbc as integer


MsgBox "La Macro va s'initialiser, pour cela il faut que vous ayez rempli les feuilles Terrains de stage et Liste etudiants correctement"
	
  
	'Définit la 1ere feuille 
	oCurseur = ThisComponent.Sheets(0).createCursor  
	oCurseur.gotoEndOfUsedArea( False ) 
  
	'L'index de la première ligne = 0  
	NBe = oCurseur.RangeAddress.EndRow

	'Définit la 2eme feuille 
	oCurseur = ThisComponent.Sheets(1).createCursor  
	oCurseur.gotoEndOfUsedArea( False ) 
  
	'L'index de la première colonne = 0  
	NBs = oCurseur.RangeAddress.EndColumn-1

'Fin d'initialisation
Doc = ThisComponent

''''''''Feuille Coefficient Période X''''''''

Dim Nb_periode as integer
Dim Nb_periode_Max as integer
    NB_periode_Max = InputBox ("Veuillez entrer le nombre de périodes : ","Chère utilisatrice, cher utilisateur")
    MsgBox ( Nb_periode_Max , 64, "Confirmation de formule")


For Nb_periode=1 to Nb_periode_Max

	'Ajoute une feuille, la nomme et place l'onglet en 3eme position 
	Doc.GetSheets.insertNewByName("Coef " & Nb_periode,Nb_periode+1)
	
 'Commande suivante devant être executé depuis la feuille : erreur depuis la console
 
Sheet = Doc.Sheets(Nb_periode+1)
Dim Range as object
		  
		Range = Sheet.getCellRangeByPosition( 0 , 0 , Nbs+3 , 0 ) 
		Range.Merge( True ) 

		Cell = Sheet.getCellByPosition(0,0)   
		With Cell 
		  .setString( "Notes Brutes" ) 'insére du texte dans la cellule 
		  .CellBackColor = RGB(255,215,0) 'indique la couleur de fond 
		  '.paraAdjust = com.sun.star.style.ParagraphAdjust.CENTER 'alignement centré 
		  '.RotateAngle = 9000 'Rotation 9000 = 90° 
		  .CharColor = RGB(0,0,0) 'couleur des caractères 
		  .CharHeight = 14 'Taille catactères 
		  .CharWeight = com.sun.star.awt.FontWeight.BOLD 'gras 
		  .CharPosture = com.sun.star.awt.FontSlant.ITALIC 'italique 
		  .CharFontName = "Arial" 'Font 
		 ' .CharUnderline = com.sun.star.awt.FontUnderline.DOUBLE 'souligné double 
		End With 
		  
		  
		'Ajuste la hauteur de la 5eme ligne au contenu des cellules.  
		Sheet.Rows(0).OptimalHeight = True 
		
 'Formule des étudiants pour brut
 Sheet = Doc.Sheets(Nb_periode+1)
	Cell = Sheet.getCellByPosition(0,1)
		Cell.formula = "=$'Liste Etudiants'.A1"
	Range = Sheet.getCellRangeByPosition(0,1,2,nbe+1)
		Range.fillSeries(com.sun.star.sheet.FillDirection.TO_BOTTOM,com.sun.star.sheet.FillMode.SIMPLE,0,0,0)
		Range.fillSeries(com.sun.star.sheet.FillDirection.TO_RIGHT,com.sun.star.sheet.FillMode.SIMPLE,0,0,0)
Cell = Sheet.getCellByPosition(3,1) 
Cell.String = "Coefficients" 		
'formule des terrains pour brut
 
 Sheet = Doc.Sheets(Nb_periode+1)
	Cell = Sheet.getCellByPosition(4,1)
		Cell.formula = "=$'Place Stage'.C1"
	Range = Sheet.getCellRangeByPosition(4,1,3+Nbs,1)
		Range.fillSeries(com.sun.star.sheet.FillDirection.TO_RIGHT,com.sun.star.sheet.FillMode.SIMPLE,0,0,0)


	'Définit la 1ere feuille 
	oCurseur = Doc.Sheets(Nb_periode+1).createCursor  
	oCurseur.gotoEndOfUsedArea( False ) 
  
	'L'index de la première ligne = 0  
	nbl = oCurseur.RangeAddress.EndRow+1
  
	'L'index de la première colonne = 0  
	nbc = oCurseur.RangeAddress.EndColumn+1
	

	Sheet = Doc.Sheets(Nb_periode+1)
	Range = Sheet.getCellRangeByPosition(4,1,3+Nbs,1)   
		With Range 
		  '.setString( "Notes Bruts" ) 'insére du texte dans la cellule 
		  .CellBackColor = RGB(218,165,32) 'indique la couleur de fond 
		  .paraAdjust = com.sun.star.style.ParagraphAdjust.CENTER 'alignement centré 
		  '.RotateAngle = 9000 'Rotation 9000 = 90° 
		  .CharColor = RGB(0,0,0) 'couleur des caractères 
		  .CharHeight = 10 'Taille catactères 
		  '.CharWeight = com.sun.star.awt.FontWeight.BOLD 'gras 
		  '.CharPosture = com.sun.star.awt.FontSlant.ITALIC 'italique 
		  .CharFontName = "Arial" 'Font 
		 ' .CharUnderline = com.sun.star.awt.FontUnderline.DOUBLE 'souligné double 
		End With 
	
	Sheet.Rows(1).OptimalHeight = True 
		
	Sheet = Doc.Sheets(Nb_periode+1)
	Range = Sheet.getCellRangeByPosition(0,2,3,NBe+1)   
		With Range 
		  '.setString( "Notes Bruts" ) 'insére du texte dans la cellule 
		  .CellBackColor = RGB(240,248,255) 'indique la couleur de fond 
		  '.paraAdjust = com.sun.star.style.ParagraphAdjust.CENTER 'alignement centré 
		  '.RotateAngle = 9000 'Rotation 9000 = 90° 
		  .CharColor = RGB(0,0,0) 'couleur des caractères 
		  .CharHeight = 10 'Taille catactères 
		  '.CharWeight = com.sun.star.awt.FontWeight.BOLD 'gras 
		  .CharPosture = com.sun.star.awt.FontSlant.ITALIC 'italique 
		  .CharFontName = "Arial" 'Font 
		 ' .CharUnderline = com.sun.star.awt.FontUnderline.DOUBLE 'souligné double 
		End With 
		Dim numl as integer
	For numl =1 to NBe-1
	Sheet.Rows(1+numl).OptimalHeight = True
	Next numl

	For numl=0 to 3
		Sheet.columns(numl).OptimalWidth = True 
	Next numl
		

'copie tableau normal
Sheet = Doc.Sheets(Nb_periode+1)


Plage = Sheet.getCellRangeByPosition( 0 , nbl , Nbs+3 , nbl ) 
		Plage.Merge( True ) 
		Cell = Sheet.getCellByPosition(0,nbl)  
		With Cell 
		  .setString( "Notes Normalisées" ) 'insére du texte dans la cellule 
		  .CellBackColor = RGB(255,215,0) 'indique la couleur de fond 
		  '.paraAdjust = com.sun.star.style.ParagraphAdjust.CENTER 'alignement centré 
		  '.RotateAngle = 9000 'Rotation 9000 = 90° 
		  .CharColor = RGB(0,0,0) 'couleur des caractères 
		  .CharHeight = 14 'Taille catactères 
		  .CharWeight = com.sun.star.awt.FontWeight.BOLD 'gras 
		  .CharPosture = com.sun.star.awt.FontSlant.ITALIC 'italique 
		  .CharFontName = "Arial" 'Font 
		 ' .CharUnderline = com.sun.star.awt.FontUnderline.DOUBLE 'souligné double 
		End With  
		Sheet.Rows(nbl).OptimalHeight = True 
		
		
 Sheet = Doc.Sheets(Nb_periode+1)
	Cell = Sheet.getCellByPosition(0,nbl+1)
		Cell.formula = "=A2"
	Range = Sheet.getCellRangeByPosition(0,nbl+1,Nbs+3,2*nbl-1)
		Range.fillSeries(com.sun.star.sheet.FillDirection.TO_BOTTOM,com.sun.star.sheet.FillMode.SIMPLE,0,0,0)
		Range.fillSeries(com.sun.star.sheet.FillDirection.TO_RIGHT,com.sun.star.sheet.FillMode.SIMPLE,0,0,0)
		
	Range = Sheet.getCellRangeByPosition(0,nbl+1,3,2*nbl-1)
	With Range 
		  '.setString( "Notes Normalisées" ) 'insére du texte dans la cellule 
		  .CellBackColor = RGB(240,248,255) 'indique la couleur de fond 
		  '.paraAdjust = com.sun.star.style.ParagraphAdjust.CENTER 'alignement centré 
		  '.RotateAngle = 9000 'Rotation 9000 = 90° 
		  .CharColor = RGB(0,0,0) 'couleur des caractères 
		 ' .CharHeight = 10 'Taille catactères 
		 ' .CharWeight = com.sun.star.awt.FontWeight.BOLD 'gras 
		  .CharPosture = com.sun.star.awt.FontSlant.ITALIC 'italique 
		  .CharFontName = "Arial" 'Font 
		 ' .CharUnderline = com.sun.star.awt.FontUnderline.DOUBLE 'souligné double 
		End With  
		
	Range = Sheet.getCellRangeByPosition(4,nbl+1,3+Nbs,nbl+1)
	With Range 
		  '.setString( "Notes Normalisées" ) 'insére du texte dans la cellule 
		  .CellBackColor = RGB(218,165,32) 'indique la couleur de fond 
		  '.paraAdjust = com.sun.star.style.ParagraphAdjust.CENTER 'alignement centré 
		  '.RotateAngle = 9000 'Rotation 9000 = 90° 
		  .CharColor = RGB(0,0,0) 'couleur des caractères 
		 ' .CharHeight = 10 'Taille catactères 
		  .CharWeight = com.sun.star.awt.FontWeight.BOLD 'gras 
		  .CharPosture = com.sun.star.awt.FontSlant.ITALIC 'italique 
		  .CharFontName = "Arial" 'Font 
		 ' .CharUnderline = com.sun.star.awt.FontUnderline.DOUBLE 'souligné double 
		End With  
 	
'Ecriture de la formule cellule 1,1 Normalisé

	Sheet = Doc.Sheets(Nb_periode+1)
	Cell = Sheet.getCellByPosition(4,nbl+2) 
	form = "=SI(ESTERREUR( E3 * $D"& nbl+3 & "*10/Max(3:3));0; E3 * $D"& nbl+3 & "*10/Max(3:3))"
	Cell.formulalocal = form 
	
' filler formule dans le reste du tableau

Cell = Sheet.getCellByPosition(4,nbl+2) 

Range = Sheet.getCellRangeByPosition(4,nbl+2,Nbs+3,nbl+NBe+1)

Range.fillSeries(com.sun.star.sheet.FillDirection.TO_BOTTOM,com.sun.star.sheet.FillMode.SIMPLE,0,0,0)
Range.fillSeries(com.sun.star.sheet.FillDirection.TO_RIGHT,com.sun.star.sheet.FillMode.SIMPLE,0,0,0)

Next Nb_periode


''''''''Feuille Stage Période X''''''''

For Nb_periode=1 to Nb_periode_Max

 
	oDoc=ThisComponent 
	  
	'Ajoute une feuille, la nomme et place l'onglet en 3eme position 
	oDoc.GetSheets.insertNewByName("Stage " & Nb_periode,Nb_periode_Max + Nb_periode+1)

'Ecriture des en têtes 

	Sheet = Doc.Sheets(Nb_periode_Max + Nb_periode+1)
	
	Plage = Sheet.getCellRangeByPosition( 1 , 0 , Nbs+3 , 0 ) 
		Plage.Merge( True ) 

		Cell = Sheet.getCellByPosition(1,0) 
		With Cell 
		  .setString( "Tableau Binaire" ) 'insére du texte dans la cellule 
		  .CellBackColor = RGB(255,215,0) 'indique la couleur de fond 
		  '.paraAdjust = com.sun.star.style.ParagraphAdjust.CENTER 'alignement centré 
		  '.RotateAngle = 9000 'Rotation 9000 = 90° 
		  .CharColor = RGB(0,0,0) 'couleur des caractères 
		  .CharHeight = 14 'Taille catactères 
		  .CharWeight = com.sun.star.awt.FontWeight.BOLD 'gras 
		  .CharPosture = com.sun.star.awt.FontSlant.ITALIC 'italique 
		  .CharFontName = "Arial" 'Font 
		 ' .CharUnderline = com.sun.star.awt.FontUnderline.DOUBLE 'souligné double 
		End With 
		  
		'Ajuste la hauteur de la 5eme ligne au contenu des cellules.  
		Sheet.Rows(0).OptimalHeight = True 
	
	
	Cell = Sheet.getCellByPosition(0,0) 
		With Cell 
		  .setString( "Fonction Solveur" ) 'insére du texte dans la cellule 
		  .CellBackColor = RGB(60,179,113) 'indique la couleur de fond 
		  .paraAdjust = com.sun.star.style.ParagraphAdjust.CENTER 'alignement centré 
		  '.RotateAngle = 9000 'Rotation 9000 = 90° 
		  .CharColor = RGB(0,0,0) 'couleur des caractères 
		  .CharHeight = 14 'Taille catactères 
		  .CharWeight = com.sun.star.awt.FontWeight.BOLD 'gras 
		  '.CharPosture = com.sun.star.awt.FontSlant.ITALIC 'italique 
		  .CharFontName = "Arial" 'Font 
		 ' .CharUnderline = com.sun.star.awt.FontUnderline.DOUBLE 'souligné double 
		End With 
		  
	Cell = Sheet.getCellByPosition(0,1) 
	
	Cell.formulalocal = "=Adresse("& 2*NBe+4 & ";" & Nbs+4 & ";4)"
	Dim Extract_1 as string
	Dim Extract_2 as string 
	Extract_1 = cell.string
	Cell.formulalocal = "=Adresse("& NBe+5 & ";" & Nbs+1 & ";4)"
	Extract_2 = cell.string
	
	Cell = Sheet.getCellByPosition(0,1)
	Cell.Formulalocal = "=SOMMEPROD($'Coef " & Nb_periode & "'.E" & (Nbe+5) & ":$'Coef " & Nb_periode & "'." & Extract_1 & ";B6:" & Extract_2 &")"
	
		With Cell 
	'	  .setFormula("=SOMMEPROD($'Coef " & Nb_periode & "'.D" & (Nbe+4) & ":$'Coef " & Nb_periode & "'." & Extract_1 & ";B6:" & Extract_2 &")")
		  .CellBackColor = RGB(144,238,144) 'indique la couleur de fond 
		  .paraAdjust = com.sun.star.style.ParagraphAdjust.CENTER 'alignement centré 
		  '.RotateAngle = 9000 'Rotation 9000 = 90° 
		  .CharColor = RGB(0,0,0) 'couleur des caractères 
		  .CharHeight = 14 'Taille catactères 
		  .CharWeight = com.sun.star.awt.FontWeight.BOLD 'gras 
		  '.CharPosture = com.sun.star.awt.FontSlant.ITALIC 'italique 
		  .CharFontName = "Arial" 'Font 
		 ' .CharUnderline = com.sun.star.awt.FontUnderline.DOUBLE 'souligné double 
		End With 
	
	
	Cell = Sheet.getCellByPosition(0,2) 
		
		With Cell 
		  .setString( "Minimum" ) 'insére du texte dans la cellule 
		  .CellBackColor = RGB(240,128,128) 'indique la couleur de fond 
		  '.paraAdjust = com.sun.star.style.ParagraphAdjust.CENTER 'alignement centré 
		  '.RotateAngle = 9000 'Rotation 9000 = 90° 
		  .CharColor = RGB(0,0,0) 'couleur des caractères 
		  .CharHeight = 10 'Taille catactères 
		  .CharWeight = com.sun.star.awt.FontWeight.BOLD 'gras 
		  '.CharPosture = com.sun.star.awt.FontSlant.ITALIC 'italique 
		  .CharFontName = "Arial" 'Font 
		 ' .CharUnderline = com.sun.star.awt.FontUnderline.DOUBLE 'souligné double 
		End With 
	
	
	Cell = Sheet.getCellByPosition(0,3)
	
		With Cell 
		  .setString( "Maximum" ) 'insére du texte dans la cellule 
		  .CellBackColor = RGB(205,92,92) 'indique la couleur de fond 
		  '.paraAdjust = com.sun.star.style.ParagraphAdjust.CENTER 'alignement centré 
		  '.RotateAngle = 9000 'Rotation 9000 = 90° 
		  .CharColor = RGB(0,0,0) 'couleur des caractères 
		  .CharHeight = 10 'Taille catactères 
		  .CharWeight = com.sun.star.awt.FontWeight.BOLD 'gras 
		  '.CharPosture = com.sun.star.awt.FontSlant.ITALIC 'italique 
		  .CharFontName = "Arial" 'Font 
		 ' .CharUnderline = com.sun.star.awt.FontUnderline.DOUBLE 'souligné double 
		End With  
	
	Cell = Sheet.getCellByPosition(0,4)
	With Cell 
		  .setString( "Affectés" ) 'insére du texte dans la cellule 
		  .CellBackColor = RGB(178,34,34) 'indique la couleur de fond 
		  '.paraAdjust = com.sun.star.style.ParagraphAdjust.CENTER 'alignement centré 
		  '.RotateAngle = 9000 'Rotation 9000 = 90° 
		  .CharColor = RGB(255,255,255) 'couleur des caractères 
		  .CharHeight = 10 'Taille catactères 
		  .CharWeight = com.sun.star.awt.FontWeight.BOLD 'gras 
		  '.CharPosture = com.sun.star.awt.FontSlant.ITALIC 'italique 
		  .CharFontName = "Arial" 'Font 
		 ' .CharUnderline = com.sun.star.awt.FontUnderline.DOUBLE 'souligné double 
		End With  
	
	form = "=$'Place Stage'.C$1"
	Cell = Sheet.getCellByPosition(1,1)
	Cell.formula = form
	
	With Cell 
		  '.setString( "Nombre d'étudiants sur le terrain" ) 'insére du texte dans la cellule 
		  CellBackColor = RGB(139,0,0) 'indique la couleur de fond 
		  '.paraAdjust = com.sun.star.style.ParagraphAdjust.CENTER 'alignement centré 
		  '.RotateAngle = 9000 'Rotation 9000 = 90° 
		  .CharColor = RGB(255,255,255) 'couleur des caractères 
		  .CharHeight = 10 'Taille catactères 
		  .CharWeight = com.sun.star.awt.FontWeight.BOLD 'gras 
		  '.CharPosture = com.sun.star.awt.FontSlant.ITALIC 'italique 
		  .CharFontName = "Arial" 'Font 
		 ' .CharUnderline = com.sun.star.awt.FontUnderline.DOUBLE 'souligné double 
		End With  
	Cell.CellBackColor = RGB(139,0,0)
	
dim offset as integer

	form = "=$'Place Stage'.C" & 3+offset
	Cell = Sheet.getCellByPosition(1,2)
	Cell.formula = form
		With Cell 
		  '.setString( "Minimum" ) 'insére du texte dans la cellule 
		  .CellBackColor = RGB(240,128,128) 'indique la couleur de fond 
		  '.paraAdjust = com.sun.star.style.ParagraphAdjust.CENTER 'alignement centré 
		  '.RotateAngle = 9000 'Rotation 9000 = 90° 
		  .CharColor = RGB(0,0,0) 'couleur des caractères 
		  .CharHeight = 10 'Taille catactères 
		 ' .CharWeight = com.sun.star.awt.FontWeight.BOLD 'gras 
		  '.CharPosture = com.sun.star.awt.FontSlant.ITALIC 'italique 
		  .CharFontName = "Arial" 'Font 
		 ' .CharUnderline = com.sun.star.awt.FontUnderline.DOUBLE 'souligné double 
		End With 
		
	form = "=$'Place Stage'.C" & 4+offset
	Cell = Sheet.getCellByPosition(1,3)
	Cell.formula = form
	
	With Cell 
		 ' .setString( "Maximum" ) 'insére du texte dans la cellule 
		  .CellBackColor = RGB(205,92,92) 'indique la couleur de fond 
		  '.paraAdjust = com.sun.star.style.ParagraphAdjust.CENTER 'alignement centré 
		  '.RotateAngle = 9000 'Rotation 9000 = 90° 
		  .CharColor = RGB(0,0,0) 'couleur des caractères 
		  .CharHeight = 10 'Taille catactères 
		 ' .CharWeight = com.sun.star.awt.FontWeight.BOLD 'gras 
		  '.CharPosture = com.sun.star.awt.FontSlant.ITALIC 'italique 
		  .CharFontName = "Arial" 'Font 
		 ' .CharUnderline = com.sun.star.awt.FontUnderline.DOUBLE 'souligné double 
		End With
	
	Cell = Sheet.getCellByPosition(1,4)
	Cell.formula = "=SUM(B6:B" & 5+NBe & ")"
		With Cell 
		  '.setString( "Nombre d'étudiants sur le terrain" ) 'insére du texte dans la cellule 
		  .CellBackColor = RGB(178,34,34) 'indique la couleur de fond 
		  '.paraAdjust = com.sun.star.style.ParagraphAdjust.CENTER 'alignement centré 
		  '.RotateAngle = 9000 'Rotation 9000 = 90° 
		  .CharColor = RGB(255,255,255) 'couleur des caractères 
		  .CharHeight = 10 'Taille catactères 
		 ' .CharWeight = com.sun.star.awt.FontWeight.BOLD 'gras 
		  '.CharPosture = com.sun.star.awt.FontSlant.ITALIC 'italique 
		  .CharFontName = "Arial" 'Font 
		 ' .CharUnderline = com.sun.star.awt.FontUnderline.DOUBLE 'souligné double 
		End With  
	
	Sheet.columns(0).OptimalWidth = True
	
	Cell = Sheet.getCellByPosition(4,nbl+2) 
	
	
Range = Sheet.getCellRangeByPosition(1,1,Nbs,4)
Range.fillSeries(com.sun.star.sheet.FillDirection.TO_RIGHT,com.sun.star.sheet.FillMode.SIMPLE,0,0,0)

	Cell = Sheet.getCellByPosition(0,5)
	Form = "=CONCAT($'Liste Etudiants'.$B2 ;" & chr(34) & chr(0160) & chr(34) &"; $'Liste Etudiants'.$C2)"
	Cell.formula = form
	
	With Cell 
		 ' .setString( "Nombre d'étudiants sur le terrain" ) 'insére du texte dans la cellule 
		  .CellBackColor = RGB(255,250,205) 'indique la couleur de fond 
		  '.paraAdjust = com.sun.star.style.ParagraphAdjust.CENTER 'alignement centré 
		  '.RotateAngle = 9000 'Rotation 9000 = 90° 
		  .CharColor = RGB(0,0,0) 'couleur des caractères 
		  .CharHeight = 10 'Taille catactères 
		 ' .CharWeight = com.sun.star.awt.FontWeight.BOLD 'gras 
		 .CharPosture = com.sun.star.awt.FontSlant.ITALIC 'italique 
		  .CharFontName = "Arial" 'Font 
		 ' .CharUnderline = com.sun.star.awt.FontUnderline.DOUBLE 'souligné double 
		End With  
		
	Range = Sheet.getCellRangeByPosition(0,5,0,NBe+4)
	Range.fillSeries(com.sun.star.sheet.FillDirection.TO_BOTTOM,com.sun.star.sheet.FillMode.SIMPLE,0,0,0)

dim extractform as string
dim extractform2 as string

	Cell = Sheet.getCellByPosition(Nbs+1,5)
	Cell.formulalocal = "=ADRESSE(6;" & Nbs+1 & ";4;; )"
	extractform = Cell.string
	Cell = Sheet.getCellByPosition(Nbs+1,5)
	Cell.formulalocal = "=ADRESSE(6;" & Nbs+3 & ";4;; )"
	extractform2 = Cell.string 
	Cell.formula = "=SUM(B6:" & extractform & ")+" & extractform2
		
	Cell = Sheet.getCellByPosition(Nbs+2,4)
			With Cell 
		  .setString( "Hors Algo" ) 'insére du texte dans la cellule 
		  .CellBackColor = RGB(26, 188, 156) 'indique la couleur de fond 
		  '.paraAdjust = com.sun.star.style.ParagraphAdjust.CENTER 'alignement centré 
		  '.RotateAngle = 9000 'Rotation 9000 = 90° 
		  .CharColor = RGB(255,255,255) 'couleur des caractères 
		  .CharHeight = 10 'Taille catactères 
		  .CharWeight = com.sun.star.awt.FontWeight.BOLD 'gras 
		  '.CharPosture = com.sun.star.awt.FontSlant.ITALIC 'italique 
		  .CharFontName = "Arial" 'Font 
		 ' .CharUnderline = com.sun.star.awt.FontUnderline.DOUBLE 'souligné double 
		End With  
		
		Cell = Sheet.getCellByPosition(Nbs+3,4)
			With Cell 
		  .setString( "Commentaire" ) 'insére du texte dans la cellule 
		  .CellBackColor = RGB(26, 188, 156) 'indique la couleur de fond 
		  '.paraAdjust = com.sun.star.style.ParagraphAdjust.CENTER 'alignement centré 
		  '.RotateAngle = 9000 'Rotation 9000 = 90° 
		  .CharColor = RGB(255,255,255) 'couleur des caractères 
		  .CharHeight = 10 'Taille catactères 
		  .CharWeight = com.sun.star.awt.FontWeight.BOLD 'gras 
		  '.CharPosture = com.sun.star.awt.FontSlant.ITALIC 'italique 
		  .CharFontName = "Arial" 'Font 
		 ' .CharUnderline = com.sun.star.awt.FontUnderline.DOUBLE 'souligné double 
		End With  
Dim horsalgo as string
	Cell = Sheet.getCellByPosition(Nbs+2,5)
	Cell.formulalocal = "=ADRESSE(6;" & Nbs+4 & ";3;; )"
	horsalgo = Cell.string
	Cell.formulalocal = "=SI(ESTVIDE(" & horsalgo & ");0;1)"
			With Cell 
		  '.setString( "Hors Algo" ) 'insére du texte dans la cellule 
		  .CellBackColor = RGB(130, 224, 170) 'indique la couleur de fond 
		  '.paraAdjust = com.sun.star.style.ParagraphAdjust.CENTER 'alignement centré 
		  '.RotateAngle = 9000 'Rotation 9000 = 90° 
		  .CharColor = RGB(0,0,0) 'couleur des caractères 
		  .CharHeight = 10 'Taille catactères 
		 ' .CharWeight = com.sun.star.awt.FontWeight.BOLD 'gras 
		  '.CharPosture = com.sun.star.awt.FontSlant.ITALIC 'italique 
		  .CharFontName = "Arial" 'Font 
		 ' .CharUnderline = com.sun.star.awt.FontUnderline.DOUBLE 'souligné double 
		End With  
	Cell = Sheet.getCellByPosition(Nbs+3,5)
		With Cell 
		  '.setString( "Hors Algo" ) 'insére du texte dans la cellule 
		  .CellBackColor = RGB(130, 224, 170) 'indique la couleur de fond 
		  '.paraAdjust = com.sun.star.style.ParagraphAdjust.CENTER 'alignement centré 
		  '.RotateAngle = 9000 'Rotation 9000 = 90° 
		  .CharColor = RGB(0,0,0) 'couleur des caractères 
		  .CharHeight = 10 'Taille catactères 
		 ' .CharWeight = com.sun.star.awt.FontWeight.BOLD 'gras 
		  '.CharPosture = com.sun.star.awt.FontSlant.ITALIC 'italique 
		  .CharFontName = "Arial" 'Font 
		 ' .CharUnderline = com.sun.star.awt.FontUnderline.DOUBLE 'souligné double 
		End With 
	Range = Sheet.getCellRangeByPosition(Nbs+1,5,Nbs+3,NBe+4)
	Range.fillSeries(com.sun.star.sheet.FillDirection.TO_BOTTOM,com.sun.star.sheet.FillMode.SIMPLE,0,0,0)
		
	
offset = offset +2

	Sheet = Doc.Sheets(Nb_periode_Max + Nb_periode+1)
	
	Plage = Sheet.getCellRangeByPosition( 0 , NBe+5 , Nbs+3 , NBe+5 ) 
	Plage.Merge( True ) 	
	
	Cell = Sheet.getCellByPosition(0,NBe+5)
		With Cell 
		  .setString( "Tableau Satisfaction" ) 'insére du texte dans la cellule 
		  .CellBackColor = RGB(255,215,0) 'indique la couleur de fond 
		  '.paraAdjust = com.sun.star.style.ParagraphAdjust.CENTER 'alignement centré 
		  '.RotateAngle = 9000 'Rotation 9000 = 90° 
		  .CharColor = RGB(0,0,0) 'couleur des caractères 
		  .CharHeight = 14 'Taille catactères 
		  .CharWeight = com.sun.star.awt.FontWeight.BOLD 'gras 
		.CharPosture = com.sun.star.awt.FontSlant.ITALIC 'italique 
		  .CharFontName = "Arial" 'Font 
		 ' .CharUnderline = com.sun.star.awt.FontUnderline.DOUBLE 'souligné double 
		End With  
	  
	Cell = Sheet.getCellByPosition(1,NBe+6) 
	Cell.formula = "=$'Place Stage'.C$1"
	
	With Cell 
		  '.setString( "Nombre d'étudiants sur le terrain" ) 'insére du texte dans la cellule 
		  CellBackColor = RGB(139,0,0) 'indique la couleur de fond 
		  '.paraAdjust = com.sun.star.style.ParagraphAdjust.CENTER 'alignement centré 
		  '.RotateAngle = 9000 'Rotation 9000 = 90° 
		  .CharColor = RGB(255,255,255) 'couleur des caractères 
		  .CharHeight = 10 'Taille catactères 
		  .CharWeight = com.sun.star.awt.FontWeight.BOLD 'gras 
		  '.CharPosture = com.sun.star.awt.FontSlant.ITALIC 'italique 
		  .CharFontName = "Arial" 'Font 
		 ' .CharUnderline = com.sun.star.awt.FontUnderline.DOUBLE 'souligné double 
		End With  
	Cell.CellBackColor = RGB(139,0,0)
	
	Range = Sheet.getCellRangeByPosition(1,NBe+6,Nbs,NBe+6)
	Range.fillSeries(com.sun.star.sheet.FillDirection.TO_RIGHT,com.sun.star.sheet.FillMode.SIMPLE,0,0,0)

	Cell = Sheet.getCellByPosition(0,NBe+7)
	Form = "=CONCAT($'Liste Etudiants'.$B2 ;" & chr(34) & chr(0160) & chr(34) &"; $'Liste Etudiants'.$C2)"
	Cell.formula = form
		With Cell 
		 ' .setString( "Nombre d'étudiants sur le terrain" ) 'insére du texte dans la cellule 
		  .CellBackColor = RGB(255,250,205) 'indique la couleur de fond 
		  '.paraAdjust = com.sun.star.style.ParagraphAdjust.CENTER 'alignement centré 
		  '.RotateAngle = 9000 'Rotation 9000 = 90° 
		  .CharColor = RGB(0,0,0) 'couleur des caractères 
		  .CharHeight = 10 'Taille catactères 
		 ' .CharWeight = com.sun.star.awt.FontWeight.BOLD 'gras 
		 .CharPosture = com.sun.star.awt.FontSlant.ITALIC 'italique 
		  .CharFontName = "Arial" 'Font 
		 ' .CharUnderline = com.sun.star.awt.FontUnderline.DOUBLE 'souligné double 
		End With  
	Range = Sheet.getCellRangeByPosition(0,NBe+7,0,2*NBe+7)
	Range.fillSeries(com.sun.star.sheet.FillDirection.TO_BOTTOM,com.sun.star.sheet.FillMode.SIMPLE,0,0,0)
	
	Cell = Sheet.getCellByPosition(1,NBe+7)
	Form = "=SI(B6=1;$'Coef " & Nb_periode &"'.E"& nbl+3 &";0)"
	Cell.formulalocal = form
	Range = Sheet.getCellRangeByPosition(1,NBe+7,Nbs,2*NBe+6)
	Range.fillSeries(com.sun.star.sheet.FillDirection.TO_BOTTOM,com.sun.star.sheet.FillMode.SIMPLE,0,0,0)
	Range.fillSeries(com.sun.star.sheet.FillDirection.TO_RIGHT,com.sun.star.sheet.FillMode.SIMPLE,0,0,0)

	Cell = Sheet.getCellByPosition(Nbs+1,NBe+7)
	Cell.formulalocal = "=ADRESSE("& NBe+8 &";" & Nbs+1 & ";4;;)"
	extractform = Cell.string
	Cell.formulalocal = "=ADRESSE("& NBe+8 &";" & Nbs+3 & ";3;;)"
	extractform2 = Cell.string
	Cell.formulalocal = "=SOMME(B"& NBe+8 &":" & extractform & ")/$'Coef " & Nb_periode &"'.D" & NBe+5 & "+" & extractform2
	Cell = Sheet.getCellByPosition(Nbs+2,NBe+7)
	Cell.formulalocal = "=SI(ESTVIDE(" & horsalgo & ");0;10)"
	Range = Sheet.getCellRangeByPosition(Nbs+1,NBe+7,Nbs+2,2*NBe+6)
	Range.fillSeries(com.sun.star.sheet.FillDirection.TO_BOTTOM,com.sun.star.sheet.FillMode.SIMPLE,0,0,0)
		
	Sheet = Doc.Sheets(Nb_periode_Max + Nb_periode+1)
	
	Plage = Sheet.getCellRangeByPosition( 0 , 2*NBe+7 , Nbs+3 , 2*NBe+7 ) 
	Plage.Merge( True ) 	
	
	Cell = Sheet.getCellByPosition(0,2*NBe+7)
		With Cell 
		  .setString( "Tableau Stage" ) 'insére du texte dans la cellule 
		  .CellBackColor = RGB(255,215,0) 'indique la couleur de fond 
		  '.paraAdjust = com.sun.star.style.ParagraphAdjust.CENTER 'alignement centré 
		  '.RotateAngle = 9000 'Rotation 9000 = 90° 
		  .CharColor = RGB(0,0,0) 'couleur des caractères 
		  .CharHeight = 14 'Taille catactères 
		  .CharWeight = com.sun.star.awt.FontWeight.BOLD 'gras 
		.CharPosture = com.sun.star.awt.FontSlant.ITALIC 'italique 
		  .CharFontName = "Arial" 'Font 
		 ' .CharUnderline = com.sun.star.awt.FontUnderline.DOUBLE 'souligné double 
		End With  
	
	Cell = Sheet.getCellByPosition(1,2*NBe+8) 
	Cell.formula = "=$'Place Stage'.C$1"
	With Cell 
		  '.setString( "Nombre d'étudiants sur le terrain" ) 'insére du texte dans la cellule 
		  CellBackColor = RGB(139,0,0) 'indique la couleur de fond 
		  '.paraAdjust = com.sun.star.style.ParagraphAdjust.CENTER 'alignement centré 
		  '.RotateAngle = 9000 'Rotation 9000 = 90° 
		  .CharColor = RGB(255,255,255) 'couleur des caractères 
		  .CharHeight = 10 'Taille catactères 
		  .CharWeight = com.sun.star.awt.FontWeight.BOLD 'gras 
		  '.CharPosture = com.sun.star.awt.FontSlant.ITALIC 'italique 
		  .CharFontName = "Arial" 'Font 
		 ' .CharUnderline = com.sun.star.awt.FontUnderline.DOUBLE 'souligné double 
		End With  
	Cell.CellBackColor = RGB(139,0,0)
	
	Range = Sheet.getCellRangeByPosition(1,2*NBe+8,Nbs,2*NBe+8)
	Range.fillSeries(com.sun.star.sheet.FillDirection.TO_RIGHT,com.sun.star.sheet.FillMode.SIMPLE,0,0,0)

	Cell = Sheet.getCellByPosition(0,2*NBe+9)
	Form = "=CONCAT($'Liste Etudiants'.$B2 ;" & chr(34) & chr(0160) & chr(34) &"; $'Liste Etudiants'.$C2)"
	Cell.formula = form
		With Cell 
		 ' .setString( "Nombre d'étudiants sur le terrain" ) 'insére du texte dans la cellule 
		  .CellBackColor = RGB(255,250,205) 'indique la couleur de fond 
		  '.paraAdjust = com.sun.star.style.ParagraphAdjust.CENTER 'alignement centré 
		  '.RotateAngle = 9000 'Rotation 9000 = 90° 
		  .CharColor = RGB(0,0,0) 'couleur des caractères 
		  .CharHeight = 10 'Taille catactères 
		 ' .CharWeight = com.sun.star.awt.FontWeight.BOLD 'gras 
		 .CharPosture = com.sun.star.awt.FontSlant.ITALIC 'italique 
		  .CharFontName = "Arial" 'Font 
		 ' .CharUnderline = com.sun.star.awt.FontUnderline.DOUBLE 'souligné double 
		End With  
	Range = Sheet.getCellRangeByPosition(0,2*NBe+9,0,3*NBe+8)
	Range.fillSeries(com.sun.star.sheet.FillDirection.TO_BOTTOM,com.sun.star.sheet.FillMode.SIMPLE,0,0,0)
	
	Cell = Sheet.getCellByPosition(1,2*NBe+9)
	Form = "=SI(B6=1;B$2;" & chr(34) & chr(34) &")"
	Cell.formulalocal = form
	Range = Sheet.getCellRangeByPosition(1,2*NBe+9,Nbs,3*NBe+8)
	Range.fillSeries(com.sun.star.sheet.FillDirection.TO_BOTTOM,com.sun.star.sheet.FillMode.SIMPLE,0,0,0)
	Range.fillSeries(com.sun.star.sheet.FillDirection.TO_RIGHT,com.sun.star.sheet.FillMode.SIMPLE,0,0,0)

	Cell = Sheet.getCellByPosition(Nbs+1,2*NBe+9)
	Cell.formulalocal = "=ADRESSE("& 2*NBe+10 &";" & Nbs+1 & ";4;;)"
	extractform = Cell.string
	Cell.formulalocal = "=ADRESSE("& 2*NBe+10 &";" & Nbs+3 & ";4;;)"
	extractform2 = Cell.string
	Cell.formula = "=CONCAT(B"& 2*NBe+10 &":" & extractform & ";" & extractform2 & ")"
	
	Cell = Sheet.getCellByPosition(Nbs+2,2*NBe+9)
	Cell.formulalocal = "=ADRESSE(6;" & Nbs+4 & ";3;; )"
	extractform2 = Cell.string
	Cell.formulalocal = "=SI(ESTVIDE(" & horsalgo & ");" & chr(34) & chr(34) & ";" & extractform2 & ")"
	Range = Sheet.getCellRangeByPosition(Nbs+1,2*NBe+9,Nbs+2,3*NBe+8)
	Range.fillSeries(com.sun.star.sheet.FillDirection.TO_BOTTOM,com.sun.star.sheet.FillMode.SIMPLE,0,0,0)

Next Nb_periode
'''''''''''''''''''''' Feuille Réouverture''''''''''''''''''''''''''''''''''''''''

	Doc=ThisComponent 
	  
	'Ajoute une feuille, la nomme et place l'onglet en 3eme position 
	Doc.GetSheets.insertNewByName("Réouverture",2*Nb_periode_Max+3)
	
	Sheet = Doc.Sheets.GetbyName("Réouverture")
'''''''''''''''''''''	
	Range = Sheet.getCellRangeByPosition( 0 , 0 , Nbs+3 , 0 ) 
		Range.Merge( True ) 

		Cell = Sheet.getCellByPosition(0,0)   
		With Cell 
		  .setString( "Réouverture" ) 'insére du texte dans la cellule 
		  .CellBackColor = RGB(255,215,0) 'indique la couleur de fond 
		  '.paraAdjust = com.sun.star.style.ParagraphAdjust.CENTER 'alignement centré 
		  '.RotateAngle = 9000 'Rotation 9000 = 90° 
		  .CharColor = RGB(0,0,0) 'couleur des caractères 
		  .CharHeight = 14 'Taille catactères 
		  .CharWeight = com.sun.star.awt.FontWeight.BOLD 'gras 
		  .CharPosture = com.sun.star.awt.FontSlant.ITALIC 'italique 
		  .CharFontName = "Arial" 'Font 
		 ' .CharUnderline = com.sun.star.awt.FontUnderline.DOUBLE 'souligné double 
		End With 
		  
		  
		'Ajuste la hauteur de la 5eme ligne au contenu des cellules.  
		Sheet.Rows(0).OptimalHeight = True 
		


	Cell = Sheet.getCellByPosition(0,2)
		Cell.formula = "=$'Liste Etudiants'.A1"
	Range = Sheet.getCellRangeByPosition(0,2,2,nbe+2)
		Range.fillSeries(com.sun.star.sheet.FillDirection.TO_BOTTOM,com.sun.star.sheet.FillMode.SIMPLE,0,0,0)
		Range.fillSeries(com.sun.star.sheet.FillDirection.TO_RIGHT,com.sun.star.sheet.FillMode.SIMPLE,0,0,0)
		
 
 
	Cell = Sheet.getCellByPosition(3,1)
	Cell.formula = "=$'Place Stage'.C1"
	Range = Sheet.getCellRangeByPosition(3,1,2+Nbs,2)
	Range.fillSeries(com.sun.star.sheet.FillDirection.TO_BOTTOM,com.sun.star.sheet.FillMode.SIMPLE,0,0,0)
	Range.fillSeries(com.sun.star.sheet.FillDirection.TO_RIGHT,com.sun.star.sheet.FillMode.SIMPLE,0,0,0)


		Range = Sheet.getCellRangeByPosition(3,1,2+Nbs,2)   
		With Range 
		  '.setString( "Notes Bruts" ) 'insére du texte dans la cellule 
		  .CellBackColor = RGB(218,165,32) 'indique la couleur de fond 
		  .paraAdjust = com.sun.star.style.ParagraphAdjust.CENTER 'alignement centré 
		  '.RotateAngle = 9000 'Rotation 9000 = 90° 
		  .CharColor = RGB(0,0,0) 'couleur des caractères 
		  .CharHeight = 10 'Taille catactères 
		  '.CharWeight = com.sun.star.awt.FontWeight.BOLD 'gras 
		  '.CharPosture = com.sun.star.awt.FontSlant.ITALIC 'italique 
		  .CharFontName = "Arial" 'Font 
		 ' .CharUnderline = com.sun.star.awt.FontUnderline.DOUBLE 'souligné double 
		End With 
	
	Sheet.Rows(1).OptimalHeight = True 
		

	Range = Sheet.getCellRangeByPosition(0,2,2,NBe+2)   
		With Range 
		  '.setString( "Notes Bruts" ) 'insére du texte dans la cellule 
		  .CellBackColor = RGB(240,248,255) 'indique la couleur de fond 
		  '.paraAdjust = com.sun.star.style.ParagraphAdjust.CENTER 'alignement centré 
		  '.RotateAngle = 9000 'Rotation 9000 = 90° 
		  .CharColor = RGB(0,0,0) 'couleur des caractères 
		  .CharHeight = 10 'Taille catactères 
		  '.CharWeight = com.sun.star.awt.FontWeight.BOLD 'gras 
		  .CharPosture = com.sun.star.awt.FontSlant.ITALIC 'italique 
		  .CharFontName = "Arial" 'Font 
		 ' .CharUnderline = com.sun.star.awt.FontUnderline.DOUBLE 'souligné double 
		End With 
	For numl =1 to NBe-1
	Sheet.Rows(1+numl).OptimalHeight = True
	Next numl

	For numl=0 to 3
		Sheet.columns(numl).OptimalWidth = True 
	Next numl
		
'''''''''''''''''''''''''
	
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''' clipboard 1
	
	
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Cell = Sheet.getCellByPosition(3,3)
	Cell.Value = 0
	
	Range = Sheet.getCellRangeByPosition(3,3,Nbs+2,NBe+2)
	Range.fillSeries(com.sun.star.sheet.FillDirection.TO_RIGHT,com.sun.star.sheet.FillMode.SIMPLE,0,0,0)
	Range.fillSeries(com.sun.star.sheet.FillDirection.TO_BOTTOM,com.sun.star.sheet.FillMode.SIMPLE,0,0,0)
	
Cell = Sheet.getCellByPosition(Nbs+3,1)
	Cell.formula = "=MAX($'Place Stage'.2:2)"
	NbGRP = cell.value
	
	IF NbGRP>0 Then

	For NumGRP = 1 to NbGRP
	
	'
		Cell = Sheet.getCellByPosition(Nbs+2+NumGRP,1)
		Cell.String = "Regroupement N°"
		With Cell 
		  '.setString( "Nombre d'étudiants sur le terrain" ) 'insére du texte dans la cellule 
		  .CellBackColor = RGB(222,184,135) 'indique la couleur de fond 
		  '.paraAdjust = com.sun.star.style.ParagraphAdjust.CENTER 'alignement centré 
		  '.RotateAngle = 9000 'Rotation 9000 = 90° 
		  .CharColor = RGB(0,0,0) 'couleur des caractères 
		  .CharHeight = 10 'Taille catactères 
		  .CharWeight = com.sun.star.awt.FontWeight.BOLD 'gras 
		  '.CharPosture = com.sun.star.awt.FontSlant.ITALIC 'italique 
		  .CharFontName = "Arial" 'Font 
		 ' .CharUnderline = com.sun.star.awt.FontUnderline.DOUBLE 'souligné double 
		End With
		Cell = Sheet.getCellByPosition(Nbs+2+NumGRP,2)
		Cell.Value = NumGRP
		With Cell 
		  '.setString( "Nombre d'étudiants sur le terrain" ) 'insére du texte dans la cellule 
		  .CellBackColor = RGB(222,184,135) 'indique la couleur de fond 
		  '.paraAdjust = com.sun.star.style.ParagraphAdjust.CENTER 'alignement centré 
		  '.RotateAngle = 9000 'Rotation 9000 = 90° 
		  .CharColor = RGB(0,0,0) 'couleur des caractères 
		  .CharHeight = 10 'Taille catactères 
		  .CharWeight = com.sun.star.awt.FontWeight.BOLD 'gras 
		  '.CharPosture = com.sun.star.awt.FontSlant.ITALIC 'italique 
		  .CharFontName = "Arial" 'Font 
		 ' .CharUnderline = com.sun.star.awt.FontUnderline.DOUBLE 'souligné double 
		End With
		Sheet.columns(Nbs+2+NumGRP).OptimalWidth = True
	Next NumGRP
	Sheet.columns(0).OptimalWidth = True

	
	Cell = Sheet.getCellByPosition(Nbs+1,3)
		Cell.Value = 0
			
	Range = Sheet.getCellRangeByPosition(Nbs+1,3,Nbs+2+NumGRP-1,NBe+2)
	Range.fillSeries(com.sun.star.sheet.FillDirection.TO_RIGHT,com.sun.star.sheet.FillMode.SIMPLE,0,0,0)
	Range.fillSeries(com.sun.star.sheet.FillDirection.TO_BOTTOM,com.sun.star.sheet.FillMode.SIMPLE,0,0,0)	
		
	End If
	Sheet.columns(0).OptimalWidth = True



	
'''''''''''''''''''''' Feuille Regroupement''''''''''''''''''''''''''''''''''''''''

	Doc=ThisComponent 
	  
	'Ajoute une feuille, la nomme et place l'onglet en 3eme position 
	Doc.GetSheets.insertNewByName("Regroupement",2*Nb_periode_Max+4)
	
	Sheet = Doc.Sheets.GetByNAme("Regroupement")
	
	
		Range = Sheet.getCellRangeByPosition( 0 , 0 , Nbs+2 , 0 ) 
		Range.Merge( True ) 

		Cell = Sheet.getCellByPosition(0,0)   
		With Cell 
		  .setString( "Regroupement" ) 'insére du texte dans la cellule 
		  .CellBackColor = RGB(255,215,0) 'indique la couleur de fond 
		  '.paraAdjust = com.sun.star.style.ParagraphAdjust.CENTER 'alignement centré 
		  '.RotateAngle = 9000 'Rotation 9000 = 90° 
		  .CharColor = RGB(0,0,0) 'couleur des caractères 
		  .CharHeight = 14 'Taille catactères 
		  .CharWeight = com.sun.star.awt.FontWeight.BOLD 'gras 
		  .CharPosture = com.sun.star.awt.FontSlant.ITALIC 'italique 
		  .CharFontName = "Arial" 'Font 
		 ' .CharUnderline = com.sun.star.awt.FontUnderline.DOUBLE 'souligné double 
		End With 
		  
		  
		'Ajuste la hauteur de la 5eme ligne au contenu des cellules.  
		Sheet.Rows(0).OptimalHeight = True 
		
	Cell = Sheet.getCellByPosition(0,2)
		Cell.formula = "=$'Liste Etudiants'.A1"
	Range = Sheet.getCellRangeByPosition(0,2,2,nbe+2)
		Range.fillSeries(com.sun.star.sheet.FillDirection.TO_BOTTOM,com.sun.star.sheet.FillMode.SIMPLE,0,0,0)
		Range.fillSeries(com.sun.star.sheet.FillDirection.TO_RIGHT,com.sun.star.sheet.FillMode.SIMPLE,0,0,0)
		
 
 
	Cell = Sheet.getCellByPosition(3,1)
	Cell.formula = "=$'Place Stage'.C1"
	Range = Sheet.getCellRangeByPosition(3,1,2+Nbs,2)
	Range.fillSeries(com.sun.star.sheet.FillDirection.TO_BOTTOM,com.sun.star.sheet.FillMode.SIMPLE,0,0,0)
	Range.fillSeries(com.sun.star.sheet.FillDirection.TO_RIGHT,com.sun.star.sheet.FillMode.SIMPLE,0,0,0)


		Range = Sheet.getCellRangeByPosition(3,1,2+Nbs,2)   
		With Range 
		  '.setString( "Notes Bruts" ) 'insére du texte dans la cellule 
		  .CellBackColor = RGB(218,165,32) 'indique la couleur de fond 
		  .paraAdjust = com.sun.star.style.ParagraphAdjust.CENTER 'alignement centré 
		  '.RotateAngle = 9000 'Rotation 9000 = 90° 
		  .CharColor = RGB(0,0,0) 'couleur des caractères 
		  .CharHeight = 10 'Taille catactères 
		  '.CharWeight = com.sun.star.awt.FontWeight.BOLD 'gras 
		  '.CharPosture = com.sun.star.awt.FontSlant.ITALIC 'italique 
		  .CharFontName = "Arial" 'Font 
		 ' .CharUnderline = com.sun.star.awt.FontUnderline.DOUBLE 'souligné double 
		End With 
	
	Sheet.Rows(1).OptimalHeight = True 
		

	Range = Sheet.getCellRangeByPosition(0,2,2,NBe+2)   
		With Range 
		  '.setString( "Notes Bruts" ) 'insére du texte dans la cellule 
		  .CellBackColor = RGB(240,248,255) 'indique la couleur de fond 
		  '.paraAdjust = com.sun.star.style.ParagraphAdjust.CENTER 'alignement centré 
		  '.RotateAngle = 9000 'Rotation 9000 = 90° 
		  .CharColor = RGB(0,0,0) 'couleur des caractères 
		  .CharHeight = 10 'Taille catactères 
		  '.CharWeight = com.sun.star.awt.FontWeight.BOLD 'gras 
		  .CharPosture = com.sun.star.awt.FontSlant.ITALIC 'italique 
		  .CharFontName = "Arial" 'Font 
		 ' .CharUnderline = com.sun.star.awt.FontUnderline.DOUBLE 'souligné double 
		End With 
	For numl =1 to NBe-1
	Sheet.Rows(1+numl).OptimalHeight = True
	Next numl

	For numl=0 to 3
		Sheet.columns(numl).OptimalWidth = True 
	Next numl
	
	
	
	
	
	
	
	
	Dim FormuleBoucle as string
	FormuleBoucle = "="
	For Nb_periode = 1 to Nb_periode_Max
		If Nb_periode < Nb_periode_Max 	then FormuleBoucle = FormuleBoucle & "$'Stage " & Nb_periode & "'.B6 +" 	else FormuleBoucle = FormuleBoucle & "$'Stage " & Nb_periode & "'.B6-$Réouverture.D4+$Précédents.D4"
	Next Nb_periode

	Cell = Sheet.getCellByPosition(3,3)
	Cell.formula = FormuleBoucle
	
	Range = Sheet.getCellRangeByPosition(3,3,Nbs+2,NBe+2)
	Range.fillSeries(com.sun.star.sheet.FillDirection.TO_RIGHT,com.sun.star.sheet.FillMode.SIMPLE,0,0,0)
	Range.fillSeries(com.sun.star.sheet.FillDirection.TO_BOTTOM,com.sun.star.sheet.FillMode.SIMPLE,0,0,0)
	
	
	Cell = Sheet.getCellByPosition(Nbs+3,1)
	Cell.formula = "=MAX($'Place Stage'.2:2)"
	NbGRP = cell.value

dim born1 as string
dim born2 as string
dim born3 as string
dim born4 as string
	
	IF NbGRP>0 Then

	For NumGRP = 1 to NbGRP
	
	'
		Cell = Sheet.getCellByPosition(Nbs+NumGRP+2,1)
		Cell.String = "Regroupement N°"
		With Cell 
		  '.setString( "Nombre d'étudiants sur le terrain" ) 'insére du texte dans la cellule 
		  .CellBackColor = RGB(222,184,135) 'indique la couleur de fond 
		  '.paraAdjust = com.sun.star.style.ParagraphAdjust.CENTER 'alignement centré 
		  '.RotateAngle = 9000 'Rotation 9000 = 90° 
		  .CharColor = RGB(0,0,0) 'couleur des caractères 
		  .CharHeight = 10 'Taille catactères 
		  .CharWeight = com.sun.star.awt.FontWeight.BOLD 'gras 
		  '.CharPosture = com.sun.star.awt.FontSlant.ITALIC 'italique 
		  .CharFontName = "Arial" 'Font 
		 ' .CharUnderline = com.sun.star.awt.FontUnderline.DOUBLE 'souligné double 
		End With
		Cell = Sheet.getCellByPosition(Nbs+NumGRP+2,2)
		Cell.Value = NumGRP
		With Cell 
		  '.setString( "Nombre d'étudiants sur le terrain" ) 'insére du texte dans la cellule 
		  .CellBackColor = RGB(222,184,135) 'indique la couleur de fond 
		  '.paraAdjust = com.sun.star.style.ParagraphAdjust.CENTER 'alignement centré 
		  '.RotateAngle = 9000 'Rotation 9000 = 90° 
		  .CharColor = RGB(0,0,0) 'couleur des caractères 
		  .CharHeight = 10 'Taille catactères 
		  .CharWeight = com.sun.star.awt.FontWeight.BOLD 'gras 
		  '.CharPosture = com.sun.star.awt.FontSlant.ITALIC 'italique 
		  .CharFontName = "Arial" 'Font 
		 ' .CharUnderline = com.sun.star.awt.FontUnderline.DOUBLE 'souligné double 
		End With
		Sheet.columns(Nbs+NumGRP+3).OptimalWidth = True
		
		Cell = Sheet.getCellByPosition(Nbs+2+NumGRP,3)
		Cell.formulalocal = "=ADRESSE(3;" & Nbs+3 & ";1;;)"
		born1 = Cell.string
		Cell.formulalocal = "=ADRESSE(3;" & Nbs+3+NumGRP & ";2;;)"
		born2 = Cell.string
		Cell.formulalocal = "=ADRESSE(4;" & Nbs+3 & ";3;;)"
		born3 = Cell.string
		Cell.formulalocal = "=ADRESSE(4;" & Nbs+2+NumGRP+1 & ";4;;)"
		born4 = Cell.string
		Cell.formulalocal = "=SOMME.SI($D$3:" & born1 & ";"& born2 &";$D4:" & born3 &")-$Réouverture." & born4
		
		Sheet.columns(Nbs+2+NumGRP).OptimalWidth = True
	Next NumGRP

	Range = Sheet.getCellRangeByPosition(Nbs+3,3,Nbs+NumGRP+2,NBe+2)

	Range.fillSeries(com.sun.star.sheet.FillDirection.TO_BOTTOM,com.sun.star.sheet.FillMode.SIMPLE,0,0,0)	
		
	End If
	
	
''''''''''''''''Feuille Synthèse '''''''''''''''''''''''''''''''''
	
	oDoc=ThisComponent 
	  
	'Ajoute une feuille, la nomme et place l'onglet en 3eme position 
	oDoc.GetSheets.insertNewByName("Synthèse",2*Nb_periode_Max+5)
	
	Sheet = Doc.Sheets.GetByName("Synthèse")
	
	Cell = Sheet.getCellByPosition(0,0)
	Cell.string = "Sous Groupe"
		With Cell 
		 ' .setString( "Nombre d'étudiants sur le terrain" ) 'insére du texte dans la cellule 
		  .CellBackColor = RGB(205,133,63) 'indique la couleur de fond 
		  '.paraAdjust = com.sun.star.style.ParagraphAdjust.CENTER 'alignement centré 
		  '.RotateAngle = 9000 'Rotation 9000 = 90° 
		  .CharColor = RGB(0,0,0) 'couleur des caractères 
		  .CharHeight = 10 'Taille catactères 
		 ' .CharWeight = com.sun.star.awt.FontWeight.BOLD 'gras 
		 .CharPosture = com.sun.star.awt.FontSlant.ITALIC 'italique 
		  .CharFontName = "Arial" 'Font 
		 ' .CharUnderline = com.sun.star.awt.FontUnderline.DOUBLE 'souligné double 
		End With  
	Cell = Sheet.getCellByPosition(0,1)
	Cell.formula = "=$'Liste Etudiants'.$A2 "
		With Cell 
		 ' .setString( "Nombre d'étudiants sur le terrain" ) 'insére du texte dans la cellule 
		  .CellBackColor = RGB(255,250,205) 'indique la couleur de fond 
		  '.paraAdjust = com.sun.star.style.ParagraphAdjust.CENTER 'alignement centré 
		  '.RotateAngle = 9000 'Rotation 9000 = 90° 
		  .CharColor = RGB(0,0,0) 'couleur des caractères 
		  .CharHeight = 10 'Taille catactères 
		 ' .CharWeight = com.sun.star.awt.FontWeight.BOLD 'gras 
		 .CharPosture = com.sun.star.awt.FontSlant.ITALIC 'italique 
		  .CharFontName = "Arial" 'Font 
		 ' .CharUnderline = com.sun.star.awt.FontUnderline.DOUBLE 'souligné double 
		End With  
	Range = Sheet.getCellRangeByPosition(0,1,0,NBe)
	Range.fillSeries(com.sun.star.sheet.FillDirection.TO_BOTTOM,com.sun.star.sheet.FillMode.SIMPLE,0,0,0)
	
	Cell = Sheet.getCellByPosition(1,0)
	Cell.string = "Etudiant"
	With Cell 
		 ' .setString( "Nombre d'étudiants sur le terrain" ) 'insére du texte dans la cellule 
		  .CellBackColor = RGB(205,133,63) 'indique la couleur de fond 
		  '.paraAdjust = com.sun.star.style.ParagraphAdjust.CENTER 'alignement centré 
		  '.RotateAngle = 9000 'Rotation 9000 = 90° 
		  .CharColor = RGB(0,0,0) 'couleur des caractères 
		  .CharHeight = 10 'Taille catactères 
		 ' .CharWeight = com.sun.star.awt.FontWeight.BOLD 'gras 
		 .CharPosture = com.sun.star.awt.FontSlant.ITALIC 'italique 
		  .CharFontName = "Arial" 'Font 
		 ' .CharUnderline = com.sun.star.awt.FontUnderline.DOUBLE 'souligné double 
		End With  
		
	Cell = Sheet.getCellByPosition(1,1)
	Form = "=CONCAT($'Liste Etudiants'.$B2 ;" & chr(34) & chr(0160) & chr(34) &"; $'Liste Etudiants'.$C2)"
	Cell.formula = form
	With Cell 
		 ' .setString( "Nombre d'étudiants sur le terrain" ) 'insére du texte dans la cellule 
		  .CellBackColor = RGB(255,250,205) 'indique la couleur de fond 
		  '.paraAdjust = com.sun.star.style.ParagraphAdjust.CENTER 'alignement centré 
		  '.RotateAngle = 9000 'Rotation 9000 = 90° 
		  .CharColor = RGB(0,0,0) 'couleur des caractères 
		  .CharHeight = 10 'Taille catactères 
		 ' .CharWeight = com.sun.star.awt.FontWeight.BOLD 'gras 
		 .CharPosture = com.sun.star.awt.FontSlant.ITALIC 'italique 
		  .CharFontName = "Arial" 'Font 
		 ' .CharUnderline = com.sun.star.awt.FontUnderline.DOUBLE 'souligné double 
		End With  
		
	Range = Sheet.getCellRangeByPosition(0,1,1,NBe)
	Range.fillSeries(com.sun.star.sheet.FillDirection.TO_BOTTOM,com.sun.star.sheet.FillMode.SIMPLE,0,0,0)
	Sheet.columns(1).OptimalWidth = True
	
		For Nb_periode = 1 to Nb_periode_Max
		
		Sheet = Doc.Sheets.GetByName("Synthèse")	
		Cell = Sheet.getCellByPosition(1+Nb_periode,0)
		Cell.String = "Période " & Nb_periode
		With Cell 
		 ' .setString( "Nombre d'étudiants sur le terrain" ) 'insére du texte dans la cellule 
		  .CellBackColor = RGB(205,133,63) 'indique la couleur de fond 
		  '.paraAdjust = com.sun.star.style.ParagraphAdjust.CENTER 'alignement centré 
		  '.RotateAngle = 9000 'Rotation 9000 = 90° 
		  .CharColor = RGB(0,0,0) 'couleur des caractères 
		  .CharHeight = 10 'Taille catactères 
		 ' .CharWeight = com.sun.star.awt.FontWeight.BOLD 'gras 
		 .CharPosture = com.sun.star.awt.FontSlant.ITALIC 'italique 
		  .CharFontName = "Arial" 'Font 
		 ' .CharUnderline = com.sun.star.awt.FontUnderline.DOUBLE 'souligné double 
		End With  
		
		Cell = Sheet.getCellByPosition(Nb_periode+1,1)
		Cell.formulalocal = "=ADRESSE("& 2*NBe+10 &";" & Nbs+2 & ";4;;)"
		extractform = Cell.string
		'Cell.formula = "=COB"& 2*NBe+10 &":" & extractform & ")"
		Cell.Formula = "=$'Stage "& Nb_periode &"'." & extractform
		With Cell 
		 ' .setString( "Nombre d'étudiants sur le terrain" ) 'insére du texte dans la cellule 
		  .CellBackColor = RGB(255,250,205) 'indique la couleur de fond 
		  '.paraAdjust = com.sun.star.style.ParagraphAdjust.CENTER 'alignement centré 
		  '.RotateAngle = 9000 'Rotation 9000 = 90° 
		  .CharColor = RGB(0,0,0) 'couleur des caractères 
		  .CharHeight = 10 'Taille catactères 
		 ' .CharWeight = com.sun.star.awt.FontWeight.BOLD 'gras 
		 .CharPosture = com.sun.star.awt.FontSlant.ITALIC 'italique 
		  .CharFontName = "Arial" 'Font 
		 ' .CharUnderline = com.sun.star.awt.FontUnderline.DOUBLE 'souligné double 
		End With  
		
		Sheet.columns(1+Nb_periode).OptimalWidth = True
		
		Range = Sheet.getCellRangeByPosition(1+Nb_periode,1,1+Nb_periode,NBe)
		Range.fillSeries(com.sun.star.sheet.FillDirection.TO_BOTTOM,com.sun.star.sheet.FillMode.SIMPLE,0,0,0)
			
	Next Nb_periode

''''Ecriture des coef ''''''''
Dim extractor as string

For Nb_periode =1 to Nb_periode_Max

Sheet = Doc.Sheets(Nb_periode+1)

Cell = Sheet.getCellByPosition(3,nbl+2)
Cell.formulalocal = "=Adresse(" & NBe+8 & ";" & Nbs+2 & ";4)"
Extractor = Cell.string
If Nb_periode = 1 Then Cell.value = 1 Else Cell.formulalocal = "=EXP(MOYENNE($'Coef " & Nb_periode-1 & "'.3:3)*(10/MAX($'Coef " & Nb_periode-1 & "'.3:3)) - $'Stage " & Nb_periode-1 & "'." & extractor & ")*$'Coef " & Nb_periode-1 &"'.D" & NBe+5
Next Nb_periode

Sheet = Doc.Sheets(2)

	For ligne = NBe+4 to 2*NBe+3
	Cell = Sheet.getCellByPosition(3,ligne)
	Cell.value = 1
	next ligne

For Nb_periode =2 to Nb_periode_Max
	Sheet = Doc.Sheets(Nb_periode+1)
	Range = Sheet.getCellRangeByPosition(3,NBe+4,3,2*NBe+3)
	Range.fillSeries(com.sun.star.sheet.FillDirection.TO_BOTTOM,com.sun.star.sheet.FillMode.SIMPLE,0,0,0)
Next Nb_periode

'''''Feuille de travail '''''''''''''''''''''

	Doc.GetSheets.insertNewByName("Travail",2*Nb_periode_Max+6)
	Sheet = Doc.Sheets.GetByName("Travail")

		Range = Sheet.getCellRangeByPosition( 0 , 0 , Nbs+2 , 0 ) 
		Range.Merge( True ) 

		Cell = Sheet.getCellByPosition(0,0)   
		With Cell 
		  .setString( "Notes Brutes" ) 'insére du texte dans la cellule 
		  .CellBackColor = RGB(255,215,0) 'indique la couleur de fond 
		  '.paraAdjust = com.sun.star.style.ParagraphAdjust.CENTER 'alignement centré 
		  '.RotateAngle = 9000 'Rotation 9000 = 90° 
		  .CharColor = RGB(0,0,0) 'couleur des caractères 
		  .CharHeight = 14 'Taille catactères 
		  .CharWeight = com.sun.star.awt.FontWeight.BOLD 'gras 
		  .CharPosture = com.sun.star.awt.FontSlant.ITALIC 'italique 
		  .CharFontName = "Arial" 'Font 
		 ' .CharUnderline = com.sun.star.awt.FontUnderline.DOUBLE 'souligné double 
		End With 
		 
		'Ajuste la hauteur de la 5eme ligne au contenu des cellules.  
		Sheet.Rows(0).OptimalHeight = True 
		
		 'Formule des étudiants pour brut
		Cell = Sheet.getCellByPosition(0,1)
		Cell.formula = "=$'Liste Etudiants'.A1"
			With Cell 
		  '.setString( "Notes Brutes" ) 'insére du texte dans la cellule 
		  .CellBackColor = RGB(151, 200, 255) 'indique la couleur de fond 
		  '.paraAdjust = com.sun.star.style.ParagraphAdjust.CENTER 'alignement centré 
		  '.RotateAngle = 9000 'Rotation 9000 = 90° 
		  .CharColor = RGB(0,0,0) 'couleur des caractères 
		  .CharHeight = 10 'Taille catactères 
		  '.CharWeight = com.sun.star.awt.FontWeight.BOLD 'gras 
		  '.CharPosture = com.sun.star.awt.FontSlant.ITALIC 'italique 
		  .CharFontName = "Arial" 'Font 
		 ' .CharUnderline = com.sun.star.awt.FontUnderline.DOUBLE 'souligné double 
		End With 
		Range = Sheet.getCellRangeByPosition(0,1,2,nbe+1)
		Range.fillSeries(com.sun.star.sheet.FillDirection.TO_BOTTOM,com.sun.star.sheet.FillMode.SIMPLE,0,0,0)
		Range.fillSeries(com.sun.star.sheet.FillDirection.TO_RIGHT,com.sun.star.sheet.FillMode.SIMPLE,0,0,0)
		
		Cell = Sheet.getCellByPosition(3,1)
		Cell.formula = "=$'Place Stage'.C1"
			With Cell 
		  '.setString( "Notes Brutes" ) 'insére du texte dans la cellule 
		  .CellBackColor = RGB(161, 47, 0) 'indique la couleur de fond 
		  '.paraAdjust = com.sun.star.style.ParagraphAdjust.CENTER 'alignement centré 
		  '.RotateAngle = 9000 'Rotation 9000 = 90° 
		  .CharColor = RGB(255,255,255) 'couleur des caractères 
		  .CharHeight = 10 'Taille catactères 
		  .CharWeight = com.sun.star.awt.FontWeight.BOLD 'gras 
		  '.CharPosture = com.sun.star.awt.FontSlant.ITALIC 'italique 
		  .CharFontName = "Arial" 'Font 
		 ' .CharUnderline = com.sun.star.awt.FontUnderline.DOUBLE 'souligné double 
		End With 
		Range = Sheet.getCellRangeByPosition(3,1,2+Nbs,1)
		Range.fillSeries(com.sun.star.sheet.FillDirection.TO_RIGHT,com.sun.star.sheet.FillMode.SIMPLE,0,0,0)
		  
		Range = Sheet.getCellRangeByPosition( 0 , nbe+2 , Nbs+2 , nbe+2 ) 
		Range.Merge( True ) 

		Cell = Sheet.getCellByPosition(0,nbe+2)   
		With Cell 
		  .setString( "Check Stage et Blocs" ) 'insére du texte dans la cellule 
		  .CellBackColor = RGB(255,215,0) 'indique la couleur de fond 
		  '.paraAdjust = com.sun.star.style.ParagraphAdjust.CENTER 'alignement centré 
		  '.RotateAngle = 9000 'Rotation 9000 = 90° 
		  .CharColor = RGB(0,0,0) 'couleur des caractères 
		  .CharHeight = 14 'Taille catactères 
		  .CharWeight = com.sun.star.awt.FontWeight.BOLD 'gras 
		  .CharPosture = com.sun.star.awt.FontSlant.ITALIC 'italique 
		  .CharFontName = "Arial" 'Font 
		 ' .CharUnderline = com.sun.star.awt.FontUnderline.DOUBLE 'souligné double 
		End With 
		 
		'Ajuste la hauteur de la 5eme ligne au contenu des cellules.  
		Sheet.Rows(nbe+2).OptimalHeight = True 
		
		 'Formule des étudiants pour brut
		Cell = Sheet.getCellByPosition(0,nbe+3)
		Cell.formula = "=$'Liste Etudiants'.A1"
		With Cell 
		  '.setString( "Notes Brutes" ) 'insére du texte dans la cellule 
		  .CellBackColor = RGB(151, 200, 255) 'indique la couleur de fond 
		  '.paraAdjust = com.sun.star.style.ParagraphAdjust.CENTER 'alignement centré 
		  '.RotateAngle = 9000 'Rotation 9000 = 90° 
		  .CharColor = RGB(0,0,0) 'couleur des caractères 
		  .CharHeight = 10 'Taille catactères 
		  '.CharWeight = com.sun.star.awt.FontWeight.BOLD 'gras 
		  '.CharPosture = com.sun.star.awt.FontSlant.ITALIC 'italique 
		  .CharFontName = "Arial" 'Font 
		 ' .CharUnderline = com.sun.star.awt.FontUnderline.DOUBLE 'souligné double 
		End With  
		
		Range = Sheet.getCellRangeByPosition(0,nbe+3,2,2*nbe+3)
		Range.fillSeries(com.sun.star.sheet.FillDirection.TO_BOTTOM,com.sun.star.sheet.FillMode.SIMPLE,0,0,0)
		Range.fillSeries(com.sun.star.sheet.FillDirection.TO_RIGHT,com.sun.star.sheet.FillMode.SIMPLE,0,0,0)
		Cell = Sheet.getCellByPosition(3,nbe+3)
		Cell.formula = "=$'Place Stage'.C1"
		With Cell 
		  '.setString( "Notes Brutes" ) 'insére du texte dans la cellule 
		  .CellBackColor = RGB(161, 47, 0) 'indique la couleur de fond 
		  '.paraAdjust = com.sun.star.style.ParagraphAdjust.CENTER 'alignement centré 
		  '.RotateAngle = 9000 'Rotation 9000 = 90° 
		  .CharColor = RGB(255,255,255) 'couleur des caractères 
		  .CharHeight = 10 'Taille catactères 
		  .CharWeight = com.sun.star.awt.FontWeight.BOLD 'gras 
		  '.CharPosture = com.sun.star.awt.FontSlant.ITALIC 'italique 
		  .CharFontName = "Arial" 'Font 
		 ' .CharUnderline = com.sun.star.awt.FontUnderline.DOUBLE 'souligné double 
		End With 
		
		Range = Sheet.getCellRangeByPosition(3,nbe+3,2+Nbs,nbe+3)
		Range.fillSeries(com.sun.star.sheet.FillDirection.TO_RIGHT,com.sun.star.sheet.FillMode.SIMPLE,0,0,0)
	IF NbGRP = 0 Then 
	Cell = Sheet.getCellByPosition(3,nbe+4)	
	
		Cell.formulalocal ="=SI($Regroupement.B3>=1;" &  chr(34) & chr(34) & ";D3)"
		With Cell 
		  '.setString( "Notes Brutes" ) 'insére du texte dans la cellule 
		  .CellBackColor = RGB(211,211,211) 'indique la couleur de fond 
		  '.paraAdjust = com.sun.star.style.ParagraphAdjust.CENTER 'alignement centré 
		  '.RotateAngle = 9000 'Rotation 9000 = 90° 
		  .CharColor = RGB(0,0,0) 'couleur des caractères 
		  .CharHeight = 10 'Taille catactères 
		  .CharWeight = com.sun.star.awt.FontWeight.BOLD 'gras 
		  '.CharPosture = com.sun.star.awt.FontSlant.ITALIC 'italique 
		  .CharFontName = "Arial" 'Font 
		 ' .CharUnderline = com.sun.star.awt.FontUnderline.DOUBLE 'souligné double 
		End With
	Else
	Cell = Sheet.getCellByPosition(3,nbe+4)	
		Cell.formulalocal = "=Adresse(3;" & Nbs+4 & ";1)"
		Born1 = Cell.String
		Cell.formulalocal = "=Adresse(" & 3+Nbe & ";" & Nbs+3+NbGRP & ";1)"
		Born2 = Cell.String
		Cell.formulalocal = "=Adresse(3;" & Nbs+3+NbGRP & ";1)"
		Born5 = Cell.String
	'	Cell.formulalocal ="=SI(RECHERCHEH($'Place Stage'.C$2;$Regroupement."& born1 &":"& born2 &";LIGNE($Regroupement.A3)-1;)>=1;" & chr(34) & chr(34) &" ;SI($Regroupement.B3>=1;" &  chr(34) & chr(34) & ";D3))"
				
		'Formule initiale : Cell.formulalocal = "=SI(ESTERREUR(RECHERCHEH($'Place Stage'.C$2;$Regroupement." & born1 & ":" & born2 & ";LIGNE($Regroupement.A3)-1;)>=1);SI($Regroupement.D4>=1;" &  chr(34) & chr(34) & ";D3);SI(RECHERCHEH($'Place Stage'.C$2;$Regroupement." & born1 &":"& born2 &";LIGNE($Regroupement.A3)-1;)>=1;" &  chr(34) & chr(34) & ";SI($Regroupement.D4>=1;" &  chr(34) & chr(34) & ";D3)))"
		Cell.formulalocal = "=SI(ESTERREUR(RECHERCHEH($'Place Stage'.C$2;$Regroupement." & born1 & ":" & born2 & ";LIGNE($Regroupement.A3)-1;)>=SI(SOMME.SI($Regroupement." & born1 & ":" & born5 & ";$'Place Stage'.C$2;$Regroupement." & born1 & ":" & born5 &")>0;SOMME.SI($Regroupement." & born1 & ":" & born5 & ";$'Place Stage'.C$2;$Regroupement." & born1 & ":" & born5 & ");1)" & ");SI($Regroupement.D4>=1;" &  chr(34) & chr(34) & ";D3);SI(RECHERCHEH($'Place Stage'.C$2;$Regroupement." & born1 &":"& born2 &";LIGNE($Regroupement.A3)-1;)>=SI(SOMME.SI($Regroupement."& born1 & ":" & born5 & ";$'Place Stage'.C$2;$Regroupement." &born1 & ":" & born5 &")>0;SOMME.SI($Regroupement." & born1 & ":" & born5 & ";$'Place Stage'.C$2;$Regroupement." & born1 & ":" & born5 & ");1)" &" ;" &  chr(34) & chr(34) & ";SI($Regroupement.D4>=1;" &  chr(34) & chr(34) & ";D3)))"
	
		With Cell 
		  '.setString( "Notes Brutes" ) 'insére du texte dans la cellule 
		  .CellBackColor = RGB(211,211,211) 'indique la couleur de fond 
		  '.paraAdjust = com.sun.star.style.ParagraphAdjust.CENTER 'alignement centré 
		  '.RotateAngle = 9000 'Rotation 9000 = 90° 
		  .CharColor = RGB(0,0,0) 'couleur des caractères 
		  .CharHeight = 10 'Taille catactères 
		  .CharWeight = com.sun.star.awt.FontWeight.BOLD 'gras 
		  '.CharPosture = com.sun.star.awt.FontSlant.ITALIC 'italique 
		  .CharFontName = "Arial" 'Font 
		 ' .CharUnderline = com.sun.star.awt.FontUnderline.DOUBLE 'souligné double 
		End With 		
	End IF	
	Range = Sheet.getCellRangeByPosition(3,nbe+4,2+NbS,2*nbe+3)
		Range.fillSeries(com.sun.star.sheet.FillDirection.TO_BOTTOM,com.sun.star.sheet.FillMode.SIMPLE,0,0,0)
		Range.fillSeries(com.sun.star.sheet.FillDirection.TO_RIGHT,com.sun.star.sheet.FillMode.SIMPLE,0,0,0)	

'''''''''''''''''''''' Feuille Précédents ''''''''''''''''''''''''''''''''''''''''

	Doc=ThisComponent 
	  
	'Ajoute une feuille, la nomme et place l'onglet en 3eme position 
Doc.GetSheets.insertNewByName("Précédents",2*Nb_periode_Max+7)
	
Sheet = Doc.Sheets.GetByNAme("Précédents")
	
	
		Range = Sheet.getCellRangeByPosition( 0 , 0 , Nbs+2 , 0 ) 
		Range.Merge( True ) 

		Cell = Sheet.getCellByPosition(0,0)   
		With Cell 
.setString( "Précédents" ) 'insére du texte dans la cellule 
		  .CellBackColor = RGB(255,215,0) 'indique la couleur de fond 
		  '.paraAdjust = com.sun.star.style.ParagraphAdjust.CENTER 'alignement centré 
		  '.RotateAngle = 9000 'Rotation 9000 = 90° 
		  .CharColor = RGB(0,0,0) 'couleur des caractères 
		  .CharHeight = 14 'Taille catactères 
		  .CharWeight = com.sun.star.awt.FontWeight.BOLD 'gras 
		  .CharPosture = com.sun.star.awt.FontSlant.ITALIC 'italique 
		  .CharFontName = "Arial" 'Font 
		 ' .CharUnderline = com.sun.star.awt.FontUnderline.DOUBLE 'souligné double 
		End With 
		  
		  
		'Ajuste la hauteur de la 5eme ligne au contenu des cellules.  
		Sheet.Rows(0).OptimalHeight = True 
		
	Cell = Sheet.getCellByPosition(0,2)
		Cell.formula = "=$'Liste Etudiants'.A1"
	Range = Sheet.getCellRangeByPosition(0,2,2,nbe+2)
		Range.fillSeries(com.sun.star.sheet.FillDirection.TO_BOTTOM,com.sun.star.sheet.FillMode.SIMPLE,0,0,0)
		Range.fillSeries(com.sun.star.sheet.FillDirection.TO_RIGHT,com.sun.star.sheet.FillMode.SIMPLE,0,0,0)
		
 
 
	Cell = Sheet.getCellByPosition(3,1)
	Cell.formula = "=$'Place Stage'.C1"
	Range = Sheet.getCellRangeByPosition(3,1,2+Nbs,2)
	Range.fillSeries(com.sun.star.sheet.FillDirection.TO_BOTTOM,com.sun.star.sheet.FillMode.SIMPLE,0,0,0)
	Range.fillSeries(com.sun.star.sheet.FillDirection.TO_RIGHT,com.sun.star.sheet.FillMode.SIMPLE,0,0,0)


		Range = Sheet.getCellRangeByPosition(3,1,2+Nbs,2)   
		With Range 
		  '.setString( "Notes Bruts" ) 'insére du texte dans la cellule 
		  .CellBackColor = RGB(218,165,32) 'indique la couleur de fond 
		  .paraAdjust = com.sun.star.style.ParagraphAdjust.CENTER 'alignement centré 
		  '.RotateAngle = 9000 'Rotation 9000 = 90° 
		  .CharColor = RGB(0,0,0) 'couleur des caractères 
		  .CharHeight = 10 'Taille catactères 
		  '.CharWeight = com.sun.star.awt.FontWeight.BOLD 'gras 
		  '.CharPosture = com.sun.star.awt.FontSlant.ITALIC 'italique 
		  .CharFontName = "Arial" 'Font 
		 ' .CharUnderline = com.sun.star.awt.FontUnderline.DOUBLE 'souligné double 
		End With 
	
	Sheet.Rows(1).OptimalHeight = True 
		

	Range = Sheet.getCellRangeByPosition(0,2,2,NBe+2)   
		With Range 
		  '.setString( "Notes Bruts" ) 'insére du texte dans la cellule 
		  .CellBackColor = RGB(240,248,255) 'indique la couleur de fond 
		  '.paraAdjust = com.sun.star.style.ParagraphAdjust.CENTER 'alignement centré 
		  '.RotateAngle = 9000 'Rotation 9000 = 90° 
		  .CharColor = RGB(0,0,0) 'couleur des caractères 
		  .CharHeight = 10 'Taille catactères 
		  '.CharWeight = com.sun.star.awt.FontWeight.BOLD 'gras 
		  .CharPosture = com.sun.star.awt.FontSlant.ITALIC 'italique 
		  .CharFontName = "Arial" 'Font 
		 ' .CharUnderline = com.sun.star.awt.FontUnderline.DOUBLE 'souligné double 
		End With 
	For numl =1 to NBe-1
	Sheet.Rows(1+numl).OptimalHeight = True
	Next numl

	For numl=0 to 3
		Sheet.columns(numl).OptimalWidth = True 
	Next numl

End Sub
