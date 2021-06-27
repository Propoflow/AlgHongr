REM  *****  BASIC  *****

Sub OptVersMacro
	'Cahier de variables :
	Dim NBe as integer 'Nombre d'étudiants : basé sur la liste sur feuille #1
	Dim Nbs as integer 'Nombre de terrains de stages : basé sur la liste des stages sur la feuille #2
	Dim cell_dest as integer
	Dim nbl as integer 'Nombre de lignes d'un tableau = nombre d'étudiants + Lignes d'en têtes 
	Dim nbc as integer 'Nombre de colonnes d'un tableau = nombre de stage + Colonne d'en têtes
	
	Dim Nb_periode as integer 'Compteur de la période en cours de travail
	Dim Nb_periode_Max as integer 'Nombre de périodes pour lesquels créer des feuilles = demandé à l'utilisateur
	Dim Num_Feuille as integer 'Numéro de la feuille en cours de travail
	Num_Feuille = 3
	
	' Feuille 1 : Liste étudiants
	' Feuille 2 : Liste stages
	' Feuille 2 + Nb_periode à 2 + Nb_periode_Max : Feuille Coef 1 à Nb_periode_Max
	' Feuille 2 + Nb_periode_Max à 2 + 2* Nb_periode_Max : Feuille Stage 1 à Nb_periode_Max
	' Feuille 2 + 2* Nb_periode_Max + 1 : Feuille Réouverture
	' Feuille 2 + 2* Nb_periode_Max + 2 : Feuille Regroupement
	' Feuille 2 + 2* Nb_periode_Max + 3 : Feuille Synthèse
	' Feuille 2 + 2* Nb_periode_Max + 4 : Feuille Travail
	' Feuille 2 + 2* Nb_periode_Max + 5 : Feuille Précédents
	Dim numl as integer 'Compteur de la ligne en cours de modification
	
	Dim Extract_1 as string 'Element de stockage de coordonnées sous forme Ax pour formules
	Dim Extract_2 as string 'Element de stockage de coordonnées sous forme Ax pour formules
	Dim Extract_HA as string 'Element de stockage de coordonnées sous forme Ax pour formules
	dim offset as integer 'Décalage sur le renvoi vers la feuille stage entre les période 1 et n 
	
	dim born1 as string 'Element de stockage de coordonnées sous forme Ax pour formules
	dim born2 as string 'Element de stockage de coordonnées sous forme Ax pour formules
	dim born3 as string 'Element de stockage de coordonnées sous forme Ax pour formules
	dim born4 as string 'Element de stockage de coordonnées sous forme Ax pour formules
	dim born5 as string 'Element de stockage de coordonnées sous forme Ax pour formules
	
	Dim FormuleBoucle as string ' Formule en boucle de la feuille regroupement
	'Références relatives dans la macro
	Dim Doc As Object
	Dim Sheet As Object
	Dim Cell As object
	Dim oCurseur As Object 
	Dim feuille_dest as object
	Dim Range as object
	Dim CellRangeAddress As New com.sun.star.table.CellRangeAddress
	Dim CellAddress As New com.sun.star.table.CellAddress

	'Initilisation :
	MsgBox "La Macro va s'initialiser, pour cela il faut que vous ayez rempli les feuilles Terrains de stage et Liste etudiants correctement"
	' Nombre de périodes de stages 
	NB_periode_Max = InputBox ("Veuillez entrer le nombre de périodes : ","Chère utilisatrice, cher utilisateur")
	MsgBox ( Nb_periode_Max , 64, "Confirmation de formule")
	'Nombre d'étudiants et de stages
	oCurseur = ThisComponent.Sheets(0).createCursor  
	oCurseur.gotoEndOfUsedArea( False )   
	NBe = oCurseur.RangeAddress.EndRow
	oCurseur = ThisComponent.Sheets(1).createCursor  
	oCurseur.gotoEndOfUsedArea( False ) 
	NBs = oCurseur.RangeAddress.EndColumn-1
	Doc = ThisComponent
	'Fin d'initialisation
	
' Feuille Coefficient Période X
For Nb_periode=1 to Nb_periode_Max
	Doc.GetSheets.insertNewByName("Coef " & Nb_periode,Num_Feuille) 'Ajoute une feuille, la nomme et place l'onglet en position Num_feuille
	'Commande suivante devant être executé depuis la feuille : erreur depuis la console
	Sheet = Doc.sheets.getByName("Coef " & Nb_periode)
	
	'Selectionne les cellules de A1 à la colonne correspondant à (Nombre de Stages + colonnes d'en-tête) puis définit la ligne de titre du tableau
		Range = Sheet.getCellRangeByPosition( 0 , 0 , Nbs+3 , 0 ) 
		Range.Merge( True ) 'Les fusionne
		Cell = Sheet.getCellByPosition(0,0)   
		With Cell 
		  .setString( "Notes Brutes" ) 'insére du texte dans la cellule 
		  .CellBackColor = RGB(255,215,0) 'indique la couleur de fond 
		  .CharColor = RGB(0,0,0) 'couleur des caractères 
		  .CharHeight = 14 'Taille catactères 
		  .CharWeight = com.sun.star.awt.FontWeight.BOLD 'gras 
		  .CharPosture = com.sun.star.awt.FontSlant.ITALIC 'italique 
		  .CharFontName = "Arial" 'Font 
		End With 
		Sheet.Rows(0).OptimalHeight = True 
		
	'Création des en-tête du tableau brut et généralisation des formules
	' Liste des étudiants :
		Cell = Sheet.getCellByPosition(0,1)
		Cell.formula = "=$'Liste Etudiants'.A1"
		Range = Sheet.getCellRangeByPosition(0,1,2,nbe+1)
		Range.fillSeries(com.sun.star.sheet.FillDirection.TO_BOTTOM,com.sun.star.sheet.FillMode.SIMPLE,0,0,0)
		Range.fillSeries(com.sun.star.sheet.FillDirection.TO_RIGHT,com.sun.star.sheet.FillMode.SIMPLE,0,0,0)
	' Colonne coefficient :
		Cell = Sheet.getCellByPosition(3,1) 
		Cell.String = "Coefficients" 		
	'Liste des Terrains de stages :
		Cell = Sheet.getCellByPosition(4,1)
		Cell.formula = "=$'Place Stage'.C1"
		Range = Sheet.getCellRangeByPosition(4,1,3+Nbs,1)
		Range.fillSeries(com.sun.star.sheet.FillDirection.TO_RIGHT,com.sun.star.sheet.FillMode.SIMPLE,0,0,0)

	'Définition des nombres de lignes de de colonne d'un tableau 
	oCurseur = Sheet.createCursor  
	oCurseur.gotoEndOfUsedArea( False ) 
  	nbl = oCurseur.RangeAddress.EndRow+1
  	nbc = oCurseur.RangeAddress.EndColumn+1
	
	'Mise en forme en-tête stage et étudiants
	Range = Sheet.getCellRangeByPosition(4,1,3+Nbs,1)   
		With Range 
		  .CellBackColor = RGB(218,165,32) 'indique la couleur de fond 
		  .paraAdjust = com.sun.star.style.ParagraphAdjust.CENTER 'alignement centré 
		  .CharColor = RGB(0,0,0) 'couleur des caractères 
		  .CharHeight = 10 'Taille catactères 
		  .CharFontName = "Arial" 'Font 
		End With 
		Sheet.Rows(1).OptimalHeight = True 
	Range = Sheet.getCellRangeByPosition(0,2,3,NBe+1)   
		With Range 
		  .CellBackColor = RGB(240,248,255) 'indique la couleur de fond 
		  .CharColor = RGB(0,0,0) 'couleur des caractères 
		  .CharHeight = 10 'Taille catactères 
		  .CharPosture = com.sun.star.awt.FontSlant.ITALIC 'italique 
		  .CharFontName = "Arial" 'Font 
		End With 
	For numl =1 to NBe-1
		Sheet.Rows(1+numl).OptimalHeight = True
	Next numl
	For numl=0 to 3
		Sheet.columns(numl).OptimalWidth = True 
	Next numl
	
	'Tableau "normalisé"
		Plage = Sheet.getCellRangeByPosition( 0 , nbl , Nbs+3 , nbl ) 
		Plage.Merge( True ) 
	Cell = Sheet.getCellByPosition(0,nbl)  'Création titre du tableau "normalisé"
		With Cell 
			  .setString( "Notes Normalisées" ) 'insére du texte dans la cellule 
			  .CellBackColor = RGB(255,215,0) 'indique la couleur de fond 
			  .CharColor = RGB(0,0,0) 'couleur des caractères 
			  .CharHeight = 14 'Taille catactères 
			  .CharWeight = com.sun.star.awt.FontWeight.BOLD 'gras 
			  .CharPosture = com.sun.star.awt.FontSlant.ITALIC 'italique 
			  .CharFontName = "Arial" 'Font 
		End With  
		Sheet.Rows(nbl).OptimalHeight = True 
	
	'Renvoi vers le premier tableau
	Cell = Sheet.getCellByPosition(0,nbl+1)
		Cell.formula = "=A2"
	Range = Sheet.getCellRangeByPosition(0,nbl+1,Nbs+3,2*nbl-1)
		Range.fillSeries(com.sun.star.sheet.FillDirection.TO_BOTTOM,com.sun.star.sheet.FillMode.SIMPLE,0,0,0)
		Range.fillSeries(com.sun.star.sheet.FillDirection.TO_RIGHT,com.sun.star.sheet.FillMode.SIMPLE,0,0,0)
	
	'Mise en forme en têtes tableau normalisé
	Range = Sheet.getCellRangeByPosition(0,nbl+1,3,2*nbl-1)
		With Range 
			.CellBackColor = RGB(240,248,255) 'indique la couleur de fond 
			.CharColor = RGB(0,0,0) 'couleur des caractères 
			.CharPosture = com.sun.star.awt.FontSlant.ITALIC 'italique 
			.CharFontName = "Arial" 'Font 
		End With  
	Range = Sheet.getCellRangeByPosition(4,nbl+1,3+Nbs,nbl+1)
		With Range 
			.CellBackColor = RGB(218,165,32) 'indique la couleur de fond 
			.CharColor = RGB(0,0,0) 'couleur des caractères 
			.CharWeight = com.sun.star.awt.FontWeight.BOLD 'gras 
			.CharPosture = com.sun.star.awt.FontSlant.ITALIC 'italique 
			.CharFontName = "Arial" 'Font 
		End With  
	
	'Ecriture de la formule cellule 1,1 Normalisé + Généralisation au reste du tableau
	Cell = Sheet.getCellByPosition(4,nbl+2) 
	form = "=SI(ESTERREUR( E3 * $D"& nbl+3 & "*10/Max(3:3));0; E3 * $D"& nbl+3 & "*10/Max(3:3))"
	Cell.formulalocal = form 
	
	Range = Sheet.getCellRangeByPosition(4,nbl+2,Nbs+3,nbl+NBe+1)
		Range.fillSeries(com.sun.star.sheet.FillDirection.TO_BOTTOM,com.sun.star.sheet.FillMode.SIMPLE,0,0,0)
		Range.fillSeries(com.sun.star.sheet.FillDirection.TO_RIGHT,com.sun.star.sheet.FillMode.SIMPLE,0,0,0)

Num_Feuille = Num_Feuille + 1 ' Incrémentation de la feuille de travail
Next Nb_periode

' Feuille Stage Période X
For Nb_periode=1 to Nb_periode_Max
	  
	Doc.GetSheets.insertNewByName("Stage " & Nb_periode,Num_Feuille) 'Ajoute une feuille, la nomme et place l'onglet en position 2 + Nb_periode_Max +1
	Sheet = Doc.sheets.getByName("Stage " & Nb_periode)
	'Tableau binaire
	'Ecriture des en-têtes et titre du tableau
		Plage = Sheet.getCellRangeByPosition( 1 , 0 , Nbs+3 , 0 ) 
			Plage.Merge( True ) 
		Cell = Sheet.getCellByPosition(1,0) 
		With Cell 
			.setString( "Tableau Binaire" ) 'insére du texte dans la cellule 
			.CellBackColor = RGB(255,215,0) 'indique la couleur de fond 
			.CharColor = RGB(0,0,0) 'couleur des caractères 
			.CharHeight = 14 'Taille catactères 
			.CharWeight = com.sun.star.awt.FontWeight.BOLD 'gras 
			.CharPosture = com.sun.star.awt.FontSlant.ITALIC 'italique 
			.CharFontName = "Arial" 'Font  
		End With 
		Sheet.Rows(0).OptimalHeight = True 
	
		Cell = Sheet.getCellByPosition(0,0) 
		With Cell 
			.setString( "Fonction Solveur" ) 'insére du texte dans la cellule 
			.CellBackColor = RGB(60,179,113) 'indique la couleur de fond 
			.paraAdjust = com.sun.star.style.ParagraphAdjust.CENTER 'alignement centré 
			.CharColor = RGB(0,0,0) 'couleur des caractères 
			.CharHeight = 14 'Taille catactères 
			.CharWeight = com.sun.star.awt.FontWeight.BOLD 'gras
			.CharFontName = "Arial" 'Font 
		End With  
	
	'Récupération des coordonnées de la cellule de fin du tableau pour formule de la fonction solveur 
		Cell = Sheet.getCellByPosition(0,1) 
			Cell.formulalocal = "=Adresse("& 2*NBe+4 & ";" & Nbs+4 & ";4)"
				Extract_1 = cell.string
			Cell.formulalocal = "=Adresse("& NBe+5 & ";" & Nbs+1 & ";4)"
				Extract_2 = cell.string
			Cell.Formulalocal = "=SOMMEPROD($'Coef " & Nb_periode & "'.E" & (Nbe+5) & ":$'Coef " & Nb_periode & "'." & Extract_1 & ";B6:" & Extract_2 &")"

		With Cell 
			.CellBackColor = RGB(144,238,144) 'indique la couleur de fond 
			.paraAdjust = com.sun.star.style.ParagraphAdjust.CENTER 'alignement centré 
			.CharColor = RGB(0,0,0) 'couleur des caractères 
			.CharHeight = 14 'Taille catactères 
			.CharWeight = com.sun.star.awt.FontWeight.BOLD 'gras 
			.CharFontName = "Arial" 'Font 
		 End With 
	'Suite des en-têtes
		Cell = Sheet.getCellByPosition(0,2) 
		With Cell 
			.setString( "Minimum" ) 'insére du texte dans la cellule 
			.CellBackColor = RGB(240,128,128) 'indique la couleur de fond 
			.CharColor = RGB(0,0,0) 'couleur des caractères 
			.CharHeight = 10 'Taille catactères 
			.CharWeight = com.sun.star.awt.FontWeight.BOLD 'gras 
			.CharFontName = "Arial" 'Font 
		End With 
		Cell = Sheet.getCellByPosition(0,3)
		With Cell 
			.setString( "Maximum" ) 'insére du texte dans la cellule 
			.CellBackColor = RGB(205,92,92) 'indique la couleur de fond 
			.CharColor = RGB(0,0,0) 'couleur des caractères 
			.CharHeight = 10 'Taille catactères 
			.CharWeight = com.sun.star.awt.FontWeight.BOLD 'gras 
			.CharFontName = "Arial" 'Font 
		End With  
		Cell = Sheet.getCellByPosition(0,4)
		With Cell 
			.setString( "Affectés" ) 'insére du texte dans la cellule 
			.CellBackColor = RGB(178,34,34) 'indique la couleur de fond 
			.CharColor = RGB(255,255,255) 'couleur des caractères 
			.CharHeight = 10 'Taille catactères 
			.CharWeight = com.sun.star.awt.FontWeight.BOLD 'gras 
			.CharFontName = "Arial" 'Font 
		End With  
		Cell = Sheet.getCellByPosition(1,1)
			form = "=$'Place Stage'.C$1"
			Cell.formula = form
		With Cell 
			.CellBackColor = RGB(139,0,0) 'indique la couleur de fond 
			.CharColor = RGB(255,255,255) 'couleur des caractères 
			.CharHeight = 10 'Taille catactères 
			.CharWeight = com.sun.star.awt.FontWeight.BOLD 'gras 
			.CharFontName = "Arial" 'Font 
		End With  
		Cell = Sheet.getCellByPosition(1,2)
			form = "=$'Place Stage'.C" & 3+offset
			Cell.formula = form
		With Cell 
			.CellBackColor = RGB(240,128,128) 'indique la couleur de fond 
			.CharColor = RGB(0,0,0) 'couleur des caractères 
			.CharHeight = 10 'Taille catactères 
			.CharFontName = "Arial" 'Font 
		End With 
		Cell = Sheet.getCellByPosition(1,3)
			form = "=$'Place Stage'.C" & 4+offset	
			Cell.formula = form
		With Cell 
			.CellBackColor = RGB(205,92,92) 'indique la couleur de fond 
			.CharColor = RGB(0,0,0) 'couleur des caractères 
			.CharHeight = 10 'Taille catactères 
			.CharFontName = "Arial" 'Font 
		End With
		Cell = Sheet.getCellByPosition(1,4)
			Cell.formula = "=SUM(B6:B" & 5+NBe & ")"
		With Cell 
			.CellBackColor = RGB(178,34,34) 'indique la couleur de fond 
			.CharColor = RGB(255,255,255) 'couleur des caractères 
			.CharHeight = 10 'Taille catactères 
			.CharFontName = "Arial" 'Font 
		End With  
		Sheet.columns(0).OptimalWidth = True
	
	'Généralisation formule en-tête	stage et étudiants
	Range = Sheet.getCellRangeByPosition(1,1,Nbs,4)
	Range.fillSeries(com.sun.star.sheet.FillDirection.TO_RIGHT,com.sun.star.sheet.FillMode.SIMPLE,0,0,0)

		Cell = Sheet.getCellByPosition(0,5)
		Form = "=CONCAT($'Liste Etudiants'.$B2 ;" & chr(34) & chr(0160) & chr(34) &"; $'Liste Etudiants'.$C2)"
		Cell.formula = form
		With Cell 
			.CellBackColor = RGB(255,250,205) 'indique la couleur de fond 
			.CharColor = RGB(0,0,0) 'couleur des caractères 
			.CharHeight = 10 'Taille catactères 
			.CharPosture = com.sun.star.awt.FontSlant.ITALIC 'italique 
			.CharFontName = "Arial" 'Font 
		End With  
	Range = Sheet.getCellRangeByPosition(0,5,0,NBe+4)
	Range.fillSeries(com.sun.star.sheet.FillDirection.TO_BOTTOM,com.sun.star.sheet.FillMode.SIMPLE,0,0,0)

	'Récupération des coordonnées de la dernières cellule de la ligne & de la cellule hors algo pour formule de la somme de la ligne 
		Cell = Sheet.getCellByPosition(Nbs+1,5)
			Cell.formulalocal = "=ADRESSE(6;" & Nbs+1 & ";4;; )"
			Extract_HA = Cell.string
			Cell.formulalocal = "=ADRESSE(6;" & Nbs+3 & ";4;; )"
			Extract_2 = Cell.string 
		Cell.formula = "=SUM(B6:" & Extract_1 & ")+" & Extract_2
	
	'Création des colonnes Hors algo
		Cell = Sheet.getCellByPosition(Nbs+2,4)
		With Cell 
			.setString( "Hors Algo" ) 'insére du texte dans la cellule 
			.CellBackColor = RGB(26, 188, 156) 'indique la couleur de fond 
			.CharColor = RGB(255,255,255) 'couleur des caractères 
			.CharHeight = 10 'Taille catactères 
			.CharWeight = com.sun.star.awt.FontWeight.BOLD 'gras 
			.CharFontName = "Arial" 'Font 
		End With  
		Cell = Sheet.getCellByPosition(Nbs+3,4)
		With Cell 
			.setString( "Commentaire" ) 'insére du texte dans la cellule 
			.CellBackColor = RGB(26, 188, 156) 'indique la couleur de fond 
			.CharColor = RGB(255,255,255) 'couleur des caractères 
			.CharHeight = 10 'Taille catactères 
			.CharWeight = com.sun.star.awt.FontWeight.BOLD 'gras 
			.CharFontName = "Arial" 'Font 
		End With  
		Cell = Sheet.getCellByPosition(Nbs+2,5)
			Cell.formulalocal = "=ADRESSE(6;" & Nbs+4 & ";3;; )"
			Extract_1 = Cell.string
		Cell.formulalocal = "=SI(ESTVIDE(" & Extract_1 & ");0;1)"
		With Cell 
			.CellBackColor = RGB(130, 224, 170) 'indique la couleur de fond 
			.CharColor = RGB(0,0,0) 'couleur des caractères 
			.CharHeight = 10 'Taille catactères 
			.CharFontName = "Arial" 'Font 
		End With  
		Cell = Sheet.getCellByPosition(Nbs+3,5)
		With Cell 
			.CellBackColor = RGB(130, 224, 170) 'indique la couleur de fond 
			.CharColor = RGB(0,0,0) 'couleur des caractères 
			.CharHeight = 10 'Taille catactères 
			.CharFontName = "Arial" 'Font 
		End With 
	'Généralisation à la colonne hors algo
	Range = Sheet.getCellRangeByPosition(Nbs+1,5,Nbs+3,NBe+4)
	Range.fillSeries(com.sun.star.sheet.FillDirection.TO_BOTTOM,com.sun.star.sheet.FillMode.SIMPLE,0,0,0)	

	'Taleau Satisafaction	
	'En têtes du tableau
		Plage = Sheet.getCellRangeByPosition( 0 , NBe+5 , Nbs+3 , NBe+5 ) 
			Plage.Merge( True ) 	
		Cell = Sheet.getCellByPosition(0,NBe+5)
		With Cell 
			.setString( "Tableau Satisfaction" ) 'insére du texte dans la cellule 
			.CellBackColor = RGB(255,215,0) 'indique la couleur de fond 
			.CharColor = RGB(0,0,0) 'couleur des caractères 
			.CharHeight = 14 'Taille catactères 
			.CharWeight = com.sun.star.awt.FontWeight.BOLD 'gras 
			.CharPosture = com.sun.star.awt.FontSlant.ITALIC 'italique 
			.CharFontName = "Arial" 'Font 
		End With  
		Cell = Sheet.getCellByPosition(1,NBe+6) 
			Cell.formula = "=$'Place Stage'.C$1"
		With Cell 
		.CellBackColor = RGB(139,0,0) 'indique la couleur de fond 
		.CharColor = RGB(255,255,255) 'couleur des caractères 
		.CharHeight = 10 'Taille catactères 
		.CharWeight = com.sun.star.awt.FontWeight.BOLD 'gras 
		.CharFontName = "Arial" 'Font 
		End With 
	Range = Sheet.getCellRangeByPosition(1,NBe+6,Nbs,NBe+6)
	Range.fillSeries(com.sun.star.sheet.FillDirection.TO_RIGHT,com.sun.star.sheet.FillMode.SIMPLE,0,0,0)

		Cell = Sheet.getCellByPosition(0,NBe+7)
			Form = "=CONCAT($'Liste Etudiants'.$B2 ;" & chr(34) & chr(0160) & chr(34) &"; $'Liste Etudiants'.$C2)"
		Cell.formula = form
		With Cell 
			.CellBackColor = RGB(255,250,205) 'indique la couleur de fond 
			.CharColor = RGB(0,0,0) 'couleur des caractères 
			.CharHeight = 10 'Taille catactères 
			.CharPosture = com.sun.star.awt.FontSlant.ITALIC 'italique 
			.CharFontName = "Arial" 'Font 
		End With  
	Range = Sheet.getCellRangeByPosition(0,NBe+7,0,2*NBe+7)
	Range.fillSeries(com.sun.star.sheet.FillDirection.TO_BOTTOM,com.sun.star.sheet.FillMode.SIMPLE,0,0,0)
	'Formule de renvoie de la note obtenue si stage correspondant affecté + Généralisation
		Cell = Sheet.getCellByPosition(1,NBe+7)
			Form = "=SI(B6=1;$'Coef " & Nb_periode &"'.E"& nbl+3 &";0)"
		Cell.formulalocal = form
	Range = Sheet.getCellRangeByPosition(1,NBe+7,Nbs,2*NBe+6)
	Range.fillSeries(com.sun.star.sheet.FillDirection.TO_BOTTOM,com.sun.star.sheet.FillMode.SIMPLE,0,0,0)
	Range.fillSeries(com.sun.star.sheet.FillDirection.TO_RIGHT,com.sun.star.sheet.FillMode.SIMPLE,0,0,0)
	
	'Recupération des coordonnées de dernière cellule de chaque ligne pour somme de la ligne + Généralisation de la formule
		Cell = Sheet.getCellByPosition(Nbs+1,NBe+7)
			Cell.formulalocal = "=ADRESSE("& NBe+8 &";" & Nbs+1 & ";4;;)"
			Extract_1 = Cell.string
			Cell.formulalocal = "=ADRESSE("& NBe+8 &";" & Nbs+3 & ";3;;)"
			Extract_2 = Cell.string
			Cell.formulalocal = "=SOMME(B"& NBe+8 &":" & Extract_1 & ")/$'Coef " & Nb_periode &"'.D" & NBe+5 & "+" & Extract_2
		Cell = Sheet.getCellByPosition(Nbs+2,NBe+7)
			Cell.formulalocal = "=SI(ESTVIDE(" & Extract_HA & ");0;10)"
	Range = Sheet.getCellRangeByPosition(Nbs+1,NBe+7,Nbs+2,2*NBe+6)
	Range.fillSeries(com.sun.star.sheet.FillDirection.TO_BOTTOM,com.sun.star.sheet.FillMode.SIMPLE,0,0,0)
	
	'Tableau stage obtenu
	'En-têtes
	Plage = Sheet.getCellRangeByPosition( 0 , 2*NBe+7 , Nbs+3 , 2*NBe+7 ) 
		Plage.Merge( True ) 	
		Cell = Sheet.getCellByPosition(0,2*NBe+7)
		With Cell 
			.setString( "Tableau Stage" ) 'insére du texte dans la cellule 
			.CellBackColor = RGB(255,215,0) 'indique la couleur de fond 
			.CharColor = RGB(0,0,0) 'couleur des caractères 
			.CharHeight = 14 'Taille catactères 
			.CharWeight = com.sun.star.awt.FontWeight.BOLD 'gras 
			.CharPosture = com.sun.star.awt.FontSlant.ITALIC 'italique 
			.CharFontName = "Arial" 'Font 
		End With  
		Cell = Sheet.getCellByPosition(1,2*NBe+8) 
			Cell.formula = "=$'Place Stage'.C$1"
		With Cell 
			.CellBackColor = RGB(139,0,0) 'indique la couleur de fond 
			.CharColor = RGB(255,255,255) 'couleur des caractères 
			.CharHeight = 10 'Taille catactères 
			.CharWeight = com.sun.star.awt.FontWeight.BOLD 'gras 
			.CharFontName = "Arial" 'Font 
		End With
	Range = Sheet.getCellRangeByPosition(1,2*NBe+8,Nbs,2*NBe+8)
	Range.fillSeries(com.sun.star.sheet.FillDirection.TO_RIGHT,com.sun.star.sheet.FillMode.SIMPLE,0,0,0)

		Cell = Sheet.getCellByPosition(0,2*NBe+9)
			Form = "=CONCAT($'Liste Etudiants'.$B2 ;" & chr(34) & chr(0160) & chr(34) &"; $'Liste Etudiants'.$C2)"
			Cell.formula = form
		With Cell 
			.CellBackColor = RGB(255,250,205) 'indique la couleur de fond 
			.CharColor = RGB(0,0,0) 'couleur des caractères 
			.CharHeight = 10 'Taille catactères 
			.CharPosture = com.sun.star.awt.FontSlant.ITALIC 'italique 
			.CharFontName = "Arial" 'Font 
		End With  
	Range = Sheet.getCellRangeByPosition(0,2*NBe+9,0,3*NBe+8)
	Range.fillSeries(com.sun.star.sheet.FillDirection.TO_BOTTOM,com.sun.star.sheet.FillMode.SIMPLE,0,0,0)
	'Renvoie du nom du stage si ce dernier est affecté + Généralisation
		Cell = Sheet.getCellByPosition(1,2*NBe+9)
			Form = "=SI(B6=1;B$2;" & chr(34) & chr(34) &")"
			Cell.formulalocal = form
	Range = Sheet.getCellRangeByPosition(1,2*NBe+9,Nbs,3*NBe+8)
	Range.fillSeries(com.sun.star.sheet.FillDirection.TO_BOTTOM,com.sun.star.sheet.FillMode.SIMPLE,0,0,0)
	Range.fillSeries(com.sun.star.sheet.FillDirection.TO_RIGHT,com.sun.star.sheet.FillMode.SIMPLE,0,0,0)
	'Coordonnées de la dernière cellule de la ligne pour formule de somme de la ligne
		Cell = Sheet.getCellByPosition(Nbs+1,2*NBe+9)
			Cell.formulalocal = "=ADRESSE("& 2*NBe+10 &";" & Nbs+1 & ";4;;)"
			Extract_1 = Cell.string
			Cell.formulalocal = "=ADRESSE("& 2*NBe+10 &";" & Nbs+3 & ";4;;)"
			Extract_2 = Cell.string
			Cell.formula = "=CONCAT(B"& 2*NBe+10 &":" & Extract_1 & ";" & Extract_2 & ")"
		Cell = Sheet.getCellByPosition(Nbs+2,2*NBe+9)
			Cell.formulalocal = "=ADRESSE(6;" & Nbs+4 & ";3;; )"
			Extract_2 = Cell.string
			Cell.formulalocal = "=SI(ESTVIDE(" & Extract_HA & ");" & chr(34) & chr(34) & ";" & Extract_2 & ")"
	Range = Sheet.getCellRangeByPosition(Nbs+1,2*NBe+9,Nbs+2,3*NBe+8)
	Range.fillSeries(com.sun.star.sheet.FillDirection.TO_BOTTOM,com.sun.star.sheet.FillMode.SIMPLE,0,0,0)

offset = offset +2 'incrément du decalage correspondant aux deux lignes "minimum" et "maximum" pour chaque période
Num_Feuille = Num_Feuille + 1 ' Incrémentation de la feuille de travail
Next Nb_periode

' Feuille Réouverture
	Doc.GetSheets.insertNewByName("Réouverture",Num_Feuille)'Ajoute une feuille, la nomme et place l'onglet en position 2 + 2* Nb_periode_Max + 1
	Sheet = Doc.Sheets.GetbyName("Réouverture")
	'En tête & Mise en forme
		Range = Sheet.getCellRangeByPosition( 0 , 0 , Nbs+3 , 0 ) 
		Range.Merge( True ) 
		Cell = Sheet.getCellByPosition(0,0)   
		With Cell 
			.setString( "Réouverture" ) 'insére du texte dans la cellule 
			.CellBackColor = RGB(255,215,0) 'indique la couleur de fond 
			.CharColor = RGB(0,0,0) 'couleur des caractères 
			.CharHeight = 14 'Taille catactères 
			.CharWeight = com.sun.star.awt.FontWeight.BOLD 'gras 
			.CharPosture = com.sun.star.awt.FontSlant.ITALIC 'italique 
			.CharFontName = "Arial" 'Font 
		End With 
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
			.CellBackColor = RGB(218,165,32) 'indique la couleur de fond 
			.paraAdjust = com.sun.star.style.ParagraphAdjust.CENTER 'alignement centré 
			.CharColor = RGB(0,0,0) 'couleur des caractères 
			.CharHeight = 10 'Taille catactères 
			.CharFontName = "Arial" 'Font 
		End With 
	Sheet.Rows(1).OptimalHeight = True 
	Range = Sheet.getCellRangeByPosition(0,2,2,NBe+2)   
		With Range 
		.CellBackColor = RGB(240,248,255) 'indique la couleur de fond 
		.CharColor = RGB(0,0,0) 'couleur des caractères 
		.CharHeight = 10 'Taille catactères 
		.CharPosture = com.sun.star.awt.FontSlant.ITALIC 'italique 
		.CharFontName = "Arial" 'Font 
		End With 
	For numl =1 to NBe-1
		Sheet.Rows(1+numl).OptimalHeight = True
	Next numl
	For numl=0 to 3
		Sheet.columns(numl).OptimalWidth = True 
	Next numl
	'Remplissage du tableau avec par défaut 0 partout
		Cell = Sheet.getCellByPosition(3,3)
			Cell.Value = 0
	Range = Sheet.getCellRangeByPosition(3,3,Nbs+2,NBe+2)
	Range.fillSeries(com.sun.star.sheet.FillDirection.TO_RIGHT,com.sun.star.sheet.FillMode.SIMPLE,0,0,0)
	Range.fillSeries(com.sun.star.sheet.FillDirection.TO_BOTTOM,com.sun.star.sheet.FillMode.SIMPLE,0,0,0)
	
	'Récupération du nombre de regroupement de stage et création en fonction ou non du nombre de colonne associées	
		Cell = Sheet.getCellByPosition(Nbs+3,1)
			Cell.formula = "=MAX($'Place Stage'.2:2)"
			NbGRP = cell.value
	
	IF NbGRP>0 Then
	For NumGRP = 1 to NbGRP
		Cell = Sheet.getCellByPosition(Nbs+2+NumGRP,1)
			Cell.String = "Regroupement N°"
		With Cell 
			.CellBackColor = RGB(222,184,135) 'indique la couleur de fond 
			.CharColor = RGB(0,0,0) 'couleur des caractères 
			.CharHeight = 10 'Taille catactères 
			.CharWeight = com.sun.star.awt.FontWeight.BOLD 'gras 
			.CharFontName = "Arial" 'Font 
		End With
		Cell = Sheet.getCellByPosition(Nbs+2+NumGRP,2)
			Cell.Value = NumGRP
		With Cell 
			.CellBackColor = RGB(222,184,135) 'indique la couleur de fond 
			.CharColor = RGB(0,0,0) 'couleur des caractères 
			.CharHeight = 10 'Taille catactères 
			.CharWeight = com.sun.star.awt.FontWeight.BOLD 'gras 
			.CharFontName = "Arial" 'Font 
		End With
		Sheet.columns(Nbs+2+NumGRP).OptimalWidth = True
	Next NumGRP
	Sheet.columns(0).OptimalWidth = True
	'Remplissage du tableau avec par défaut 0 partout	
		Cell = Sheet.getCellByPosition(Nbs+1,3)
			Cell.Value = 0
	Range = Sheet.getCellRangeByPosition(Nbs+1,3,Nbs+2+NumGRP-1,NBe+2)
	Range.fillSeries(com.sun.star.sheet.FillDirection.TO_RIGHT,com.sun.star.sheet.FillMode.SIMPLE,0,0,0)
	Range.fillSeries(com.sun.star.sheet.FillDirection.TO_BOTTOM,com.sun.star.sheet.FillMode.SIMPLE,0,0,0)	
	End If
	Sheet.columns(0).OptimalWidth = True
Num_Feuille = Num_Feuille + 1 ' Incrémentation de la feuille de travail
'Feuille Regroupement
Doc.GetSheets.insertNewByName("Regroupement",Num_Feuille)'Ajoute une feuille, la nomme et place l'onglet en position 2 + 2* Nb_periode_Max + 2
	Sheet = Doc.Sheets.GetByNAme("Regroupement")
	'En-tête
	Range = Sheet.getCellRangeByPosition( 0 , 0 , Nbs+2 , 0 ) 
	Range.Merge( True ) 
		Cell = Sheet.getCellByPosition(0,0)   
		With Cell 
			.setString( "Regroupement" ) 'insére du texte dans la cellule 
			.CellBackColor = RGB(255,215,0) 'indique la couleur de fond 
			.CharColor = RGB(0,0,0) 'couleur des caractères 
			.CharHeight = 14 'Taille catactères 
			.CharWeight = com.sun.star.awt.FontWeight.BOLD 'gras 
			.CharPosture = com.sun.star.awt.FontSlant.ITALIC 'italique 
			.CharFontName = "Arial" 'Font 
		End With 
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
			.CellBackColor = RGB(218,165,32) 'indique la couleur de fond 
			.paraAdjust = com.sun.star.style.ParagraphAdjust.CENTER 'alignement centré 
			.CharColor = RGB(0,0,0) 'couleur des caractères 
			.CharHeight = 10 'Taille catactères 
			.CharFontName = "Arial" 'Font 
		End With 
		Sheet.Rows(1).OptimalHeight = True 
	Range = Sheet.getCellRangeByPosition(0,2,2,NBe+2)   
		With Range 
			.CellBackColor = RGB(240,248,255) 'indique la couleur de fond 
			.CharColor = RGB(0,0,0) 'couleur des caractères 
			.CharHeight = 10 'Taille catactères 
			.CharPosture = com.sun.star.awt.FontSlant.ITALIC 'italique 
			.CharFontName = "Arial" 'Font 
		End With 
	For numl =1 to NBe-1
		Sheet.Rows(1+numl).OptimalHeight = True
	Next numl
	For numl=0 to 3
		Sheet.columns(numl).OptimalWidth = True 
	Next numl
	
	'Création des colonnes selon les regroupements de stages et leur numéro
		Cell = Sheet.getCellByPosition(Nbs+3,1)
			Cell.formula = "=MAX($'Place Stage'.2:2)"
			NbGRP = cell.value
	IF NbGRP>0 Then
	For NumGRP = 1 to NbGRP
		Cell = Sheet.getCellByPosition(Nbs+NumGRP+2,1)
			Cell.String = "Regroupement N°"
		With Cell 
			.CellBackColor = RGB(222,184,135) 'indique la couleur de fond 
			.CharColor = RGB(0,0,0) 'couleur des caractères 
			.CharHeight = 10 'Taille catactères 
			.CharWeight = com.sun.star.awt.FontWeight.BOLD 'gras 
			.CharFontName = "Arial" 'Font 
		End With
		Cell = Sheet.getCellByPosition(Nbs+NumGRP+2,2)
			Cell.Value = NumGRP
		With Cell 
			.CellBackColor = RGB(222,184,135) 'indique la couleur de fond 
			.CharColor = RGB(0,0,0) 'couleur des caractères 
			.CharHeight = 10 'Taille catactères 
			.CharWeight = com.sun.star.awt.FontWeight.BOLD 'gras 
			.CharFontName = "Arial" 'Font 
		End With
	'Récupération des coordonnées pour élaboration formules regroupement + Généralisation formule
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
Num_Feuille = Num_Feuille + 1 ' Incrémentation de la feuille de travail
' Feuille Synthèse
	Doc.GetSheets.insertNewByName("Synthèse",Num_Feuille) 'Ajoute une feuille, la nomme et place l'onglet en position 2 + 2* Nb_periode_Max + 3
	Sheet = Doc.Sheets.GetByName("Synthèse")
	'En-tête
		Cell = Sheet.getCellByPosition(0,0)
			Cell.string = "Sous Groupe"
		With Cell 
			.CellBackColor = RGB(205,133,63) 'indique la couleur de fond 
			.CharColor = RGB(0,0,0) 'couleur des caractères 
			.CharHeight = 10 'Taille catactères 
			.CharPosture = com.sun.star.awt.FontSlant.ITALIC 'italique 
			.CharFontName = "Arial" 'Font 
		End With  
		Cell = Sheet.getCellByPosition(0,1)
			Cell.formula = "=$'Liste Etudiants'.$A2 "
		With Cell 
			.CellBackColor = RGB(255,250,205) 'indique la couleur de fond 
			.CharColor = RGB(0,0,0) 'couleur des caractères 
			.CharHeight = 10 'Taille catactères 
			.CharPosture = com.sun.star.awt.FontSlant.ITALIC 'italique 
			.CharFontName = "Arial" 'Font 
		End With  
	Range = Sheet.getCellRangeByPosition(0,1,0,NBe)
	Range.fillSeries(com.sun.star.sheet.FillDirection.TO_BOTTOM,com.sun.star.sheet.FillMode.SIMPLE,0,0,0)
		Cell = Sheet.getCellByPosition(1,0)
			Cell.string = "Etudiant"
		With Cell 
			.CellBackColor = RGB(205,133,63) 'indique la couleur de fond 
			.CharColor = RGB(0,0,0) 'couleur des caractères 
			.CharHeight = 10 'Taille catactères 
			.CharPosture = com.sun.star.awt.FontSlant.ITALIC 'italique 
			.CharFontName = "Arial" 'Font 
		End With  
		'Renvois vers nom des étudiants
		Cell = Sheet.getCellByPosition(1,1)
			Form = "=CONCAT($'Liste Etudiants'.$B2 ;" & chr(34) & chr(0160) & chr(34) &"; $'Liste Etudiants'.$C2)"
			Cell.formula = form
		With Cell 
			.CellBackColor = RGB(255,250,205) 'indique la couleur de fond 
			.CharColor = RGB(0,0,0) 'couleur des caractères 
			.CharHeight = 10 'Taille catactères 
			.CharPosture = com.sun.star.awt.FontSlant.ITALIC 'italique 
			.CharFontName = "Arial" 'Font 
		End With  
	Range = Sheet.getCellRangeByPosition(0,1,1,NBe)
	Range.fillSeries(com.sun.star.sheet.FillDirection.TO_BOTTOM,com.sun.star.sheet.FillMode.SIMPLE,0,0,0)
	Sheet.columns(1).OptimalWidth = True
	'Renvois vers les résultats de chaque affectation selon le nombre de période générées
	For Nb_periode = 1 to Nb_periode_Max
		Cell = Sheet.getCellByPosition(1+Nb_periode,0)
			Cell.String = "Période " & Nb_periode
		With Cell 
			.CellBackColor = RGB(205,133,63) 'indique la couleur de fond 
			.CharColor = RGB(0,0,0) 'couleur des caractères 
			.CharHeight = 10 'Taille catactères 
			.CharPosture = com.sun.star.awt.FontSlant.ITALIC 'italique 
			.CharFontName = "Arial" 'Font 
		End With  
		'Renvoie vers le stage associé à la période dans le tableau Stage Affecté 
		Cell = Sheet.getCellByPosition(Nb_periode+1,1)
			Cell.formulalocal = "=ADRESSE("& 2*NBe+10 &";" & Nbs+2 & ";4;;)"
			Extract_1 = Cell.string
			Cell.Formula = "=$'Stage "& Nb_periode &"'." & Extract_1
		With Cell 
			.CellBackColor = RGB(255,250,205) 'indique la couleur de fond 
			.CharColor = RGB(0,0,0) 'couleur des caractères 
			.CharHeight = 10 'Taille catactères 
			.CharPosture = com.sun.star.awt.FontSlant.ITALIC 'italique 
			.CharFontName = "Arial" 'Font 
		End With  
		Sheet.columns(1+Nb_periode).OptimalWidth = True
		Range = Sheet.getCellRangeByPosition(1+Nb_periode,1,1+Nb_periode,NBe)
		Range.fillSeries(com.sun.star.sheet.FillDirection.TO_BOTTOM,com.sun.star.sheet.FillMode.SIMPLE,0,0,0)
			
	Next Nb_periode
Num_Feuille = Num_Feuille + 1 ' Incrémentation de la feuille de travail
'Ecriture des coefficients sur les feuilles Coef X
For Nb_periode =1 to Nb_periode_Max
Sheet = Doc.Sheets(Nb_periode+1)

		Cell = Sheet.getCellByPosition(3,nbl+2)
			Cell.formulalocal = "=Adresse(" & NBe+8 & ";" & Nbs+2 & ";4)"
			Extract_1 = Cell.string
If Nb_periode = 1 Then Cell.value = 1 Else Cell.formulalocal = "=EXP(MOYENNE($'Coef " & Nb_periode-1 & "'.3:3)*(10/MAX($'Coef " & Nb_periode-1 & "'.3:3)) - $'Stage " & Nb_periode-1 & "'." & Extract_1 & ")*$'Coef " & Nb_periode-1 &"'.D" & NBe+5
Next Nb_periode
Sheet = Doc.Sheets(2)
	For ligne = NBe+4 to 2*NBe+3
		Cell = Sheet.getCellByPosition(3,ligne)
		Cell.value = 1
	Next ligne

For Nb_periode =2 to Nb_periode_Max
	Sheet = Doc.Sheets(Nb_periode+1)
	Range = Sheet.getCellRangeByPosition(3,NBe+4,3,2*NBe+3)
	Range.fillSeries(com.sun.star.sheet.FillDirection.TO_BOTTOM,com.sun.star.sheet.FillMode.SIMPLE,0,0,0)
Next Nb_periode

'Feuille de travail
	Doc.GetSheets.insertNewByName("Travail",Num_Feuille)
	Sheet = Doc.Sheets.GetByName("Travail")
'En-tête Tableau 1
	Range = Sheet.getCellRangeByPosition( 0 , 0 , Nbs+2 , 0 ) 
		Range.Merge( True ) 
		Cell = Sheet.getCellByPosition(0,0)   
		With Cell 
			.setString( "Notes Brutes" ) 'insére du texte dans la cellule 
			.CellBackColor = RGB(255,215,0) 'indique la couleur de fond 
			.CharColor = RGB(0,0,0) 'couleur des caractères 
			.CharHeight = 14 'Taille catactères 
			.CharWeight = com.sun.star.awt.FontWeight.BOLD 'gras 
			.CharPosture = com.sun.star.awt.FontSlant.ITALIC 'italique 
			.CharFontName = "Arial" 'Font 
		End With 
		Sheet.Rows(0).OptimalHeight = True 
		Cell = Sheet.getCellByPosition(0,1)
			Cell.formula = "=$'Liste Etudiants'.A1"
		With Cell 
			.CellBackColor = RGB(151, 200, 255) 'indique la couleur de fond 
			.CharColor = RGB(0,0,0) 'couleur des caractères 
			.CharHeight = 10 'Taille catactères 
			.CharFontName = "Arial" 'Font 
		End With 
		Range = Sheet.getCellRangeByPosition(0,1,2,nbe+1)
		Range.fillSeries(com.sun.star.sheet.FillDirection.TO_BOTTOM,com.sun.star.sheet.FillMode.SIMPLE,0,0,0)
		Range.fillSeries(com.sun.star.sheet.FillDirection.TO_RIGHT,com.sun.star.sheet.FillMode.SIMPLE,0,0,0)
		Cell = Sheet.getCellByPosition(3,1)
			Cell.formula = "=$'Place Stage'.C1"
		With Cell 
			.CellBackColor = RGB(161, 47, 0) 'indique la couleur de fond 
			.CharColor = RGB(255,255,255) 'couleur des caractères 
			.CharHeight = 10 'Taille catactères 
			.CharWeight = com.sun.star.awt.FontWeight.BOLD 'gras 
			.CharFontName = "Arial" 'Font 
		End With 
		Range = Sheet.getCellRangeByPosition(3,1,2+Nbs,1)
		Range.fillSeries(com.sun.star.sheet.FillDirection.TO_RIGHT,com.sun.star.sheet.FillMode.SIMPLE,0,0,0)
		Range = Sheet.getCellRangeByPosition( 0 , nbe+2 , Nbs+2 , nbe+2 ) 
			Range.Merge( True ) 
		Cell = Sheet.getCellByPosition(0,nbe+2)   
		With Cell 
			.setString( "Check Stage et Blocs" ) 'insére du texte dans la cellule 
			.CellBackColor = RGB(255,215,0) 'indique la couleur de fond 
			.CharColor = RGB(0,0,0) 'couleur des caractères 
			.CharHeight = 14 'Taille catactères 
			.CharWeight = com.sun.star.awt.FontWeight.BOLD 'gras 
			.CharPosture = com.sun.star.awt.FontSlant.ITALIC 'italique 
			.CharFontName = "Arial" 'Font 
		End With 
		Sheet.Rows(nbe+2).OptimalHeight = True 
	'En-tête Tableau 2	
		Cell = Sheet.getCellByPosition(0,nbe+3)
			Cell.formula = "=$'Liste Etudiants'.A1"
		With Cell 
			.CellBackColor = RGB(151, 200, 255) 'indique la couleur de fond 
			.CharColor = RGB(0,0,0) 'couleur des caractères 
			.CharHeight = 10 'Taille catactères 
			.CharFontName = "Arial" 'Font 
		End With  
		Range = Sheet.getCellRangeByPosition(0,nbe+3,2,2*nbe+3)
		Range.fillSeries(com.sun.star.sheet.FillDirection.TO_BOTTOM,com.sun.star.sheet.FillMode.SIMPLE,0,0,0)
		Range.fillSeries(com.sun.star.sheet.FillDirection.TO_RIGHT,com.sun.star.sheet.FillMode.SIMPLE,0,0,0)
		Cell = Sheet.getCellByPosition(3,nbe+3)
			Cell.formula = "=$'Place Stage'.C1"
		With Cell 
			.CellBackColor = RGB(161, 47, 0) 'indique la couleur de fond 
			.CharColor = RGB(255,255,255) 'couleur des caractères 
			.CharHeight = 10 'Taille catactères 
			.CharWeight = com.sun.star.awt.FontWeight.BOLD 'gras 
			.CharFontName = "Arial" 'Font 
		End With 
		Range = Sheet.getCellRangeByPosition(3,nbe+3,2+Nbs,nbe+3)
		Range.fillSeries(com.sun.star.sheet.FillDirection.TO_RIGHT,com.sun.star.sheet.FillMode.SIMPLE,0,0,0)
	
	' Formule si pas de regroupements	
	IF NbGRP = 0 Then 
		Cell = Sheet.getCellByPosition(3,nbe+4)	
			Cell.formulalocal ="=SI($Regroupement.B3>=1;" &  chr(34) & chr(34) & ";D3)"
		With Cell 
		.CellBackColor = RGB(211,211,211) 'indique la couleur de fond 
		.CharColor = RGB(0,0,0) 'couleur des caractères 
		.CharHeight = 10 'Taille catactères 
		.CharWeight = com.sun.star.awt.FontWeight.BOLD 'gras 
		.CharFontName = "Arial" 'Font 
		End With
	' Formule si regroupements
	Else
		Cell = Sheet.getCellByPosition(3,nbe+4)	
			Cell.formulalocal = "=Adresse(3;" & Nbs+4 & ";1)"
			Born1 = Cell.String
			Cell.formulalocal = "=Adresse(" & 3+Nbe & ";" & Nbs+3+NbGRP & ";1)"
			Born2 = Cell.String
			Cell.formulalocal = "=Adresse(3;" & Nbs+3+NbGRP & ";1)"
			Born5 = Cell.String
		Cell.formulalocal = "=SI(ESTERREUR(RECHERCHEH($'Place Stage'.C$2;$Regroupement." & born1 & ":" & born2 & ";LIGNE($Regroupement.A3)-1;)>=SI(SOMME.SI($Regroupement." & born1 & ":" & born5 & ";$'Place Stage'.C$2;$Regroupement." & born1 & ":" & born5 &")>0;SOMME.SI($Regroupement." & born1 & ":" & born5 & ";$'Place Stage'.C$2;$Regroupement." & born1 & ":" & born5 & ");1)" & ");SI($Regroupement.D4>=1;" &  chr(34) & chr(34) & ";D3);SI(RECHERCHEH($'Place Stage'.C$2;$Regroupement." & born1 &":"& born2 &";LIGNE($Regroupement.A3)-1;)>=SI(SOMME.SI($Regroupement."& born1 & ":" & born5 & ";$'Place Stage'.C$2;$Regroupement." &born1 & ":" & born5 &")>0;SOMME.SI($Regroupement." & born1 & ":" & born5 & ";$'Place Stage'.C$2;$Regroupement." & born1 & ":" & born5 & ");1)" &" ;" &  chr(34) & chr(34) & ";SI($Regroupement.D4>=1;" &  chr(34) & chr(34) & ";D3)))"
		With Cell 
			.CellBackColor = RGB(211,211,211) 'indique la couleur de fond 
			.CharColor = RGB(0,0,0) 'couleur des caractères 
			.CharHeight = 10 'Taille catactères 
			.CharWeight = com.sun.star.awt.FontWeight.BOLD 'gras 
			.CharFontName = "Arial" 'Font 
		End With 		
	End IF	
	Range = Sheet.getCellRangeByPosition(3,nbe+4,2+NbS,2*nbe+3)
		Range.fillSeries(com.sun.star.sheet.FillDirection.TO_BOTTOM,com.sun.star.sheet.FillMode.SIMPLE,0,0,0)
		Range.fillSeries(com.sun.star.sheet.FillDirection.TO_RIGHT,com.sun.star.sheet.FillMode.SIMPLE,0,0,0)	

Num_Feuille = Num_Feuille + 1 ' Incrémentation de la feuille de travail
'Feuille Précédents 
	Doc.GetSheets.insertNewByName("Précédents",Num_Feuille)'Ajoute une feuille, la nomme et place l'onglet en position 2 + 2* Nb_periode_Max + 5
	Sheet = Doc.Sheets.GetByNAme("Précédents")
		Range = Sheet.getCellRangeByPosition( 0 , 0 , Nbs+2 , 0 ) 
			Range.Merge( True ) 
			Cell = Sheet.getCellByPosition(0,0)   
			With Cell 
				.setString( "Précédents" ) 'insére du texte dans la cellule 
				.CellBackColor = RGB(255,215,0) 'indique la couleur de fond 
				.CharColor = RGB(0,0,0) 'couleur des caractères 
				.CharHeight = 14 'Taille catactères 
				.CharWeight = com.sun.star.awt.FontWeight.BOLD 'gras 
				.CharPosture = com.sun.star.awt.FontSlant.ITALIC 'italique 
				.CharFontName = "Arial" 'Font 
			End With 
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
			.CellBackColor = RGB(218,165,32) 'indique la couleur de fond 
			.paraAdjust = com.sun.star.style.ParagraphAdjust.CENTER 'alignement centré 
			.CharColor = RGB(0,0,0) 'couleur des caractères 
			.CharHeight = 10 'Taille catactères 
			.CharFontName = "Arial" 'Font 
		End With 
		Sheet.Rows(1).OptimalHeight = True 
		Range = Sheet.getCellRangeByPosition(0,2,2,NBe+2)   
		With Range 
			.CellBackColor = RGB(240,248,255) 'indique la couleur de fond 
			.CharColor = RGB(0,0,0) 'couleur des caractères 
			.CharHeight = 10 'Taille catactères 
			.CharPosture = com.sun.star.awt.FontSlant.ITALIC 'italique 
			.CharFontName = "Arial" 'Font 
		End With 
	For numl =1 to NBe-1
		Sheet.Rows(1+numl).OptimalHeight = True
	Next numl
	For numl=0 to 3
		Sheet.columns(numl).OptimalWidth = True 
	Next numl

'Ecriture formaule feuille regroupement """
Sheet = Doc.Sheets.GetByNAme("Regroupement")
	FormuleBoucle = "="
	For Nb_periode = 1 to Nb_periode_Max
		If Nb_periode < Nb_periode_Max 	then FormuleBoucle = FormuleBoucle & "$'Stage " & Nb_periode & "'.B6 +" 	else FormuleBoucle = FormuleBoucle & "$'Stage " & Nb_periode & "'.B6-$Réouverture.D4+$Précédents.D4"
	Next Nb_periode

		Cell = Sheet.getCellByPosition(3,3)
			Cell.formula = FormuleBoucle
	Range = Sheet.getCellRangeByPosition(3,3,Nbs+2,NBe+2)
	Range.fillSeries(com.sun.star.sheet.FillDirection.TO_RIGHT,com.sun.star.sheet.FillMode.SIMPLE,0,0,0)
	Range.fillSeries(com.sun.star.sheet.FillDirection.TO_BOTTOM,com.sun.star.sheet.FillMode.SIMPLE,0,0,0)

End Sub
