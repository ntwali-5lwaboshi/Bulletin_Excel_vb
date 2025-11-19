Sub CreerFicheAvecDonnees()
    Dim wbNew As Workbook
    Dim wsSource As Worksheet
    Dim wsDest As Worksheet
    Dim dossierDestination As String
    Dim nomFichier As String
    Dim cheminComplet As String
    Dim fd As FileDialog
    Dim dataEleves As Variant
    Dim i As Long

    ' 1?? Choisir dossier
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    With fd
        .title = "Choisissez où enregistrer la fiche"
        If .Show <> -1 Then Exit Sub
        dossierDestination = .SelectedItems(1)
    End With

    ' 2?? Créer nouveau classeur
    Set wbNew = Workbooks.Add
    Set wsDest = wbNew.Sheets(1)
    wsDest.Name = "Fiche"

    ' 3?? Feuille modèle
    Set wsSource = ThisWorkbook.Sheets("DashBoard")

    ' 4?? Copier le tableau du modèle (B6:F7)
    wsSource.range("B6:F6").Copy
    With wsDest.range("A6")
        .PasteSpecial Paste:=xlPasteColumnWidths
        .PasteSpecial Paste:=xlPasteAll
    End With
    With wsDest.range("B:B").Select
        Selection.Delete Shift:=xlToLeft
    End With
    
    Application.CutCopyMode = False

    ' 5?? Ajouter les en-têtes pour les données élèves
    wsDest.range("A6").Value = "N°"
    wsDest.range("B6").Value = "Nom Post-nom & Prénom"
    wsDest.range("C6").Value = "Sexe"
    wsDest.range("D6").Value = "Lieu et date de naissance"


    ' 8?? Mise en forme : bordures automatiques
    With wsDest.range("A6:D" & 110).Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = 0
    End With

    ' 9?? Enregistrer le fichier
    nomFichier = "Fiche_Eleves_" & Format(Now, "yyyymmdd_hhmmss") & ".xlsx"
    cheminComplet = dossierDestination & "\" & nomFichier
    wbNew.SaveAs filename:=cheminComplet, FileFormat:=xlOpenXMLWorkbook

    MsgBox "? Fiche générée avec données et sauvegardée :" & vbCrLf & cheminComplet, vbInformation
    wbNew.Activate
End Sub

Sub ImporterEtMatriculer()
    Dim fd As FileDialog
    Dim cheminFichier As String
    Dim wbSource As Workbook
    Dim wsSource As Worksheet
    Dim wsDest As Worksheet
    Dim lastRowSource As Long
    Dim lastRowDest As Long
    Dim i As Long
    Dim numeroMatricule As Long
    Dim suffixeAnnee As String
    
    ' ?? Paramètres
    suffixeAnnee = Sheets("DashBoard").range("U5").Value ' <-- suffixe année scolaire
    numeroMatricule = Sheets("Eleves").range("I7").Value      ' <-- numéro de départ

    ' 1?? Sélectionner le fichier externe
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .title = "Sélectionnez le fichier source des élèves"
        .Filters.Add "Fichiers Excel", "*.xls; *.xlsx; *.xlsm"
        .AllowMultiSelect = False
        If .Show <> -1 Then Exit Sub
        cheminFichier = .SelectedItems(1)
    End With

    ' 2?? Ouvrir le fichier externe en lecture seule
    Set wbSource = Workbooks.Open(cheminFichier, ReadOnly:=True)
    Set wsSource = wbSource.Sheets(1)

    ' 3?? Dernière ligne du fichier source
    lastRowSource = wsSource.Cells(wsSource.Rows.Count, "B").End(xlUp).Row

    ' 4?? Feuille de destination dans le classeur principal
    Set wsDest = ThisWorkbook.Sheets("Eleves") ' <-- à adapter

    ' 5?? Déterminer la première ligne vide dans la feuille destination
    lastRowDest = wsDest.Cells(wsDest.Rows.Count, "A").End(xlUp).Row
    If lastRowDest = 1 And wsDest.Cells(1, 1).Value = "" Then lastRowDest = 0
    
    'numérotation des lignes eleves
    num = 0

    ' 6?? Copier ligne par ligne et cellule par cellule
    Dim j As Long
    For i = 8 To lastRowSource  ' saute les en-têtes
        
        num = num + 1
        lastRowDest = lastRowDest + 1
        'numéroter dans le prémières colonnes
        wsDest.range("A" & lastRowDest).Value = num
        
        ' Colonne E = matricule unique
        wsDest.range("B" & lastRowDest).Value = Format(numeroMatricule, "0000") & "/" & suffixeAnnee
        numeroMatricule = numeroMatricule + 1
        
        ' Colonne A à D
        For j = 3 To 5
            wsDest.Cells(lastRowDest, j).Value = wsSource.Cells(i, j - 1).Value
        Next j
    Next i


    ' 8?? Fermer le fichier externe
    wbSource.Close False
    
    'appel du fonction pour appeler les données aussi sur le dashboard
    Call actuealiser_list_dash
    
    MsgBox num & " Elèves enregistrés et matricules générés pour chaque élève !", vbInformation
End Sub

Sub Add_eleve()
Add_student_form.Show
End Sub

Sub actuealiser_list_dash()
    Dim wsDest As Worksheet
    Dim wsSource As Worksheet
    

    Set wsSource = Sheets("Eleves")
    
    ' 3?? Dernière ligne du fichier source
    lastRowSource = wsSource.Cells(wsSource.Rows.Count, "B").End(xlUp).Row
    
    ' 4?? Feuille de destination dans le classeur principal
    Set wsDest = ThisWorkbook.Sheets("DashBoard") ' <-- à adapter
    
    'arranger l'espace de destination
    wsDest.range("B7:F110").ClearContents

    ' 5?? Déterminer la première ligne vide dans la feuille destination
    lastRowDest = wsDest.Cells(wsDest.Rows.Count, "B").End(xlUp).Row
    If lastRowDest = 1 And wsDest.Cells(1, 1).Value = "" Then lastRowDest = 0
    
     
    ' 6?? Copier ligne par ligne et cellule par cellule
    Dim j As Long
    For i = 8 To lastRowSource  ' saute les en-têtes
        If wsSource.Cells(i, 6).Value <> "Abandon" Then
            'incrementer que quand le critère est valide
            lastRowDest = lastRowDest + 1
            
            'numéroter dans le prémières colonnes
            wsDest.range("B" & lastRowDest).Value = lastRowDest - 6
            
            ' Colonne A à D
            For j = 3 To 6
                wsDest.Cells(lastRowDest, j).Value = wsSource.Cells(i, j - 1).Value
            Next j
        End If
    Next i
End Sub
Sub ModeleCotes()

    Dim wbNew As Workbook
    Dim wsSource As Worksheet
    Dim wsSource_d As Worksheet
    Dim wsS_dash As Worksheet
    Dim wsDest As Worksheet
    Dim dossierDestination As String
    Dim nomFichier As String
    Dim cheminComplet As String
    Dim fd As FileDialog
    Dim dataEleves As Variant
    Dim i As Long
    
    'cours
    Dim c_line As range
    
    'recupérer le dashboard pour y récupérer le cours et autre infos
    Set wsS_dash = ThisWorkbook.Sheets("DashBoard")

    'recherche des information sur le cours
    Dim cours As String
    cours = wsS_dash.range("I3").Value
    
    If Trim(cours) = "" Then
        MsgBox "Sélectionner un cours"
        wsS_dash.range("I3").Activate
        Exit Sub
    End If
    
    ' Chercher le cours dans la liste de cours sur le dashbord
    Set c_line = wsS_dash.Columns("X:X").Find(What:=cours, LookIn:=xlValues, LookAt:=xlWhole)

    If c_line Is Nothing Then
        MsgBox "le cours sélectionné n'est pas parmis le cours " & cours
        Exit Sub
    ElseIf wsS_dash.range("Z" & c_line.Row).Value = "Non" Then
        MsgBox "le cours sélectionné est présent sur le bulletin pour cette classe, mais il n'est pas étudié, unitile de lui faire un fiche si le cours est étudié vérifier parmis les cours si par erreurs le cours n'est pas cohé comme non étudié : (" & cours & ")", vbInformation
        Exit Sub
    End If
    

    ' Choisir dossier
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    With fd
        .title = "Choisir le dossier qui pour le model de fiche de cotes"
        If .Show <> -1 Then Exit Sub
        dossierDestination = .SelectedItems(1)
    End With
    
    
    ' 2?? Créer nouveau classeur
    Set wbNew = Workbooks.Add
    Set wsDest = wbNew.Sheets(1)
    wsDest.Name = "Fiche"

    ' 3?? Feuille modèle
    Set wsSource = ThisWorkbook.Sheets("ModelFiche")
    Set wsSource_d = ThisWorkbook.Sheets("Eleves")
 

    ' 4?? Copier le tableau du modèle (B6:F7)
    wsSource.range("B1:K6").Copy
    With wsDest.range("A6")
        .PasteSpecial Paste:=xlPasteColumnWidths
        .PasteSpecial Paste:=xlPasteAll
    End With
    
    Application.CutCopyMode = False
    
    
    'trouver le dernier ligne cintenant le donnée dans le destination
    Dim lastRow As Long
    lastRow = wsDest.Cells(wsDest.Rows.Count, "A").End(xlUp).Row
    
    'collection des informations entêtes nom enseignant cours et classe
    nom_ens = InputBox("Nom du titulaire :", "Titulaire")
    classe = wsS_dash.range("S7").Value
    annee = wsS_dash.range("S5").Value
    
    'complexion des informations entêtes
    wsDest.range("D" & lastRow - 5).Value = annee
    wsDest.range("D" & lastRow - 4).Value = classe
    wsDest.range("D" & lastRow - 3).Value = nom_ens
    wsDest.range("D" & lastRow - 2).Value = cours
    
    'completer les pondération sur la ligne des entêtes
    ponderation = wsS_dash.range("AB" & c_line.Row).Value
    'semestre 1
    wsDest.range("E" & lastRow).Value = ponderation
    wsDest.range("F" & lastRow).Value = ponderation
    wsDest.range("G" & lastRow).Value = ponderation * 2

    'semestre 2
    wsDest.range("H" & lastRow).Value = ponderation
    wsDest.range("I" & lastRow).Value = ponderation
    wsDest.range("J" & lastRow).Value = ponderation * 2
    
    Call get_data_eleve(wsDest, wsSource_d, lastRow + 0, "A", False, 1, 4, True, False)
    
    lastRow = wsDest.Cells(wsDest.Rows.Count, "A").End(xlUp).Row
    With wsDest.range("A10:J" & lastRow).Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = 0
    End With

    ' 9?? Enregistrer le fichier
    nomFichier = "Fiche_" & cours & "_" & classe & "_" & annee & "_" & Format(Now, "yyyymmdd_hhmmss") & ".xlsx"
    cheminComplet = dossierDestination & "\" & nomFichier
    wbNew.SaveAs filename:=cheminComplet, FileFormat:=xlOpenXMLWorkbook

    MsgBox "? Fiche générée:" & vbCrLf & cheminComplet, vbInformation
    wbNew.Close
    
    
End Sub


Sub get_data_eleve(wsDest As Worksheet, wsSource As Worksheet, skipLine As Integer, foundCol As String, abandon As Boolean, firstColDest As Integer, lastColData As Integer, numerotation As Boolean, actualiser As Boolean)

    ' 3?? Dernière ligne du fichier source
    lastRowSource = wsSource.Cells(wsSource.Rows.Count, "B").End(xlUp).Row
    
    
    ' 5?? Déterminer la première ligne vide dans la feuille destination
    lastRowDest = wsDest.Cells(wsDest.Rows.Count, foundCol).End(xlUp).Row
    If lastRowDest = 1 And wsDest.Cells(1, 1).Value = "" Then lastRowDest = 0
    
    'recuperer le premier colone si la numerotation n'est pas activée
    fCVNum = 0
    If numerotation = False Then
        firstColDest = firstColDest - 1
        fCVNum = fCVNum + 1
    End If
     
    ' 6?? Copier ligne par ligne et cellule par cellule
    Dim j As Long
    Dim matricule_row
    Dim Matricule As String
    
    'compteur des modifications
    c_modif = 0
    
    For i = 8 To lastRowSource  ' saute les en-têtes
        If wsSource.Cells(i, 6).Value <> "Abandon" Or abandon Then
            If Not actualiser Then
                'incrementer que quand le critère est valide
                lastRowDest = lastRowDest + 1
            
                If numerotation = True Then
                    'numéroter dans le prémières colonnes
                    wsDest.Cells(lastRowDest, firstColDest).Value = lastRowDest - skipLine
                End If
                
                ' Colonne A à D
                For j = firstColDest + 1 To lastColData
                    wsDest.Cells(lastRowDest, j).Value = wsSource.Cells(i, j + fCVNum).Value
                Next j
            Else
                Matricule = wsSource.Cells(i, 2).Value
                        
                ' Chercher le matricule dans la fiche periode
                Set matricule_row = wsDest.Columns("A:A").Find(What:=Matricule, LookIn:=xlValues, LookAt:=xlWhole)
                'MsgBox matricule & " linge " & matricule_Row.Row & " col " & matricule_Row.Column
                If Not matricule_row Is Nothing Then
                    'si le matricule est trouvé, modification de saligne
                    'si nous constantons de changements

                    For k = firstColDest + 1 To lastColData
                        If wsDest.Cells(matricule_row.Row, k).Value <> wsSource.Cells(i, k + fCVNum).Value Then
                            wsDest.Cells(matricule_row.Row, k).Value = wsSource.Cells(i, k + fCVNum).Value
                        End If
                    Next k
                
                Else
                    'incrementer que quand le critère est valide
                    lastRowDest = lastRowDest + 1
                    
                    MsgBox Matricule & " Introuvable on va donc ajouter"
                    
                    'si le matricule n'est pas trouvé donc c'est une nouvelle enregistrement
                    If numerotation = True Then
                        'numéroter dans le prémières colonnes
                        wsDest.Cells(lastRowDest, firstColDest).Value = lastRowDest - skipLine
                    End If
                    
                    ' Colonne A à D
                    For j = firstColDest + 1 To lastColData
                        wsDest.Cells(lastRowDest, j).Value = wsSource.Cells(i, j + fCVNum).Value
                    Next j
                End If
                

            End If
            
        End If
    Next i
End Sub

Sub initialiser_fiche_periodes_data()
Call initialiser_fiche_periodes(False)
Call save
End Sub
Sub actualiser_fiche_periodes_data()
Call initialiser_fiche_periodes(True)
Call save
End Sub

Sub initialiser_fiche_periodes(actualiser As Boolean)
    Dim wsSource As Worksheet
    Dim wsDest As Worksheet
    Dim f As Variant
    
    'compteur des erreurs,
    c_erreurs = 0
    c = 0
    
    'source des données
    Set wsSource = ThisWorkbook.Sheets("Eleves")
    
    'Liste des fiche par périodes, examen, semestres
    periodes = Array("P1", "P2", "EXAM1", "P3", "P4", "EXAM2", "SEM1", "SEM2", "TOT")
    
    'parcours de tous les fiches pour y metre les données
    For Each f In periodes
        
        
        'recupérer chaque fiche un par un
        Set wsDest = ThisWorkbook.Sheets(f)
        
        'Déterminer la première ligne vide dans la feuille destination
        lastRowDest = wsDest.Cells(wsDest.Rows.Count, "A").End(xlUp).Row
        
        If lastRowDest > 16 And actualiser = False Then
            c_erreurs = c_erreurs + 1
        Else
            'insertion des donnes, sur la fiche car tous les fiches on une structure similaire
            Call get_data_eleve(wsDest, wsSource, lastRowDest + 0, "A", True, 1, 5, False, actualiser)
            c = c + 1
        End If
    Next f
    MsgBox c_erreurs & " Echec,/" & c & " Opérations"
End Sub
Sub importer_cotes_sur_fiche_sigle_file()
Call importer_cotes_sur_fiche("")
End Sub

Function chercher_cours(cours As String) As String
    ' Chercher le cours dans la fiche periode
    Dim cours_c
    Set cours_c = ThisWorkbook.Sheets("DashBoard").Columns("X:X").Find(What:=cours, LookIn:=xlValues, LookAt:=xlWhole)
    If Not cours_c Is Nothing Then
        chercher_cours = ThisWorkbook.Sheets("DashBoard").range("X" & cours_c.Row).Value
    Else
        chercher_cours = ""
    End If
End Function
Sub importer_cotes_sur_fiche(wsS As String)
    Dim fd As FileDialog
    Dim cheminFichier As String
    Dim wbSource As Workbook
    Dim wsSource As Worksheet
    Dim wsDest As Worksheet
    Dim lastRowSource As Long
    Dim lastRowDest As Long
    Dim i As Long
    Dim numeroMatricule As Long
    Dim suffixeAnnee As String
    
    
    ' ?? Paramètres
    'Liste des fiche par périodes, examen, semestres
    fichesPeriode = Array("P1", "P2", "EXAM1", "P3", "P4", "EXAM2")
    
    
    'compteur d 'oppérations
    c = 0
    
    'compteur d'erreurs
    e_c = 0
    
    'la colone du période sélectionné sur la fiche
    p_fiche_col = 0
    
    'chaine des erreurs
    erreurs = " Erreur :"
    
    'verificateur si la période a été rétrouvé
    c_p = 0
    
    'compter les colonne pour savoir si on a réussi à trouver la colonne spécifié
    p_fiche_col = 0
    
    
    'cours sélectionnée
    cours = ThisWorkbook.Sheets("DashBoard").range("I3").Value
    
    'chercher le fichier source que s'il n'a pas été passé en parametre
    
    If wsS = "" Then
        '1 Sélectionner le fichier externe
        Set fd = Application.FileDialog(msoFileDialogFilePicker)
        With fd
            .title = "Sélectionnez le fichier source des élèves"
            .Filters.Add "Fichiers Excel", "*.xls; *.xlsx; *.xlsm"
            .AllowMultiSelect = False
            If .Show <> -1 Then Exit Sub
            cheminFichier = .SelectedItems(1)
        End With
    Else
        'prendre le nom dépuis le parametre
        cheminFichier = wsS
        
 
    End If
    
    

    ' 2?? Ouvrir le fichier externe en lecture seule
    Set wbSource = Workbooks.Open(cheminFichier, ReadOnly:=True)
    Set wsSource = wbSource.Sheets(1)
    
    'recupéré le cours source pour la comparaisons
    cours_source = wsSource.range("D9").Value
    
    'VERIFICATION DU COURS ET PERIODE
    '__________________________________
    
    If wsS <> "" Then
        'on part chercher le cours parmis les cours car nous sommes en mode insertion multiple
        cours = chercher_cours(CStr(cours_source))
    End If

    periode = ThisWorkbook.Sheets("DashBoard").range("I7").Value      ' <-- Période
    act = ThisWorkbook.Sheets("DashBoard").range("K7").Value     ' <-- Action d'ajouter ou de remplacer le cotes
    
    'verifier si le cours n'a pas un examen, et que la période sélectionné est un examen
    ' Chercher le cours dans la liste de cours sur le dashbord
    Dim c_line
    Set c_line = ThisWorkbook.Sheets("DashBoard").Columns("X:X").Find(What:=cours, LookIn:=xlValues, LookAt:=xlWhole)
    
    'recupérer un true ou false si le cours est examen ou pas
    a_exam = ThisWorkbook.Sheets("DashBoard").range("Y" & c_line.Row).Value
    ponderation = ThisWorkbook.Sheets("DashBoard").range("AB" & c_line.Row).Value
    
    'sortir du sub si la condition n'est pas respécter
    If periode = "EXAM1" Or periode = "EXAM2" And a_exam = "Nom" Then
        MsgBox "Le cours : " & cours & " N'a pas d'examen"
        Exit Sub
    End If
    

    If cours = "" Or periode = "" Or act = "" Then
        MsgBox "sélectionner le cours, periode, et action pour continuer", vbInformation
        Exit Sub
    End If

    ' 3?? Dernière ligne du fichier source
    lastRowSource = wsSource.Cells(wsSource.Rows.Count, "B").End(xlUp).Row
    
    
    If cours_source <> cours Then
        If MsgBox("Le cours sur la fiche est différent du cours sélectionné, il pourait y avoir les problème de podération, assurez vous que le deux cours on une la même pondération si vous souhaitez quand même continuer", vbYesNo + vbQuestion, "Confirmation") = vbNo Then
            MsgBox "Opération annulée", vbOKOnly, "Information"
            Exit Sub
        End If
    End If
    

    'parcours de tous les fiches pour choisir de prériode
    For Each f In fichesPeriode
        If f = periode Then
            'recupérer chaque fiche un par un
            Set wsDest = ThisWorkbook.Sheets(f)
            
            'compter si l'on a trouvé la période séléctionné parmis les fiches périodes
            c_p = c_p + 1
            
            'parcourir les colonnes du fiche pour localiser la colonne du période sélectionné
            For i = 5 To 10
                If wsSource.Cells(10, i).Value = f Then
                    p_fiche_col = p_fiche_col + 1
                    
                    
                    'si la colone trouvé alors passe à la recherche du destination de chauqe cote
                    '____________________________________________________________________________
                    
                    'parcourir tous les élèves sur la fiche et chercher par matricule sur la fiche période
                    Matricule = ""
                    Dim matricule_row
                    Dim cours_Col
                    For j = 12 To lastRowSource
                        Matricule = wsSource.Cells(j, 2).Value
                        
                        ' Chercher le matricule dans la fiche periode
                        Set matricule_row = wsDest.Columns("A:A").Find(What:=Matricule, LookIn:=xlValues, LookAt:=xlWhole)
                        If matricule_row Is Nothing Then
                            'enregistrer l'erreurs
                            erreurs = erreurs & " - l'élève: " & Matricule & " - " & wsSource.Cells(j, 3).Value & "non trouvé"
                            
                            'compter cette erreurs
                            e_c = e_c + 1
                        End If
                        
                        ' Chercher le cours dans la fiche periode
                        Set cours_Col = wsDest.Rows("14:14").Find(What:=cours, LookIn:=xlValues, LookAt:=xlWhole)
                        If cours_Col Is Nothing Then
                            'enregistrer l'erreurs
                            erreurs = erreurs & " - le cours : " & cours & " non trouvé"
                            
                            'compter cette erreurs
                            e_c = e_c + 1
                        End If
                        
                        
                        'nous avons déjè la linge et la colonne pour cours et matricule
                        '_______________________________________________________________
                        
                        'place à la verification de l'option choisie si remplacement, on remplace la cote
                        'même si elle existait, sinon on saute
                        
                        'coteSource recupérer la cote en cours pour un code propre
                        coteSource = wsSource.Cells(j, i)
                        
                        'verifier si la source est vide
                        If Trim(coteSource) <> "" Then
                        
                            'verifier si la cote source est un nombre
                            If IsNumeric(coteSource) Then
                                'prendre la dernière cote enregistrer pour cet élève dans ce cours
                                last_cote = wsDest.Cells(matricule_row.Row, cours_Col.Column).Value
                                
                                'verifier si le cours cote ne depasse pas la poderation du cours
                                If coteSource > poderation And coteSource < 0 Then
                                    'enregistrer l'erreurs
                                    erreurs = erreurs & " - Cote non valide, la cote ne doit pas être supérieur au ponderation ni inférieur à 0 : " & Matricule & " - " & wsSource.Cells(j, 3).Value & " > (" & coteSource & ")"
                                    
                                    'compter cette erreurs
                                    e_c = e_c + 1
                                    
                                    MsgBox "Source :" & coteSource & " ponderation :" & ponderation & " tests 2: cotesource < 0 (" & coteSource < 0 & ") "
                                Else
                                    If act = "Ajouter" Then
                                      If Not last_cote > 0 Then
                                        'compter l'opperation
                                         c = c + 1
                                         
                                        'insertion du cote
                                        wsDest.Cells(matricule_row.Row, cours_Col.Column).Value = wsSource.Cells(j, i)
                                      End If
                                    ElseIf act = "Remplacer" Then
                                        'compter l'opperation
                                         c = c + 1
                                         
                                        'insertion du cote
                                        wsDest.Cells(matricule_row.Row, cours_Col.Column).Value = wsSource.Cells(j, i)
                                    Else
                                        'enregistrer l'erreurs
                                        erreurs = erreurs & " - Action selectionner inconu, : " & act & " non trouvé"
                                        
                                        'compter cette erreurs
                                        e_c = e_c + 1
                                    End If
                                End If
                            Else
                                'enregistrer l'erreurs
                                erreurs = erreurs & " - Cote non valide, la cote doit être un nombre positif : " & Matricule & " - " & wsSource.Cells(j, 3).Value & " > (" & coteSource & ")"
                                
                                'compter cette erreurs
                                e_c = e_c + 1
                            End If
                        End If
                    Next j
                End If
            Next i
        End If
        
    Next f
    If c_p = 0 Then
        MsgBox "Nous n'avons pas trouve la période séléctionné, vérifier sons orthographe", vbInformation
    End If
    
    If p_fiche_col = 0 Then
        MsgBox "Format du fichier source nom conforme, vérifier s'il n'a pas été modifié", vbInformation
    End If
    
    
    
    MsgBox c & " Opération effectués " & e_c & erreurs, vbInformation
    
    'sauvegarder modif
    Call save
    
End Sub

Sub save()
    ThisWorkbook.save ' Enregistre le classeur actif
End Sub

Sub completer_semestres(periodes As Variant, semestre As String)
    Dim wsSource As Worksheet
    Dim wsDest As Worksheet
    
    'initialiser la destination
    Set wsDest = ThisWorkbook.Sheets(semestre)
    
    'la colonne du Periode au destination
    Dim periode_Col
    
    'la colonne du Periode au destination
    Dim max
    
    'collection des erreurs
    erreurs = "Erreur(s)"
    
    'compteur d'erreur
    e_c = 0
    
    'compteur d 'opperations
    c = 0
    
    'parcourir les période pour
    For Each f In periodes
        Set wsSource = ThisWorkbook.Sheets(f)
        
        'colonne à charcher
        col_max = "MAXIMA GEN."
        
        
        'chercher la colone qu'aucuper le manmum en cours dans la feille période destination
        Set max = wsSource.Rows("14:14").Find(What:=col_max, LookIn:=xlValues, LookAt:=xlWhole)
        If max Is Nothing Then
            
            MsgBox "La colonne de maximum n'est pas rétrouvé sur la période : " & f & " , vérifier si cette colonne existe avant de continuer, "
            Exit Sub
        End If
        
        
        'chercher la colone qu'aucuper la période en cours dans la feille période destination
        Set periode_Col = wsDest.Rows("14:14").Find(What:=f, LookIn:=xlValues, LookAt:=xlWhole)
        
        If Not periode_Col Is Nothing Then
            'Dernière ligne du fichier source dans la colonne A celui des matricueles
            lastRowSource = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row
            
            'parcourir la liste des étudiant au source
            
            Matricule = ""
            For i = 17 To lastRowSource
                Matricule = wsSource.Cells(i, 1).Value
                
                If Trim(Matricule) = "" Then
                    'enregistrer l'erreurs
                    erreurs = erreurs & " - la période : " & f & " non trouvé"
                    
                    'compter cette erreurs
                    e_c = e_c + 1
                Else
                    'chercher le matricule de l'élève dans la fiche de destination semestre
                    Dim matricule_row
                    Set matricule_row = wsDest.Columns("A:A").Find(What:=Matricule, LookIn:=xlValues, LookAt:=xlWhole)
                    If Not matricule_row Is Nothing Then
                        
                        'insertion direct du maxima
                        maxima = wsSource.Cells(i, max.Column).Value
                        If IsNumeric(maxima) Then
                            'insertion
                            wsDest.Cells(matricule_row.Row, periode_Col.Column).Value = maxima
                        Else
                            'enregistrer l'erreurs
                            erreurs = erreurs & " - Format invalide : " & maxima & " n'est pas un nombre valide"
                        End If
                    End If
                End If
            Next i
             
        Else
            'enregistrer l'erreurs
            erreurs = erreurs & " - la période : " & f & " non trouvé"
            
            'compter cette erreurs
            e_c = e_c + 1
        End If
    Next f
    MsgBox e_c & " " & erreurs
End Sub
Sub completer_tout_semestre()
    'completer la semestre 1
    Call completer_semestres(Array("P1", "P2", "EXAM1"), "SEM1")
    
    'completer la semestre 2
    Call completer_semestres(Array("P3", "P4", "EXAM2"), "SEM2")
    
    'COMPLETER TOT QUI CONTIENT LE DEUX
    'completer tot
    Call completer_semestres(Array("P1", "P2", "EXAM1", "P3", "P4", "EXAM2"), "TOT")
    MsgBox "mis à jour"
End Sub
Sub actualiser_tout()
    Call actualiser_fiche_periodes_data
    Call actuealiser_list_dash
    Call completer_tout_semestre

End Sub

Sub navigate(sheet As String)
 ThisWorkbook.Sheets(sheet).Activate
End Sub
Sub liste_proclamation(periode As String, list_proc As String)
    Dim wsSource As Worksheet
    Dim wsDest As Worksheet
    
    'initialiser la destination
    Set wsDest = ThisWorkbook.Sheets(list_proc)
    
    'la colonne du Periode au destination
    Dim max
    
    'collection des erreurs
    erreurs = "Erreur(s)"
    
    'compteur d'erreur
    e_c = 0
    
    'compteur d 'opperations
    c = 0
    
    
    
    'chercher les echecs
    
    
    Set wsSource = ThisWorkbook.Sheets(periode)
    
    'colonne à charcher partie source
    col_max = "MAXIMA GEN."
    
    
    'chercher la colone qu'aucuper le manmum en cours dans la feille période destination
    Set max = wsSource.Rows("14:14").Find(What:=col_max, LookIn:=xlValues, LookAt:=xlWhole)
    If max Is Nothing Then
        
        MsgBox "La colonne de maximum n'est pas rétrouvé sur la période : " & f & " , vérifier si cette colonne existe avant de continuer, "
        Exit Sub
    End If
    
    'Dernière ligne du fichier source dans la colonne A celui des matricueles
    lastRowSource = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row
    
    'delimitateur
    cama = ", "
    
    Matricule = ""
    For i = 17 To lastRowSource
        Matricule = wsSource.Cells(i, 1).Value
        
        If Trim(Matricule) = "" Then
            'enregistrer l'erreurs
            erreurs = erreurs & " - la période : " & f & " non trouvé"
            
            'compter cette erreurs
            e_c = e_c + 1
        Else
            'chercher le matricule de l'élève dans la fiche de destination semestre
            Dim matricule_row
            Set matricule_row = wsDest.Columns("A:A").Find(What:=Matricule, LookIn:=xlValues, LookAt:=xlWhole)
            If Not matricule_row Is Nothing Then
                
                'compteur d'échecs
                echec_c = 0
                
                'garder les cours sur lesquelles l'élève a échoues
                cours_echcs = ""
                
                'chercher les echecs
                
                For j = 6 To 25
                    cote = wsSource.Cells(i, j).Value
                    pond = wsSource.Cells(15, j).Value
                    If Not pond = "" Then
                        If cote < pond / 2 Then
                            'compter echec
                            echec_c = echec_c + 1
                            
                            'augmenter le cours aux cours échoués
                            cours_echcs = cours_echcs & wsSource.Cells(14, j).Value & IIf(j < 25, cama, "")
                        End If
                    End If
                Next j
                
                'insertion direct du maxima
                maxima = wsSource.Cells(i, max.Column).Value
                pourc = wsSource.Cells(i, max.Column + 1).Value
                
                If IsNumeric(maxima) Then
                    'insertion
                    wsDest.Cells(matricule_row.Row, 4).Value = maxima
                    wsDest.Cells(matricule_row.Row, 5).Value = pourc
                    wsDest.Cells(matricule_row.Row, 7).Value = echec_c
                    wsDest.Cells(matricule_row.Row, 9).Value = cours_echcs
                    
                Else
                    'enregistrer l'erreurs
                    erreurs = erreurs & " - Format invalide : " & maxima & " n'est pas un nombre valide"
                    
                    'compter cette erreurs
                    e_c = e_c + 1
                End If
            End If
        End If
    Next i
         
    MsgBox e_c & " " & erreurs
End Sub
Sub proclamer()
    periode = ThisWorkbook.Sheets("DashBoard").range("N5").Value
    
    Call liste_proclamation(CStr(periode), "PROC1")
End Sub

Option Explicit

'---------------------------------------------------------
'  Boîte pour choisir un dossier
'---------------------------------------------------------
Function ChoisirDossier() As String
    Dim dlg As FileDialog
    Set dlg = Application.FileDialog(msoFileDialogFolderPicker)

    With dlg
        .title = "Choisissez l'emplacement pour l'archive"
        .AllowMultiSelect = False

        If .Show = -1 Then
            ChoisirDossier = .SelectedItems(1)
        Else
            ChoisirDossier = ""
        End If
    End With
End Function


'---------------------------------------------------------
'  MACRO PRINCIPALE - TOUT EN UN
'---------------------------------------------------------
Sub GenererArchiveEtPDF()

    Dim CheminBase As String
    Dim AnneeScolaire As String
    Dim DossierArchive As String
    Dim DossierBulletins As String
    Dim wbArchive As Workbook
    Dim FeuillesArchive As Variant
    Dim f As Variant

    Dim wsList As Worksheet
    Dim wsBulletin As Worksheet
    Dim MatriculeCell As range
    Dim DerniereLigne As Long
    Dim Matricule As String, NomEleve As String
    Dim FichierPDF As String
    Dim i As Long

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' 1. Demander l'année scolaire
    '---------------------------------------------------------
    classe = ThisWorkbook.Sheets("DashBoard").range("S7").Value
    AnneeScolaire = ThisWorkbook.Sheets("DashBoard").range("U5").Value

    If Trim(AnneeScolaire) = "" Then
        MsgBox "Année scolaire invalide.", vbCritical
        Exit Sub
    End If


    ' 2. Choisir l’emplacement de sauvegarde
    '---------------------------------------------------------
    CheminBase = ChoisirDossier()

    If CheminBase = "" Then
        MsgBox "Aucun emplacement choisi.", vbExclamation
        Exit Sub
    End If

    ' 3. Créer dossier Archive_AnnéeScolaire
    '---------------------------------------------------------
    DossierArchive = CheminBase & "\Archive_" & classe & "_" & AnneeScolaire
    If Dir(DossierArchive, vbDirectory) = "" Then MkDir DossierArchive

    ' 4. Créer sous-dossier des bulletins PDF
    '---------------------------------------------------------
    DossierBulletins = DossierArchive & "\Bulletins"
    If Dir(DossierBulletins, vbDirectory) = "" Then MkDir DossierBulletins

    ' 5. Copier les feuilles dans un nouveau classeur
    '---------------------------------------------------------
    FeuillesArchive = Array("Eleves", "P1", "P2", "P3", "P4", "EXAM1", "EXAM2", "SEM1", "SEM2", "TOT", "Bulletin")

    Set wbArchive = Workbooks.Add

    Do While wbArchive.Sheets.Count > 1
        wbArchive.Sheets(1).Delete
    Loop

    For Each f In FeuillesArchive
        ThisWorkbook.Sheets(f).Copy After:=wbArchive.Sheets(wbArchive.Sheets.Count)
    Next f

    wbArchive.Sheets(1).Delete

    wbArchive.SaveAs filename:=DossierArchive & "\Archive_" & AnneeScolaire & ".xlsx"
    wbArchive.Close SaveChanges:=True


    '---------------------------------------------------------
    ' 6. Générer les bulletins PDF
    '---------------------------------------------------------
    

    Set wsBulletin = Sheets("Bulletin") ' feuille avec le bulletin
    Set MatriculeCell = wsBulletin.range("K13") ' Cellule avec la liste déroulante
    Set wsList = Sheets("Eleves") ' La liste complète des noms à parcourir

    DerniereLigne = wsList.range("A" & Rows.Count).End(xlUp).Row

    For i = 8 To DerniereLigne

        Matricule = wsList.Cells(i, 2).Value
        NomEleve = wsList.Cells(i, 3).Value

        MatriculeCell.Value = Matricule
        DoEvents
        On Error GoTo GestionErreur ' Gestion d'erreur de base
        
        FichierPDF = DossierBulletins & "\" & "B_" & NomEleve & "_" & classe & "_" & AnneeScolaire & ".pdf"

        wsBulletin.range("C2:Q63").ExportAsFixedFormat _
            Type:=xlTypePDF, _
            filename:=FichierPDF, _
            Quality:=xlQualityStandard, _
            IncludeDocProperties:=True, _
            IgnorePrintAreas:=False, _
            OpenAfterPublish:=False
        

    Next i

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox "ARCHIVE ET BULLETINS GÉNÉRÉS AVEC SUCCÈS !" & vbCrLf & vbCrLf & _
           "Dossier : " & DossierArchive, vbInformation
    
GestionErreur:     ' En cas d'erreur
    MsgBox "Erreur lors de la création du PDF : " & Err.Description, vbCritical

End Sub

Sub ImprimerTousLesBulletinspOut()
    Dim listeNoms As range
    Dim nomCellule As range
    Dim celluleListe As range
    Dim feuille As Worksheet
    Dim celluleListeDeroulante As range
    
    ' Définis ici :
    Set feuille = Sheets("Bulletin") ' Remplace par le nom de ta feuille avec le bulletin
    Set celluleListeDeroulante = feuille.range("K13") ' Cellule avec la liste déroulante
    Set listeNoms = Sheets(" P1").range("E9:E66") ' La liste complète des noms à parcourir

    Application.ScreenUpdating = False
    
    reponse = MsgBox("Voulez vous vraiment imprimer tous les bulletins", vbYesNo + vbQuestion, "Confirmation")
    If reponse = vbNo Then
        MsgBox "Opération annulé", vbInformation
        Exit Sub
    End If
    
    For Each nomCellule In listeNoms
        If nomCellule.Value <> "" Then
            celluleListeDeroulante.Value = nomCellule.Value
            DoEvents ' Laisse le temps à Excel de mettre à jour les données du bulletin
            
            ' Imprime la feuille active ou une zone d'impression
            feuille.PrintOut
            
        End If
    Next nomCellule

    Application.ScreenUpdating = True
    MsgBox "Tous les bulletins ont été imprimés avec succès."
End Sub

Sub ExporterPDFConcise()
    Dim ws As Worksheet, plagePDF As range, cheminComplet As String, nomFichier As String

    Set ws = ThisWorkbook.Sheets("Facture") ' Nom de votre feuille
    Set plagePDF = ws.range("A1:E23")       ' Votre zone d'impression

    ' Nom du fichier PDF (ex: Facture_123.pdf)
    nomFichier = "Facture_" & ws.range("B2").Value & ".pdf"
    ' Chemin complet du fichier (dossier du classeur + nom du fichier)
    cheminComplet = ThisWorkbook.Path & "\" & nomFichier

    On Error GoTo GestionErreur ' Gestion d'erreur de base

    ' Exporte la plage sélectionnée en PDF
    plagePDF.ExportAsFixedFormat Type:=xlTypePDF, _
                                 filename:=cheminComplet, _
                                 OpenAfterPublish:=True ' Ouvre le PDF après création

    MsgBox "PDF créé avec succès : " & cheminComplet, vbInformation
    Exit Sub ' Quitte la sub pour ne pas passer par la gestion d'erreur

GestionErreur: ' En cas d'erreur
    MsgBox "Erreur lors de la création du PDF : " & Err.Description, vbCritical
End Sub

Option Explicit

Sub ImportationProComplete()

    Dim dlg As FileDialog
    Dim dossier As String
    Dim fichier As String
    Dim wbSource As Workbook, wsSource As Worksheet
    Dim wsDest As Worksheet, wsLog As Worksheet
    Dim lastRowDest As Long, lastRow As Long
    Dim lastLog As Long
    Dim ligne As Long
    Dim nbLignesImp As Long
    Dim colImport As Long
    Dim cell As range

    ' 1. Sélection du dossier
    '---------------------------------------------
    Set dlg = Application.FileDialog(msoFileDialogFolderPicker)
    dlg.title = "Sélectionner le dossier des fichiers Excel"

    If dlg.Show <> -1 Then Exit Sub
    dossier = dlg.SelectedItems(1) & "\"

    ' 2. Feuille destination
    '---------------------------------------------
    Set wsDest = ThisWorkbook.Sheets("Feuil1") '?? Modifier si besoin
    
    ' 3. Feuille LOG
    '---------------------------------------------
    On Error Resume Next
    Set wsLog = ThisWorkbook.Sheets("LOG")
    On Error GoTo 0

    If wsLog Is Nothing Then
        Set wsLog = ThisWorkbook.Worksheets.Add
        wsLog.Name = "LOG"
        wsLog.range("A1:F1") = Array("Fichier", "Feuille", "Lignes importées", "Date", "Heure", "Chemin complet")
        wsLog.Rows(1).Font.Bold = True
    End If

    ' 4. Parcours de tous les fichiers
    '---------------------------------------------
    fichier = Dir(dossier & "*.xlsx")

    Do While fichier <> ""

        Set wbSource = Workbooks.Open(dossier & fichier)

        '---------------------------------------------
        ' 5. Parcours de toutes les feuilles du fichier
        '---------------------------------------------
        For Each wsSource In wbSource.Worksheets

            nbLignesImp = 0

            ' Trouver dernière colonne et dernière ligne
            lastRow = wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).Row

            If lastRow < 2 Then GoTo FeuilleSuivante

            ' Vérifier si colonne "Importé" existe
            colImport = 0
            On Error Resume Next
            colImport = wsSource.Rows(1).Find("Importé", LookAt:=xlWhole).Column
            On Error GoTo 0

            ' Si elle n'existe pas ? créer
            If colImport = 0 Then
                colImport = wsSource.Cells(1, wsSource.Columns.Count).End(xlToLeft).Column + 1
                wsSource.Cells(1, colImport).Value = "Importé"
            End If
            
            ' 6. Importation ligne par ligne
            '---------------------------------------------
            For ligne = 2 To lastRow

                ' Sauter si déjà importé
                If wsSource.Cells(ligne, colImport).Value <> "" Then GoTo LigneSuivante

                ' Copier la ligne complète
                lastRowDest = wsDest.Cells(wsDest.Rows.Count, 1).End(xlUp).Row + 1
                wsSource.Rows(ligne).Copy wsDest.Rows(lastRowDest)

                ' Marquer dans le fichier source
                wsSource.Cells(ligne, colImport).Value = "Importé le : " & Format(Now, "dd/mm/yyyy HH:NN")

                nbLignesImp = nbLignesImp + 1

LigneSuivante:
            Next ligne

            ' 7. Log de cette feuille
            '---------------------------------------------
            lastLog = wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Row + 1

            wsLog.Cells(lastLog, 1).Value = fichier
            wsLog.Cells(lastLog, 2).Value = wsSource.Name
            wsLog.Cells(lastLog, 3).Value = nbLignesImp
            wsLog.Cells(lastLog, 4).Value = Date
            wsLog.Cells(lastLog, 5).Value = Time
            wsLog.Cells(lastLog, 6).Value = dossier & fichier

FeuilleSuivante:
        Next wsSource

        wbSource.Close SaveChanges:=True

        fichier = Dir()
    Loop

    ' 8. Mise en forme LOG (expert)
    '---------------------------------------------
    With wsLog
        .Columns("A:F").AutoFit
        .Rows(1).Font.Bold = True
        .range("A1:F" & .Cells(.Rows.Count, 1).End(xlUp).Row).Borders.LineStyle = xlContinuous
    End With

    MsgBox "Importation professionnelle terminée !", vbInformation

End Sub

Sub insertion_cotes_automatique_miltifier()

    Dim dlg As FileDialog
    Dim dossier As String
    Dim fichier As String
    Dim wbSource As Workbook, wsSource As Worksheet
    Dim wsDest As Worksheet, wsLog As Worksheet
    Dim lastRowDest As Long, lastRow As Long
    Dim lastLog As Long
    Dim ligne As Long
    Dim nbLignesImp As Long
    Dim colImport As Long
    Dim cell As range

    ' 1. Sélection du dossier
    '---------------------------------------------
    Set dlg = Application.FileDialog(msoFileDialogFolderPicker)
    dlg.title = "Sélectionner le dossier des fichiers Excel"

    If dlg.Show <> -1 Then Exit Sub
    dossier = dlg.SelectedItems(1) & "\"

    ' 3. Feuille LOG
    '---------------------------------------------
    On Error Resume Next
    Set wsLog = ThisWorkbook.Sheets("LOG")
    On Error GoTo 0

    If wsLog Is Nothing Then
        Set wsLog = ThisWorkbook.Worksheets.Add
        wsLog.Name = "LOG"
        wsLog.range("A1:E1") = Array("Fichier", "Date", "Heure", "Chemin complet")
        wsLog.Rows(1).Font.Bold = True
    End If

    ' 4. Parcours de tous les fichiers
    '---------------------------------------------
    fichier = Dir(dossier & "*.xlsx")

    Do While fichier <> ""
       
       
       
        Call importer_cotes_sur_fiche(dossier & fichier)

        ' 7. Log de cette feuille
        '---------------------------------------------
        lastLog = wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Row + 1

        wsLog.Cells(lastLog, 1).Value = fichier
        wsLog.Cells(lastLog, 3).Value = Date
        wsLog.Cells(lastLog, 4).Value = Time
        wsLog.Cells(lastLog, 5).Value = dossier & fichier

        fichier = Dir()
    Loop

    ' 8. Mise en forme LOG (expert)
    '---------------------------------------------
    With wsLog
        .Columns("A:E").AutoFit
        .Rows(1).Font.Bold = True
        .range("A1:E" & .Cells(.Rows.Count, 1).End(xlUp).Row).Borders.LineStyle = xlContinuous
    End With

    MsgBox "Importation professionnelle terminée !", vbInformation

End Sub
Function STAT_ELEVE(genre As String, critere As String) As Integer
    Dim wsSource As Worksheet
    

    Set wsSource = Sheets("Eleves")
    
    ' 3?? Dernière ligne du fichier source
    lastRowSource = wsSource.Cells(wsSource.Rows.Count, "B").End(xlUp).Row
 
    tot = 0

    For i = 8 To lastRowSource  ' saute les en-têtes
        If critere = "" Then
            If wsSource.Cells(i, 4).Value = genre Then
                'incrementer que quand le critère est valide
                tot = tot + 1
            End If
        ElseIf wsSource.Cells(i, 6).Value = critere And wsSource.Cells(i, 4).Value Then
            'incrementer que quand le critère est valide
            tot = tot + 1
            
        End If
    Next i
    
    STAT_ELEVE = tot
End Function

Sub Bulletin_to_generate_pdf(file_name As String)
    
    Dim wsList As Worksheet
    Dim wsBulletin As Worksheet
    Dim MatriculeCell As range
    
    Set wsBulletin = Sheets("Bulletin") ' feuille avec le bulletin
    Set MatriculeCell = wsBulletin.range("K13") ' Cellule avec la liste déroulante
    Set wsList = Sheets("Eleves") ' La liste complète des noms à parcourir

    DerniereLigne = wsList.range("A" & Rows.Count).End(xlUp).Row

    For i = 8 To DerniereLigne

        Matricule = wsList.Cells(i, 2).Value
        NomEleve = wsList.Cells(i, 3).Value

        MatriculeCell.Value = Matricule
        DoEvents
        On Error GoTo GestionErreur ' Gestion d'erreur de base
        
        If file_name <> "" Then
        End If
        
        FichierPDF = DossierBulletins & "\" & "B_" & NomEleve & "_" & classe & "_" & AnneeScolaire & ".pdf"

        If file_name <> "" Then
            FichierPDF = file_name
        Else
            ' 1. Sélection du dossier
            '---------------------------------------------
            Set dlg = Application.FileDialog(msoFileDialogFolderPicker)
            dlg.title = "Sélectionner le dossier des fichiers Excel"
        
            If dlg.Show <> -1 Then Exit Sub
            dossier = dlg.SelectedItems(1) & "\"
                End If
        
        wsBulletin.range("C2:Q63").ExportAsFixedFormat _
            Type:=xlTypePDF, _
            filename:=FichierPDF, _
            Quality:=xlQualityStandard, _
            IncludeDocProperties:=True, _
            IgnorePrintAreas:=False, _
            OpenAfterPublish:=False
        

    Next i

End Sub
