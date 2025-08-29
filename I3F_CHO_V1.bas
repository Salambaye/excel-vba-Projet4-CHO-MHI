Attribute VB_Name = "Module1"
'Salamata Nourou MBAYE - 29/08/2025 - Version 2.1 Simplifiée
'Projet 4 - Programme 2 - CHO pour I3F

Sub CHO()

    '---------------------- Optimisation pour accélérer la macro --------------------------
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    
    ' ________________________________ETAPE 1 : Déclaration des variables ______________________________________________________
    
    Dim fdlg As FileDialog
    Dim nomFichier As String
    Dim cheminFichierTxt As String
    Dim cheminFichierExcel As String
    Dim cheminSortie As String
    Dim contenu As String
    Dim lignes As Variant
    Dim i, j As Long
    Dim numFichier As Long
    
    Dim dossierSauvegarde As String
    Dim fdlgDossier As FileDialog
    
    ' Variables pour les fichiers de sortie
    Dim lignesPrestatTiers As String
    Dim lignesAucunePrestation As String
    Dim lignesSDCNonGeneres As String
    Dim lignesRemiseEDI As String
    
    ' Variables pour le traitement
    Dim codeAgence As String
    Dim uex As String
    Dim ligne As String
    Dim wsSDC As Worksheet
    Dim wbExcel As Workbook
    Dim ligneExcel As Long
    Dim marcheValue As String
    Dim uexValue As String
    
    ' Arrays simples pour stocker les données Excel
    Dim uexSDC() As String ' Toutes les UEX de la feuille SDC
    Dim uexDelete() As String ' Les UEX avec MARCHE = DELETE
    Dim nbUexSDC As Long
    Dim nbUexDelete As Long
    
    ' Compteurs pour le rapport
    Dim compteurPrestatTiers As Long
    Dim compteurAucunePrestation As Long
    Dim compteurSDCNonGeneres As Long
    Dim compteurRemiseEDI As Long
    
    ' _____________________________ Etape 2 : Sélection du fichier TXT d'entrée ________________________________________
    
    MsgBox "Sélectionner le fichier TXT d'entrée"
    Set fdlg = Application.FileDialog(msoFileDialogFilePicker)
    fdlg.Title = "Sélection du fichier TXT d'entrée"
    fdlg.Filters.Clear
    fdlg.Filters.Add "Fichiers TXT", "*.txt"
    fdlg.AllowMultiSelect = False

    If fdlg.Show <> -1 Then
        MsgBox "Sélection annulée par l'utilisateur.", vbInformation
        GoTo Fin
    End If
    
    cheminFichierTxt = fdlg.SelectedItems(1)
    
    ' --------------- Vérification du fichier TXT -------------
    If Dir(cheminFichierTxt) = "" Then
        MsgBox "Le fichier TXT n'existe pas : " & cheminFichierTxt, vbCritical
        GoTo Fin
    End If
    
    ' _____________________________ Etape 3 : Sélection du fichier Excel "découpage CHO-MHI.xlsx" ________________________________________
    
    MsgBox "Sélectionner le fichier Excel 'découpage CHO-MHI.xlsx'"
    Set fdlg = Application.FileDialog(msoFileDialogFilePicker)
    fdlg.Title = "Sélection du fichier découpage CHO-MHI.xlsx"
    fdlg.Filters.Clear
    fdlg.Filters.Add "Fichiers Excel", "*.xlsx"
    fdlg.AllowMultiSelect = False

    If fdlg.Show <> -1 Then
        MsgBox "Sélection annulée par l'utilisateur.", vbInformation
        GoTo Fin
    End If
    
    cheminFichierExcel = fdlg.SelectedItems(1)
    
    ' --------------- Vérification du fichier Excel -------------
    If Dir(cheminFichierExcel) = "" Then
        MsgBox "Le fichier Excel n'existe pas : " & cheminFichierExcel, vbCritical
        GoTo Fin
    End If
    
    ' ____________ Etape 4 : Sélection du dossier de sauvegarde ______________
    
    MsgBox "Sélectionner le dossier de sauvegarde des fichiers"
    Set fdlgDossier = Application.FileDialog(msoFileDialogFolderPicker)
    With fdlgDossier
        .Title = "Sélectionner le dossier de sauvegarde des fichiers"
        .AllowMultiSelect = False
        .InitialFileName = Environ("USERPROFILE") & "\DESKTOP\"
    End With
    
    If fdlgDossier.Show <> -1 Then
        MsgBox "Sélection du dossier annulée par l'utilisateur.", vbInformation
        GoTo Fin
    End If
    
    dossierSauvegarde = fdlgDossier.SelectedItems(1)
    
    ' Vérifier que le dossier existe et est accessible
    If Dir(dossierSauvegarde, vbDirectory) = "" Then
        MsgBox "Le dossier sélectionné n'est pas accessible : " & dossierSauvegarde, vbCritical
        GoTo Fin
    End If
    
    ' ________________________________ ETAPE 5 : Lecture du fichier Excel et stockage des données ________________________________
    
    Set wbExcel = Workbooks.Open(cheminFichierExcel, ReadOnly:=True)
    Set wsSDC = wbExcel.Worksheets("SDC-MARCHE NOK")
    
    ' Initialiser les compteurs
    nbUexSDC = 0
    nbUexDelete = 0
    
    ' Première passe : compter les lignes pour dimensionner les tableaux
    ligneExcel = 2 ' Commencer à la ligne 2 (ligne 1 = en-têtes)
    Do While wsSDC.Cells(ligneExcel, 2).Value <> "" ' Tant qu'il y a des UEX en colonne B
        nbUexSDC = nbUexSDC + 1
        If UCase(Trim(wsSDC.Cells(ligneExcel, 4).Value)) = "DELETE" Then ' Colonne D = MARCHE
            nbUexDelete = nbUexDelete + 1
        End If
        ligneExcel = ligneExcel + 1
    Loop
    
    ' Redimensionner les tableaux
    ReDim uexSDC(1 To nbUexSDC)
    If nbUexDelete > 0 Then ReDim uexDelete(1 To nbUexDelete)
    
    ' Deuxième passe : remplir les tableaux
    ligneExcel = 2
    Dim indexSDC As Long, indexDelete As Long
    indexSDC = 1
    indexDelete = 1
    
    Do While wsSDC.Cells(ligneExcel, 2).Value <> ""
        uexValue = Trim(wsSDC.Cells(ligneExcel, 2).Value) ' Colonne B = UEX
        marcheValue = Trim(wsSDC.Cells(ligneExcel, 4).Value) ' Colonne D = MARCHE
        
        ' Stocker toutes les UEX
        uexSDC(indexSDC) = uexValue
        indexSDC = indexSDC + 1
        
        ' Stocker les UEX avec MARCHE = DELETE
        If UCase(marcheValue) = "DELETE" Then
            uexDelete(indexDelete) = uexValue
            indexDelete = indexDelete + 1
        End If
        
        ligneExcel = ligneExcel + 1
    Loop
    
    ' Fermer le fichier Excel
    wbExcel.Close SaveChanges:=False
    
    ' ________________________________ ETAPE 6 : Lecture du fichier TXT ________________________________
    
    numFichier = FreeFile
    Open cheminFichierTxt For Input As #numFichier
    contenu = Input$(LOF(numFichier), numFichier)
    Close #numFichier
    
    lignes = Split(contenu, vbCrLf)
    
    ' ________________________________ ETAPE 7 : Traitement des lignes du fichier TXT ________________________________
    
    ' Initialiser les variables de sortie
    lignesPrestatTiers = ""
    lignesAucunePrestation = ""
    lignesSDCNonGeneres = ""
    lignesRemiseEDI = ""
    
    ' Traiter chaque ligne
    For i = 0 To UBound(lignes)
        ligne = lignes(i)
        
        If Len(ligne) >= 10 Then ' S'assurer que la ligne a au moins 10 caractères
            ' Extraire le code d'agence (2 caractères à partir de la colonne 3, donc positions 3-4)
            codeAgence = Mid(ligne, 3, 2)
            
            ' Extraire l'UEX (6 caractères suivant le code d'agence, donc positions 5-10)
            uex = Mid(ligne, 5, 6)
            
            ' Variable pour savoir si la ligne a été traitée
            Dim ligneTraitee As Boolean
            ligneTraitee = False
            
            ' --- FICHIER 1 : Presta Tiers.txt ---
            ' Codes d'agence : 08, 10, 15, 30, 74, 75, 93
            If codeAgence = "08" Or codeAgence = "10" Or codeAgence = "15" Or _
               codeAgence = "30" Or codeAgence = "74" Or codeAgence = "75" Or codeAgence = "93" Then
                lignesPrestatTiers = lignesPrestatTiers & ligne & vbCrLf
                compteurPrestatTiers = compteurPrestatTiers + 1
                ligneTraitee = True
            End If
            
            ' --- FICHIER 2 : Aucune prestation pour ce programme.txt ---
            If Not ligneTraitee Then
                ' Codes d'agence 03 et 07
                If codeAgence = "03" Or codeAgence = "07" Then
                    lignesAucunePrestation = lignesAucunePrestation & ligne & vbCrLf
                    compteurAucunePrestation = compteurAucunePrestation + 1
                    ligneTraitee = True
                Else
                    ' Vérifier si l'UEX correspond à un MARCHE = DELETE
                    For j = 1 To nbUexDelete
                        If uex = uexDelete(j) Then
                            lignesAucunePrestation = lignesAucunePrestation & ligne & vbCrLf
                            compteurAucunePrestation = compteurAucunePrestation + 1
                            ligneTraitee = True
                            Exit For ' Sortir de la boucle dès qu'on trouve
                        End If
                    Next j
                End If
            End If
            
            ' --- FICHIER 3 : SDC non générés.txt ---
            If Not ligneTraitee Then
                ' Vérifier si l'UEX est présent dans la feuille SDC
                For j = 1 To nbUexSDC
                    If uex = uexSDC(j) Then
                        lignesSDCNonGeneres = lignesSDCNonGeneres & ligne & vbCrLf
                        compteurSDCNonGeneres = compteurSDCNonGeneres + 1
                        ligneTraitee = True
                        Exit For ' Sortir de la boucle dès qu'on trouve
                    End If
                Next j
            End If
            
            ' --- FICHIER 4 : RemiseEDI.txt ---
            ' Toutes les autres lignes
            If Not ligneTraitee Then
                lignesRemiseEDI = lignesRemiseEDI & ligne & vbCrLf
                compteurRemiseEDI = compteurRemiseEDI + 1
            End If
        End If
    Next i
    
    ' ________________________________ ETAPE 8 : Sauvegarde des fichiers de sortie ________________________________
    
    Dim timestamp As String
    timestamp = Format(Now, "yyyymmdd_hhmmss")
    
    ' Fichier 1 : Presta Tiers.txt
    If compteurPrestatTiers > 0 Then
        cheminSortie = dossierSauvegarde & "\Presta_Tiers_" & timestamp & ".txt"
        numFichier = FreeFile
        Open cheminSortie For Output As #numFichier
        Print #numFichier, Left(lignesPrestatTiers, Len(lignesPrestatTiers) - 2) ' Enlever le dernier vbCrLf
        Close #numFichier
    End If
    
    ' Fichier 2 : Aucune prestation pour ce programme.txt
    If compteurAucunePrestation > 0 Then
        cheminSortie = dossierSauvegarde & "\Aucune_prestation_pour_ce_programme_" & timestamp & ".txt"
        numFichier = FreeFile
        Open cheminSortie For Output As #numFichier
        Print #numFichier, Left(lignesAucunePrestation, Len(lignesAucunePrestation) - 2)
        Close #numFichier
    End If
    
    ' Fichier 3 : SDC non générés.txt
    If compteurSDCNonGeneres > 0 Then
        cheminSortie = dossierSauvegarde & "\SDC_non_generes_" & timestamp & ".txt"
        numFichier = FreeFile
        Open cheminSortie For Output As #numFichier
        Print #numFichier, Left(lignesSDCNonGeneres, Len(lignesSDCNonGeneres) - 2)
        Close #numFichier
    End If
    
    ' Fichier 4 : RemiseEDI.txt
    If compteurRemiseEDI > 0 Then
        cheminSortie = dossierSauvegarde & "\RemiseEDI_" & timestamp & ".txt"
        numFichier = FreeFile
        Open cheminSortie For Output As #numFichier
        Print #numFichier, Left(lignesRemiseEDI, Len(lignesRemiseEDI) - 2)
        Close #numFichier
    End If
    
    ' ________________________________ ETAPE 9 : Création du rapport d'anomalies (Fichier 5) ________________________________
    
    Call CreerRapportAnomalies(dossierSauvegarde, timestamp, compteurPrestatTiers, _
                              compteurAucunePrestation, compteurSDCNonGeneres, compteurRemiseEDI)
    
    ' ------------------------------- Message de fin de traitement --------------------------
    MsgBox "Traitement terminé avec succès !" & vbCrLf & vbCrLf & _
           "Résumé :" & vbCrLf & _
           "• Fichier 1 (Presta Tiers) : " & compteurPrestatTiers & " lignes" & vbCrLf & _
           "• Fichier 2 (Aucune prestation) : " & compteurAucunePrestation & " lignes" & vbCrLf & _
           "• Fichier 3 (SDC non générés) : " & compteurSDCNonGeneres & " lignes" & vbCrLf & _
           "• Fichier 4 (RemiseEDI) : " & compteurRemiseEDI & " lignes" & vbCrLf & _
           "• Fichier 5 (Rapport) : Créé", vbInformation

    ' Ouvrir le dossier contenant les fichiers créés
    Shell "explorer.exe """ & dossierSauvegarde & """", vbNormalFocus

Fin:
    ' ------------------------ Nettoyer les références ------------------------------------
    Set fdlg = Nothing
    Set fdlgDossier = Nothing
    Set wbExcel = Nothing
    Set wsSDC = Nothing
    
    ' ----------------------------------- Restaurer les paramètres --------------------------------
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayAlerts = True

End Sub

' ________________________________ PROCEDURE POUR CREER LE RAPPORT (FICHIER 5) ________________________________

Sub CreerRapportAnomalies(dossierSauvegarde As String, timestamp As String, _
                         compteurPrestatTiers As Long, compteurAucunePrestation As Long, _
                         compteurSDCNonGeneres As Long, compteurRemiseEDI As Long)
    
    Dim wbRapport As Workbook
    Dim wsRapport As Worksheet
    Dim cheminRapport As String
    Dim totalLignes As Long
    
    ' Calculer le total
    totalLignes = compteurPrestatTiers + compteurAucunePrestation + compteurSDCNonGeneres + compteurRemiseEDI
    
    ' Créer un nouveau classeur
    Set wbRapport = Workbooks.Add
    Set wsRapport = wbRapport.ActiveSheet
    
    ' Configurer le rapport
    With wsRapport
        .Name = "Rapport_Anomalies"
        
        ' Titre
        .Cells(1, 1).Value = "RAPPORT D'ANOMALIES - CHO"
        .Cells(1, 1).Font.Bold = True
        .Cells(1, 1).Font.Size = 16
        .Range("A1:B1").Merge
        
        ' Date/heure
        .Cells(3, 1).Value = "Date/Heure du traitement :"
        .Cells(3, 2).Value = Format(Now, "dd/mm/yyyy hh:mm:ss")
        .Cells(3, 1).Font.Bold = True
        
        ' En-têtes du tableau
        .Cells(5, 1).Value = "Fichier de sortie"
        .Cells(5, 2).Value = "Nombre de lignes"
        .Range("A5:B5").Font.Bold = True
        .Range("A5:B5").Interior.Color = RGB(200, 200, 200)
        
        ' Données
        .Cells(6, 1).Value = "1. Presta Tiers"
        .Cells(6, 2).Value = compteurPrestatTiers
        
        .Cells(7, 1).Value = "2. Aucune prestation pour ce programme"
        .Cells(7, 2).Value = compteurAucunePrestation
        
        .Cells(8, 1).Value = "3. SDC non générés"
        .Cells(8, 2).Value = compteurSDCNonGeneres
        
        .Cells(9, 1).Value = "4. RemiseEDI"
        .Cells(9, 2).Value = compteurRemiseEDI
        
        ' Total
        .Cells(11, 1).Value = "TOTAL lignes traitées :"
        .Cells(11, 2).Value = totalLignes
        .Range("A11:B11").Font.Bold = True
        .Range("A11:B11").Interior.Color = RGB(255, 255, 0)
        
        ' Mise en forme
        .Columns("A:B").AutoFit
        .Range("A5:B11").Borders.LineStyle = xlContinuous
        .Range("A1:B11").HorizontalAlignment = xlLeft
    End With
    
    ' Sauvegarder le rapport
    cheminRapport = dossierSauvegarde & "\Rapport_ANOMALIES_" & timestamp & ".xlsx"
    wbRapport.SaveAs cheminRapport
    wbRapport.Close
    
    Set wbRapport = Nothing
    Set wsRapport = Nothing
    
End Sub



