Attribute VB_Name = "Module1"
'Salamata Nourou MBAYE - 29/08/2025 - Version 1.0
'Projet 4 - Programme 2 - CHO pour I3F


Sub CHO()

  '---------------------- Optimisation pour acc�l�rer la macro --------------------------
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    
    
' ________________________________ETAPE 1 : Initialiser les rapports d'anomalies et d'activit�s______________________________________________

'    Call InitialiserRapportActivites
'    Call DemarrerNouveauTraitement
'    Call InitialiserRapportAnomalies


' ________________________________ETAPE 2 : D�claration des variables ______________________________________________________

    Dim fdlg As FileDialog
    Dim nomFichier As String
    Dim cheminFichier As String
    Dim cheminSortie As String
    Dim contenu As String
    Dim contenuModifie As String
    Dim lignes As Variant
    Dim i, j As Long
    Dim numFichier As Long
    
    Dim dossierSauvegarde As String
    Dim fdlgDossier As FileDialog
    
    'Variables pour multi-fichiers
    Dim fichiersSelectionnes() As String
    Dim compteurLignes As Long
    Dim compteurLignesTotal As Long
    Dim compteurFichiers As Long
    Dim numFichierCourant As Long
    Dim dernierCheminSortie As String



' _____________________________ Etape 3 : S�lection du ou des fichiers SYM (RFC et/ou CET) ________________________________________

    MsgBox "S�lectionner le(les) fichier(s) SYM de type csv"
    Set fdlg = Application.FileDialog(msoFileDialogFilePicker)
    fdlg.Title = "S�lection du fichier SYM RFC ou CET de type csv"
    fdlg.Filters.Clear
    fdlg.Filters.Add "Fichiers CSV", "*.csv"
    fdlg.AllowMultiSelect = True

    If fdlg.Show <> -1 Then
        MsgBox "S�lection annul�e par l'utilisateur.", vbInformation
        GoTo Fin
    End If
    
    ' V�rification qu'au moins un fichier est s�lectionn�
    If fdlg.SelectedItems.Count = 0 Then
        MsgBox "Aucun fichier s�lectionn�.", vbInformation
        GoTo Fin
    End If
    
    cheminFichier = fdlg.SelectedItems(1)

 ' --------------- V�rification du fichier -------------
    If Dir(cheminFichier) = "" Then
        MsgBox "Le fichier n'existe pas : " & cheminFichier, vbCritical
        GoTo Fin
    End If
    
    ' Stocker tous les fichiers s�lectionn�s
    ReDim fichiersSelectionnes(1 To fdlg.SelectedItems.Count)
    For i = 1 To fdlg.SelectedItems.Count
        fichiersSelectionnes(i) = fdlg.SelectedItems(i)
    Next i
    
    compteurFichiers = UBound(fichiersSelectionnes)

' ____________ Etape 4 : S�lection du dossier de sauvegarde du fichier final csv et des rapports d'anomalies et d'activit�s ______________

    MsgBox "S�lectionner le dossier de sauvegarde des fichiers "
    Set fdlgDossier = Application.FileDialog(msoFileDialogFolderPicker)
    With fdlgDossier
        .Title = "S�lectionner le dossier de sauvegarde des fichiers "
        .AllowMultiSelect = False
        .InitialFileName = Environ("USERPROFILE") & "\DESKTOP\"
    End With
    
    If fdlgDossier.Show <> -1 Then
        MsgBox "S�lection du dossier annul�e par l'utilisateur.", vbInformation
        GoTo Fin
    End If
    
    dossierSauvegarde = fdlgDossier.SelectedItems(1)
    
    ' Lib�rer la variable fdlgDossier
    Set fdlgDossier = Nothing
    
    ' V�rifier que le dossier existe et est accessible
    If Dir(dossierSauvegarde, vbDirectory) = "" Then
        MsgBox "Le dossier s�lectionn� n'est pas accessible : " & dossierSauvegarde, vbCritical
        GoTo Fin
    End If
    


    ' ____________________________________ BOUCLE PRINCIPALE : TRAITER CHAQUE FICHIER _________________________________
    
    For numFichierCourant = 1 To compteurFichiers
        cheminFichier = fichiersSelectionnes(numFichierCourant)
        nomFichierCourant = Replace(Dir(cheminFichier), ".csv", "")   ' Stocker le nom du fichier
        
        ' R�initialiser le compteur pour ce fichier
        compteurLignes = 0
        contenuModifie = ""
        
        ' ---------------------------- Lecture fichier -----------------------------------------
        numFichier = FreeFile
        Open cheminFichier For Input As #numFichier
        contenu = Replace(Input$(LOF(numFichier), numFichier), vbCrLf, vbLf)
        Close #numFichier

        lignes = Split(contenu, vbCrLf)

        
        ' --------------------- Sauvegarde du fichier modifi� ----------------------
        nomFichier = Replace(Dir(cheminFichier), ".csv", "")
        cheminSortie = dossierSauvegarde & "\I3F_" & nomFichier & "_REDI_" & Format(Now, "yyyymmdd_hhmmss") & ".csv"
        
        numFichier = FreeFile
        Open cheminSortie For Output As #numFichier
        Print #numFichier, contenu  ' Pour l'instant, on sauvegarde le contenu original
        Close #numFichier
        
        ' Sauvegarder le dernier chemin de sortie pour l'ouverture finale
         dernierCheminSortie = cheminSortie

    Next numFichierCourant
    
    
    

    '------------------------------- Message de fin de traitement --------------------------
    MsgBox "Traitement termin�. " & compteurFichiers & " fichier(s) trait�(s).", vbInformation

    ' Ouvrir le dossier contenant les fichiers cr��s
    If dernierCheminSortie <> "" Then
        Shell "explorer.exe /select,""" & dernierCheminSortie & """", vbNormalFocus
    End If


Fin:

    ' ------------------------ Nettoyer la r�f�rence au dialog ------------------------------------
    Set fdlg = Nothing
    
    ' ----------------------------------- Restautrer les param�tres --------------------------------
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayAlerts = True

End Sub

