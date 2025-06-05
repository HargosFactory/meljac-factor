Sub ImporterFichierCSV()
    Dim ChoisirFichier As FileDialog
    Dim CheminFichier As String
    Dim FeuilleImport As Worksheet
    Dim DerniereLigne As Long

    ' Ouvrir la boîte de dialogue pour choisir le fichier à importer
    Set ChoisirFichier = Application.FileDialog(msoFileDialogFilePicker)
    With ChoisirFichier
        .Title = "Sélectionner le fichier CSV à importer"
        .Filters.Clear
        .Filters.Add "Fichiers CSV", "*.csv"
        .AllowMultiSelect = False
        If .Show = -1 Then
            CheminFichier = .SelectedItems(1)
        Else
            Exit Sub ' Annuler si aucun fichier n'a été sélectionné
        End If
    End With

    ' Déterminer la feuille de calcul "Import" dans le classeur actif
    Set FeuilleImport = ThisWorkbook.Sheets("Import")
    
    ' Copier les entêtes de colonnes de la feuille import
    FeuilleImport.Range("A1:T1").Cut FeuilleImport.Range("W1:AP1")

    ' Trouver la dernière ligne dans la feuille "Import"
    DerniereLigne = FeuilleImport.Cells(FeuilleImport.Rows.Count, 1).End(xlUp).row + 1 ' Dernière ligne + 1 pour ajouter les données sous la dernière ligne existante

    ' Importer les données du fichier CSV
    With FeuilleImport.QueryTables.Add(Connection:="TEXT;" & CheminFichier, Destination:=FeuilleImport.Cells(DerniereLigne, 1))
        .TextFileParseType = xlDelimited
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = False
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = False
        .TextFileOtherDelimiter = ";"
        .TextFileColumnDataTypes = Array(1, 2, 2, 2, 1, 2, 2, 2, 1, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2) ' Types de données des colonnes
        .TextFileStartRow = 2 ' Ligne de départ des données (après les entêtes)
        .Refresh BackgroundQuery:=True
    End With

    ' Supprimer les colonnes indésirables
     With FeuilleImport
         .Columns("V:V").Delete ' Supprimer La colonne V (Clearing Journal Entry)
         .Columns("U:U").Delete ' Supprimer la colonne U (Due Net (Symbol))
         .Columns("T:T").Delete ' Supprimer la colonne T (Special G/L)
         .Columns("S:S").Delete ' Supprimer la colonne S (Clearing Status)
         .Columns("A:A").Delete ' Supprimer la colonne A (Company Code)
     End With
     
    ' Coller les entêtes de colonnes sur la première ligne
    FeuilleImport.Range("R1:AK1").Cut FeuilleImport.Range("A1:T1")

    ' Réajuster les colonnes restantes
    FeuilleImport.Columns.AutoFit
    
    ' Désactiver l'actualisation automatique des données externes
    With FeuilleImport.QueryTables(1)
        .EnableRefresh = False
    End With

    MsgBox "Les données ont été importées avec succès dans la feuille ""Import""."
End Sub

Sub AjouterFormuleColonneST()
    Dim Feuille As Worksheet
    Dim DerniereLigne As Long
    Dim i As Long
    
    ' Définir la feuille de calcul à utiliser
    Set Feuille = ThisWorkbook.Sheets("Import")
    
    ' Trouver la dernière ligne avec des données dans la colonne P
    DerniereLigne = Feuille.Cells(Feuille.Rows.Count, 16).End(xlUp).row
    
    ' Parcourir chaque ligne avec des données à partir de la deuxième ligne
    For i = 2 To DerniereLigne
        ' Vérifier si la cellule de la colonne P est non vide
        If Not IsEmpty(Feuille.Cells(i, 16).Value) Then
            ' Ajouter la formule dans la cellule de la colonne S
            Feuille.Cells(i, 19).FormulaLocal = "=SI(E" & i & "=""FR"";""DOM"";""EXP"")"
            ' Ajouter status pour selection colonne (T)
            Feuille.Cells(i, 20).Value = 1
        End If
    Next i
End Sub

Sub SupprimerLignesColonnesAT()
    Dim Feuille As Worksheet
    Dim DerniereLigne As Long
    
    ' Définir la feuille de calcul à utiliser
    Set Feuille = ThisWorkbook.Sheets("Import")
    
    ' Si A2 est vide, alors ne rien faire
    If Feuille.Range("A2").Value = "" Then
        Exit Sub
    End If
    
    ' Trouver la dernière ligne avec des données dans la première colonne
    DerniereLigne = Feuille.Cells(Feuille.Rows.Count, 1).End(xlUp).row
        
    ' Supprimer les lignes des colonnes A à T
    Feuille.Range("A2:T" & DerniereLigne).EntireRow.Delete
End Sub

Sub SupprimerLignesColonnesHRemise()
    
    Dim Feuille As Worksheet
    Dim DerniereLigne As Long
    
    ' Définir la feuille de calcul à utiliser
    Set Feuille = ThisWorkbook.Sheets("REMISE DOMESTIQUE")
    
    ' Si H13 est vide, alors ne rien faire
    If Feuille.Range("H13").Value = "" Then
        Exit Sub
    End If
    
    ' Trouver la dernière ligne avec des données dans la premiere colonne
    DerniereLigne = Feuille.Cells(Feuille.Rows.Count, 1).End(xlUp).row
    Feuille.Range("H13:H" & DerniereLigne).Delete
End Sub

Sub AjouterCheckBoxesDynamiques()
    Dim Feuille As Worksheet
    Dim DerniereLigne As Long
    Dim Cellule As Range
    Dim CheckBox As CheckBox
    Dim i As Long
    
    ' Définir la feuille de calcul à utiliser
    Set Feuille = ThisWorkbook.Sheets("Import")
    
    ' Trouver la dernière ligne avec des données dans la première colonne
    DerniereLigne = Feuille.Cells(Feuille.Rows.Count, 1).End(xlUp).row + 1
    
    ' Parcourir chaque ligne avec des données
    For i = 2 To DerniereLigne
        ' Ajouter une case à cocher dans la colonne U
        Set CheckBox = Feuille.CheckBoxes.Add(Feuille.Cells(i, 21).Left, Feuille.Cells(i, 20).Top, 30, 6)
        ' retirer texte
        CheckBox.Caption = ""
        ' Cocher la case à cocher par défaut
        CheckBox.Value = True
        ' Assigner la macro de gestion des clics à la case à cocher
        AssignerMacro CheckBox, "GestionCheckBox"
    Next i
End Sub

Sub AssignerMacro(ctrl As Object, nomMacro As String)
    Dim i As Long

    ' Obtenir le numéro de contrôle
    i = Right(ctrl.Name, Len(ctrl.Name) - 10)
    
    ' Assigner la macro à la case à cocher
    ctrl.OnAction = "'" & ThisWorkbook.Name & "'!GestionCheckBox"
End Sub

Sub RedimensionnerTableauCroiseDynamiqueDomestique()
    Dim FeuilleSource As Worksheet
    Dim FeuilleTCD As Worksheet
    Dim TableauCroise As PivotTable
    Dim CacheTCD As PivotCache
    Dim PlageDonnees As Range
    Dim TailleImport As Long
    
    ' Définir la feuille de calcul source contenant la nouvelle plage de données
    Set FeuilleSource = ThisWorkbook.Sheets("Import")
    
    ' Définir la feuille de calcul contenant le tableau croisé dynamique
    Set FeuilleTCD = ThisWorkbook.Sheets("TCDDOM")
    
    ' Définir le tableau croisé dynamique
    Set TableauCroise = FeuilleTCD.PivotTables("tcddomestique")
    
    ' Définir taille de l'import (10 = col J)
    TailleImport = FeuilleSource.Cells(FeuilleSource.Rows.Count, "J").End(xlUp).row
    
    ' Définir la nouvelle plage de données
    Set PlageDonnees = FeuilleSource.Range("A1:T" & TailleImport)
    
    ' Créer un nouveau cache de tableau croisé dynamique avec la nouvelle plage de données
    Set CacheTCD = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=PlageDonnees)
    
    ' Assigner le nouveau cache de tableau croisé dynamique au tableau croisé dynamique existant
    TableauCroise.ChangePivotCache CacheTCD
End Sub

Sub ModifierCellulesColonneL()
    Dim FeuilleSource As Worksheet
    Dim Cellule As Range
    Dim DerniereLigne As Long
    Dim i As Long
    
    ' Spécifie la feuille source
    Set FeuilleSource = ThisWorkbook.Sheets("Import")
    
    ' Trouve la dernière ligne de la colonne L
    DerniereLigne = FeuilleSource.Cells(Rows.Count, "L").End(xlUp).row
    
    ' Parcourt toutes les cellules de la colonne L
    For i = 2 To DerniereLigne
        Set Cellule = FeuilleSource.Cells(i, "L")
        
        ' Récupère le contenu de la cellule
        Contenu = Cellule.Value
        
        ' Supprime les espaces et le texte "EUR"
        NouveauContenu = replace(Contenu, " EUR", "")
        NouveauContenu = replace(NouveauContenu, " ", "")
        
        ' Remplace la virgule par un point décimal
        'NouveauContenu = replace(NouveauContenu, ",", ".")
        
        ' Met à jour le contenu de la cellule
        Cellule.Value = NouveauContenu
        
        ' Applique le format personnalisé
        Cellule.NumberFormat = "General"
    Next i
End Sub

Sub SupprimerCheckBoxesFeuille(ByRef Feuille As Worksheet)
    Dim cb As CheckBox
    
    ' Parcourir toutes les cases à cocher de la feuille spécifiée et les supprimer
    For Each cb In Feuille.CheckBoxes
        cb.Delete
    Next cb
End Sub

Sub SupprimerToutesLesCheckBoxes()
    Dim Feuille As Worksheet
    Set Feuille = ThisWorkbook.Sheets("Import")
    SupprimerCheckBoxesFeuille Feuille
End Sub

Sub import_button_sap()
    ImporterFichierCSV
    AjouterFormuleColonneST
    ModifierCellulesColonneL
    'AjouterCheckBoxesDynamiques
    'RedimensionnerTableauCroiseDynamiqueDomestique
End Sub

Sub reset_import_button_sap()
    SupprimerToutesLesCheckBoxes
    SupprimerLignesColonnesAT
    SupprimerLignesColonnesHRemise
End Sub

Sub reload_button_domestique()
    RedimensionnerTableauCroiseDynamiqueDomestique
End Sub

