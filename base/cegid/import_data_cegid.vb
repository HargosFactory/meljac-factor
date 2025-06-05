' Language: VBA
' Category: Accounting & Factoring
' Description:
' Author: FLORENTIN William
' Organization: HARGOS
' Version: 1.0
' Date: 2024-040-08
' Last update: 2024-04-12

' Module 1 : Gestion des cases à cocher
Sub GestionCheckBox()
    Dim CheckBox As CheckBox
    Dim Ligne As Long
    
    ' Récupérer la ligne où se trouve la case à cocher
    Ligne = ActiveSheet.CheckBoxes(Application.Caller).TopLeftCell.row
    
    ' Modifier la valeur de la cellule dans la deuxième colonne (colonne B) en fonction de l'état de la case à cocher
    ActiveSheet.Cells(Ligne, 21).value = IIf(ActiveSheet.CheckBoxes(Application.Caller).value = xlOn, 1, 0)
End Sub

' Module 2 : Création des cases à cocher
Sub SupprimerCheckBoxesFeuille(ByRef Feuille As Worksheet)
    Dim cb As CheckBox
    
    ' Parcourir toutes les cases à cocher de la feuille spécifiée et les supprimer
    For Each cb In Feuille.CheckBoxes
        cb.Delete
    Next cb
End Sub

' Module 3 : Ajout des cases à cocher
Sub AjouterCheckBoxesDynamiques()
    Dim Feuille As Worksheet
    Dim DerniereLigne As Long
    Dim Cellule As Range
    Dim CheckBox As CheckBox
    Dim i As Long
    
    ' Définir la feuille de calcul à utiliser
    Set Feuille = ThisWorkbook.Sheets("CONVERT")
    
    ' Trouver la dernière ligne avec des données dans la première colonne
    DerniereLigne = Feuille.Cells(Feuille.Rows.Count, 1).End(xlUp).row + 1
    
    ' Parcourir chaque ligne avec des données
    For i = 2 To DerniereLigne
        ' Ajouter une case à cocher dans la troisième colonne (colonne C)
        Set CheckBox = Feuille.CheckBoxes.Add(Feuille.Cells(i, 22).Left, Feuille.Cells(i, 22).Top, 30, 6)
        ' retirer texte
        CheckBox.Caption = ""
        ' Cocher la case à cocher par défaut
        CheckBox.value = True
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

Sub SupprimerToutesLesCheckBoxes()
    Dim Feuille As Worksheet
    Set Feuille = ThisWorkbook.Sheets("CONVERT")
    SupprimerCheckBoxesFeuille Feuille
End Sub

' Module 4 : Importer des données à partir d'un fichier CSV sans en-têtes
Sub ImporterCSVSansEntetes()
    Dim CheminFichier As Variant
    Dim FeuilleActuelle As Worksheet
    Dim DerniereLigne As Long

    ' Demander à l'utilisateur de sélectionner un fichier CSV
    CheminFichier = Application.GetOpenFilename("Fichiers CSV (*.csv), *.csv", , "Sélectionner un fichier CSV")

    ' Vérifier si un fichier a été sélectionné
    If CheminFichier <> False Then
        ' Enregistrer la feuille de calcul active
        Set FeuilleActuelle = ThisWorkbook.Sheets("CONVERT")
        
        ' Déterminer la prochaine ligne vide pour commencer à écrire les données
        DerniereLigne = FeuilleActuelle.Cells(FeuilleActuelle.Rows.Count, 1).End(xlUp).row + 1
        
        ' Ouvrir le fichier CSV et importer les données dans la feuille de calcul à partir de la deuxième ligne
        With FeuilleActuelle.QueryTables.Add(Connection:="TEXT;" & CheminFichier, Destination:=FeuilleActuelle.Cells(DerniereLigne, 1))
            .TextFileParseType = xlDelimited
            .TextFileConsecutiveDelimiter = False
            .TextFileTabDelimiter = False
            .TextFileSemicolonDelimiter = False
            .TextFileCommaDelimiter = False
            .TextFileOtherDelimiter = ";"
            .TextFileStartRow = 2 ' Commencer l'importation à partir de la deuxième ligne du fichier CSV
            .TextFileTextQualifier = xlTextQualifierDoubleQuote
            .TextFilePlatform = xlWindows
            .Refresh BackgroundQuery:=False
        End With
        
        ' Enregistrer le numéro de la dernière ligne avec des données
        DerniereLigne = FeuilleActuelle.Cells(FeuilleActuelle.Rows.Count, 1).End(xlUp).row
        
        ' Supprimer les lignes vides à la fin des données importées
        If DerniereLigne > 1 Then
            FeuilleActuelle.Rows(DerniereLigne + 1 & ":" & FeuilleActuelle.Rows.Count).Delete
        End If
    End If
End Sub

' Module 5 : Supprimer les lignes des colonnes A à U
Sub SupprimerLignesColonnesAU()
    Dim Feuille As Worksheet
    Dim DerniereLigne As Long
    
    ' Définir la feuille de calcul à utiliser
    Set Feuille = ThisWorkbook.Sheets("CONVERT")
    
    ' Trouver la dernière ligne avec des données dans la première colonne
    DerniereLigne = Feuille.Cells(Feuille.Rows.Count, 1).End(xlUp).row
    
    ' Supprimer les lignes des colonnes A à T
    Feuille.Range("A2:U" & DerniereLigne).EntireRow.Delete
End Sub

' Module 6 : Ajouter une formule dans la colonne T en fonction des valeurs des colonnes S et Q
Sub AjouterFormuleColonneT()
    Dim Feuille As Worksheet
    Dim DerniereLigne As Long
    Dim i As Long
    
    ' Définir la feuille de calcul à utiliser
    Set Feuille = ThisWorkbook.Sheets("CONVERT")
    
    ' Trouver la dernière ligne avec des données dans la colonne S
    DerniereLigne = Feuille.Cells(Feuille.Rows.Count, 19).End(xlUp).row ' Colonne S = colonne 19
    
    ' Parcourir chaque ligne avec des données à partir de la deuxième ligne
    For i = 2 To DerniereLigne
        ' Vérifier si la cellule de la colonne S est non vide
        If Not IsEmpty(Feuille.Cells(i, 19).value) Then
            ' Ajouter la formule dans la cellule de la colonne T
            Feuille.Cells(i, 20).FormulaLocal = "=SI(E" & i & "=""FRANCE""; SI(Q" & i & "=""""; ""NAN""; ""DOM""); SI(Q" & i & "=""""; ""NAN""; ""EXP""))"
            Feuille.Cells(i, 21).value = 1
        End If
    Next i
End Sub

' Module 7 : Créer un tableau croisé dynamique à partir des données
Sub RedimensionnerTableauCroiseDynamique()
    Dim FeuilleSource As Worksheet
    Dim FeuilleTCD As Worksheet
    Dim TableauCroise As PivotTable
    Dim CacheTCD As PivotCache
    Dim PlageDonnees As Range
    
    ' Définir la feuille de calcul source contenant la nouvelle plage de données
    Set FeuilleSource = ThisWorkbook.Sheets("CONVERT")
    
    ' Définir la feuille de calcul contenant le tableau croisé dynamique
    Set FeuilleTCD = ThisWorkbook.Sheets("TCD DOM")
    
    ' Définir le tableau croisé dynamique
    Set TableauCroise = FeuilleTCD.PivotTables("tcdDom")
    
    ' Définir la nouvelle plage de données
    Set PlageDonnees = FeuilleSource.Range("A1:U175")
    
    ' Créer un nouveau cache de tableau croisé dynamique avec la nouvelle plage de données
    Set CacheTCD = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=PlageDonnees)
    
    ' Assigner le nouveau cache de tableau croisé dynamique au tableau croisé dynamique existant
    TableauCroise.ChangePivotCache CacheTCD
End Sub

' Module 8 : Exécuter toutes les étapes
Sub test_button()
    ImporterCSVSansEntetes
    AjouterFormuleColonneT
    AjouterCheckBoxesDynamiques
    RedimensionnerTableauCroiseDynamique
End Sub

' Module 9 : Réinitialiser l'importation
Sub reset_import()
    SupprimerToutesLesCheckBoxes
    SupprimerLignesColonnesAU
End Sub

' Module 10 : Rafraîchir le tableau croisé dynamique
Sub reload_tcd()
    RedimensionnerTableauCroiseDynamique
End Sub
    