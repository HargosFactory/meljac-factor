Sub GestionCheckBox()
    Dim CheckBox As CheckBox
    Dim Ligne As Long
    
    ' Récupérer la ligne où se trouve la case à cocher
    Ligne = ActiveSheet.CheckBoxes(Application.Caller).TopLeftCell.row
    
    ' Modifier la valeur de la cellule dans la deuxième colonne (colonne B) en fonction de l'état de la case à cocher
    ActiveSheet.Cells(Ligne, 19).Value = IIf(ActiveSheet.CheckBoxes(Application.Caller).Value = xlOn, 1, 0)
End Sub