Option Explicit
' Language: VBA
' Category: Accounting & Factoring
' Description:
' Author: FLORENTIN William
' Organization: HARGOS
' Version: 1.0
' Date: 2024-040-09
' Last update: 2024-04-09

' Uncomment the function make_db() to create the database and run it once
'Public Sub make_db()
'    Dim req As String

    ' create export table
'   req = "CREATE TABLE export (id AUTOINCREMENT PRIMARY KEY, status TEXT, created_at DATETIME);"
'    insert_value ".\src\database\db.accdb", req
'    ' create export_line table
'    req = "CREATE TABLE export_line (id AUTOINCREMENT PRIMARY KEY, Raison_sociale TEXT, Code_postal TEXT, Ville TEXT, Tiers_facture TEXT, Pays TEXT, Nature TEXT, Numero TEXT, Date_piece DATETIME, Date_Eche DATETIME, Devise TEXT, Total_TTC DOUBLE, Acompte DOUBLE, Net_a_payer DOUBLE, Conditions_de_reglement TEXT, Code_SIRET TEXT NULL, Secteur_activite TEXT, Commentaire TEXT NULL, Tiers TEXT, Modifiable TEXT, status TEXT, date_trait DATETIME NULL, date_reception_factor DATE NULL, date_compta DATE NULL, export_id INTEGER, CONSTRAINT FK_export_line_export FOREIGN KEY (export_id) REFERENCES export(id) ON DELETE CASCADE);"
'    insert_value ".\src\database\db.accdb", req
'End Sub

Function construct_str_conn(db As String) As String
    construct_str_conn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & db & ";"
End Function

Function get_values_from_excel_file(file_path As String, sheet_name As String, Optional start_index As String = "A1", Optional end_index As String = "S1000") As Variant
    Dim xlApp As Object
    Dim xlBook As Object
    Dim xlSheet As Object
    Dim data_range As Object
    Dim data As Variant
    
    ' Open sheet
    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.Open(file_path)
    Set xlSheet = xlBook.Sheets(sheet_name)
    
    ' get data range
    Set data_range = xlSheet.Range(start_index & ":" & end_index)
    data = data_range.Value
    
    ' close workbook
    xlBook.Close False
    xlApp.Quit
    
    ' free memory
    Set data_range = Nothing
    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set xlApp = Nothing
    
    get_values_from_excel_file = data
End Function

Function get_value(db As String, strSQL As String) As Variant
    Dim cn As Object
    Dim rs As Object
    Dim strConn As String

    strConn = construct_str_conn(db)

    On Error GoTo ErrorHandler

    Set cn = CreateObject("ADODB.Connection")
    Set rs = CreateObject("ADODB.Recordset")

    cn.Open strConn

    rs.Open strSQL, cn

    If Not rs.EOF Then
        get_value = rs '.Fields(0).Value
    Else
        get_value = Null
    End If

    rs.Close
    cn.Close
    Set rs = Nothing
    Set cn = Nothing

    Exit Function
ErrorHandler:
    Debug.Print ("Erreur lors de la recuperation de la valeur : " & Err.Description)
    Debug.Print ("Requête : " & strSQL)
    get_value = Null
End Function

Function insert_value(db As String, strSQL As String) As Boolean
    Dim cn As Object
    Dim strConn As String

    strConn = construct_str_conn(db)

    On Error GoTo ErrorHandler

    Set cn = CreateObject("ADODB.Connection")

    cn.Open strConn

    cn.Execute strSQL

    cn.Close
    Set cn = Nothing

    insert_value = True
    Exit Function
ErrorHandler:
    Debug.Print ("Erreur lors de l'execution de la requête : " & Err.Description)
    Debug.Print ("Requête : " & strSQL)
    insert_value = False
End Function

Function regex_replace(text As String, pattern As String, replace As String) As String
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.pattern = pattern
    regex_replace = regex.replace(text, replace)
End Function

Function delete_file(path As String) As Boolean
    On Error GoTo ErrorHandler
    Kill (path)
    Debug.Print "File is deleted: " & path
    delete_file = True
    Exit Function
ErrorHandler:
    Debug.Print ("Erreur lors de la suppression du fichier : " & Err.Description)
    delete_file = False
End Function

Function search_file(path As String) As Boolean
    Dim fs As Object
    Set fs = CreateObject("Scripting.FileSystemObject")
    search_file = fs.FileExists(path)
End Function

Function insert_data(db_path As String, data As Variant) As Boolean
    Dim export_id As Long
    Dim query As String
    Dim i As Long

    On Error GoTo ErrorHandler

    insert_value db_path, "INSERT INTO export (status, created_at) VALUES (" & "'EXPORT'" & ", " & "#" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "#);"
    export_id = get_value(db_path, "SELECT MAX(id) FROM export")

    For i = LBound(data) To UBound(data)
        query = "INSERT INTO export_line (Raison_sociale, Code_postal, Ville, Tiers_facture, Pays, Nature, Numero, Date_piece, Date_Eche, Devise, Total_TTC, Acompte, Net_a_payer, Conditions_de_reglement, Code_SIRET, Secteur_activite, Commentaire, Tiers, Modifiable, status, export_id) VALUES ("
        query = query & "'" & regex_replace(CStr(data(i, 1)), "[^\w\s]", "") & "', '" & data(i, 2) & "', '" & regex_replace(CStr(data(i, 3)), "[^\w\s]", "") & "', '" & data(i, 4) & "', '" & regex_replace(CStr(data(i, 5)), "[^\w\s]", "") & "', '" & data(i, 6) & "', '" & data(i, 7) & "', #" & Format(data(i, 8), "yyyy-mm-dd") & "#, #" & Format(data(i, 9), "yyyy-mm-dd") & "#, '" & data(i, 10) & "', " & regex_replace(CStr(data(i, 11)), ",", ".") & ", " & regex_replace(CStr(data(i, 12)), ",", ".") & ", " & regex_replace(CStr(data(i, 13)), ",", ".") & ", '" & data(i, 14) & "', '" & data(i, 15) & "', '" & data(i, 16) & "', '" & data(i, 17) & "', '" & data(i, 18) & "', '" & data(i, 19) & "', "
        query = query & "'EXPORT', " & export_id & ");"

        insert_value db_path, query
    Next i
    Debug.Print "===================================="
    Debug.Print "#  Insertion des données réussie   #"
    Debug.Print "===================================="

    insert_data = True
    Exit Function
ErrorHandler:
    insert_value db_path, "DELETE FROM export WHERE id = " & export_id & ";"
    Debug.Print ("Erreur lors de l'insertion des données : " & Err.Description)
    insert_data = False
End Function

Function get_values(db As String, strSQL As String) As Variant
    Dim cn As Object
    Dim rs As Object
    Dim strConn As String
    Dim values As Collection
    Dim rowData As Variant
    Dim i As Long
    
    ' Initialisation de la connexion et du recordset
    strConn = construct_str_conn(db)
    On Error GoTo ErrorHandler
    Set cn = CreateObject("ADODB.Connection")
    Set rs = CreateObject("ADODB.Recordset")
    cn.Open strConn
    rs.Open strSQL, cn

    ' Initialisation de la collection pour stocker les données
    Set values = New Collection
    
    ' Vérification si le recordset est vide
    If Not rs.EOF Then
        ' Parcours des enregistrements
        Do While Not rs.EOF
            ' Initialisation du tableau pour stocker les valeurs de l'enregistrement actuel
            ReDim rowData(1 To rs.Fields.Count)
            ' Stockage des valeurs de l'enregistrement actuel dans le tableau
            For i = 1 To rs.Fields.Count
                ' Vérification si la valeur est Null avant de l'affecter au tableau
                If Not IsNull(rs.Fields(i - 1).Value) Then
                    rowData(i) = rs.Fields(i - 1).Value
                Else
                    rowData(i) = ""
                End If
            Next i
            ' Ajout du tableau de données à la collection
            values.Add rowData
            rs.MoveNext
        Loop
    End If

    ' Fermeture du recordset et de la connexion
    rs.Close
    cn.Close
    Set rs = Nothing
    Set cn = Nothing

    ' Conversion de la collection en tableau à deux dimensions
    get_values = CollectionToArray(values)
    Exit Function

ErrorHandler:
    ' Gestion des erreurs
    Debug.Print ("Erreur lors de la récupération des valeurs : " & Err.Description)
    Debug.Print ("Requête : " & strSQL)
    ' Retourne une valeur Null en cas d'erreur
    get_values = Null
End Function

Function CollectionToArray(col As Collection) As Variant
    Dim arr() As Variant
    Dim i As Long, j As Long
    Dim item As Variant
    
    ' Initialisation du tableau
    ReDim arr(1 To col.Count, 1 To col(1).Count)
    
    ' Copie des données de la collection dans le tableau
    For i = 1 To col.Count
        For j = 1 To col(i).Count
            arr(i, j) = col(i)(j)
        Next j
    Next i
    
    ' Retourne le tableau
    CollectionToArray = arr
End Function

    
Public Sub main()

    Dim folder_path As String
    Dim sheet_name As String
    Dim data As Variant
    Dim end_line As Long
    Dim error_message As String
    Dim res As Boolean

    ' folder path
    folder_path = ThisWorkbook.path & "\"
    sheet_name = "CONVERT"
    end_line = 175
    
    ' get data from excel file
    data = get_values_from_excel_file(folder_path & "FAC.xlsx", sheet_name, "A2", "S" & end_line)
    On Error GoTo ErrorHandler
        ' insert data into access database
        res = insert_data(folder_path & "src\database\db.accdb", data)
        If res = False Then
            error_message = "Erreur lors de l'insertion des données"
            Exit Sub
        End If
        ' Delete export file
        delete_file folder_path & "FAC.xlsx"
        If search_file(folder_path & "FAC.xlsx") = True Then
            error_message = "Erreur lors de la suppression du fichier"
            Exit Sub
        End If

        Dim remise As Long ' change datatype to match the data type in the database (its possible that the data type is not long (possible string))
        Dim convert As Variant
        Dim ws As Worksheet

        remise = get_value(folder_path & "src\database\db.accdb", "SELECT MAX(id) FROM export")
        convert = get_value(folder_path & "src\database\db.accdb", "SELECT Raison_sociale, Code_postal, Ville, Tiers_facture, Pays, Nature, Numero, Date_piece, Date_Eche, Devise, Total_TTC, Acompte, Net_a_payer, Conditions_de_reglement, Code_SIRET, Secteur_activite, Commentaire, Tiers, Modifiable FROM export_line WHERE export_id = " & remise & ";")
        Set ws = ThisWorkbook.Sheets("test")
        
        'ws.Range("A2").CopyFromRecordset convert
        
        'Debug.Print convert
        'ws.Cells(2, 1).CopyFromRecordset convert

        MsgBox "Les données ont été importées avec succès", vbInformation
    Exit Sub
ErrorHandler:
    MsgBox "Une erreur s'est produite. Arrêt de l'exécution du programme: " & error_message, vbCritical
End Sub


Sub ExportDataWithCondition()

    Dim ws As Worksheet
    Dim conn As Object
    Dim rs As Object
    Dim strSQL As String
    Dim i As Integer
    Dim remise As Integer
    
    
    
    ' Spécifie la feuille de calcul dans laquelle les données seront collées
    Set ws = ThisWorkbook.Sheets("Feuil3") ' Remplacez "Nom de la feuille" par le nom de votre feuille
    
    ' Connexion à la base de données Access
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\taf\project-factory\src\database\db.accdb;"
    
    ' Requête SQL avec condition WHERE
    strSQL = "SELECT Raison_sociale, Code_postal, Ville, Tiers_facture, Pays, Nature, Numero, Date_piece, Date_Eche, Devise, Total_TTC, Acompte, Net_a_payer, Conditions_de_reglement, Code_SIRET, Secteur_activite, Commentaire, Tiers, Modifiable " & _
             "FROM export_line " & _
             "WHERE export_id = 25;" ' Remplacez "VotreCondition" et "VotreValeur" par votre condition
    
    ' Exécute la requête SQL
    Set rs = conn.Execute(strSQL)
    
    ws.Range("A2").CopyFromRecordset rs ' Données
    
    ' Ferme la connexion et libère la mémoire
    rs.Close
    conn.Close
    Set rs = Nothing
    Set conn = Nothing

End Sub

Sub ImporterCSVSansEntetes()
    Dim CheminFichier As Variant
    Dim FeuilleActuelle As Worksheet
    Dim DerniereLigne As Long

    ' Demander à l'utilisateur de sélectionner un fichier CSV
    CheminFichier = Application.GetOpenFilename("Fichiers CSV (*.csv), *.csv", , "Sélectionner un fichier CSV")

    ' Vérifier si un fichier a été sélectionné
    If CheminFichier <> False Then
        ' Enregistrer la feuille de calcul active
        Set FeuilleActuelle = ThisWorkbook.Sheets("Feuil1")
        
        ' Déterminer la prochaine ligne vide pour commencer à écrire les données
        DerniereLigne = FeuilleActuelle.Cells(FeuilleActuelle.Rows.Count, 1).End(xlUp).Row + 1
        
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
        DerniereLigne = FeuilleActuelle.Cells(FeuilleActuelle.Rows.Count, 1).End(xlUp).Row
        
        ' Supprimer les lignes vides à la fin des données importées
        If DerniereLigne > 1 Then
            FeuilleActuelle.Rows(DerniereLigne + 1 & ":" & FeuilleActuelle.Rows.Count).Delete
        End If
    End If
End Sub
    

    Worksheets("Sheet1").Range("A1:G37").Locked = False
    Worksheets("Sheet1").Protect
    


Sub RedimensionnerTableauCroiseDynamique()
    Dim Feuille As Worksheet
    Dim TableauCroise As PivotTable
    Dim PlageDonnees As Range
    Dim NouvellePlage As Range
    
    ' Définir la feuille de calcul contenant le tableau croisé dynamique
    Set Feuille = ThisWorkbook.Sheets("Feuil1") ' Remplacez "Feuil1" par le nom de votre feuille
    
    ' Définir le tableau croisé dynamique
    Set TableauCroise = Feuille.PivotTables("test") ' Remplacez "TableauCroisé1" par le nom de votre tableau croisé dynamique
    
    ' Définir la nouvelle plage de données
    Set PlageDonnees = Feuille.Range("A1:U100")
    
    ' Redimensionner le tableau croisé dynamique
    TableauCroise.ChangePivotCache ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=PlageDonnees, _
        Version:=xlPivotTableVersion15)
End Sub