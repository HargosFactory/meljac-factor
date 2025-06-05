Option Explicit
' Language: VBA
' Category: Accounting & Factoring
' Description:
' Author: FLORENTIN William
' Organization: HARGOS
' Version: 1.0
' Date: 2024-040-08
' Last update: 2024-04-12

Function define_end_sheet(ws As Worksheet, column As Range) As Integer
    define_end_sheet = ws.Cells(ws.Rows.Count, column.column).End(xlUp).Row
End Function

Function write_in_txt_file(path As String, value As String) As Boolean
    Dim file As Integer
    file = FreeFile
    Open path For Append As file
    Print #file, value
    Close file
    write_in_txt_file = True
End Function

Function create_file(path As String) As Boolean
    Dim fs As Object
    Dim a As Object
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set a = fs.CreateTextFile(path, True)
    a.Close
    create_file = True
End Function

Function search_end_point_in_column(ws As Worksheet, column As Range, value As String) As Long
    Dim i As Long
    Dim end_ws As Long
    end_ws = define_end_sheet(ws, column)
    For i = 1 To end_ws
        If ws.Cells(i, column.column).value = value Then
            search_end_point_in_column = i
            Exit For
        End If
    Next i
End Function

Function extract_variant(ws As Worksheet, ByVal start As String, ByVal halt As String) As Variant
    Dim start_range As Range
    Dim end_range As Range
    Dim dataRange As Range
    Dim dataArray As Variant
    Dim i As Long, j As Long
    
    Set start_range = ws.Range(start)
    Set end_range = ws.Range(halt)
    
    If start_range Is Nothing Or end_range Is Nothing Then
        extract_variant = "Error selecting range"
        Exit Function
    End If
    
    Set dataRange = ws.Range(start_range, end_range)
    dataArray = dataRange.value
    
    '' Delete the first row of the variant
    Dim rowsCount As Long, colsCount As Long
    rowsCount = UBound(dataArray, 1) - LBound(dataArray, 1) + 1
    colsCount = UBound(dataArray, 2) - LBound(dataArray, 2) + 1
    
    Dim resultArray() As Variant
    ReDim resultArray(1 To rowsCount, 1 To colsCount)
    
    For i = 1 To rowsCount
        For j = 1 To colsCount
            resultArray(i, j) = dataArray(i, j)
        Next j
    Next i
    
    extract_variant = resultArray
End Function

Function delete_file(path As String) As Boolean
    Kill (path)
    delete_file = True
End Function

Function search_file(path As String) As Boolean
    Dim fs As Object
    Set fs = CreateObject("Scripting.FileSystemObject")
    search_file = fs.FileExists(path)
End Function

Function regex_replace(text As String, pattern As String, replace As String) As String
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.pattern = pattern
    regex_replace = regex.replace(text, replace)
End Function

Function format_reference(ByVal reference As Variant) As String
    Dim formattedReference As String
    Dim charToAdd As Integer
    Dim ref As String
    ref = CStr(reference)

    If regex(ref, "A") = False Then
        ' Ajouter des zéros à gauche pour obtenir une référence de 14 caractères
        charToAdd = 14 - Len(reference)
        formattedReference = String(charToAdd, "0") & ref
    Else
        ' Ajouter des espace à droite pour obtenir une référence de 14 caractères
        charToAdd = 14 - Len(reference)
        formattedReference = reference & String(charToAdd, " ")
    End If
    
    format_reference = formattedReference
End Function

Function format_valeur(ByVal valeur As Double) As String
    Dim formattedValue As String
    
    If valeur < 0 Then
        valeur = Abs(valeur)
    End If
    formattedValue = Format(valeur, "00000000000000")
    
    format_valeur = formattedValue
End Function

Function regex(text As String, pattern As String) As Boolean
    Dim regexObj As Object
    Set regexObj = CreateObject("VBScript.RegExp")
    regexObj.Global = True
    regexObj.pattern = pattern
    regex = regexObj.test(text)
End Function
    

Function GetType(ByVal variable As Variant) As String
    If IsEmpty(variable) Then
        GetType = "Empty"
    ElseIf IsNull(variable) Then
        GetType = "Null"
    ElseIf IsNumeric(variable) Then
        GetType = "Numeric"
    ElseIf IsDate(variable) Then
        GetType = "Date"
    ElseIf IsObject(variable) Then
        GetType = "Object"
    ElseIf IsArray(variable) Then
        GetType = "Array"
    ElseIf TypeName(variable) = "String" Then
        GetType = "String"
    ElseIf TypeName(variable) = "Boolean" Then
        GetType = "Boolean"
    Else
        GetType = "Unknown"
    End If
End Function

Function get_line_nb(ws As Worksheet, col As Integer, val As Variant) As Long
    Dim rng As Range
    Dim line As Long
    
    Set rng = ws.Columns(col).Find(What:=val, LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not rng Is Nothing Then
        line = rng.Row
    Else
        line = 0
    End If
    
    get_line_nb = line
End Function

Function get_value_of_cell(ws As Worksheet, row As Integer, column As Integer) As Variant
    get_value_of_cell = ws.Cells(row, column).value
End Function

Function facture_or_not(text As String) As String
    If InStr(text, "Facture") > 0 Then
        facture_or_not = "F"
    Else
        facture_or_not = "A"
    End If
End Function

Function format_tiers(ByVal tiers As String) As String
    Dim formattedTiers As String
    Dim charToAdd As Integer

    charToAdd = 15 - Len(tiers)
    formattedTiers = tiers & String(charToAdd, "0")

    format_tiers = formattedTiers
End Function

Public Sub main_domestique()
    Dim ws As Worksheet
    Dim convert As Worksheet
    Dim end_ws As Integer
    Dim remise As String
    Dim remise_date As String
    Dim client As String
    Dim devise As String
    Dim file_path As String
    Dim data As Variant
    Dim start_index As String
    
    ' select sheet
    Set ws = Worksheets("DECLARATIF DOMESTIQUE")
    Set convert = Worksheets("CONVERT")
    ws.Activate
    ' fixed value
    start_index = "A14"
    remise = ws.Range("B7").value
    remise_date = ws.Range("B8").value
    client = Mid(ws.Range("D6").value, 9)
    devise = ws.Range("B9").value
    file_path = ws.Range("B5").value & "\" & Mid(remise, 1, 8) & ".txt"
    ' transform data selection to send factor in variant
    end_ws = search_end_point_in_column(ws, ws.Range(start_index), "Total général")
    data = extract_variant(ws, start_index, "F" & (end_ws - 1))
    ' create and write in txt file
    If search_file(file_path) = False Then
        create_file (file_path)
    Else
        delete_file (file_path)
        create_file (file_path)
    End If
    
    Dim i As Long
    For i = LBound(data, 1) To UBound(data, 1)
        ' Debug.Print
        write_in_txt_file file_path, Mid(remise, 4, 5) & ";" & client & ";" & regex_replace(remise_date, "/", "") & ";" & "D" & ";" & "511" & ";" & facture_or_not(get_value_of_cell(convert, get_line_nb(convert, 7, data(i, 3)), 6)) & ";" & devise & ";" & regex_replace(CStr(data(i, 1)), "[^0-9]", "") & ";" & format_tiers(get_value_of_cell(convert, get_line_nb(convert, 7, data(i, 3)), 4)) & ";" & "                       ;" & format_reference(data(i, 3)) & ";" & format_valeur(data(i, 4)) & ";" & regex_replace(CStr(data(i, 2)), "/", "") & ";" & regex_replace(CStr(data(i, 5)), "/", "") & ";" & CStr(data(i, 6)) & ";0123456789;                         ;              ;                                                   ;   ;"
    Next i

    MsgBox "Fichier " & Mid(remise, 1, 8) & ".txt créé avec succès !"
    
End Sub


    =SI(SOMME(SI('Import'!$L$2:$L$175=1;SI('Import'!$S$2:$S$175="DOM";'Import'!$T$2:$T$175;0);0);0)<>B11;"ERREUR";"OK")

    Debug.Print Mid(remise, 4, 5)
    Debug.Print client
    Debug.Print regex_replace(remise_date, "/", "")
    Debug.Print "D"
    Debug.Print "511"
    Debug.Print facture_or_not(get_value_of_cell(convert, get_line_nb(convert, 7, data(i, 3)), 6))
    Debug.Print devise
    Debug.Print regex_replace(CStr(data(i, 1)), "[^0-9]", "")
    Debug.Print format_tiers(get_value_of_cell(convert, get_line_nb(convert, 7, data(i, 3)), 4))
    Debug.Print " "
    Debug.Print format_reference(data(i, 3))
    Debug.Print format_valeur(data(i, 4))
    Debug.Print regex_replace(CStr(data(i, 2)), "/", "")
    Debug.Print regex_replace(CStr(data(i, 5)), "/", "")
    Debug.Print CStr(data(i, 6))