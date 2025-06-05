Option Explicit
' Language: VBA
' Category: Accounting & Factoring
' Description:
' Author: FLORENTIN William
' Organization: HARGOS
' Version: 1.0
' Date: 2024-04-08
' Last update: 2024-04-18

' #################################################################################################################
' #                                                   FUNCTIONS                                                   #
' #################################################################################################################

' Description: This function defines the last row of a worksheet.
' Parameters:
'   - ws: The worksheet to analyze.
'   - column: The column to analyze.
' Returns: The last row of the worksheet.
Function define_end_sheet(ws As Worksheet, column As Range) As Integer
    define_end_sheet = ws.Cells(ws.Rows.Count, column.column).End(xlUp).row
End Function

' Description: This function writes a value in a text file.
' Parameters:
'   - path: The path of the file.
'   - value: The value to write.
' Returns: True if the value was written successfully, False otherwise.
Function write_in_txt_file(path As String, value As String) As Boolean
    Dim file As Integer
    file = FreeFile
    Open path For Append As file
    Print #file, value
    Close file
    write_in_txt_file = True
End Function

' Description: This function creates a text file.
' Parameters:
'   - path: The path of the file.
' Returns: True if the file was created successfully, False otherwise.
Function create_file(path As String) As Boolean
    Dim fs As Object
    Dim a As Object
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set a = fs.CreateTextFile(path, True)
    a.Close
    create_file = True
End Function

' Description: This function searches for a value in a column.
' Parameters:
'   - ws: The worksheet to analyze.
'   - column: The column to analyze.
'   - Value: The value to search for.
' Returns: The row number of the value.
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

' Description: This function extracts a range of data from a worksheet.
' Parameters:
'   - ws: The worksheet to analyze.
'   - start: The starting cell of the range.
'   - halt: The ending cell of the range.
' Returns: The extracted data as a variant.
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

' Description: This function deletes a file.
' Parameters:
'   - path: The path of the file.
' Returns: True if the file was deleted successfully, False otherwise.
Function delete_file(path As String) As Boolean
    Kill (path)
    delete_file = True
End Function

' Description: This function searches for a file.
' Parameters:
'   - path: The path of the file.
' Returns: True if the file exists, False otherwise.
Function search_file(path As String) As Boolean
    Dim fs As Object
    Set fs = CreateObject("Scripting.FileSystemObject")
    search_file = fs.FileExists(path)
End Function

' Description: This function replaces a pattern in a text with a replacement.
' Parameters:
'   - text: The text to search.
'   - pattern: The pattern to replace.
'   - replace: The replacement text.
' Returns: The modified text.
Function regex_replace(text As String, pattern As String, replace As String) As String
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.pattern = pattern
    regex_replace = regex.replace(text, replace)
End Function

' Description: This function formats a reference.
' Parameters:
'   - reference: The reference to format.
Function format_reference(ByVal reference As Variant) As String
    Dim formattedReference As String
    Dim charToAdd As Integer
        ' Ajouter des espace à droite pour obtenir une référence de 14 caractères
        charToAdd = 14 - Len(reference)
        formattedReference = reference & String(charToAdd, " ")
    
    format_reference = formattedReference
End Function

' Description: This function formats a value.
' Parameters:
'   - valeur: The value to format.
' Returns: The formatted value.
Function format_valeur(ByVal valeur As Double) As String
    Dim formattedValue As String
    
    If valeur < 0 Then
        valeur = Abs(valeur)
    End If
    formattedValue = Format(valeur, "00000000000000")
    
    format_valeur = formattedValue
End Function

' Description: This function checks if a pattern is present in a text.
' Parameters:
'   - text: The text to search.
'   - pattern: The pattern to search for.
' Returns: True if the pattern is found, False otherwise.
Function regex(text As String, pattern As String) As Boolean
    Dim regexObj As Object
    Set regexObj = CreateObject("VBScript.RegExp")
    regexObj.Global = True
    regexObj.pattern = pattern
    regex = regexObj.test(text)
End Function

' Description: This function get the type of a variable.
' Parameters:
'   - variable: The variable to analyze.
' Returns: The type of the variable.
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

' Description: This function gets the line number of a value in a column.
' Parameters:
'   - ws: The worksheet to analyze.
'   - col: The column to search.
'   - val: The value to search for.
' Returns: The line number of the value.
Function get_line_nb(ws As Worksheet, col As Integer, val As Variant) As Long
    Dim rng As Range
    Dim line As Long
    
    Set rng = ws.Columns(col).Find(What:=val, LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not rng Is Nothing Then
        line = rng.row
    Else
        line = 0
    End If
    
    get_line_nb = line
End Function

' Description: This function gets the value of a cell.
' Parameters:
'   - ws: The worksheet to analyze.
'   - row: The row of the cell.
'   - column: The column of the cell.
' Returns: The value of the cell.
Function get_value_of_cell(ws As Worksheet, row As Integer, column As Integer) As Variant
    get_value_of_cell = ws.Cells(row, column).value
End Function

' Description: This function checks if a text contains "Facture".
' Parameters:
'   - text: The text to search.
' Returns: "F" if the text contains "Facture", "A" otherwise.
Function facture_or_not(text As String) As String
    If InStr(text, "S") > 0 Then
        facture_or_not = "F"
    Else
        facture_or_not = "A"
    End If
End Function

' Description: This function formats a Customer
' Parameters:
'   - tiers: The Customer to format.
' Returns: The formatted Customer.
Function format_tiers(ByVal tiers As String) As String
    Dim formattedTiers As String
    Dim charToAdd As Integer

    charToAdd = 15 - Len(tiers)
    formattedTiers = tiers & String(charToAdd, "0")

    format_tiers = formattedTiers
End Function

' Description: This function changes the format of a date.
' Parameters:
'   - dateStr: The date to format.
' Returns: The formatted date.
Function ChangeFormatDate(ByVal dateStr As String) As String
    Dim pattern As String
    Dim replaceStr As String
    
    ' Définir le modèle de recherche pour les chiffres
    pattern = "(\d{2})\.(\d{2})\.(\d{4})"
    
    ' Définir la chaîne de remplacement
    replaceStr = "$3$2$1"
    
    ' Appliquer la fonction regex_replace pour changer le format
    ChangeFormatDate = regex_replace(dateStr, pattern, replaceStr)
End Function

' Description: This function changes the format of a date for SAP.
' Parameters:
'   - dateStr: The date to format.
' Returns: The formatted date.
Function FDateSap(ByVal dateStr As String) As String
    Dim pattern As String
    Dim replaceStr As String
    
    ' Définir le modèle de recherche pour les chiffres
    pattern = "(\d{2})\.(\d{2})\.(\d{4})"
    
    ' Définir la chaîne de remplacement
    replaceStr = "$3-$2-$1"
    
    ' Appliquer la fonction regex_replace pour changer le format
    FDateSap = regex_replace(dateStr, pattern, replaceStr)
End Function

' Description: This function reverse the amount where the nature.
' Parameters:
'   - montant: The amount to reverse.
'   - Nature: The nature of the amount.
' Returns: The reversed amount.
Function inverse_montant(ByVal montant As String, ByVal Nature As String) As Double
    If Nature = "S" Then
        inverse_montant = "-" & montant
    Else
        inverse_montant = regex_replace(montant, "-", "")
    End If
End Function

' Description: This function reverse the nature.
' Parameters:
'   - Nature: The nature to reverse.
' Returns: The reversed nature.
Function inverse_nature(ByVal Nature As String) As String
    If Nature = "S" Then
        inverse_nature = "H"
    Else
        inverse_nature = "S"
    End If
End Function

' Description: This function write the value in a cell.
' Parameters:
'   - ws: The worksheet to analyze.
'   - line: The line of the cell.
'   - column: The column of the cell.
'   - value: The value to write.
' Returns: True if the value was written successfully, False otherwise.
Function report_value(ws As Worksheet, ByVal line As Long, ByVal column As Long, ByVal value As String) As Boolean
    Dim cell As Range

    Set cell = ws.Cells(line, column)
    If cell.value = "" Then
        cell.value = value
        report_value = True
    Else
        report_value = False
    End If
End Function


' #################################################################################################################
' #                                                   MAIN FUNCTION                                               #
' #################################################################################################################
' Description: This function is the main function for the domestic factor.
' It's use for make file for factor and send request to API SAP.
Public Sub main_domestique()
    Dim ws As Worksheet
    Dim convert As Worksheet
    Dim end_ws As Integer
    Dim remise As String
    Dim remise_date As String
    Dim client As String
    Dim Devise As String
    Dim file_path As String
    Dim data As Variant
    Dim start_index As String
    Dim GLAccount_write_factor As String
    Dim GLAccount_write_stat As String
    Dim res As String
    
    ' select sheet
    Set ws = Worksheets("REMISE DOMESTIQUE")
    Set convert = Worksheets("Import")
    ws.Activate
    ' fixed value
    start_index = "A13"
    remise = ws.Range("B7").value
    remise_date = ws.Range("B8").valueRGB(169, 208, 142)
    client = Mid(ws.Range("D6").value, 9)
    Devise = ws.Range("B9").value
    file_path = ws.Range("B5").value & "\" & Mid(remise, 1, 8) & ".txt"
    GLAccount_write_factor = "46710000"
    GLAccount_write_stat = "41100000"

    ' transform data selection to send factor in variant
    end_ws = search_end_point_in_column(ws, ws.Range(start_index), "Total général")
    data = extract_variant(ws, start_index, "F" & (end_ws - 1))
    ' create and write in txt file
    If search_file(file_path) = False Then
        create_file (file_path)
    End If
    'Else
        'delete_file (file_path)
        'create_file (file_path)
    'End If
    
    Dim i As Long
    For i = LBound(data, 1) To UBound(data, 1)
        If Len(get_value_of_cell(ws, i + 12, 8)) < 1 Then
            ' send request to API
            res = make_and_send_xml_request(CStr(i), _
            "doc.ref", _
            "XYZ", _
            "RFBU", _
            get_value_of_cell(convert, get_line_nb(convert, 16, data(i, 3)), 6), data(i, 1), _
            "htext", _
            "CB9980000080", _
            "1000", _
            FDateSap(data(i, 2)), FDateSap(data(i, 5)), _
            "Ref1", _
            "Ref2", _
            GLAccount_write_factor, _
            data(i, 4), _
            get_value_of_cell(convert, get_line_nb(convert, 16, data(i, 3)), 7), _
            get_value_of_cell(convert, get_line_nb(convert, 16, data(i, 3)), 17), _
            "1", _
            inverse_montant(data(i, 4), get_value_of_cell(convert, get_line_nb(convert, 16, data(i, 3)), 7)), _
            inverse_nature(get_value_of_cell(convert, get_line_nb(convert, 16, data(i, 3)), 7)), _
            "A0", _
            "MWS", _
            "0,00", _
            get_value_of_cell(convert, get_line_nb(convert, 16, data(i, 3)), 7), _
            data(i, 4), _
            get_value_of_cell(convert, get_line_nb(convert, 16, data(i, 3)), 4), _
            get_value_of_cell(convert, get_line_nb(convert, 16, data(i, 3)), 13))

            If res = "0000000000" Then
                MsgBox "Erreur lors de l'envoi de la requête à l'API"
                Exit Sub
            Else
                report_value ws, i + 12, 8, res
            End If
        End If
    Next i

    Dim b As Long
    For b = LBound(data, 1) To UBound(data, 1)
        If Len(get_value_of_cell(ws, b + 12, 8)) < 1 Then
            ' send request to API
            res = make_and_send_request_stat(CStr(b), _
            "doc.ref", _
            "XYZ", _
            "RFBU", _
            get_value_of_cell(convert, get_line_nb(convert, 16, data(b, 3)), 6), data(b, 1), _
            "htext", _
            "CB9980000080", _
            "1000", _
            FDateSap(data(b, 2)), FDateSap(data(b, 5)), _
            "Ref1", _
            "Ref2", _
            GLAccount_write_stat, _
            data(b, 4), _
            get_value_of_cell(convert, get_line_nb(convert, 16, data(b, 3)), 7), _
            get_value_of_cell(convert, get_line_nb(convert, 16, data(b, 3)), 17), _
            "1", _
            data(b, 4), _
            get_value_of_cell(convert, get_line_nb(convert, 16, data(b, 3)), 7), _
            "A0", _
            "MWS", _
            "0,00", _
            inverse_nature(get_value_of_cell(convert, get_line_nb(convert, 16, data(b, 3)), 7)), _
            inverse_montant(data(b, 4), get_value_of_cell(convert, get_line_nb(convert, 16, data(b, 3)), 7)), _
            get_value_of_cell(convert, get_line_nb(convert, 16, data(b, 3)), 4), _
            get_value_of_cell(convert, get_line_nb(convert, 16, data(b, 3)), 13))

            If res = "0000000000" Then
                MsgBox "Erreur lors de l'envoi de la requête à l'API"
                Exit Sub
            Else
                report_value ws, b + 12, 9, res
            End If
        End If
    Next b

    Dim J As Long
    For J = LBound(data, 1) To UBound(data, 1)
        If Len(get_value_of_cell(ws, j + 12, 8)) > 0 And Len(get_value_of_cell(ws, J + 12, 9)) > 0 Then
        ' write in txt file
            'Debug.Print Mid(remise, 4, 5) & ";" & client & ";" & regex_replace(remise_date, "/", "") & ";" & "D" & ";" & "511" & ";" & facture_or_not(get_value_of_cell(convert, get_line_nb(convert, 16, data(j, 3)), 6)) & ";" & Devise & ";" & regex_replace(CStr(data(j, 1)), "[^0-9]", "") & ";" & format_tiers(get_value_of_cell(convert, get_line_nb(convert, 16, data(j, 3)), 4)) & ";                       ;" & format_reference(data(j, 3)) & ";" & format_valeur(data(j, 4)) & ";" & ChangeFormatDate(data(j, 2)) & ";" & ChangeFormatDate(data(j, 5)) & ";" & CStr(data(j, 6)) & ";0123456789;                         ;              ;                                                   ;   ;"
            write_in_txt_file file_path, Mid(remise, 4, 5) & ";" & client & ";" & regex_replace(remise_date, "/", "") & ";" & "D" & ";" & "511" & ";" & facture_or_not(get_value_of_cell(convert, get_line_nb(convert, 16, data(j, 3)), 7)) & ";" & Devise & ";" & regex_replace(CStr(data(j, 1)), "[^0-9]", "") & ";" & format_tiers(get_value_of_cell(convert, get_line_nb(convert, 16, data(j, 3)), 4)) & ";" & "                       ;" & format_reference(data(j, 3)) & ";" & format_valeur(data(j, 4)) & ";" & ChangeFormatDate(data(j, 2)) & ";" & ChangeFormatDate(data(j, 5)) & ";" & CStr(data(j, 6)) & ";0123456789;                         ;              ;                                                   ;   ;"
        End If
    Next J
    MsgBox "Fichier " & Mid(remise, 1, 8) & ".txt créé avec succès !"
End Sub

