Function search_end_point_in_column_with_start(ws As Worksheet, column As Range, value As String, startRow As Long) As Long
    Dim i As Long
    Dim end_ws As Long
    end_ws = define_end_sheet(ws, column)
    For i = startRow To end_ws
        If value = "" Then
            If IsEmpty(ws.Cells(i, column.column).value) Then
                search_end_point_in_column_with_start = i
                Exit Function
            End If
        Else
            If ws.Cells(i, column.column).value = value Then
                search_end_point_in_column_with_start = i
                Exit Function
            End If
        End If
    Next i
    search_end_point_in_column_with_start = i ' bug du language
End Function

Function FDateSapForPayment(ByVal dateStr As String) As String
    Dim pattern As String
    Dim replaceStr As String
    
    pattern = "(\d{2})\/(\d{2})\/(\d{4})"
    replaceStr = "$3-$2-$1"
    FDateSap = regex_replace(dateStr, pattern, replaceStr)

End Function

Public Sub main_paiement_domestique()
    Dim ws As Worksheet
    Set ws = Worksheets("PAIEMENT DOMESTIQUE")
    ws.Activate
    
    Dim GLAccount As String
    Dim end_ws As Integer
    Dim start_index As String
    Dim data As Variant
    Dim manual_client As String
    Dim response_api As String
    
    start_index = "A5"
    GLAccount = "46710000"
    manual_client = "7777777"
    end_ws = search_end_point_in_column_with_start(ws, ws.Range(start_index), "", 5)
    data = extract_variant(ws, start_index, "Q" & (end_ws - 1))
    
    Dim i As Long
    For i = LBound(data, 1) To UBound(data, 1)
        If get_value_of_cell(ws, i + 4, 4) <> manual_client Then


            response_api = make_and_send_request_stat_willi(CStr(i), _
                "doc.ref", _
                "XYZ", _
                "RFBU", _
                "UE", _
                "htext", _
                "CB9980000080", _
                "1000", _
                FDateSapForPayment(data(i, 7)), FDateSapForPayment(data(i, 7)), "2587690", _
                "Ref1", _
                "Ref2", _
                GLAccount, _
                data(i, 9), _
                "S", _
                data(i, 17), _
                "1", _
                inverse_montant(data(i, 9), "S"), _
                "H", _
                "A0", _
                "MWS", _
                "0,00", _
                "H", _
                inverse_montant(data(i, 9), "S"), _
                "truc", _
                "truc")


            If CStr(res) = "0000000000" Then
                MsgBox "Erreur lors de l'envoi de la requête à l'API payment"
                Exit Sub
            Else
                report_value ws, b + 12, 9, res
            End If
        End If
    Next i
    
End Sub