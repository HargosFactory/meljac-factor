' Language: VBA
' Category: Accounting & Factoring
' Description: This script creates a journal entry in SAP S/4HANA Cloud using the Journal Entry API.
' API: Journal Entry - PÖST (Async) doc: https://api.sap.com/api/JOURNALENTRYBULKCREATIONREQUES/resource
' Author: FLORENTIN William
' Organization: HARGOS
' Version: 1.0
' Date: 2024-03-25
' Last update: 2024-04-18

' #################################################################################################################
' #                                                   FUNCTIONS                                                   #
' #################################################################################################################

' Description: This function sends a request to the API.
' Parameters:
'   - url: The URL of the API.
'   - xml_body: The XML body of the request.
'   - username: The username for authentication.
'   - password: The password for authentication.
' Returns: The response from the API.
' Example: send_request("https://my404630-api.s4hana.cloud.sap/sap/bc/srt/scs_ext/sap/journalentrycreaterequestconfi", xml_body, "user", "password")
Function send_request(ByVal url As String, ByVal xml_body As String, username As String, password As String) As String
    Dim xmlhttp As Object
    Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
    
    ' Envoyer la requête
    xmlhttp.Open "POST", url, False
    xmlhttp.setRequestHeader "Content-Type", "text/xml"
    ' Add authentication headers if necessary
    xmlhttp.setRequestHeader "Authorization", "Basic " & Base64Encode(username & ":" & password)
    
    xmlhttp.send xml_body
    
    ' Vérifier le statut de la réponse
    If xmlhttp.Status <> 200 Then
        MsgBox "Request failed with status " & xmlhttp.Status, vbCritical
    End If
    
    ' Renvoyer la réponse
    send_request = xmlhttp.responseText
End Function


Function Base64Encode(ByVal sText As String) As String
    Dim arrData() As Byte
    arrData = StrConv(sText, vbFromUnicode)
    Dim objXML As Object
    Dim objNode As Object
    Set objXML = CreateObject("MSXML2.DOMDocument")
    Set objNode = objXML.createElement("b64")
    objNode.DataType = "bin.base64"
    objNode.nodeTypedValue = arrData
    Base64Encode = objNode.text
    Set objNode = Nothing
    Set objXML = Nothing
End Function


' Description: This function constructs the XML payload for the request.
' Parameters:
'   - item: The item to create.
' Returns: The XML payload for the request.
Function construct_payload_xml_request(item As Object) As String
    Dim xml As String
    xml = "<?xml version=""1.0"" encoding=""UTF-8""?>" & _
      "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:sfin=""http://sap.com/xi/SAPSCORE/SFIN"">" & _
          "<soapenv:Header>" & _
              "<msgID:messageId xmlns:msgID=""http://www.sap.com/webas/640/soap/features/messageId/"">uuid:00163e02-a5da-1ee1-a7d5-e24f8067c623</msgID:messageId>" & _
          "</soapenv:Header>" & _
          "<soapenv:Body>" & _
              "<sfin:JournalEntryBulkCreateRequest xmlns:sfin=""http://sap.com/xi/SAPSCORE/SFIN"">" & _
                  "<MessageHeader>" & _
                      "<CreationDateTime>" & Format(DateTime.Now, "yyyy-MM-ddThh:mm:ssZ") & "</CreationDateTime>" & _
                      "<ID>MSG_" & Format(DateTime.Now, "yyyy-MM_dd") & "_" & item.GetId & "</ID>" & _
                  "</MessageHeader>"
    xml = xml & "<JournalEntryCreateRequest>" & _
                      "<MessageHeader>" & _
                          "<CreationDateTime>" & Format(DateTime.Now, "yyyy-MM-ddThh:mm:ssZ") & "</CreationDateTime>" & _
                          "<ID>SUBMSG_" & Format(DateTime.Now, "yyyy-MM-dd") & "_" & item.GetId & "</ID>" & _
                      "</MessageHeader>" & _
                      "<JournalEntry>" & _
                          "<OriginalReferenceDocumentType>BKPFF</OriginalReferenceDocumentType>" & _
                          "<OriginalReferenceDocument>" & item.GetOriginalReferenceDocument & "</OriginalReferenceDocument>" & _
                          "<OriginalReferenceDocumentLogicalSystem>" & item.GetOriginalReferenceDocumentLogicalSystem & "</OriginalReferenceDocumentLogicalSystem>" & _
                          "<BusinessTransactionType>" & item.GetBusinessTransactionType & "</BusinessTransactionType>" & _
                          "<AccountingDocumentType>" & item.GetAccountingDocumentType & "</AccountingDocumentType>" & _
                          "<DocumentReferenceID>" & item.GetDocumentReferenceID & "</DocumentReferenceID>" & _
                          "<DocumentHeaderText>" & item.GetDocumentHeaderText & "</DocumentHeaderText>" & _
                          "<CreatedByUser>" & item.GetCreatedByUser & "</CreatedByUser>" & _
                          "<CompanyCode>" & item.GetCompanyCode & "</CompanyCode>" & _
                          "<DocumentDate>" & item.GetDocumentDate & "</DocumentDate>" & _
                          "<PostingDate>" & Format(DateTime.Now, "yyyy-MM-dd") & "</PostingDate>" & _
                          "<TaxDeterminationDate>" & item.GetTaxDeterminationDate & "</TaxDeterminationDate>" & _
                          "<Reference1InDocumentHeader>" & item.GetReference1InDocumentHeader & "</Reference1InDocumentHeader>" & _
                          "<Reference2InDocumentHeader>" & item.GetReference2InDocumentHeader & "</Reference2InDocumentHeader>" & _
                          "<Item>"
    xml = xml & "<GLAccount>" & item.GetGLAccount & "</GLAccount>" & _
                              "<AmountInTransactionCurrency currencyCode=""" & item.GetDevise & """>" & item.GetItemAmountInTransactionCurrency & "</AmountInTransactionCurrency>" & _
                              "<DebitCreditCode>" & item.GetItemDebitCreditCode & "</DebitCreditCode>" & _
                              "<DocumentItemText>" & item.GetItemDocumentItemText & "</DocumentItemText>" & _
                          "</Item>" & _
                          "<DebtorItem>"
    xml = xml & "<ReferenceDocumentItem>" & item.GetCreditItemReferenceDocumentItem & "</ReferenceDocumentItem>" & _
                              "<Debtor>" & item.GetDebtor & "</Debtor>" & _
                              "<AmountInTransactionCurrency currencyCode=""" & item.GetDevise & """>" & item.GetCreditItemAmountInTransactionCurrency & "</AmountInTransactionCurrency>" & _
                              "<DebitCreditCode>" & item.GetCreditItemDebitCreditCode & "</DebitCreditCode>" & _
                          "</DebtorItem>" & _
                          "<ProductTaxItem>"
    xml = xml & "<TaxCode>" & item.GetProductTaxItemTaxCode & "</TaxCode>" & _
                              "<TaxItemClassification>" & item.GetProductTaxItemTaxItemClassification & "</TaxItemClassification>" & _
                              "<IsDirectTaxPosting>true</IsDirectTaxPosting>" & _
                              "<AmountInTransactionCurrency currencyCode=""" & item.GetDevise & """>" & item.GetProductTaxItemAmountInTransactionCurrency & "</AmountInTransactionCurrency>" & _
                              "<DebitCreditCode>" & item.GetProductTaxItemDebitCreditCode & "</DebitCreditCode>" & _
                              "<TaxBaseAmountInTransCrcy currencyCode=""" & item.GetDevise & """>" & item.GetProductTaxItemTaxBaseAmountInTransCrcy & "</TaxBaseAmountInTransCrcy>" & _
                          "</ProductTaxItem>" & _
                      "</JournalEntry>" & _
                  "</JournalEntryCreateRequest>" & _
              "</sfin:JournalEntryBulkCreateRequest>" & _
          "</soapenv:Body>" & _
      "</soapenv:Envelope>"
      
    construct_payload_xml_request = xml
End Function

Function construct_payload_xml_request_stat(item As Object) As String
    Dim xml As String
    xml = "<?xml version=""1.0"" encoding=""UTF-8""?>" & _
      "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:sfin=""http://sap.com/xi/SAPSCORE/SFIN"">" & _
          "<soapenv:Header>" & _
              "<msgID:messageId xmlns:msgID=""http://www.sap.com/webas/640/soap/features/messageId/"">uuid:00163e02-a5da-1ee1-a7d5-e24f8067c623</msgID:messageId>" & _
          "</soapenv:Header>" & _
          "<soapenv:Body>" & _
              "<sfin:JournalEntryBulkCreateRequest xmlns:sfin=""http://sap.com/xi/SAPSCORE/SFIN"">" & _
                  "<MessageHeader>" & _
                      "<CreationDateTime>" & Format(DateTime.Now, "yyyy-MM-ddThh:mm:ssZ") & "</CreationDateTime>" & _
                      "<ID>MSG_" & Format(DateTime.Now, "yyyy-MM_dd") & "_" & item.GetId & "</ID>" & _
                  "</MessageHeader>"
    xml = xml & "<JournalEntryCreateRequest>" & _
                      "<MessageHeader>" & _
                          "<CreationDateTime>" & Format(DateTime.Now, "yyyy-MM-ddThh:mm:ssZ") & "</CreationDateTime>" & _
                          "<ID>SUBMSG_" & Format(DateTime.Now, "yyyy-MM-dd") & "_" & item.GetId & "</ID>" & _
                      "</MessageHeader>" & _
                      "<JournalEntry>" & _
                          "<OriginalReferenceDocumentType>BKPFF</OriginalReferenceDocumentType>" & _
                          "<OriginalReferenceDocument>" & item.GetOriginalReferenceDocument & "</OriginalReferenceDocument>" & _
                          "<OriginalReferenceDocumentLogicalSystem>" & item.GetOriginalReferenceDocumentLogicalSystem & "</OriginalReferenceDocumentLogicalSystem>" & _
                          "<BusinessTransactionType>" & item.GetBusinessTransactionType & "</BusinessTransactionType>" & _
                          "<AccountingDocumentType>" & item.GetAccountingDocumentType & "</AccountingDocumentType>" & _
                          "<DocumentReferenceID>" & item.GetDocumentReferenceID & "</DocumentReferenceID>" & _
                          "<DocumentHeaderText>" & item.GetDocumentHeaderText & "</DocumentHeaderText>" & _
                          "<CreatedByUser>" & item.GetCreatedByUser & "</CreatedByUser>" & _
                          "<CompanyCode>" & item.GetCompanyCode & "</CompanyCode>" & _
                          "<DocumentDate>" & item.GetDocumentDate & "</DocumentDate>" & _
                          "<PostingDate>" & Format(DateTime.Now, "yyyy-MM-dd") & "</PostingDate>" & _
                          "<TaxDeterminationDate>" & item.GetTaxDeterminationDate & "</TaxDeterminationDate>" & _
                          "<Reference1InDocumentHeader>" & item.GetReference1InDocumentHeader & "</Reference1InDocumentHeader>" & _
                          "<Reference2InDocumentHeader>" & item.GetReference2InDocumentHeader & "</Reference2InDocumentHeader>" & _
                          "<Item>"
    xml = xml & "<GLAccount>" & item.GetGLAccount & "</GLAccount>" & _
                              "<AmountInTransactionCurrency currencyCode=""" & item.GetDevise & """>" & item.GetItemAmountInTransactionCurrency & "</AmountInTransactionCurrency>" & _
                              "<DebitCreditCode>" & item.GetItemDebitCreditCode & "</DebitCreditCode>" & _
                              "<DocumentItemText>" & item.GetItemDocumentItemText & "</DocumentItemText>" & _
                          "</Item>" & _
                          "<CreditorItem>"
    xml = xml & "<ReferenceDocumentItem>" & item.GetCreditItemReferenceDocumentItem & "</ReferenceDocumentItem>" & _
                              "<Creditor>" & item.GetCreditor & "</Creditor>" & _
                              "<AmountInTransactionCurrency currencyCode=""" & item.GetDevise & """>" & item.GetCreditItemAmountInTransactionCurrency & "</AmountInTransactionCurrency>" & _
                              "<DebitCreditCode>" & item.GetCreditItemDebitCreditCode & "</DebitCreditCode>" & _
                          "</CreditorItem>" & _
                          "<ProductTaxItem>"
    xml = xml & "<TaxCode>" & item.GetProductTaxItemTaxCode & "</TaxCode>" & _
                              "<TaxItemClassification>" & item.GetProductTaxItemTaxItemClassification & "</TaxItemClassification>" & _
                              "<IsDirectTaxPosting>true</IsDirectTaxPosting>" & _
                              "<AmountInTransactionCurrency currencyCode=""" & item.GetDevise & """>" & item.GetProductTaxItemAmountInTransactionCurrency & "</AmountInTransactionCurrency>" & _
                              "<DebitCreditCode>" & item.GetProductTaxItemDebitCreditCode & "</DebitCreditCode>" & _
                              "<TaxBaseAmountInTransCrcy currencyCode=""" & item.GetDevise & """>" & item.GetProductTaxItemTaxBaseAmountInTransCrcy & "</TaxBaseAmountInTransCrcy>" & _
                          "</ProductTaxItem>" & _
                      "</JournalEntry>" & _
                  "</JournalEntryCreateRequest>" & _
              "</sfin:JournalEntryBulkCreateRequest>" & _
          "</soapenv:Body>" & _
      "</soapenv:Envelope>"
      
    construct_payload_xml_request_stat = xml
End Function

' Description: This function extracts the accounting document from the response XML.
' Parameters:
'   - xml: The XML response from the API.
' Returns: The accounting document extracted from the XML.
Function ExtractAccountingDocument(xml As String) As String
    Dim xmlDoc As Object
    Set xmlDoc = CreateObject("MSXML2.DOMDocument")
    
    ' Charger la chaîne XML dans l'objet DOMDocument
    xmlDoc.LoadXML xml
    
    ' Vérifier si le chargement s'est effectué avec succès
    If xmlDoc.parseError.ErrorCode <> 0 Then
        MsgBox "Erreur de chargement XML: " & xmlDoc.parseError.reason
        Exit Function
    End If
    
    ' Rechercher l'élément <AccountingDocument>
    Dim accountingDocNode As Object
    Set accountingDocNode = xmlDoc.SelectSingleNode("//AccountingDocument")
    
    ' Vérifier si l'élément a été trouvé
    If Not accountingDocNode Is Nothing Then
        ' Récupérer la valeur de l'élément <AccountingDocument>
        ExtractAccountingDocument = accountingDocNode.text
    Else
        ' Si l'élément n'a pas été trouvé, retourner une chaîne vide ou une valeur par défaut
        ExtractAccountingDocument = ""
    End If
End Function


' Description: This function creates and sends an XML request to the API.
' Returns: The accounting document created.
Public Function make_and_send_xml_request(ByVal ID As String, ByVal Doc As String, ByVal LogicalSystem As String, ByVal BTType As String, ByVal AcctType As String, ByVal RefID As String, ByVal HeaderText As String, ByVal CreatedBy As String, ByVal Company As String, ByVal DocDate As String, ByVal TaxDate As String, ByVal Ref1 As String, ByVal Ref2 As String, ByVal Acc As String, ByVal Amount As String, ByVal DebitCredit As String, ByVal DocItemText As String, ByVal CreditItemRefItem As String, ByVal CreditItemAmount As String, ByVal CreditItemDebitCredit As String, ByVal ProdTaxCode As String, ByVal ProdTaxClass As String, ByVal ProdTaxAmount As String, ByVal ProdTaxDebitCredit As String, ByVal ProdTaxBaseAmount As String, ByVal Debtor As String, ByVal Devise As String)
    Dim api_url As String
    Dim username As String
    Dim password As String
    Dim req As item
    Dim xml_body As String
    Dim res As String

    api_url = "https://my404630-api.s4hana.cloud.sap/sap/bc/srt/scs_ext/sap/journalentrycreaterequestconfi"
    username = "API_USER"
    password = "mpsSREjdeEdfFTPyUlCPkPnmWQE9G+QyTLqhAyjp"

    Set req = New item
    req.SetId ID
    req.SetOriginalReferenceDocument Doc
    req.SetOriginalReferenceDocumentLogicalSystem LogicalSystem
    req.SetBusinessTransactionType BTType
    req.SetAccountingDocumentType AcctType
    req.SetDocumentReferenceID RefID
    req.SetDocumentHeaderText HeaderText
    req.SetCreatedByUser CreatedBy
    req.SetCompanyCode Company
    req.SetDocumentDate DocDate
    req.SetTaxDeterminationDate TaxDate
    req.SetReference1InDocumentHeader Ref1
    req.SetReference2InDocumentHeader Ref2
    req.SetGLAccount Acc
    req.SetItemAmountInTransactionCurrency Amount
    req.SetItemDebitCreditCode DebitCredit
    req.SetItemDocumentItemText DocItemText
    req.SetCreditItemReferenceDocumentItem CreditItemRefItem
    req.SetCreditItemAmountInTransactionCurrency CreditItemAmount
    req.SetCreditItemDebitCreditCode CreditItemDebitCredit
    req.SetProductTaxItemTaxCode ProdTaxCode
    req.SetProductTaxItemTaxItemClassification ProdTaxClass
    req.SetProductTaxItemAmountInTransactionCurrency ProdTaxAmount
    req.SetProductTaxItemDebitCreditCode ProdTaxDebitCredit
    req.SetProductTaxItemTaxBaseAmountInTransCrcy ProdTaxBaseAmount
    req.SetDebtor Debtor
    req.SetDevise Devise

    ' Make sure the XML body is constructed correctly
    xml_body = construct_payload_xml_request(req)

    ' Send the request and extract the accounting document a response API and return it
    make_and_send_xml_request = ExtractAccountingDocument(send_request(api_url, xml_body, username, password))
End Function

Public Function make_and_send_request_stat(ByVal ID As String, ByVal Doc As String, ByVal LogicalSystem As String, ByVal BTType As String, ByVal AcctType As String, ByVal RefID As String, ByVal HeaderText As String, ByVal CreatedBy As String, ByVal Company As String, ByVal DocDate As String, ByVal TaxDate As String, ByVal Ref1 As String, ByVal Ref2 As String, ByVal Acc As String, ByVal Amount As String, ByVal DebitCredit As String, ByVal DocItemText As String, ByVal CreditItemRefItem As String, ByVal CreditItemAmount As String, ByVal CreditItemDebitCredit As String, ByVal ProdTaxCode As String, ByVal ProdTaxClass As String, ByVal ProdTaxAmount As String, ByVal ProdTaxDebitCredit As String, ByVal ProdTaxBaseAmount As String, ByVal Creditor As String, ByVal Devise As String)
    Dim api_url As String
    Dim username As String
    Dim password As String
    Dim req As item_stat
    Dim xml_body As String
    Dim res As String

    api_url = "https://my404630-api.s4hana.cloud.sap/sap/bc/srt/scs_ext/sap/journalentrycreaterequestconfi"
    username = "API_USER"
    password = "mpsSREjdeEdfFTPyUlCPkPnmWQE9G+QyTLqhAyjp"

    Set req = New item_stat
    req.SetId ID
    req.SetOriginalReferenceDocument Doc
    req.SetOriginalReferenceDocumentLogicalSystem LogicalSystem
    req.SetBusinessTransactionType BTType
    req.SetAccountingDocumentType AcctType
    req.SetDocumentReferenceID RefID
    req.SetDocumentHeaderText HeaderText
    req.SetCreatedByUser CreatedBy
    req.SetCompanyCode Company
    req.SetDocumentDate DocDate
    req.SetTaxDeterminationDate TaxDate
    req.SetReference1InDocumentHeader Ref1
    req.SetReference2InDocumentHeader Ref2
    req.SetGLAccount Acc
    req.SetItemAmountInTransactionCurrency Amount
    req.SetItemDebitCreditCode DebitCredit
    req.SetItemDocumentItemText DocItemText
    req.SetCreditItemReferenceDocumentItem CreditItemRefItem
    req.SetCreditItemAmountInTransactionCurrency CreditItemAmount
    req.SetCreditItemDebitCreditCode CreditItemDebitCredit
    req.SetProductTaxItemTaxCode ProdTaxCode
    req.SetProductTaxItemTaxItemClassification ProdTaxClass
    req.SetProductTaxItemAmountInTransactionCurrency ProdTaxAmount
    req.SetProductTaxItemDebitCreditCode ProdTaxDebitCredit
    req.SetProductTaxItemTaxBaseAmountInTransCrcy ProdTaxBaseAmount
    req.SetCreditor Creditor
    req.SetDevise Devise

    ' Make sure the XML body is constructed correctly
    xml_body = construct_payload_xml_request_stat(req)

    ' Send the request and extract the accounting document a response API and return it
    make_and_send_xml_request_stat = ExtractAccountingDocument(send_request(api_url, xml_body, username, password))
End Function



