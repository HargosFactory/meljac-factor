' Class Module Item
Private ID As String
Private OriginalReferenceDocument As String
Private OriginalReferenceDocumentLogicalSystem As String
Private BusinessTransactionType As String
Private AccountingDocumentType As String
Private DocumentReferenceID As String
Private DocumentHeaderText As String
Private CreatedByUser As String
Private CompanyCode As String
Private DocumentDate As String
Private TaxDeterminationDate As String
Private Reference1InDocumentHeader As String
Private Reference2InDocumentHeader As String
Private GLAccount As String
Private ItemAmountInTransactionCurrency As String
Private ItemDebitCreditCode As String
Private ItemDocumentItemText As String
'Private ItemTaxCode As String
'Private ItemProfitCenter As String
Private CreditItemReferenceDocumentItem As String
Private CreditItemAmountInTransactionCurrency As String
Private CreditItemDebitCreditCode As String
Private ProductTaxItemTaxCode As String
Private ProductTaxItemTaxItemClassification As String
Private ProductTaxItemAmountInTransactionCurrency As String
Private ProductTaxItemDebitCreditCode As String
Private ProductTaxItemTaxBaseAmountInTransCrcy As String
Private Debtor As String
Private Devise As String

Private Function regex(text As String, pattern As String) As Boolean
    Dim regexObj As Object
    Set regexObj = CreateObject("VBScript.RegExp")
    regexObj.Global = True
    regexObj.pattern = pattern
    regex = regexObj.test(text)
End Function

Private Function Interval(ByVal Value As String, ByVal Min As Integer, ByVal Max As Integer) As Boolean
    If Len(Value) >= Min And Len(Value) <= Max Then
        Interval = True
    Else
        Interval = False
    End If
End Function

Private Function regex_replace(text As String, pattern As String, replace As String) As String
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.pattern = pattern
    regex_replace = regex.replace(text, replace)
End Function

' Méthods
Public Sub SetId(ByVal Value As String)
    If Interval(Value, 1, 60) Then
        ID = Value
    Else
        MsgBox "Invalid ID"
    End If
End Sub

Public Function GetId()
    GetId = ID
End Function

Public Sub SetOriginalReferenceDocument(ByVal Value As String)
    If Interval(Value, 1, 20) Then
        OriginalReferenceDocument = Value
    Else
        MsgBox "Invalid OriginalReferenceDocument"
    End If
End Sub

Public Function GetOriginalReferenceDocument()
    GetOriginalReferenceDocument = OriginalReferenceDocument
End Function

Public Sub SetOriginalReferenceDocumentLogicalSystem(ByVal Value As String)
    If Interval(Value, 1, 10) Then
        OriginalReferenceDocumentLogicalSystem = Value
    Else
        MsgBox "Invalid OriginalReferenceDocumentLogicalSystem"
    End If
End Sub

Public Function GetOriginalReferenceDocumentLogicalSystem()
    GetOriginalReferenceDocumentLogicalSystem = OriginalReferenceDocumentLogicalSystem
End Function

Public Sub SetBusinessTransactionType(ByVal Value As String)
    If Interval(Value, 1, 4) Then
        BusinessTransactionType = Value
    Else
        MsgBox "Invalid BusinessTransactionType"
    End If
End Sub

Public Function GetBusinessTransactionType()
    GetBusinessTransactionType = BusinessTransactionType
End Function

Public Sub SetAccountingDocumentType(ByVal Value As String)
    If Interval(Value, 1, 2) Then
        AccountingDocumentType = Value
    Else
        MsgBox "Invalid AccountingDocumentType"
    End If
End Sub

Public Function GetAccountingDocumentType()
    GetAccountingDocumentType = AccountingDocumentType
End Function

Public Sub SetDocumentReferenceID(ByVal Value As String)
    If Interval(Value, 1, 60) Then
        DocumentReferenceID = Value
    Else
        MsgBox "Invalid DocumentReferenceID"
    End If
End Sub

Public Function GetDocumentReferenceID()
    GetDocumentReferenceID = DocumentReferenceID
End Function

Public Sub SetDocumentHeaderText(ByVal Value As String)
    If Interval(Value, 1, 60) Then
        DocumentHeaderText = Value
    Else
        MsgBox "Invalid DocumentHeaderText"
    End If
End Sub

Public Function GetDocumentHeaderText()
    GetDocumentHeaderText = DocumentHeaderText
End Function

Public Sub SetCreatedByUser(ByVal Value As String)
    If Interval(Value, 1, 12) Then
        CreatedByUser = Value
    Else
        MsgBox "Invalid CreatedByUser"
    End If
End Sub

Public Function GetCreatedByUser()
    GetCreatedByUser = CreatedByUser
End Function

Public Sub SetCompanyCode(ByVal Value As String)
    If Interval(Value, 1, 4) Then
        CompanyCode = Value
    Else
        MsgBox "Invalid CompanyCode"
    End If
End Sub

Public Function GetCompanyCode()
    GetCompanyCode = CompanyCode
End Function

Public Sub SetDocumentDate(ByVal Value As String)
    If regex(Value, "^\d{4}-\d{2}-\d{2}$") Then
        DocumentDate = Value
    Else
        MsgBox "Invalid DocumentDate"
    End If
End Sub

Public Function GetDocumentDate()
    GetDocumentDate = DocumentDate
End Function

Public Sub SetTaxDeterminationDate(ByVal Value As String)
    If regex(Value, "^\d{4}-\d{2}-\d{2}$") Then
        TaxDeterminationDate = Value
    Else
        MsgBox "Invalid TaxDeterminationDate"
    End If
End Sub

Public Function GetTaxDeterminationDate()
    GetTaxDeterminationDate = TaxDeterminationDate
End Function

Public Sub SetReference1InDocumentHeader(ByVal Value As String)
    If Interval(Value, 1, 12) Then
        Reference1InDocumentHeader = Value
    Else
        MsgBox "Invalid Reference1InDocumentHeader"
    End If
End Sub

Public Function GetReference1InDocumentHeader()
    GetReference1InDocumentHeader = Reference1InDocumentHeader
End Function

Public Sub SetReference2InDocumentHeader(ByVal Value As String)
    If Interval(Value, 1, 12) Then
        Reference2InDocumentHeader = Value
    Else
        MsgBox "Invalid Reference2InDocumentHeader"
    End If
End Sub

Public Function GetReference2InDocumentHeader()
    GetReference2InDocumentHeader = Reference2InDocumentHeader
End Function

Public Sub SetGLAccount(ByVal Value As String)
    If regex(Value, "^\d+$") Then
        GLAccount = Value
    Else
        MsgBox "Invalid GLAccount"
    End If
End Sub

Public Function GetGLAccount()
    GetGLAccount = GLAccount
End Function

Public Sub SetItemAmountInTransactionCurrency(ByVal Value As String)
    If regex(Value, "^[-]?[0-9,]+$") Then
        ItemAmountInTransactionCurrency = regex_replace(Value, ",", ".")
    Else
        MsgBox "Invalid ItemAmountInTransactionCurrency"
    End If
End Sub

Public Function GetItemAmountInTransactionCurrency()
    GetItemAmountInTransactionCurrency = ItemAmountInTransactionCurrency
End Function

Public Sub SetItemDebitCreditCode(ByVal Value As String)
    If Interval(Value, 1, 1) And (Value = "H" Or Value = "S") Then
        ItemDebitCreditCode = Value
    Else
        MsgBox "Invalid ItemDebitCreditCode"
    End If
End Sub

Public Function GetItemDebitCreditCode()
    GetItemDebitCreditCode = ItemDebitCreditCode
End Function

Public Sub SetItemDocumentItemText(ByVal Value As String)
    If Interval(Value, 1, 60) Then
        ItemDocumentItemText = Value
    Else
        MsgBox "Invalid ItemDocumentItemText"
    End If
End Sub

Public Function GetItemDocumentItemText()
    GetItemDocumentItemText = ItemDocumentItemText
End Function

Public Sub SetItemTaxCode(ByVal Value As String)
    If Interval(Value, 1, 2) Or Value = "A0" Then
        ItemTaxCode = Value
    Else
        MsgBox "Invalid ItemTaxCode"
    End If
End Sub

Public Function GetItemTaxCode()
    GetItemTaxCode = ItemTaxCode
End Function

Public Sub SetItemProfitCenter(ByVal Value As String)
    If Interval(Value, 1, 10) Then
        ItemProfitCenter = Value
    Else
        MsgBox "Invalid ItemProfitCenter"
    End If
End Sub

Public Function GetItemProfitCenter()
    GetItemProfitCenter = ItemProfitCenter
End Function

Public Sub SetCreditItemReferenceDocumentItem(ByVal Value As String)
    If Interval(Value, 1, 10) Then
        CreditItemReferenceDocumentItem = Value
    Else
        MsgBox "Invalid CreditItemReferenceDocumentItem"
    End If
End Sub

Public Function GetCreditItemReferenceDocumentItem()
    GetCreditItemReferenceDocumentItem = CreditItemReferenceDocumentItem
End Function

Public Sub SetCreditItemAmountInTransactionCurrency(ByVal Value As String)
    If regex(Value, "^[-]?[0-9,]+$") Then
        CreditItemAmountInTransactionCurrency = regex_replace(Value, ",", ".")
    Else
        MsgBox "Invalid CreditItemAmountInTransactionCurrency"
    End If
End Sub

Public Function GetCreditItemAmountInTransactionCurrency()
    GetCreditItemAmountInTransactionCurrency = CreditItemAmountInTransactionCurrency
End Function

Public Sub SetCreditItemDebitCreditCode(ByVal Value As String)
    If Value = "H" Or Value = "S" Then
        CreditItemDebitCreditCode = Value
    Else
        MsgBox "Invalid CreditItemDebitCreditCode"
    End If
End Sub

Public Function GetCreditItemDebitCreditCode()
    GetCreditItemDebitCreditCode = CreditItemDebitCreditCode
End Function

Public Sub SetProductTaxItemTaxCode(ByVal Value As String)
    If Interval(Value, 1, 2) Or Value = "A0" Then
        ProductTaxItemTaxCode = Value
    Else
        MsgBox "Invalid ProductTaxItemTaxCode"
    End If
End Sub

Public Function GetProductTaxItemTaxCode()
    GetProductTaxItemTaxCode = ProductTaxItemTaxCode
End Function

Public Sub SetProductTaxItemTaxItemClassification(ByVal Value As String)
    If Interval(Value, 1, 3) Then
        ProductTaxItemTaxItemClassification = Value
    Else
        MsgBox "Invalid ProductTaxItemTaxItemClassification"
    End If
End Sub

Public Function GetProductTaxItemTaxItemClassification()
    GetProductTaxItemTaxItemClassification = ProductTaxItemTaxItemClassification
End Function

Public Sub SetProductTaxItemAmountInTransactionCurrency(ByVal Value As String)
    If regex(Value, "^[-]?[0-9,]+$") Then
        ProductTaxItemAmountInTransactionCurrency = regex_replace(Value, ",", ".")
    Else
        MsgBox "Invalid ProductTaxItemAmountInTransactionCurrency"
    End If
End Sub

Public Function GetProductTaxItemAmountInTransactionCurrency()
    GetProductTaxItemAmountInTransactionCurrency = ProductTaxItemAmountInTransactionCurrency
End Function

Public Sub SetProductTaxItemDebitCreditCode(ByVal Value As String)
    If Value = "H" Or Value = "S" Then
        ProductTaxItemDebitCreditCode = Value
    Else
        MsgBox "Invalid ProductTaxItemDebitCreditCode"
    End If
End Sub

Public Function GetProductTaxItemDebitCreditCode()
    GetProductTaxItemDebitCreditCode = ProductTaxItemDebitCreditCode
End Function

Public Sub SetProductTaxItemTaxBaseAmountInTransCrcy(ByVal Value As String)
    If regex(Value, "^[-]?[0-9,]+$") Then
        ProductTaxItemTaxBaseAmountInTransCrcy = regex_replace(Value, ",", ".")
    Else
        MsgBox "Invalid ProductTaxItemTaxBaseAmountInTransCrcy"
    End If
End Sub

Public Function GetProductTaxItemTaxBaseAmountInTransCrcy()
    GetProductTaxItemTaxBaseAmountInTransCrcy = ProductTaxItemTaxBaseAmountInTransCrcy
End Function

Public Sub SetDebtor(ByVal Value As String)
    If Interval(Value, 1, 10) And regex(Value, "^\d+$") Then
        Debtor = Value
    Else
        MsgBox "Invalid Debtor"
    End If
End Sub

Public Function GetDebtor()
    GetDebtor = Debtor
End Function

Public Sub SetDevise(ByVal Value As String)
    If Interval(Value, 1, 3) Then
        Devise = Value
    Else
        MsgBox "Invalid Devise"
    End If
End Sub

Public Function GetDevise() As String
    GetDevise = Devise
End Function

'Public Sub InitItem(ByVal ID As String, ByVal Doc As String, ByVal LogicalSystem As String, ByVal BTType As String, ByVal AcctType As String, ByVal RefID As String, ByVal HeaderText As String, ByVal CreatedBy As String, ByVal Company As String, ByVal DocDate As String, ByVal TaxDate As String, ByVal Ref1 As String, ByVal Ref2 As String, ByVal Acc As String, ByVal Amount As String, ByVal DebitCredit As String, ByVal DocItemText As String, ByVal TaxCode As String, ByVal ProfitCenter As String, ByVal CreditItemRefItem As String, ByVal CreditItemAmount As String, ByVal CreditItemDebitCredit As String, ByVal ProdTaxCode As String, ByVal ProdTaxClass As String, ByVal ProdTaxAmount As String, ByVal ProdTaxDebitCredit As String, ByVal ProdTaxBaseAmount As String, ByVal Debtor As String, ByVal Devise As String)
    'Me.SetId ID
    'Me.SetOriginalReferenceDocument Doc
    'Me.SetOriginalReferenceDocumentLogicalSystem LogicalSystem
    'Me.SetBusinessTransactionType BTType
    'Me.SetAccountingDocumentType AcctType
    'Me.SetDocumentReferenceID RefID
    'Me.SetDocumentHeaderText HeaderText
    'Me.SetCreatedByUser CreatedBy
    'Me.SetCompanyCode Company
    'Me.SetDocumentDate DocDate
    'Me.SetTaxDeterminationDate TaxDate
    'Me.SetReference1InDocumentHeader Ref1
    'Me.SetReference2InDocumentHeader Ref2
    'Me.SetGLAccount Acc
    'Me.SetItemAmountInTransactionCurrency Amount
    'Me.SetItemDebitCreditCode DebitCredit
    'Me.SetItemDocumentItemText DocItemText
    'Me.SetItemTaxCode TaxCode
    'Me.SetItemProfitCenter ProfitCenter
    'Me.SetCreditItemReferenceDocumentItem CreditItemRefItem
    'Me.SetCreditItemAmountInTransactionCurrency CreditItemAmount
    'Me.SetCreditItemDebitCreditCode CreditItemDebitCredit
    'Me.SetProductTaxItemTaxCode ProdTaxCode
    'Me.SetProductTaxItemTaxItemClassification ProdTaxClass
    'Me.SetProductTaxItemAmountInTransactionCurrency ProdTaxAmount
    'Me.SetProductTaxItemDebitCreditCode ProdTaxDebitCredit
    'Me.SetProductTaxItemTaxBaseAmountInTransCrcy ProdTaxBaseAmount
    'Me.SetDebtor Debtor
    'Me.SetDevise Devise
'End Sub

Public Sub DebugPrint()
    Debug.Print "ID: " & Me.GetId
    Debug.Print "OriginalReferenceDocument: " & Me.GetOriginalReferenceDocument
    Debug.Print "OriginalReferenceDocumentLogicalSystem: " & Me.GetOriginalReferenceDocumentLogicalSystem
    Debug.Print "BusinessTransactionType: " & Me.GetBusinessTransactionType
    Debug.Print "AccountingDocumentType: " & Me.GetAccountingDocumentType
    Debug.Print "DocumentReferenceID: " & Me.GetDocumentReferenceID
    Debug.Print "DocumentHeaderText: " & Me.GetDocumentHeaderText
    Debug.Print "CreatedByUser: " & Me.GetCreatedByUser
    Debug.Print "CompanyCode: " & Me.GetCompanyCode
    Debug.Print "DocumentDate: " & Me.GetDocumentDate
    Debug.Print "TaxDeterminationDate: " & Me.GetTaxDeterminationDate
    Debug.Print "Reference1InDocumentHeader: " & Me.GetReference1InDocumentHeader
    Debug.Print "Reference2InDocumentHeader: " & Me.GetReference2InDocumentHeader
    Debug.Print "GLAccount: " & Me.GetGLAccount
    Debug.Print "ItemAmountInTransactionCurrency: " & Me.GetItemAmountInTransactionCurrency
    Debug.Print "ItemDebitCreditCode: " & Me.GetItemDebitCreditCode
    Debug.Print "ItemDocumentItemText: " & Me.GetItemDocumentItemText
    'Debug.Print "ItemTaxCode: " & Me.GetItemTaxCode
    'Debug.Print "ItemProfitCenter: " & Me.GetItemProfitCenter
    Debug.Print "CreditItemReferenceDocumentItem: " & Me.GetCreditItemReferenceDocumentItem
    Debug.Print "CreditItemAmountInTransactionCurrency: " & Me.GetCreditItemAmountInTransactionCurrency
    Debug.Print "CreditItemDebitCreditCode: " & Me.GetCreditItemDebitCreditCode
    Debug.Print "ProductTaxItemTaxCode: " & Me.GetProductTaxItemTaxCode
    Debug.Print "ProductTaxItemTaxItemClassification: " & Me.GetProductTaxItemTaxItemClassification
    Debug.Print "ProductTaxItemAmountInTransactionCurrency: " & Me.GetProductTaxItemAmountInTransactionCurrency
    Debug.Print "ProductTaxItemDebitCreditCode: " & Me.GetProductTaxItemDebitCreditCode
    Debug.Print "ProductTaxItemTaxBaseAmountInTransCrcy: " & Me.GetProductTaxItemTaxBaseAmountInTransCrcy
    Debug.Print "Debtor: " & Me.GetDebtor
    Debug.Print "Devise: " & Me.GetDevise
End Sub

' Exemple
' Item.InitItem "1234", "Doc", "XYZ", "RFBU", "RN", "RefID", "htext", "U12345", "1000", "2023-12-12", "2023-12-12", "Ref1", "Ref2", "467007", "10,00", "H", "DocItemText", "A0", "CP1000", "1", "10,00", "H", "A0", "NVP", "0,00", "H", "10,00", "70100000", "EUR"
