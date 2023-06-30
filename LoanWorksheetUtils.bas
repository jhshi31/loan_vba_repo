Option Explicit

Attribute VB_Name = "LoanWorksheetUtils"

Public fixed_rate_payment As Double
Public monthly_interest_rate As Double
Public num_installments As Integer

Sub InitializeConstants()
    Dim infoWs As Worksheet
    
    Set infoWs = ThisWorkbook.Sheets("Info")
    fixed_rate_payment = infoWs.Range("C8").Value
    monthly_interest_rate = infoWs.Range("C14").Value
    num_installments = infoWs.Range("C15").Value
End Sub

Sub LoadModules()
    ImportModule("CalculateInterest")
    ImportModule("CalculatePayments")
End Sub