Option Explicit

Attribute VB_Name = "LoanUtils"

Public fixed_rate_payment As Double
Public monthly_interest_rate As Double
Public num_installments As Integer
Public constantsExists As Boolean

Function InitializeConstants()
    If Not constantsExists Then
        fixed_rate_payment = ThisWorkbook.Sheets("Info").Range("C8").Value
        monthly_interest_rate = ThisWorkbook.Sheets("Info").Range("C14").Value
        num_installments = ThisWorkbook.Sheets("Info").Range("C15").Value
        constantsExists = True
    End If
End Function