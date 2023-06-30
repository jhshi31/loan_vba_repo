Option Explicit

Attribute VB_Name = "CalculatePayments"



Private Function PaymentsChangeThisMo(principal_eom_cell As Range) As Double
    Dim principal_som As Double
    Dim principal_eom As Double
    Dim principal_payment As Double
    Dim fixed_rate_principal_payment As Double
    Dim fixed_rate_principal_eom As Double
    Dim remaining_payments As Double
    Dim fixed_rate_remaining_payments As Double

    Application.Volatile

    principal_eom = principal_eom_cell.Value
    principal_payment = principal_eom_cell.Offset(0,2).Value
    principal_som = principal_eom + principal_payment

    fixed_rate_principal_payment = fixed_rate_payment - principal_som * monthly_interest_rate
    fixed_rate_principal_eom = principal_som - fixed_rate_principal_payment

    remaining_payments = WorksheetFunction.NPer(monthly_interest_rate, -fixed_rate_payment, principal_eom)
    fixed_rate_remaining_payments = WorksheetFunction.NPer(monthly_interest_rate, -fixed_rate_payment, fixed_rate_principal_eom)

    PaymentsChangeThisMo = fixed_rate_remaining_payments - remaining_payments
End Function



Function PaymentsChangeThisYear() As Double
    Dim curWs As Worksheet
    Dim cells As Range
    Dim cell As Range
    Dim mortgage_payment As Double
    Dim payments_change As Double

    Application.Volatile

    InitializeConstants
    Set curWs = Application.Caller.Worksheet

    ' principal eom cells
    Set cells = curWs.Range("B9:B20")
    For Each cell In cells:
        mortgage_payment = cell.Offset(0,3).Value
        If mortgage_payment > fixed_rate_payment Then
            payments_change = payments_change + PaymentsChangeThisMo(cell)
        End If
    Next cell

    PaymentsChangeThisYear = WorksheetFunction.RoundDown(payments_change,0)
End Function

Function PaymentsChange() As Double
    Dim curWs As Worksheet
    Dim curWb As Workbook
    Dim sheet As Worksheet
    Dim payments_change As Double

    Application.Volatile

    Set curWs = Application.Caller.Worksheet
    Set curWb = curWs.Parent

    payments_change = 0
    For Each sheet In curWb.Sheets
        If sheet.Name <> "Info" And sheet.Name <> "Analysis" Then
            payments_change = payments_change + sheet.Range("I4").Value
        End If
    Next sheet

    PaymentsChange = payments_change
End Function