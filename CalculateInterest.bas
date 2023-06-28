Option Explicit

Attribute VB_Name = "CalculateInterest"



Private Function InterestDecreaseThisMo(mortgage_payment_cell As Range) As Double
    '
    ' end of month
    Dim principal_eom
    ' start of month
    Dim principal_som
    Dim principal_payment
    Dim fixed_rate_principal_eom
    Dim fixed_rate_principal_payment
    Dim cost_of_loan
    Dim fixed_rate_cost_of_loan

    Application.Volatile

    principal_eom = mortgage_payment_cell.Offset(0, -3).Value
    principal_payment = mortgage_payment_cell.Offset(0, -1).Value
    principal_som = principal_eom + principal_payment
    fixed_rate_principal_payment = fixed_rate_payment - (principal_som * monthly_interest_rate)
    fixed_rate_principal_eom = principal_som - fixed_rate_principal_payment

    fixed_rate_cost_of_loan = WorksheetFunction.NPer(monthly_interest_rate, -fixed_rate_payment, _
                                fixed_rate_principal_eom) * fixed_rate_payment
    cost_of_loan = WorksheetFunction.NPer(monthly_interest_rate, -fixed_rate_payment, principal_eom) * fixed_rate_payment

    InterestDecreaseThisMo = fixed_rate_cost_of_loan - cost_of_loan
End Function



Function InterestDecreaseThisYear() As Double
    '
    Dim mortgage_payment_cells
    Dim cell As Range
    Dim interest_decrease As Double
    Dim callCell As Range
    Dim callWs As Worksheet

    Application.Volatile

    InitializeConstants
    Set callCell = Application.Caller
    Set callWs = callCell.Worksheet

    Set mortgage_payment_cells = callWs.Range("E9:E20")
    interest_decrease = 0
    For Each cell In mortgage_payment_cells
        If cell.Value > fixed_rate_payment Then
            interest_decrease = interest_decrease + InterestDecreaseThisMo(cell)
        End If
    Next cell

    InterestDecreaseThisYear = WorksheetFunction.Round(interest_decrease, 2)
End Function



Function CostOfLoanSoFar() as Double
    '
    Dim ws As Worksheet
    Dim cost As Double

    Application.Volatile
    InitializeConstants

    cost = 0
    For Each ws In ThisWorkbook.Sheets
        If ws.Name <> "Info" And ws.Name <> "Analysis" Then 
            cost = cost + ws.Range("E24").Value()
        End If
    Next ws

    CostOfLoanSoFar = cost
End Function