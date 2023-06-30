Option Explicit

Attribute VB_Name = "CalculateInterest"



Private Function InterestDecreaseThisMo(mortgage_payment_cell As Range) As Double
    '
    ' end of month
    Dim principal_eom
    ' start of month
    Dim principal_som As Double
    Dim principal_payment As Double
    Dim interest_payment As Double
    Dim fixed_rate_principal_payment As Double
    Dim fixed_rate_principal_eom As Double
    Dim cost_of_loan As Double
    Dim fixed_rate_cost_of_loan As Double

    Application.Volatile

    principal_eom = mortgage_payment_cell.Offset(0, -3).Value
    principal_payment = mortgage_payment_cell.Offset(0, -1).Value
    principal_som = principal_eom + principal_payment
    interest_payment = mortgage_payment_cell.Offset(0, -2).Value
    fixed_rate_principal_payment = fixed_rate_payment - interest_payment
    fixed_rate_principal_eom = principal_som - fixed_rate_principal_payment

    fixed_rate_cost_of_loan = fixed_rate_principal_payment + WorksheetFunction.NPer(monthly_interest_rate, -fixed_rate_payment, _
                                fixed_rate_principal_eom) * fixed_rate_payment
    cost_of_loan = principal_payment + WorksheetFunction.NPer(monthly_interest_rate, -fixed_rate_payment, principal_eom) * fixed_rate_payment

    InterestDecreaseThisMo = fixed_rate_cost_of_loan - cost_of_loan
End Function



Function InterestDecreaseThisYear() As Double
    Dim curWs As Worksheet
    Dim curWb As Workbook
    Dim mortgage_payment_cells As Range
    Dim cell As Range
    Dim interest_decrease As Double    

    Application.Volatile

    InitializeConstants
    Set curWs = Application.Caller.Worksheet
    Set curWb = curWs.Parent

    Set mortgage_payment_cells = curWs.Range("E9:E20")
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
    Dim curWs As Worksheet
    Dim curWb As Workbook
    Dim ws As Worksheet
    Dim cost As Double

    Application.Volatile

    Set curWs = Application.Caller.Worksheet
    Set curWb = curWs.Parent
    cost = 0
    For Each ws In ThisWorkbook.Sheets
        If ws.Name <> "Info" And ws.Name <> "Analysis" Then 
            cost = cost + ws.Range("E24").Value()
        End If
    Next ws

    CostOfLoanSoFar = cost
End Function