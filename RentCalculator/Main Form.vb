' Name:         Rent Calculator
' Purpose:      Help calculate the money owed for rent after bills and shared expenses
' Programmer:   Eric Patrick

Option Strict On
Option Explicit On
Option Infer Off

Public Class frmMain
    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub

    Private Sub btnCalc_Click(sender As Object, e As EventArgs) Handles btnCalc.Click
        ' Declare variables as double for currency processing
        Dim dblInternet As Double
        Dim dblPower As Double
        Dim dblRent As Double
        Dim dblStorage As Double
        Dim dblTelevision As Double
        Dim dblWater As Double
        Dim dblExpenses1 As Double
        Dim dblExpenses2 As Double
        Dim dblTotal1 As Double
        Dim dblTotal2 As Double
        Dim dblHalf1 As Double
        Dim dblHalf2 As Double
        Dim dblTotalOwed As Double
        Dim strSummary As String

        ' Parse the input from the form
        Double.TryParse(txtExpenses1.Text, dblExpenses1)
        Double.TryParse(txtExpenses2.Text, dblExpenses2)
        Double.TryParse(txtInternet.Text, dblInternet)
        Double.TryParse(txtPower.Text, dblPower)
        Double.TryParse(txtRent.Text, dblRent)
        Double.TryParse(txtStorage.Text, dblStorage)
        Double.TryParse(txtTelevision.Text, dblTelevision)
        Double.TryParse(txtWater.Text, dblWater)

        ' Add up expenses for both parties
        dblTotal1 = dblExpenses1 + dblInternet + dblPower + dblStorage + dblTelevision + dblWater
        dblTotal2 = dblExpenses2 + dblRent
        dblHalf1 = dblTotal1 / 2
        dblHalf2 = dblTotal2 / 2
        dblTotalOwed = dblHalf2 - dblHalf1

        ' Display totals
        lblTotal1.Text = dblTotal1.ToString("C3")
        lblTotal2.Text = dblTotal2.ToString("C3")
        lblTotalOwed.Text = dblTotalOwed.ToString("C3")
        strSummary = "I spent " & dblTotal1.ToString("C2") & " total on:" & Environment.NewLine _
            & "- Internet: " & dblInternet.ToString("C2") & Environment.NewLine _
            & "- Power: " & dblPower.ToString("C2") & Environment.NewLine _
            & "- Storage: " & dblStorage.ToString("C2") & Environment.NewLine _
            & "- Television: " & dblTelevision.ToString("C2") & Environment.NewLine _
            & "- Water: " & dblWater.ToString("C2") & Environment.NewLine _
            & "- Expenses: " & dblExpenses1.ToString("C2") & Environment.NewLine & Environment.NewLine _
            & "You spent " & dblTotal2.ToString("C2") & " total on: " & Environment.NewLine _
            & "- Rent: " & dblRent.ToString("C2") & Environment.NewLine _
            & "- Expenses: " & dblExpenses2.ToString("C2") & Environment.NewLine & Environment.NewLine _
            & "Bills & expenses cost " & dblHalf1.ToString("C3") & " each, while rent & expenses cost " & dblHalf2.ToString("C3") & " each. The difference is " & dblTotalOwed.ToString("C3") & "."
        txtSummary.Text = strSummary
    End Sub
End Class
