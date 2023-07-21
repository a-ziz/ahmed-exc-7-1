Public Class frmFutureValue

    Private Sub btnCalculate_Click(sender As Object,
            e As EventArgs) Handles btnCalculate.Click

        Try
            If IsValidData() Then
                Dim monthlyInvestment As Decimal = CDec(txtMonthlyInvestment.Text)
                Dim yearlyInterestRate As Decimal = CDec(txtInterestRate.Text)
                Dim years As Integer = CInt(txtYears.Text)

                Dim monthlyInterestRate As Decimal = yearlyInterestRate / 12 / 100
                Dim months As Integer = years * 12

                Dim futureValue As Decimal = Me.FutureValue(monthlyInvestment, monthlyInterestRate, months)

                txtFutureValue.Text = FormatCurrency(futureValue)
            End If
            'Catch ex As InvalidCastException
            '    MessageBox.Show("Please check all entries." & " Use numeric values only.",
            '                    "Entry Error")
            'Catch ex As OverflowException
            '    MessageBox.Show("Overflow exception." & " Please enter smaller number.",
            '                    "Entry Error")
            '    Exit Sub
            'Catch ex As Exception
            '    MessageBox.Show(ex.Message & vbNewLine & ex.GetType.ToString & ex.StackTrace,
            '                    "Exception")
            '    Exit Sub
        Finally
            txtMonthlyInvestment.Select()
        End Try

    End Sub

    Private Function IsValidData() As Boolean
        ' Validate the Monthly Investment text box
        If Not IsPresent(txtMonthlyInvestment, "Monthly Investment") Then
            Return False
        End If
        If Not IsDecimal(txtMonthlyInvestment, "Monthly Investment") Then
            Return False
        End If
        If Not IsWithinRange(txtMonthlyInvestment, "Monthly Investment", 1, 1000) Then
            Return False
        End If
        ' Validate the Interest Rate text box
        If Not IsPresent(txtInterestRate, "Yearly Interest Rate") Then
            Return False
        End If
        If Not IsDecimal(txtInterestRate, "Yearly Interest Rate") Then
            Return False
        End If
        If Not IsWithinRange(txtInterestRate, "Yearly Interest Rate", 1, 15) Then
            Return False
        End If
        ' Validate the Years text box
        If Not IsPresent(txtYears, "Number of Years") Then
            Return False
        End If
        If Not IsInt32(txtYears, "Number of Years") Then
            Return False
        End If
        If Not IsWithinRange(txtYears, "Number of Years", 1, 50) Then
            Return False
        End If

        Return True
    End Function
    Private Function FutureValue(monthlyInvestment As Decimal,
            monthlyInterestRate As Decimal, ByVal months As Integer) _
            As Decimal
        For i As Integer = 1 To months
            FutureValue = (FutureValue + monthlyInvestment) *
                          (1 + monthlyInterestRate)
            'Throw New Exception
        Next
        Return FutureValue
    End Function

    Private Sub ClearFutureValue(sender As Object,
            e As EventArgs) Handles txtMonthlyInvestment.TextChanged,
            txtYears.TextChanged, txtInterestRate.TextChanged
        txtFutureValue.Text = ""
    End Sub

    Private Sub btnExit_Click(sender As Object,
            e As EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub

    Private Function IsPresent(textBox As TextBox, name As String) _
         As Boolean
        If textBox.Text = "" Then
            MessageBox.Show(name & " is a required field.", "Entry Error")
            textBox.Select()
            Return False
        Else
            Return True
        End If
    End Function

    Private Function IsDecimal(textBox As TextBox, name As String) _
            As Boolean
        Dim number As Decimal = 0
        If Decimal.TryParse(textBox.Text, number) Then
            Return True
        Else
            MessageBox.Show(name & " must be a decimal value.", "Entry Error")
            textBox.Select()
            textBox.SelectAll()
            Return False
        End If
    End Function

    Private Function IsInt32(textBox As TextBox, name As String) _
            As Boolean
        Dim number As Integer = 0
        If Int32.TryParse(textBox.Text, number) Then
            Return True
        Else
            MessageBox.Show(name & " must be an integer.", "Entry Error")
            textBox.Select()
            textBox.SelectAll()
            Return False
        End If
    End Function

    Private Function IsWithinRange(textBox As TextBox,
            name As String, min As Decimal,
            max As Decimal) As Boolean
        Dim number As Decimal = CDec(textBox.Text)
        If number < min OrElse number > max Then
            MessageBox.Show(name & " must be between " & min & " and " &
                max & ".", "Entry Error")
            textBox.Select()
            textBox.SelectAll()
            Return False
        Else
            Return True
        End If
    End Function

    Private Sub frmFutureValue_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class