# VBA-code-to-display-the-amortization-table-of-a-bond
You can calculate clean price at all dates, sensitivity, convexity and duration

Option Explicit

 

Sub bond_with_coupon()

 

Dim amount As Single, interest_rate As Double, interest_payment_frequency As Byte

Dim interest_rate_needed As Double, length_bond As Byte, issuing_date As Date

Dim i As Integer, interest_date_payment As Date

Dim Number_of_interest_payment As Integer, last_row As Integer

Dim number_of_year As Double, yield_to_maturity As Double

Dim number_of_securities_issued As Long, y As Integer

Dim bond_brice_at_issuing_date As Double, real_price As Double

 

'I ask for all needed data to the user

 

amount = InputBox("What is the total amount borrowed ?")

    While amount < 0

    amount = InputBox("What is the total amount of the issued bond ?")

    Wend

 

interest_rate = InputBox("What is the annual interest rate ? (write it 0,0.. without the %)")

 

interest_payment_frequency = InputBox("How many time(s) will the issuer pay interests in a year? 1, 2, or 4")

    While interest_payment_frequency <> 1 And interest_payment_frequency <> 2 And interest_payment_frequency <> 4

    interest_payment_frequency = InputBox("How many time(s) will the issuer pay interests in a year? 1, 2, or 4")

    Wend

 

interest_rate_needed = interest_rate / interest_payment_frequency

 

length_bond = InputBox("What is the total time of the bond? (in year(s))")

    While length_bond <= 0

    length_bond = InputBox("What is the total time of the bond? (in year(s))")

    Wend

 

issuing_date = InputBox("What is the issuing date?")

 

yield_to_maturity = InputBox("What is the yield to maturity?")

 

number_of_securities_issued = InputBox("How many securities did the company issue?")

 

number_of_year = 0

 

bond_brice_at_issuing_date = amount / number_of_securities_issued

 

'Writing all headers

 

Range("A1") = "Date"

Range("B1") = "Coupon"

Range("C1") = "Number of year"

Range("D1") = "plus 1"

Range("E1") = "Factor for convexity"

Range("F1") = "Coupon FV"

Range("G1") = "Coupon FV per security"

Range("H1") = "Amortization"

Range("I1") = "Amortization FV"

Range("J1") = "Amortization FV per security"

Range("K1") = "FV CF per security"

Range("L1") = "Clean price"

Range("M1") = "In %"

Range("N1") = "Real price (market data in %)"

Range("O1") = "Under or over-priced"

 

Number_of_interest_payment = interest_payment_frequency * length_bond

last_row = Number_of_interest_payment + 1

 

'counting variable for all interest payment dates

 

interest_date_payment = issuing_date

 

'Writing the interest amount in the second column

 

For i = 2 To last_row

    Cells(i, 2) = amount * interest_rate_needed

Next i

 

'filling the columns date, coupon, number of year,plus 1, factor for convexity, coupon FV, coupon FV per security

 

For i = 2 To last_row

    If interest_payment_frequency = 1 Then

        'column date

        interest_date_payment = DateAdd("yyyy", 1, interest_date_payment)

        Cells(i, 1) = interest_date_payment

        'column number of year

        number_of_year = number_of_year + 1

        Cells(i, 3) = number_of_year

        'column plus 1

        Cells(i, 4) = Cells(i, 3) + 1

        'column coupon FV

        Cells(i, 6) = Cells(i, 2) / (1 + yield_to_maturity / 4) ^ (4 * Cells(i, 3))

        'column coupon FV per securitiy

        Cells(i, 7) = Cells(i, 2) / number_of_securities_issued

    End If

    If interest_payment_frequency = 2 Then

        interest_date_payment = DateAdd("m", 6, interest_date_payment)

        Cells(i, 1) = interest_date_payment

        number_of_year = number_of_year + 0.5

        Cells(i, 3) = number_of_year

        Cells(i, 4) = Cells(i, 3) + 1

        Cells(i, 6) = Cells(i, 2) / (1 + yield_to_maturity / 2) ^ (2 * Cells(i, 3))

        Cells(i, 7) = Cells(i, 2) / number_of_securities_issued

    End If

    If interest_payment_frequency = 4 Then

        interest_date_payment = DateAdd("m", 3, interest_date_payment)

        Cells(i, 1) = interest_date_payment

        number_of_year = number_of_year + 0.25

        Cells(i, 3) = number_of_year

        Cells(i, 4) = Cells(i, 3) + 1

        Cells(i, 6) = Cells(i, 2) / (1 + yield_to_maturity) ^ (Cells(i, 3))

        Cells(i, 7) = Cells(i, 2) / number_of_securities_issued

    End If

Next i

 

'filling the amortization column

 

For i = 2 To (last_row - 1)

    Cells(i, 8) = 0

Next i

 

Cells(last_row, 8) = amount

 

'filling the amortization FV column and amortization FV per security, factor for convexity

 

For i = 2 To last_row

    If interest_payment_frequency = 1 Then

        Cells(i, 9) = Cells(i, 8) / (1 + yield_to_maturity) ^ Cells(i, 3)

    End If

    If interest_payment_frequency = 2 Then

        Cells(i, 9) = Cells(i, 8) / (1 + yield_to_maturity / 2) ^ (Cells(i, 3) * 2)

    End If

    If interest_payment_frequency = 4 Then

        Cells(i, 9) = Cells(i, 8) / (1 + yield_to_maturity / 4) ^ (Cells(i, 3) * 4)

    End If

    Cells(i, 10) = Cells(i, 9) / number_of_securities_issued

    Cells(i, 11) = Cells(i, 7) + Cells(i, 10)

    Cells(i, 5) = Cells(i, 3) * Cells(i, 4) * Cells(i, 11)

Next i

 

'calculating the clean price in amount and in %

 

For i = 2 To (last_row - 1)

    Cells(i, 12) = Application.WorksheetFunction.Sum(Range(Cells(i + 1, 11), Cells(last_row, 11)))

    Cells(i, 13) = Cells(i, 12) / bond_brice_at_issuing_date

Next i

Cells(last_row, 7) = 0

 

'writing all the data given by the user so he can check for any error

 

Cells(1, 17) = "Amount"

Cells(1, 18) = amount

Cells(2, 17) = "Issuing date"

Cells(2, 18) = issuing_date

Cells(3, 17) = "Length in year"

Cells(3, 18) = length_bond

Cells(4, 17) = "Yearly interest rate"

Cells(4, 18) = interest_rate

Cells(5, 17) = "Yield to maturity"

Cells(5, 18) = yield_to_maturity

Cells(6, 17) = "Number of securities issued"

Cells(6, 18) = number_of_securities_issued

 

'asking to the user the market data

 

real_price = 0

0

For i = 2 To last_row

    If Cells(i, 1) < Date Then

    real_price = InputBox("What was the real price as at " & Cells(i, 1) & " write it 0,0.. without the % (if you are not interested to know if the security is under or over-priced just put 0)")

        Cells(i, 14) = real_price

    End If

Next i

 

'use it to determine if the price is too high or low

 

For i = 2 To last_row

    If Cells(i, 13) < Cells(i, 14) Then

        Cells(i, 15) = "over-priced"

    ElseIf Cells(i, 13) > Cells(i, 14) And Cells(i, 14) <> 0 Then

        Cells(i, 15) = "under-priced"

    End If

Next i

 

'calculating the total

 

Cells(last_row + 2, 1) = "Total"

Cells(last_row + 2, 2) = Application.WorksheetFunction.Sum(Range(Cells(2, 2), Cells(last_row, 2)))

Cells(last_row + 2, 5) = Application.WorksheetFunction.Sum(Range(Cells(2, 5), Cells(last_row, 5)))

Cells(last_row + 2, 6) = Application.WorksheetFunction.Sum(Range(Cells(2, 6), Cells(last_row, 6)))

Cells(last_row + 2, 7) = Application.WorksheetFunction.Sum(Range(Cells(2, 7), Cells(last_row, 7)))

Cells(last_row + 2, 8) = Application.WorksheetFunction.Sum(Range(Cells(2, 8), Cells(last_row, 8)))

Cells(last_row + 2, 9) = Application.WorksheetFunction.Sum(Range(Cells(2, 9), Cells(last_row, 9)))

Cells(last_row + 2, 10) = Application.WorksheetFunction.Sum(Range(Cells(2, 10), Cells(last_row, 10)))

Cells(last_row + 2, 11) = Application.WorksheetFunction.Sum(Range(Cells(2, 11), Cells(last_row, 11)))

 

'calculating duration, sensibility and convexity

 

Cells(last_row + 5, 1) = "Macaulay duration"

Cells(last_row + 5, 2) = (Application.WorksheetFunction.SumProduct(Range(Cells(2, 3), Cells(last_row, 3)), Range(Cells(2, 11), Cells(last_row, 11)))) / bond_brice_at_issuing_date

Cells(last_row + 6, 1) = "Sensitivity"

Cells(last_row + 6, 2) = Cells(last_row + 5, 2) / (1 + yield_to_maturity)

Cells(last_row + 6, 3) = "If the rate increases by 0,01, the bond drops by " & Cells(last_row + 6, 2) & "%"

Cells(last_row + 7, 1) = "Convexity"

Cells(last_row + 7, 2) = Cells(last_row + 2, 5) / Cells(last_row + 2, 11)

 

End Sub