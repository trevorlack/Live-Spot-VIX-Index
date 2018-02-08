Attribute VB_Name = "Near"
'Author- Trevor Lack

Option Base 1
Function Mat(Ndays, CalcTime, Contract)
    Dim M_SettleDay As Integer
    If Contract = "Weekly" Then
        M_SettleDay = 900
    Else: M_SettleDay = 510
    End If
    H = Hour(CalcTime)
    M = Minute(CalcTime)
    T = H + M / 60
    Mat = ((24 - T) * 60 + M_SettleDay + 60 * 24 * Ndays) / 525600
End Function
Function F(RiskFree, Time)
Ndays = Range("D6").Value
Contract = Range("J8").Value

Dim LastRowCall As Integer
    LastRowCall = Range("C" & Rows.Count).End(xlUp).Row
Dim LastRowPut As Integer
    LastRowPut = Range("L" & Rows.Count).End(xlUp).Row

Dim StrikeC As Variant
    StrikeC = Range("D17:D" & LastRowCall)
Dim StrikeP As Variant
    StrikeP = Range("L17:L" & LastRowPut)

Dim CallMids As Variant
    CallMids = Range("G17:G" & LastRowCall)
Dim PutMids As Variant
    PutMids = Range("O17:O" & LastRowPut)
    
NC = Application.Count(StrikeC)
NP = Application.Count(StrikeP)

T = Mat(Ndays, Time, Contract)
k = Worksheets("VIX").Range("I12").Value

For i = 1 To NC
    If StrikeC(i, 1) = k Then
    CallOp = CallMids(i, 1)
    End If
Next i

For i = 1 To NP
    If StrikeP(i, 1) = k Then
    PutOp = PutMids(i, 1)
    End If
Next i

F = k + Exp(RiskFree * T) * Abs(CallOp - PutOp)

End Function

Function K0(RiskFree, Time)

Contract = Range("J8").Value
Dim LastRowCall As Integer
    LastRowCall = Range("C" & Rows.Count).End(xlUp).Row

Dim StrikeC As Variant
    StrikeC = Range("D17:D" & LastRowCall)

Ndays = Range("D6").Value

N = Application.Count(StrikeC)
T = Mat(Ndays, Time, Contract)
Fi = F(RiskFree, Time)

K0 = StrikeC(1, 1)
Diff = Fi - StrikeC(1, 1)

For i = 2 To N
    If (Fi - StrikeC(i, 1) < Diff) And (Fi - StrikeC(i, 1)) > 0 Then
        K0 = StrikeC(i, 1)
    End If
Next i

End Function

Function NearTermVar(Time)

Dim RiskFree As Double
Dim Ndays As Double
Dim Ko As Double
Dim F As Double
RiskFree = Range("J6").Value
Ndays = Range("D6").Value
Ko = Range("D9").Value
F = Range("D8").Value
Contract = Range("J8").Value

Dim LastRowCall As Integer
    LastRowCall = Range("C" & Rows.Count).End(xlUp).Row
Dim LastRowPut As Integer
    LastRowPut = Range("K" & Rows.Count).End(xlUp).Row
    
Dim StrikeC As Variant
    StrikeC = Range("D17:G" & LastRowCall)
Dim StrikeP As Variant
    StrikeP = Range("L17:O" & LastRowPut)

T = Mat(Ndays, Time, Contract)

''''''''''''''''
'Call Collection
''''''''''''''''
Dim N As Integer
N = UBound(StrikeC, 1)
Dim j As Integer
j = 1
Dim jj As Integer
jj = Range("D13").Value
Dim VIXCalls() As Variant
ReDim VIXCalls(1 To jj, 1 To 2) As Variant

For i = 1 To N
    
    'End Array Construction after 2 zero bids
    If StrikeC(i, 4) = "Kill" Then
        GoTo KillStop:
        Else
        'Skip Call Strikes at or below Ko or with zero bid
        If StrikeC(i, 4) = "Omit" Or StrikeC(i, 1) < Ko Or StrikeC(i, 1) = Ko Then
        GoTo OmitSkip:
        Else
            'Output of calls is in decending order [Strike, Bid-Ask Mid-Point]
            VIXCalls(j, 1) = StrikeC(i, 1)
            VIXCalls(j, 2) = StrikeC(i, 4)
            j = j + 1
        End If
    
    End If
    
OmitSkip:
    
Next i

KillStop:

j = j - 1
'Construct the Call Contribution matrix for this term
Dim VIXCallContribution() As Variant
ReDim VIXCallContribution(1 To j) As Variant

For i = 1 To j
    If i = 1 Then
        VIXCallContribution(i) = (((VIXCalls(i + 1, 1) - Ko) / 2) / (VIXCalls(i, 1) ^ 2)) * Exp(RiskFree * T) * VIXCalls(i, 2)
    Else
        If i = j Then
            VIXCallContribution(i) = ((VIXCalls(i, 1) - VIXCalls(i - 1, 1)) / (VIXCalls(i, 1) ^ 2)) * Exp(RiskFree * T) * VIXCalls(i, 2)
        Else
            VIXCallContribution(i) = (((VIXCalls(i + 1, 1) - VIXCalls(i - 1, 1)) / 2) / (VIXCalls(i, 1) ^ 2)) * Exp(RiskFree * T) * VIXCalls(i, 2)
        End If
    End If
Next i

''''''''''''''''
'Put Collection
''''''''''''''''
Dim NN As Integer
NN = UBound(StrikeP, 1)

Dim k As Integer
k = 1
Dim kk As Integer
kk = Range("L13").Value
Dim VIXPuts() As Variant
ReDim VIXPuts(1 To kk, 1 To 2) As Variant

For i = 1 To N
    
    'End Array Construction after 2 zero bids
    If StrikeP(i, 4) = "Kill" Then
        GoTo KillStop2:
        Else
        'Skip Put Strikes at or above Ko or with zero bid
        If StrikeP(i, 4) = "Omit" Or StrikeP(i, 1) > Ko Or StrikeP(i, 1) = Ko Then
        GoTo OmitSkip2:
        Else
            'Output of Puts is in decending order [Strike, Bid-Ask Mid-Point]
            VIXPuts(k, 1) = StrikeP(i, 1)
            VIXPuts(k, 2) = StrikeP(i, 4)
            k = k + 1
        End If
    
    End If
    
OmitSkip2:
    
Next i

KillStop2:

k = k - 1
'Construct the Put Contribution matrix for this term
Dim VIXPutContribution() As Variant
ReDim VIXPutContribution(1 To k) As Variant

For i = 1 To k
    If i = 1 Then
        VIXPutContribution(i) = (((Ko - VIXPuts(i + 1, 1)) / 2) / (VIXPuts(i, 1) ^ 2)) * Exp(RiskFree * T) * VIXPuts(i, 2)
    Else
        If i = k Then
            VIXPutContribution(i) = ((VIXPuts(i - 1, 1) - VIXPuts(i, 1)) / (VIXPuts(i, 1) ^ 2)) * Exp(RiskFree * T) * VIXPuts(i, 2)
        Else
            VIXPutContribution(i) = (((VIXPuts(i - 1, 1) - VIXPuts(i + 1, 1)) / 2) / (VIXPuts(i, 1) ^ 2)) * Exp(RiskFree * T) * VIXPuts(i, 2)
        End If
    End If
Next i

Dim KzeroCMid As Double
Dim KzeroPMid As Double
Dim Kzero As Double

For i = 1 To N
    If StrikeC(i, 1) = Ko Then
        KzeroCMid = StrikeC(i, 4)
    End If
Next i
For i = 1 To NN
    If StrikeP(i, 1) = Ko Then
        KzeroPMid = StrikeP(i, 4)
    End If
Next i
        
Kzero = (((VIXCalls(1, 1) - VIXPuts(1, 1)) / 2) / (Ko ^ 2)) * Exp(RiskFree * T) * ((KzeroCMid + KzeroPMid) / 2)

'Dim NearTermVar As Long
NearTermVar = 2 / T * (Application.WorksheetFunction.Sum(VIXCallContribution()) + Application.WorksheetFunction.Sum(VIXPutContribution()) + Kzero) - (F / Ko - 1) ^ 2 / T

End Function

