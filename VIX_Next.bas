Attribute VB_Name = "Next1"
'Author- Trevor Lack

Option Base 1

Function Fnext1(RiskFree1, Time)
    Application.Volatile
Ndays = Range("V6").Value
Contract = Range("AB8").Value
Dim LastRowCall As Integer
    LastRowCall = Range("U" & Rows.Count).End(xlUp).Row
Dim LastRowPut As Integer
    LastRowPut = Range("AC" & Rows.Count).End(xlUp).Row

Dim StrikeC1 As Variant
    StrikeC1 = Range("V17:V" & LastRowCall)
Dim StrikeP1 As Variant
    StrikeP1 = Range("AD17:AD" & LastRowPut)

Dim CallMids As Variant
    CallMids = Range("Y17:Y" & LastRowCall)
Dim PutMids As Variant
    PutMids = Range("AG17:AG" & LastRowPut)

    
NC = Application.Count(StrikeC1)
NP = Application.Count(StrikeP1)

T = Mat(Ndays, Time, Contract)
k = Worksheets("VIX").Range("AA12").Value


For i = 1 To NC
    If StrikeC1(i, 1) = k Then
    CallOp = CallMids(i, 1)
    End If
Next i

For i = 1 To NP
    If StrikeP1(i, 1) = k Then
    PutOp = PutMids(i, 1)
    End If
Next i


Fnext1 = k + Exp(RiskFree * T) * Abs(CallOp - PutOp)
    

End Function

Function K0_1(RiskFree, Time)
    Application.Volatile
Contract = Range("AB8").Value
Dim LastRowCall As Integer
    LastRowCall = Range("U" & Rows.Count).End(xlUp).Row

Dim StrikeC1 As Variant
    StrikeC1 = Range("V17:V" & LastRowCall)

Ndays = Range("V6").Value

N = Application.Count(StrikeC1)
T = Mat(Ndays, Time, Contract)
Fi = F(RiskFree, Time)

K0_1 = StrikeC1(1, 1)
Diff = Fi - StrikeC1(1, 1)

For i = 2 To N
    If (Fi - StrikeC1(i, 1) < Diff) And (Fi - StrikeC1(i, 1)) > 0 Then
        K0_1 = StrikeC1(i, 1)
    End If
Next i

End Function

'Sub NearTermVar2()

Function NearTermVar1(Time)
    Application.Volatile
'Dim Time As String
'Time = Range("J7")
Dim RiskFree As Double
Dim Ndays As Double
Dim Ko As Double
Dim F As Double
RiskFree = Range("AB6").Value
Ndays = Range("V6").Value
Ko = Range("V9").Value
F = Range("V8").Value
Contract = Range("AB8").Value

Dim LastRowCall As Integer
    LastRowCall = Range("U" & Rows.Count).End(xlUp).Row
Dim LastRowPut As Integer
    LastRowPut = Range("AC" & Rows.Count).End(xlUp).Row
    
Dim StrikeC1 As Variant
    StrikeC1 = Range("V17:Y" & LastRowCall)
Dim StrikeP1 As Variant
    StrikeP1 = Range("AD17:AG" & LastRowPut)

T = Mat(Ndays, Time, Contract)

''''''''''''''''
'Call Collection
''''''''''''''''
Dim N As Integer
N = UBound(StrikeC1, 1)
Dim j As Integer
j = 1
Dim jj As Integer
jj = Range("V13").Value
Dim VIXCalls() As Variant
ReDim VIXCalls(1 To jj, 1 To 2) As Variant

For i = 1 To N
    
    'End Array Construction after 2 zero bids
    If StrikeC1(i, 4) = "Kill" Then
        GoTo KillStop:
        Else
        'Skip Call Strikes at or below Ko or with zero bid
        If StrikeC1(i, 4) = "Omit" Or StrikeC1(i, 1) < Ko Or StrikeC1(i, 1) = Ko Then
        GoTo OmitSkip:
        Else
            'Output of calls is in decending order [Strike, Bid-Ask Mid-Point]
            VIXCalls(j, 1) = StrikeC1(i, 1)
            VIXCalls(j, 2) = StrikeC1(i, 4)
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
NN = UBound(StrikeP1, 1)

Dim k As Integer
k = 1
Dim kk As Integer
kk = Range("AD13").Value
Dim VIXPuts() As Variant
ReDim VIXPuts(1 To kk, 1 To 2) As Variant

For i = 1 To N
    
    'End Array Construction after 2 zero bids
    If StrikeP1(i, 4) = "Kill" Then
        GoTo KillStop2:
        Else
        'Skip Put Strikes at or above Ko or with zero bid
        If StrikeP1(i, 4) = "Omit" Or StrikeP1(i, 1) > Ko Or StrikeP1(i, 1) = Ko Then
        GoTo OmitSkip2:
        Else
            'Output of Puts is in decending order [Strike, Bid-Ask Mid-Point]
            VIXPuts(k, 1) = StrikeP1(i, 1)
            VIXPuts(k, 2) = StrikeP1(i, 4)
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
    If StrikeC1(i, 1) = Ko Then
        KzeroCMid = StrikeC1(i, 4)
    End If
Next i
For i = 1 To NN
    If StrikeP1(i, 1) = Ko Then
        KzeroPMid = StrikeP1(i, 4)
    End If
Next i
        
Kzero = (((VIXCalls(1, 1) - VIXPuts(1, 1)) / 2) / (Ko ^ 2)) * Exp(RiskFree * T) * ((KzeroCMid + KzeroPMid) / 2)

'Dim NearTermVar As Long
NearTermVar1 = 2 / T * (Application.WorksheetFunction.Sum(VIXCallContribution()) + Application.WorksheetFunction.Sum(VIXPutContribution()) + Kzero) - (F / Ko - 1) ^ 2 / T

End Function


