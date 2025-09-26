Attribute VB_Name = "IVOL"
Global Const Pi = 3.14159265358979
Option Explicit


Public Function GBlackScholesImpVolBisection(CallPutFlag As String, S As Double, _
                x As Double, T As Double, r As Double, b As Double, cm As Double) As Variant

    Dim vLow As Double, vHigh As Double, vi As Double
    Dim cLow As Double, cHigh As Double, epsilon As Double
    Dim counter As Integer
    
    vLow = 0.005
    vHigh = 4
    epsilon = 0.00000001
    cLow = GBlackScholes(CallPutFlag, S, x, T, r, b, vLow)
    cHigh = GBlackScholes(CallPutFlag, S, x, T, r, b, vHigh)
    counter = 0
    vi = vLow + (cm - cLow) * (vHigh - vLow) / (cHigh - cLow)
    While Abs(cm - GBlackScholes(CallPutFlag, S, x, T, r, b, vi)) > epsilon
        counter = counter + 1
        If counter = 10000 Then
            GBlackScholesImpVolBisection = "NA"
            Exit Function
        End If
        If GBlackScholes(CallPutFlag, S, x, T, r, b, vi) < cm Then
            vLow = vi
        Else
            vHigh = vi
        End If
        cLow = GBlackScholes(CallPutFlag, S, x, T, r, b, vLow)
        cHigh = GBlackScholes(CallPutFlag, S, x, T, r, b, vHigh)
        vi = vLow + (cm - cLow) * (vHigh - vLow) / (cHigh - cLow)
    Wend
    GBlackScholesImpVolBisection = vi
    
End Function

