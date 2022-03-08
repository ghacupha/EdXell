Option Explicit

Function logit(y As Range, xraw As Range, Optional constant, Optional stats)
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual


If IsMissing(constant) Then constant = 1
If IsMissing(stats) Then stats = 0
'Count variables
Dim i As Long, j As Long, jj As Long


'Read data dimensions
Dim K As Long, N As Long
N = y.Rows.Count
K = xraw.Columns.Count + constant

'Some error checking
If xraw.Rows.Count <> N Then MsgBox "error"


'Adding a vector of ones to the x matrix if constant=1, name xraw=x from now on
Dim x() As Double
ReDim x(1 To N, 1 To K)
For i = 1 To N
    x(i, 1) = 1
    For j = 1 + constant To K
        x(i, j) = xraw(i, j - constant)
    Next j
Next i

  
'Initializing the coefficient vector (b) and the score (bx)
Dim b() As Double, bx() As Double, ybar As Double
ReDim b(1 To K)
ReDim bx(1 To N)

ybar = Application.WorksheetFunction.Average(y)
If constant = 1 Then b(1) = Log(ybar / (1 - ybar))
For i = 1 To N
      bx(i) = b(1)
Next i



'Defining the variables used in the Newton procedure
Dim sens As Double, maxiter As Integer, iter As Integer, change As Double
Dim lambda() As Double, lnL() As Double, dlnL() As Double, hesse() As Double, hinv(), hinvg()
ReDim lambda(1 To N)

sens = 1 * 10 ^ (-11): maxiter = 50
ReDim lnL(1 To maxiter)
change = sens + 1: iter = 1: lnL(1) = 0

'Loop for Newton iteration
Do While Abs(change) > sens And iter < maxiter
    iter = iter + 1
    
    'reset derivative of log likelihood and Hessian
    Erase dlnL, hesse
    ReDim dlnL(1 To K): ReDim hesse(1 To K, 1 To K)

    'Compute prediction Lambda, gradient dlnl, Hessian hesse, and log likelihood lnl
    For i = 1 To N
        lambda(i) = 1 / (1 + Exp(-bx(i)))
        For j = 1 To K
            dlnL(j) = dlnL(j) + (y(i) - lambda(i)) * x(i, j)
            For jj = 1 To K
                hesse(jj, j) = hesse(jj, j) - lambda(i) * (1 - lambda(i)) * x(i, jj) * x(i, j)
            Next jj
        Next j
        lnL(iter) = lnL(iter) + y(i) * Log(1 / (1 + Exp(-bx(i)))) + (1 - y(i)) * Log(1 - 1 / (1 + Exp(-bx(i))))
    Next i
 
    'Compute inverse Hessian (=hinv) and multiply hinv with gradient dlnl
    hinv = Application.WorksheetFunction.MInverse(hesse)
    hinvg = Application.WorksheetFunction.MMult(dlnL, hinv)
    
    change = lnL(iter) - lnL(iter - 1)
      
   'If convergence achieved, exit now and keep the b corresponding with the estimated hessian
   If Abs(change) <= sens Then Exit Do
   
  ' Apply Newton's scheme for updating coefficients b
    For j = 1 To K
        b(j) = b(j) - hinvg(j)
    Next j



    'Compute new score (bx)
    For i = 1 To N
        bx(i) = 0
        For j = 1 To K
            bx(i) = bx(i) + b(j) * x(i, j)
        Next j
    Next i

Loop


'some error handling
If iter > maxiter Then
 MsgBox "Maximum Number of Iteration exceeded. No convergence achieved. Exiting. Sorry."
GoTo myend
End If
 

'output
Dim relogit()
ReDim relogit(1 To 1, 1 To K)
If stats = 1 Then ReDim relogit(1 To 7, 1 To K)

'Coefficients
For j = 1 To K
 relogit(1, j) = b(j)
Next j

'Additional statistics if requested
If stats = 1 Then
  For j = 1 To K
   relogit(2, j) = Sqr(-hinv(j, j))
   relogit(3, j) = relogit(1, j) / relogit(2, j)
   relogit(4, j) = (1 - Application.WorksheetFunction.NormSDist(Abs(relogit(3, j)))) * 2
   
   relogit(5, j) = "#N/A"
   relogit(6, j) = "#N/A"
   relogit(7, j) = "#N/A"
 
  Next j
 
 'ln Likelihood of model with just a constant(lnL0)
 Dim lnL0 As Double
 lnL0 = N * (ybar * Log(ybar) + (1 - ybar) * Log(1 - ybar))


 relogit(5, 1) = 1 - lnL(iter) / lnL0      'McFadden R2
 relogit(5, 2) = iter - 1           'Number of iterations
 relogit(6, 1) = 2 * (lnL(iter) - lnL0)   'LR test
 relogit(6, 2) = Application.WorksheetFunction.ChiDist(relogit(6, 1), K - 1) 'p-value for LR
 relogit(7, 1) = lnL(iter)
 relogit(7, 2) = lnL0
 
End If
 logit = relogit

GoTo myend

'Error Handler
error:
MsgBox ("Fatal Error. Reasons might be: y not {0,1}, not the same number of N for y and x's...or anything else")
myend:
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True

End Function

Function XTRANS(defaultdata As Range, x As Range, numranges As Integer)
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Dim bound, numdefaults, obs, defrate, N, j, defsum, obssum, i
ReDim bound(1 To numranges), numdefaults(1 To numranges)
ReDim obs(1 To numranges), defrate(1 To numranges)

N = x.Rows.Count

'Determining number of defaults, observations and default rates for ranges
For j = 1 To numranges
    
    bound(j) = Application.WorksheetFunction.Percentile(x, j / numranges)
    
    numdefaults(j) = Application.WorksheetFunction.SumIf(x, "<=" & bound(j), defaultdata) - defsum
    defsum = defsum + numdefaults(j)

    obs(j) = Application.WorksheetFunction.CountIf(x, "<=" & bound(j)) - obssum
    obssum = obssum + obs(j)
    
    defrate(j) = numdefaults(j) / obs(j)
Next j

'Assigning range default rates in logistic transformation
Dim transform
ReDim transform(1 To N, 1 To 1)

For i = 1 To N
    j = 1
    While x(i) - bound(j) > 0
        j = j + 1
    Wend
    transform(i, 1) = Application.WorksheetFunction.Max(defrate(j), 0.0000001)
    transform(i, 1) = Log(transform(i, 1) / (1 - transform(i, 1)))
Next i

XTRANS = transform
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True

End Function
Function WINSOR(x As Range, level As Double)
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

Dim N As Integer, i As Integer
N = x.Rows.Count

'Obtain percentiles
Dim low, up
low = Application.WorksheetFunction.Percentile(x, level)
up = Application.WorksheetFunction.Percentile(x, 1 - level)

'Pull x to percentiles
Dim result
ReDim result(1 To N, 1 To 1)
For i = 1 To N
    result(i, 1) = Application.WorksheetFunction.Max(x(i), low)
    result(i, 1) = Application.WorksheetFunction.Min(result(i, 1), up)
Next i

WINSOR = result

Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True

End Function



