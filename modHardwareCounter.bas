Attribute VB_Name = "modHardwareCounter"

Option Explicit

' ===========================================================================
' System Hardware Counter/Timer Module     Jan-2011, Randy Manning.
' ===========================================================================
'
'--About high resolution timing--
'Especially when having a go at Programming Challenges, it can be
'vital to know how long a function or an algorithm takes to run.
'
'C standard library provides the clock() function, which can be used
'to measure the elapsed time. It is a system independent C function
'declared in time.h (compatible on most operating systems), but it
'does not give accurate results, not even milli-second accuracy.
'
'The good news is that in the Intel and AMD CPUs there is a built-
'in high speed hardware counter. It increments at a rate of
'millions per second.
'
'The bad news is that the functions which can access the CPU high
'speed hardware counter are system specific. In other words, you
'have to write different codes on the different systems. Windows
'provides QueryPerformanceCounter() function, and Unix, Linux and
'Mac OS X systems have gettimeofday(), which is declared in
'sys/time.h. Both functions can measure at least 1 micro-second
'(uS) differences.
' ===========================================================================


' ===========================================================================
' The VB Currency data-type does not reference money as used here.
' ===========================================================================
'
' The VB Currency data-type is used here because it has the same
' number of bytes as the 'C' LARGE_INTEGER data-type which is
' normally used with the Performance Counter API's declared below.
' The VB Currency data-type occupies two VB Longs worth of data.
'
' Also, the Currency data-type can hold values ranging from:
' -922,337,203,685,477.5808 to 922,337,203,685,477.5807
' representing 19 significant figures 'normally' interpereted
' with 4 decimal places.
'
' But here, we must interpret the 19 digits as an integer ranging
' from: -9,223,372,036,854,775,808 to 9,223,372,036,854,775,807
' effectively ignoring the presence of the decimal point, even
' though the decimal point will still be present when we reference
' a Currency variable... We'll just handle the decimal point issue
' mathematically. i.e., 357.9545 really means 3579545 or 3,579,545.
' ===========================================================================

'Performance Counter API's
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long

Public Enum timeFormat
    fmtSeconds   ' Seconds
    fmtMillisecs ' Milliseconds - one thousandth of a second
    fmtMicrosecs ' Microseconds - one millionth of a second
End Enum

'//////////////////////////////////////////////////////
'// Set CPU_CLOCK_HZ to match speed of your machine. //
'//////////////////////////////////////////////////////
Private Const CPU_CLOCK_HZ = 2540000000#  '2.54 GHz. //
'//////////////////////////////////////////////////////

Private initlStart As Currency, initlStop As Currency
Private counterFreq As Currency
Private APIoverhead As Currency

' ===========================================================================
' Example Usage 1:
' ===========================================================================
'
' Dim c As Currency
' c = CounterStart
' ' Do something...
' txtMicroSeconds.Text = CounterStop(c, fmtMicrosecs)
'
'
' ===========================================================================
' Example Usage 2: [Paste into: Command1_Click()]
' ===========================================================================
'
' Dim cnt As Currency, elapsedTime_uS As Double
' Dim N As Long, A As Long
' 'Initialize counterFreq & APIoverhead session variables.
' Call CounterStart 'Dummy Call
' cnt = CounterStart    'Start timing.
' For N = 1 To 10000    'Do something that takes time...
'     A = N
' Next N
' elapsedTime_uS = CounterStop(cnt, fmtMicrosecs) 'Stop timing.
' 'Display result.
' Debug.Print "Elapsed Time = " & elapsedTime_uS & " Microseconds"
' 'Add a TextBox named txtMicroSeconds to the Form:
' txtMicroSeconds.Text = SigFigs(elapsedTime_uS, 7) 'Seven figures.
' Debug.Print "CPU Clocks = " & CPU_Clocks(elapsedTime_uS / 1000000)
' Debug.Print "Counter Frequency = " & CounterFrequency / 1000000 & " MHz"
' Debug.Print
' ' Notice the difference between txtMicroSeconds & Debug outputs.
' ' If the hardware counter frequency is 3.579545 MHz then the
' ' hardware timing can only be accurate to 7 significant
' ' figures. So use a 7 in the SigFigs( , 7) function call.
' ===========================================================================


' ===========================================================================
' Precise Hardware Timer - CounterStart() and CounterStop()
' ===========================================================================
'
' CounterStart() returns the current count of the high-resolution
' hardware counter in a Currency data type. This count is passed
' on to CounterStop() which subtracts it from an ending count and
' divides the count difference by the hardware counter frequency
' and returns the quotient (as an elapsed time).
'
' In CounterStop() the result is returned in the time scale
' specified, and is accurate to the maximum number of digits
' returned by QueryPerformanceCounter().
'
' Multiple hardware timers can be run concurrently if required.
'
' In the case of no high-resolution counter, CounterStart() and
' CounterStop() both return zero.
'
' ===========================================================================
Public Function CounterStart() As Currency
    ' Set the API call-overhead and the counter frequency
    ' variables on the first CounterStart() call of the session:
    If (counterFreq = 0) Or (APIoverhead = 0) Then
        QueryPerformanceCounter initlStart
        QueryPerformanceCounter initlStop
        QueryPerformanceFrequency counterFreq
        APIoverhead = initlStop - initlStart
        'Debug.Print "Happens only on first session call.  " _
        '    & counterFreq & ",  " & APIoverhead
    End If
    ' Route all subsequent CounterStart() session calls
    ' directly to here:
    If (counterFreq) Then
        'Start counting:
        'Set CounterStart to the current high-speed-counter value.
        'This is also the return value of the function.
        QueryPerformanceCounter CounterStart '<-Counter tick accurate.
    End If
End Function

Public Function CounterStop(ByVal cntStart As Currency, ByVal timeFmt As timeFormat) As Double
    Dim cntStop As Currency
    If (counterFreq) Then
        'Stop counting:
        'Set cntStop to the current high-speed-counter value.
        QueryPerformanceCounter cntStop '<-Counter tick accurate.
        Select Case timeFmt
            'Next, the four Currency variables have the effect of
            'each of their four decimal points and four decimal
            'places cancelled-out because of the following RATIO
            'operation. This same ratio operation produces the
            'desired unit of Seconds as well: Very slick move!
            'Wish I had thought of it. All I did was to include
            'the subtraction of the APIoverhead variable and
            'rename some variables to follow the logic easier.
            Case fmtSeconds
                'Leave it alone to display in second (s) units.
                CounterStop = CDbl((cntStop - cntStart - APIoverhead) / counterFreq)
            Case fmtMillisecs
                'Multiply it by 1000 to display in millisecond (ms) units.
                CounterStop = CDbl((cntStop - cntStart - APIoverhead) / counterFreq) * 1000#
            Case fmtMicrosecs
                'Multiply it by 1000000 to display in microsecond (µs) units.
                CounterStop = CDbl((cntStop - cntStart - APIoverhead) / counterFreq) * 1000000#
        End Select
    End If
End Function
' ===========================================================================

'///////////////////////////////////////////////////////////////
'===============================================================
'Return 'dblNumber' rounded to 'intSF' significant figures
'===============================================================
Public Function SigFigs(dblNumber As Double, intSF As Integer) As Double
'Only works properly for doubles in the range: (+/-)1E(+/-)308
Dim negFlag As Integer
Dim tmpDbl As Double
Dim factor As Double
Dim dblA As Double
Dim dblB As Double
Dim outNum As Double

    'dblNumber = 0 ?
    If dblNumber <> 0 Then
        'make sign of tmpDbl <- dblNumber, be positive
        If dblNumber < 0 Then
            tmpDbl = -dblNumber: negFlag = -1
        Else
            tmpDbl = dblNumber: negFlag = 0
        End If
        'get multiplication/division order-of-magnitude factor
        factor = 10 ^ (Int(Log(tmpDbl) / Log(10)) + 1)
        'dblA = tmpDbl's significant digits moved to right of
        'decimal point: 0.########
        dblA = tmpDbl / factor
        'correct dblA for sign if necessary
        If negFlag Then dblA = -dblA
        'round dblA to intSF number of decimal places
        dblB = Round(dblA, intSF)
        'restore dblB to tmpDbl's original order-of-magnitude
        outNum = dblB * factor 'outNum = (positive/negative)
        'Debug.Print tmpDbl, factor, dblA, dblB, outNum
    Else  'dblNumber = 0
        outNum = 0
    End If
    SigFigs = outNum 'return
End Function
'///////////////////////////////////////////////////////////////

' ===========================================================================
' CPU Clocks
' ===========================================================================
'
' This function is useful to comare identical code performance on
' machines with different CPU clock speeds. The number returned
' from here shouldn't vary much between machines of different speed.
'
' It's also handy when you're developing code on two or more
' different machines. You can see if you've made a performance
' improvement or not, regardless of the speed of the machine
' you're working with.
'
' I like to have it on hand just to quickly see how many clock
' cycles I can shave off of an operation by experimenting and
' compairing results.
' ===========================================================================
Public Function CPU_Clocks(dblSeconds As Double) As Currency
    'This can be a BIG number. So we use a Currency data type to
    'return it. But we only return the whole-number (Integer) part.
    '...We don't care about any 4 decimal digit fractional part
    'of a cpu clock cycle.
    CPU_Clocks = Int(CPU_CLOCK_HZ * dblSeconds)
End Function
' ===========================================================================


' ===========================================================================
' Hardware Counter Frequency - Returned in units of 'Hz'
' ===========================================================================
Public Function CounterFrequency() As Long
    Dim Freq As Currency
    QueryPerformanceFrequency Freq
    'We multiply by 10000 because, in VB, Freq is now a Currency
    'variable which has a decimal point and four decimal places
    'on the right-hand side of it. i.e., Say the API returns the
    'frequency 3579545 back to VB through our Freq (Currency)
    'variable - it just returns it in the lowest seven digits of
    'the Currency variable. Now, when we look at Freq in VB as a
    'Currency variable what we'll see is this: 357.9545, because
    'it's in Currency format (Now do you get-it?) We have to
    'multiply The Freq Currency variable by 10000 to get the real
    'frequency value that the API function returned through the
    'Currency variable to us. Namely 357.9545 * 10000 =  3579545
    '(the actual counter frequency - without the decimal point
    'stuck in the middle of it).
    CounterFrequency = CLng(Freq * 10000#)
End Function
' ===========================================================================
