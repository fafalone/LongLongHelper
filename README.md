# LongLongHelper

Provides a standard DLL for handling Currency in VB6 as actually a LongLong variable, using twinBASIC's native support for LongLong and Standard DLLs.

In VB6 when we need to use 64bit integers, we use the [tt]Currency[/tt] type, but since it's a hack, they're difficult to work with, because VB is wont to shift them back and forth by 10,000 since a Currency is a 0.0000 decimal type. This is a Standard DLL written in twinBASIC (which supports that project type natively) that treats VB6 Currency variable you're using as LongLong as actual LongLong types, and performs operations on them using twinBASIC's native language support for the LongLong type, making the code extraordinarily simple.

So you have e.g.
```
[ DllExport ]
    Public Function LongLongAnd(ByVal c1 As LongLong, ByVal c2 As LongLong) As LongLong
        On Error GoTo fail
        LongLongAnd = (c1 And c2)
        hErr = S_OK
        Exit Function
    fail:
        hErr = Err.Number
    End Function
    
    [ DllExport ]
    Public Function LongLongOr(ByVal c1 As LongLong, ByVal c2 As LongLong) As LongLong
        On Error GoTo fail
        LongLongOr = (c1 Or c2)
        hErr = S_OK
        Exit Function
    fail:
        hErr = Err.Number
    End Function
```
That you then declare in VB6 as 

`Public Declare Function LongLongAnd Lib "LngLngHelp.dll" (ByVal c1 As Currency, ByVal c2 As Currency) As Currency`

`Public Declare Function LongLongXOr Lib "LngLngHelp.dll" (ByVal c1 As Currency, ByVal c2 As Currency) As Currency`


The DLL sees them as a LongLong, so you don't need to worry about multiplying or dividing by 10,000 until you want to e.g. display the result. 


There's also a CLngLng function:

```
[ DllExport ]
    Public Function CLongLong(ByVal Value As Variant) As LongLong
        On Error GoTo fail
        CLongLong = CLngLng(Value)
        hErr = S_OK
        Exit Function
    fail:
        hErr = Err.Number
    End Function
```
`Public Declare Function CLngLng Lib "LngLngHelp.dll" Alias "CLongLong" (ByVal Value As Variant) As Currency`

The error codes you see being set can be retrieved with [tt]LongLongLastError[/tt].

Here's a complete set of declares for the DLL:
```
Public Declare Function LongLongAdd Lib "LngLngHelp.dll" (ByVal c1 As Currency, ByVal c2 As Currency) As Currency
Public Declare Function LongLongSub Lib "LngLngHelp.dll" (ByVal c1 As Currency, ByVal c2 As Currency) As Currency
Public Declare Function LongLongOr Lib "LngLngHelp.dll" (ByVal c1 As Currency, ByVal c2 As Currency) As Currency
Public Declare Function LongLongAnd Lib "LngLngHelp.dll" (ByVal c1 As Currency, ByVal c2 As Currency) As Currency
Public Declare Function LongLongXOr Lib "LngLngHelp.dll" (ByVal c1 As Currency, ByVal c2 As Currency) As Currency
Public Declare Function LongLongNot Lib "LngLngHelp.dll" (ByVal c1 As Currency, ByVal c2 As Currency) As Currency
Public Declare Function LongLongDiv Lib "LngLngHelp.dll" (ByVal c1 As Currency, ByVal c2 As Currency) As Currency
Public Declare Function LongLongMul Lib "LngLngHelp.dll" (ByVal c1 As Currency, ByVal c2 As Currency) As Currency
Public Declare Function LongLongLShift Lib "LngLngHelp.dll" (ByVal c1 As Currency, ByVal By As Byte) As Currency
Public Declare Function LongLongRShift Lib "LngLngHelp.dll" (ByVal c1 As Currency, ByVal By As Byte) As Currency
Public Declare Sub LongLongInc Lib "LngLngHelp.dll" (c1 As Currency, Optional ByVal Amount As Long = 1)
Public Declare Sub LongLongDec Lib "LngLngHelp.dll" (c1 As Currency, Optional ByVal Amount As Long = 1)
Public Declare Function CLngLng Lib "LngLngHelp.dll" Alias "CLongLong" (ByVal Value As Variant) As Currency
Public Declare Function LongLongLastError Lib "LngLngHelp.dll" () As Long 'Check if last operation errored.
```

Full source can be browsed in \Export in addition to the download of full source .twinproj file. 

Any remotely recent version of twinBASIC should work if you wish to compile/modify the DLL yourself. 

