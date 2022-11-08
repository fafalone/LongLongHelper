# LongLongHelper

Provides a standard DLL for handling Currency in VB6 as actually a LongLong variable, using twinBASIC's native support for LongLong and Standard DLLs.

Available functions in v1.0:

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
