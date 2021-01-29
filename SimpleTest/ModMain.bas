Attribute VB_Name = "ModMain"
Option Explicit
Private Const MAX_POINTS As Long = 10

Sub Main()
''This are just two quick tests, calling all procedures
  MsgBox "Simple 2D- and 3D-testroutines calling all FastGeo-procedures, if no error appears, everything then is OK."

  Call Test2D
  Call Test3D
  
  MsgBox "Test2D and Test3D OK, pogramm will quit now"
End Sub
'UD-Type und Variant:
'Nur benutzerdefinierte Typen, die in öffentlichen Objektmodulen definiert sind, können in den oder aus
'dem Typ Variant umgewandelt werden oder an eine zur Laufzeit auflösbare Funktion weitergeleitet
'werden.
