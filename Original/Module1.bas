Attribute VB_Name = "Module1"




'PARA PODER GUARDAR IMAGEN DEL FORMULARIO
Public Declare Sub keybd_event _
    Lib "user32" ( _
        ByVal bVk As Byte, _
        ByVal bScan As Byte, _
        ByVal dwFlags As Long, _
        ByVal dwExtraInfo As Long)
  
Public Type RegGraficaPrimos
  Numero As Long
  Primo As Integer
  CX As Double
  CY As Double
  Tamaño As Integer
  Color As Integer
  PCX As Double
  PCY As Double
End Type
  
Public rGraficaPrimos As RegGraficaPrimos

''Public cn As ADODB.Connection
''Public sql As String

