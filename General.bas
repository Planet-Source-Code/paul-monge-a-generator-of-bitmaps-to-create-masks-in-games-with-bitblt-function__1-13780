Attribute VB_Name = "General"


'Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

'The GetAsyncKeyState function to check when user presses Escape
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Const VK_ESCAPE = &H1B

'Gloval variables:

'The File Path and Name to be loaded as image
Public FileName As String

'X and Y coordinates for use with the point property of picture boxes for the color
Public ColX As Integer
Public ColY As Integer

'Variable to store the color tobe transparent
Public TColor As Long

'X and Y coordinates for use with the point property of picture boxes for the point position
Public Ix As Integer
Public Iy As Integer
