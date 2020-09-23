VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form MaskCreator 
   Caption         =   "Character Mask Creator By Winter"
   ClientHeight    =   6825
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7800
   Icon            =   "MaskCreator.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   455
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   520
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox ColorCont2 
      Height          =   735
      Left            =   1080
      ScaleHeight     =   675
      ScaleWidth      =   795
      TabIndex        =   19
      Top             =   2160
      Width           =   855
   End
   Begin MSComDlg.CommonDialog Dlg 
      Left            =   1560
      Top             =   4200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCreateMask 
      Caption         =   "Create Mask"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton cmdOpenFile 
      Caption         =   "Open File"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "Save Mask"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   5280
      Width           =   1215
   End
   Begin VB.PictureBox ColorCont1 
      Height          =   735
      Left            =   120
      ScaleHeight     =   675
      ScaleWidth      =   795
      TabIndex        =   13
      Top             =   2160
      Width           =   855
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse..."
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox FileLoc 
      Height          =   285
      Left            =   120
      TabIndex        =   10
      Top             =   600
      Width           =   2055
   End
   Begin VB.PictureBox ConContener2 
      Height          =   2775
      Left            =   2280
      ScaleHeight     =   181
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   357
      TabIndex        =   4
      Top             =   3600
      Width           =   5415
      Begin VB.PictureBox Contener2 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   120
         ScaleHeight     =   49
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   65
         TabIndex        =   7
         Top             =   120
         Width           =   975
      End
      Begin VB.HScrollBar HScroll2 
         Height          =   255
         Left            =   0
         TabIndex        =   6
         Top             =   2400
         Width           =   5295
      End
      Begin VB.VScrollBar VScroll2 
         Height          =   1575
         Left            =   5040
         TabIndex        =   5
         Top             =   360
         Width           =   255
      End
   End
   Begin VB.PictureBox ConContener1 
      Height          =   2775
      Left            =   2280
      ScaleHeight     =   181
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   357
      TabIndex        =   0
      Top             =   360
      Width           =   5415
      Begin VB.VScrollBar VScroll1 
         Height          =   1575
         Left            =   5040
         TabIndex        =   3
         Top             =   360
         Width           =   255
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         Left            =   0
         TabIndex        =   2
         Top             =   2400
         Width           =   5295
      End
      Begin VB.PictureBox Contener1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   120
         ScaleHeight     =   49
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   65
         TabIndex        =   1
         Top             =   120
         Width           =   975
         Begin VB.Image CSelecter 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   150
            Left            =   120
            Top             =   120
            Width           =   150
         End
      End
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Pres Escape to Cancel when generating mask!"
      Height          =   195
      Left            =   120
      TabIndex        =   20
      Top             =   6480
      Width           =   3315
   End
   Begin VB.Label Label5 
      Caption         =   "To choose a color to be transparent, just click on it on the picure window, once chosen, click the CreateMask Button."
      Height          =   1095
      Left            =   120
      TabIndex        =   15
      Top             =   3000
      Width           =   1935
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Transparent Color:"
      Height          =   195
      Left            =   120
      TabIndex        =   14
      Top             =   1920
      Width           =   1305
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "File Name:"
      Height          =   195
      Left            =   120
      TabIndex        =   11
      Top             =   360
      Width           =   750
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Mask:"
      Height          =   195
      Left            =   2280
      TabIndex        =   9
      Top             =   3360
      Width           =   435
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Picture:"
      Height          =   195
      Left            =   2280
      TabIndex        =   8
      Top             =   120
      Width           =   540
   End
End
Attribute VB_Name = "MaskCreator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************************
'* Mask Creator by Winter
'*
'* Very usefull to create masks for use on games
'* with the BitBlt function with the constants
'* vbMergePaint for the mask ans vbSrcAnd for the
'* bitmap.
'*
'*************************************************

Private Sub cmdBrowse_Click()
'Opens the dialog to browse for the bmp file
With Dlg
.DialogTitle = "Open BMP file"
.Filter = "BMP files (*.bmp)|*.bmp"
.ShowOpen
If Len(.FileName) = 0 Then
Exit Sub
End If
End With
'store the file path to the global string variable
FileLoc.Text = Dlg.FileName
FileName = FileLoc.Text
End Sub

Private Sub cmdCreateMask_Click()
'Create the mask when user clicks

'Set the global variable TColor with the color chosen
TColor = ColorCont2.BackColor

'Set the cursor to hourglass
Screen.MousePointer = vbHourglass

'Copy the picture from the bitmap and properties contener to the mask contener
With Contener2
.Width = Contener1.Width
.Height = Contener1.Height
.Picture = Contener1.Picture
End With

'Start selecting all points in the bitmap to set the color chosen as white
For Ix = 0 To Contener1.Width + 5
For Iy = 0 To Contener1.Height + 5
'If user presses Escape, stop the job
If GetAsyncKeyState(VK_ESCAPE) < 0 Then
Form_Resize
Screen.MousePointer = vbDefault
Exit Sub
End If
'If the current point is the color chosen, make it white
If Contener1.Point(Ix, Iy) = TColor Then
Contener1.PSet (Ix, Iy), vbWhite
End If
'DoEvents to show progress, you can rem this if you don't want to show progress
DoEvents
Next
Next

'Set result image as picture for bitmap 1 and copy picture to the mask contener (contener2)
Contener1.Picture = Contener1.Image
Contener2.Picture = Contener1.Picture

'Start chosing points in the mask image to turn black all those that are not white
For Ix = 0 To Contener2.Width + 5
For Iy = 0 To Contener2.Height + 5
'If user presses Escape, exit the job
If GetAsyncKeyState(VK_ESCAPE) < 0 Then
Form_Resize
Screen.MousePointer = vbDefault
Exit Sub
End If
'Turn points that are not white black
If Contener2.Point(Ix, Iy) <> vbWhite Then
Contener2.PSet (Ix, Iy), vbBlack
End If
'As before, the DoEvents
DoEvents
Next
Next

'Set the mask image as picture for contener2
Contener2.Picture = Contener2.Image

'Do the resize function and set the cursor to default
Form_Resize
Screen.MousePointer = vbDefault
End Sub

Private Sub cmdOpenFile_Click()
'Shows the picture in contener1 from the path given with the comon dialog or enered on the textbox
Contener1.Picture = LoadPicture(FileName)
'Do the form_resize function
Form_Resize
End Sub


Private Sub CmdSave_Click()
'Open the dialog to save the bitmap and mask bitmap
With Dlg
.DialogTitle = "Save Mask as 'File.bmp' and 'FileMask.bmp'"
.Filter = "Don't use extensions!|*.*"
.ShowSave
If Len(.FileName) = 0 Then
Exit Sub
End If
FileName = .FileName
End With

'Use the path an filename to save two files: File.bmp and FileMask.bmp
SavePicture Contener1.Picture, FileName & ".bmp"
SavePicture Contener2.Picture, FileName & "Mask.bmp"
End Sub

Private Sub Contener1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'This is for when the user moves the mouse over the bitmap in contener1, show the color of where the mouse is in the color contener picturebox
CSelecter.Left = X
CSelecter.Top = Y
ColX = X
ColY = Y
'Set the global variable TColor
TColor = Contener1.Point(ColX, ColY)
'Set the backcolor of the PictureBox (ColorCot1) to show the color where the mouse is on
ColorCont1.BackColor = TColor
End Sub

Private Sub CSelecter_Click()
'Select the color to be transparent when user clicks by puting it on ColorCont2 PictureBox back color
ColorCont2.BackColor = TColor
End Sub

Private Sub Form_Load()
'Set everuthing in place
Form_Resize
End Sub

Private Sub Form_Resize()
'Place everything where it belongs
With ConContener1
.Left = 155
.Width = MaskCreator.ScaleWidth - 165
.Height = MaskCreator.ScaleHeight * 0.42528
End With
With VScroll1
.Left = ConContener1.Width - 20
.Top = 0
.Height = ConContener1.Height - 21
If ConContener1.Height > Contener1.Height Then
.Min = 0
.Max = ConContener1.Height
.Visible = False
End If
If ConContener1.Height < Contener1.Height Then
.Min = 0
.Max = Contener1.Height - ConContener1.Height
.Visible = True
End If
End With
With HScroll1
.Left = 0
.Top = ConContener1.Height - 20
.Width = ConContener1.Width - 20
If ConContener1.Width > Contener1.Width Then
.Min = 0
.Max = ConContener1.Width
.Visible = False
End If
If ConContener1.Width < Contener1.Width Then
.Min = 0
.Max = Contener1.Width - ConContener1.Width
.Visible = True
End If
End With
With ConContener2
.Left = 155
.Width = MaskCreator.ScaleWidth - 165
.Height = MaskCreator.ScaleHeight * 0.42528
.Top = ConContener1.Top + ConContener1.Height + 20
End With
With VScroll2
.Left = ConContener2.Width - 20
.Top = 0
.Height = ConContener2.Height - 21
If ConContener2.Height > Contener2.Height Then
.Min = 0
.Max = ConContener2.Height
.Visible = False
End If
If ConContener2.Height < Contener2.Height Then
.Min = 0
.Max = Contener2.Height - ConContener2.Height
.Visible = True
End If
End With
With HScroll2
.Left = 0
.Top = ConContener2.Height - 20
.Width = ConContener2.Width - 20
If ConContener2.Width > Contener2.Width Then
.Min = 0
.Max = ConContener2.Width
.Visible = False
End If
If ConContener2.Width < Contener2.Width Then
.Min = 0
.Max = Contener2.Width - ConContener2.Width
.Visible = True
End If
End With
With Label1
.Left = ConContener1.Left
.Top = ConContener1.Top - 17
End With
With Label2
.Left = ConContener2.Left
.Top = ConContener2.Top - 17
End With
Contener1.Top = 0
Contener1.Left = 0
Contener2.Top = 0
Contener2.Left = 0
Label6.Top = MaskCreator.ScaleHeight - 20
End Sub

'The Scrolls will move the bitmap
Private Sub HScroll1_Change()
Contener1.Left = -HScroll1.Value
End Sub

Private Sub HScroll2_Change()
Contener2.Left = -HScroll2.Value
End Sub

Private Sub VScroll1_Change()
Contener1.Top = -VScroll1.Value
End Sub

Private Sub VScroll2_Change()
Contener2.Top = -VScroll2.Value
End Sub
