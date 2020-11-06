VERSION 5.00
Begin VB.Form VentanaImagen 
   AutoRedraw      =   -1  'True
   Caption         =   " Imagen"
   ClientHeight    =   3375
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3960
   Icon            =   "VentanaImagen.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   3375
   ScaleWidth      =   3960
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   3135
      Left            =   0
      ScaleHeight     =   3075
      ScaleWidth      =   3795
      TabIndex        =   0
      Top             =   0
      Width           =   3855
   End
End
Attribute VB_Name = "VentanaImagen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Me.Picture1.Top = 0
    Me.Picture1.Left = 0
    Me.Picture1.Height = Me.Height
    Me.Picture1.Width = Me.Width
End Sub

Private Sub Form_Resize()
    Me.Picture1.Height = Me.Height - 10
    Me.Picture1.Width = Me.Width - 10
End Sub
