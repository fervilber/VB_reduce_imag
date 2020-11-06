VERSION 5.00
Begin VB.Form Principal 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   5850
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4710
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   390
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   314
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnCerrar 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   3600
      TabIndex        =   8
      Top             =   4320
      Width           =   975
   End
   Begin VB.CommandButton btnconv1 
      Caption         =   "convertir actual"
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   4320
      Width           =   1335
   End
   Begin VB.TextBox Textrutadestino 
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   5280
      Width           =   4335
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   3960
      TabIndex        =   4
      Text            =   "60"
      Top             =   4800
      Width           =   615
   End
   Begin VB.DirListBox Dir1 
      Height          =   3465
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   2175
   End
   Begin VB.FileListBox File1 
      Height          =   3795
      Left            =   2400
      Pattern         =   "*.jpg;*.gif"
      TabIndex        =   2
      Top             =   120
      Width           =   2175
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton btnAbrirtodas 
      Caption         =   "Convertir todas"
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Top             =   4320
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "valor de compresion"
      Height          =   375
      Left            =   2400
      TabIndex        =   5
      Top             =   4800
      Width           =   1575
   End
End
Attribute VB_Name = "Principal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Cambiartamaño As Boolean
Dim objPic As Picture
Private Declare Function DIWriteJpg Lib "DIjpg.dll" (ByVal DestPath As String, ByVal quality As Long, ByVal progressive As Long) As Long
' SavePicture Me.Picture1.Image, App.Path & "\raster.bmp"
 
 

Private Sub btnAbrirtodas_Click()
'Guarda con el tamaño especificado todas las imagenes que se encuentren en el directorio
' Dim objPic As Picture
 Dim RUTA$
 Dim NOMBRE
   For n = 0 To File1.ListCount - 1
        Set objPic = Nothing
        Me.File1.ListIndex = n
        
        If Right(Me.File1.FileName, 3) = "jpg" Or Right(Me.File1.FileName, 3) = "gif" Then
            File1_Click
            'cadena = File1.Path & "\" & File1.FileName
            'Set objPic = LoadPicture(cadena)
            NOMBRE = Me.File1.FileName
            
         If UCase(Right(Me.File1.FileName, 3)) = "GIF" Then
           NOMBRE = Me.File1.FileName
           NOMBRE = Mid(NOMBRE, 1, Len(NOMBRE) - 3) & "JPG"
           Else
           NOMBRE = Me.File1.FileName
         End If
         
         RUTA = Me.Textrutadestino.Text & "\nuevos\" & NOMBRE 'Me.File1.FileName
         If Dir(Me.Textrutadestino.Text & "\nuevos\", vbDirectory) = "" Then MkDir (Me.Textrutadestino.Text & "\nuevos\")
         Me.GuardaIMG RUTA, CInt(Val(Me.Text1.Text))
         Do While Dir(RUTA) = ""
          n = n + 1
          If n > 100000 Then Exit Do
         Loop
        End If
    Next n
End Sub



Private Sub btnCerrar_Click()
'Sale de la aplicación
 End
End Sub

Private Sub btnconv1_Click()
'"Convertit uno"
'Convierte la imagen actual en otra mas pequeña
Dim RUTA$
Dim NOMBRE
    'Actualiza la imegen actual
    Set objPic = Nothing
    '
    If Right(Me.File1.FileName, 3) = "jpg" Or Right(Me.File1.FileName, 3) = "gif" Then
     File1_Click
     
        If UCase(Right(Me.File1.FileName, 3)) = "GIF" Then
           NOMBRE = Me.File1.FileName
           NOMBRE = Mid(NOMBRE, 1, Len(NOMBRE) - 3) & "JPG"
           Else
           NOMBRE = Me.File1.FileName
        End If

     RUTA = Me.Textrutadestino.Text & "\nuevos\" & NOMBRE
     
     'Crea el directorio nuevoas si no existe ya
     If Dir(Me.Textrutadestino.Text & "\nuevos\", vbDirectory) = "" Then MkDir (Me.Textrutadestino.Text & "\nuevos\")
     
     'Almacena la imagen
     Me.GuardaIMG RUTA, CInt(Val(Me.Text1.Text))
    End If
End Sub

Private Sub Dir1_Change()
Me.File1.Path = Me.Dir1.Path
End Sub

Private Sub Drive1_Change()
Me.Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_Click()
'Cuando se selecciona un gráfico en el filelist lo muestra en el picture1 ajustado
'Dim objPic As Picture
  Set objPic = Nothing
 If Right(File1.FileName, 3) = "jpg" Or Right(Me.File1.FileName, 3) = "gif" Then

   cadena = File1.Path & "\" & File1.FileName
   Set objPic = LoadPicture(cadena)
   
   VentanaImagen.Show
   Set VentanaImagen.Picture1 = Nothing
   VentanaImagen.Picture1.PaintPicture objPic, 0, 0, VentanaImagen.Picture1.ScaleWidth, VentanaImagen.Picture1.ScaleHeight

 End If
End Sub

Private Sub Form_Load()
 VentanaImagen.Show
 Me.Textrutadestino.Text = App.Path
 Dir1.Path = App.Path
 Me.Refresh
 
End Sub


Public Sub GuardaIMG(RutayNombre As String, Compresion As Integer)
Dim cadena, retval

 'Borra el fichero de imagen si existe
  
 If Dir(RutayNombre) <> "" Then
  Kill RutayNombre
 End If
 
 'Guarda la imagen en un *.bmp
 SavePicture VentanaImagen.Picture1.Image, "c:\tmp.bmp"

 'La guarda como JPEG
    retval = DIWriteJpg(RutayNombre, Compresion, 1)

    If retval = 1 Then  'Si lo hace con exito
    Else                'Si hay un error
        MsgBox "Error en la conversión a jpg"
    End If

    Kill "c:\tmp.bmp" ' borrael fichero temporal bmp

End Sub


Private Sub Form_Unload(Cancel As Integer)
 Unload VentanaImagen
 End
End Sub
