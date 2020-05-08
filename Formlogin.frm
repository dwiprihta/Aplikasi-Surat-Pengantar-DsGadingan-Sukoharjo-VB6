VERSION 5.00
Begin VB.Form formlogin 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LOGIN SISTEM APLIKASI SURAT "
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5355
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Formlogin.frx":0000
   ScaleHeight     =   4485
   ScaleWidth      =   5355
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Caption         =   "BATAL"
      Height          =   615
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3360
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      Caption         =   "MASUK"
      Height          =   615
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3360
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   615
      IMEMode         =   3  'DISABLE
      Left            =   600
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2400
      Width           =   4095
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   600
      TabIndex        =   0
      Top             =   1080
      Width           =   4095
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "PASSWORD"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   600
      TabIndex        =   5
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "USERNAME"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   600
      TabIndex        =   4
      Top             =   720
      Width           =   1455
   End
End
Attribute VB_Name = "formlogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'======================= FORM LOGIN CODE===========================
     '======================= REZA GUNAWAN ===========================
     
     
'JIKA TOMBOL LOGIN DI KLIK
Private Sub Command1_Click()

'panggil koneksi
Call Koneksi

'cek jika form masih kosong
If Text1.Text = "" Then
MsgBox "FORM USERNAME ANDA MASIH KOSONG !", vbCritical, "Perhatian"
Text1.SetFocus
ElseIf Text2.Text = "" Then
MsgBox "FORM PASSWORD ANDA MASIH KOSONG !!!", vbCritical, "Perhatian"
Text2.SetFocus
Else

'cari data login di database admin
query = "select * from login where username='" & Text1.Text & "' and password='" & Text2.Text & "'"
RS.Open (query), conn
    If RS.EOF Then
    MsgBox "USERNAME ATAU PASSWORD ANDA SALAH !", vbExclamation, "Gagal !"
    Text1.Text = ""
    Text2.Text = ""
    Text1.SetFocus
    Else
    
    'jika berhasil login masuk ke menu admin
    Unload Me
    MsgBox "ANDA BERHASIL LOGIN !", vbInformation, "LOGIN SUKSES !"
    index.Show
    End If
End If
End Sub

'JIKA TOMBOL CANCEL DIKLIK
Private Sub Command2_Click()
Unload Me
End Sub

