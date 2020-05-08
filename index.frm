VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form index 
   Caption         =   "SELAMAT DATANG ( APLIKASI SURAT PENGANTAR DESA GADINGAN, MOJOLABAN, SUKOHARJO)"
   ClientHeight    =   9165
   ClientLeft      =   165
   ClientTop       =   510
   ClientWidth     =   17625
   LinkTopic       =   "Form2"
   Picture         =   "index.frx":0000
   ScaleHeight     =   9165
   ScaleWidth      =   17625
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2040
      Top             =   0
   End
   Begin Crystal.CrystalReport CR1 
      Left            =   1440
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Image Image1 
      Height          =   2415
      Left            =   5520
      Picture         =   "index.frx":36A7
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   2415
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      Caption         =   "DATA PENDUDUK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   300
      Left            =   5625
      TabIndex        =   8
      Top             =   6600
      Width           =   2295
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "DESA GADINGAN, MOJOLABAN, SUKOHARJO"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   675
      Left            =   3705
      TabIndex        =   7
      Top             =   1560
      Width           =   12075
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "APLIKASI SURAT PENGANTAR"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   675
      Left            =   5610
      TabIndex        =   6
      Top             =   720
      Width           =   8085
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      Caption         =   "PENGATURAN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   300
      Left            =   15000
      TabIndex        =   5
      Top             =   6600
      Width           =   1845
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      Caption         =   "DATA ADMINISTRASI"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   300
      Left            =   8400
      TabIndex        =   4
      Top             =   6600
      Width           =   2745
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      Caption         =   "LAPORAN SURAT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   300
      Left            =   11760
      TabIndex        =   3
      Top             =   6600
      Width           =   2265
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      Caption         =   "TRANSAKSI SURAT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   300
      Left            =   2400
      TabIndex        =   2
      Top             =   6600
      Width           =   2505
   End
   Begin VB.Image Image5 
      Height          =   2175
      Left            =   2520
      Picture         =   "index.frx":6AEC
      Stretch         =   -1  'True
      Top             =   4200
      Width           =   2175
   End
   Begin VB.Image Image4 
      Height          =   2415
      Left            =   8520
      Picture         =   "index.frx":9894
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   2415
   End
   Begin VB.Image Image3 
      Height          =   2535
      Left            =   14640
      Picture         =   "index.frx":D245
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   2535
   End
   Begin VB.Image Image2 
      Height          =   2175
      Left            =   11760
      Picture         =   "index.frx":1068A
      Stretch         =   -1  'True
      Top             =   4200
      Width           =   2175
   End
   Begin VB.Label Label88 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      BackStyle       =   0  'Transparent
      Caption         =   "--/--/----"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   615
      Left            =   3240
      TabIndex        =   1
      Top             =   2760
      Width           =   9015
   End
   Begin VB.Label Label77 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      BackStyle       =   0  'Transparent
      Caption         =   "--:--:--"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   615
      Left            =   10440
      TabIndex        =   0
      Top             =   2760
      Width           =   2895
   End
   Begin VB.Menu keluar 
      Caption         =   "LOG-OUT"
   End
End
Attribute VB_Name = "index"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'======================= FORM HALAMAN DEPAN CODE===========================
     '======================= REZA GUNAWAN ===========================

'TAMPILKAN FORM TRANSAKSI
Private Sub Image1_Click()
penduduk.Show
End Sub

'TAMPILKAN LAPORAN KESELURUHAN DATA
Private Sub Image2_Click()
CR1.Reset
With CR1
    .ReportFileName = App.Path & "\Data_surat.rpt"
    .WindowState = crptMaximized
    .RetrieveDataFiles
    .Action = 1
End With
End Sub

Private Sub Image3_Click()
pengaturan.Show
End Sub

Private Sub Image4_Click()
user.Show
End Sub

Private Sub Image5_Click()
transaksi.Show
End Sub

'KELUAR APLIKASI
Private Sub keluar_Click()
xx = MsgBox("Apakah Anda yakin akan kelua dari aplikasi pengantar surat desa kroyo ?", vbOKCancel, "Peringatan")
            If xx = vbOK Then
                    Unload Me
                Else
                    'NO NOTIF
            End If
End Sub

'TAMPILKAN PENGATURAN (NAMA LURAH & CAMAT)
Private Sub setting_Click()
pengaturan.Show
End Sub


'SOURCE JAM BERJALAN
Private Sub Timer1_Timer()
Label77.Caption = Format(Now, "hh : mm : ss")
Label88.Caption = Format(Now, "dd MMMM yyyy")
End Sub

