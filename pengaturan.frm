VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form pengaturan 
   Caption         =   "FORM PENGATURAN"
   ClientHeight    =   2415
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   ScaleHeight     =   2415
   ScaleWidth      =   6975
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   5040
      Top             =   2400
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\APP SURAT GADINGAN\gadingan.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\APP SURAT GADINGAN\gadingan.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "combo"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "pengaturan.frx":0000
      Height          =   1215
      Left            =   120
      TabIndex        =   6
      Top             =   3000
      Visible         =   0   'False
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   2143
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      Caption         =   "PERBARUI"
      Height          =   495
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      Caption         =   "PERBARUI"
      Height          =   495
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   1
      Top             =   1680
      Width           =   3975
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   0
      Top             =   600
      Width           =   3975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "CAMAT"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "KEPALA DESA"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   720
      Width           =   1215
   End
End
Attribute VB_Name = "pengaturan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'======================= FORM PENGATURAN CODE===========================
     '======================= REZA GUNAWAN ===========================

'SIMPAN DATA (NAMA LURAH) DARI DATABASE
Private Sub Command1_Click()
Adodc1.ConnectionString = conn.ConnectionString
Adodc1.RecordSource = "select* from combo"
With Adodc1.Recordset
    !lurah = Text1.Text
.Update
End With
MsgBox "Pembaruan Disimpan", vbInformation, "Ganti nama Kepala Desa"
End Sub

'SIMPAN DATA (NAMA CAMAT) DARI DATABASE
Private Sub Command2_Click()
Adodc1.ConnectionString = conn.ConnectionString
Adodc1.RecordSource = "select* from combo"
With Adodc1.Recordset
    !camat = Text2.Text
.Update
End With
MsgBox "Pembaruan Disimpan", vbInformation, "Ganti nama Kepala Desa"
End Sub

'TAMPILAKN DATA DARI DATABSE
Private Sub Form_Load()
Koneksi
Adodc1.ConnectionString = conn.ConnectionString
Adodc1.RecordSource = "select* from combo"
With Adodc1.Recordset
    Text1 = !lurah
    Text2 = !camat
    .Update
End With

End Sub
