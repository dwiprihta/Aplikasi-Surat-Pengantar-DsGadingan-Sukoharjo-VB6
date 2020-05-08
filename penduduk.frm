VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form penduduk 
   BackColor       =   &H00FFFFFF&
   Caption         =   "FORM DATA PENDUDUK"
   ClientHeight    =   10755
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   20370
   LinkTopic       =   "Form1"
   ScaleHeight     =   10755
   ScaleWidth      =   20370
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2400
      Top             =   0
   End
   Begin VB.ComboBox Combo4 
      Appearance      =   0  'Flat
      DataSource      =   "Adodc2"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   9000
      TabIndex        =   26
      Top             =   2760
      Width           =   3615
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   10335
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   20055
      Begin VB.CommandButton Command1 
         BackColor       =   &H0000FFFF&
         Caption         =   "TAMBAH DATA"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   4680
         Width           =   1575
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H0000FFFF&
         Caption         =   "UBAH"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3840
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   4680
         Width           =   1575
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H0000FFFF&
         Caption         =   "SIMPAN"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   4680
         Width           =   1575
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H0000FFFF&
         Caption         =   "HAPUS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5520
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   4680
         Width           =   1335
      End
      Begin VB.Frame Frame4 
         Caption         =   "CARI DATA BERDASARKAN NAMA/NOMOR KTP"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   13200
         TabIndex        =   37
         Top             =   4440
         Width           =   5415
         Begin VB.TextBox Text5 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            TabIndex        =   42
            Top             =   360
            Width           =   3495
         End
         Begin VB.CommandButton Command5 
            BackColor       =   &H0000C000&
            Caption         =   "CARI"
            Height          =   495
            Left            =   3720
            Style           =   1  'Graphical
            TabIndex        =   41
            Top             =   360
            Width           =   1455
         End
      End
      Begin VB.ComboBox Combo6 
         DataSource      =   "Adodc2"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   14880
         TabIndex        =   36
         Top             =   480
         Width           =   3615
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   14880
         TabIndex        =   35
         Top             =   1200
         Width           =   3615
      End
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   14880
         TabIndex        =   34
         Top             =   1920
         Width           =   3615
      End
      Begin VB.TextBox Text8 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   14880
         TabIndex        =   33
         Top             =   2640
         Width           =   3615
      End
      Begin VB.TextBox Text9 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   14880
         TabIndex        =   32
         Top             =   3360
         Width           =   3615
      End
      Begin VB.ComboBox Combo5 
         DataSource      =   "Adodc2"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   8760
         TabIndex        =   25
         Top             =   3720
         Width           =   3615
      End
      Begin VB.ComboBox Combo2 
         DataSource      =   "Adodc2"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   8760
         TabIndex        =   24
         Top             =   480
         Width           =   3615
      End
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         DataSource      =   "Adodc2"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   8760
         TabIndex        =   23
         Top             =   1200
         Width           =   3615
      End
      Begin VB.ComboBox Combo3 
         DataSource      =   "Adodc2"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   8760
         TabIndex        =   22
         Top             =   1800
         Width           =   3615
      End
      Begin VB.ComboBox combo_agama 
         DataSource      =   "Adodc2"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   8760
         TabIndex        =   21
         Top             =   3120
         Width           =   3615
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Perempuan"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4800
         TabIndex        =   13
         Top             =   2400
         Width           =   1335
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   5
         Top             =   3000
         Width           =   3615
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   4
         Top             =   480
         Width           =   3615
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   3
         Top             =   1080
         Width           =   3615
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2520
         TabIndex        =   2
         Top             =   1800
         Width           =   3615
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Laki-Laki"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2520
         TabIndex        =   1
         Top             =   2520
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   2520
         TabIndex        =   12
         Top             =   3720
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   661
         _Version        =   393216
         Format          =   89194497
         CurrentDate     =   43327
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "penduduk.frx":0000
         Height          =   4455
         Left            =   480
         TabIndex        =   44
         Top             =   5640
         Width           =   18015
         _ExtentX        =   31776
         _ExtentY        =   7858
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
      Begin VB.Label Label13 
         Caption         =   "Status Perkawinan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   13080
         TabIndex        =   31
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "NO Paspor"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   13080
         TabIndex        =   30
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label11 
         Caption         =   "NO Kitas"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   13080
         TabIndex        =   29
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label Label10 
         Caption         =   "Nama Ayah"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   13080
         TabIndex        =   28
         Top             =   2520
         Width           =   1695
      End
      Begin VB.Label Label12 
         Caption         =   "Nama Ibu"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   13080
         TabIndex        =   27
         Top             =   3240
         Width           =   1695
      End
      Begin VB.Line Line2 
         X1              =   12840
         X2              =   12840
         Y1              =   480
         Y2              =   4200
      End
      Begin VB.Label Label9 
         Caption         =   "Status Keluarga"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6960
         TabIndex        =   20
         Top             =   3840
         Width           =   1935
      End
      Begin VB.Label Label8 
         Caption         =   "Kewarganegaraan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6960
         TabIndex        =   19
         Top             =   2520
         Width           =   1695
      End
      Begin VB.Label Label6 
         Caption         =   "Alamat"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6960
         TabIndex        =   18
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Pendidikan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   6960
         TabIndex        =   17
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Pekerjaan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   6960
         TabIndex        =   16
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label7 
         Caption         =   "Agama"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6960
         TabIndex        =   15
         Top             =   3120
         Width           =   1695
      End
      Begin VB.Line Line1 
         X1              =   6600
         X2              =   6600
         Y1              =   480
         Y2              =   4200
      End
      Begin VB.Label Label4 
         Caption         =   "Tanggal Lahir"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   11
         Top             =   3840
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Nomor KTP"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   10
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "Tempat Lahir"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   9
         Top             =   3120
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Jenis Kelamin"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   8
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Nama"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   7
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Nomor KK"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   6
         Top             =   480
         Width           =   1095
      End
   End
   Begin Crystal.CrystalReport CR1 
      Left            =   2880
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   1200
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
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
      Caption         =   "Adodc2"
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
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
      RecordSource    =   "penduduk"
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
   Begin VB.Label Label5 
      Caption         =   "Alamat"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7200
      TabIndex        =   14
      Top             =   3240
      Width           =   1815
   End
End
Attribute VB_Name = "penduduk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'========================FORM PENDUDUK CODE==========================='
     '======================= REZA GUNAWAN ==========================='
     'MENAMPILKAN DATA PADA DATABASE KE COMBO
Sub tambahcom()
Adodc2.ConnectionString = conn.ConnectionString
Adodc2.RecordSource = "select* from combo"

For Each gosong In Me.Controls
If TypeOf gosong Is ComboBox Then
gosong.Text = ""
With Adodc2.Recordset
    Do While Not .EOF
    On Error Resume Next
    Combo1.AddItem !pekerjaan
    Combo2.AddItem !pendidikan
    Combo3.AddItem !alamat
    Combo4.AddItem !kewarganegaraan
    combo_agama.AddItem !agama
    Combo5.AddItem !status_keluarga
    Combo6.AddItem !status_perkawinan
    
    'Text9.Text = !lurah
    'Text10.Text = !camat
    Text7.AddItem !keperluan
    .MoveNext
    Loop
End With
End If
Next
End Sub

Private Sub Command3_Click()
 If Combo7 = "" Then
  MsgBox "SORTIR DAHULU DATA YANG AKAN ANDA TAMPILKAN !", vbInformation, "PERHATIAN !"
  ElseIf Combo7.Text = "SEMUA" Then
    CR1.Reset
With CR1
    .ReportFileName = App.Path & "\Data_penduduk.rpt"
    .WindowState = crptMaximized
    .RetrieveDataFiles
    .Action = 1
End With
Else
'WKWK
With CR1
    .SelectionFormula = "{penduduk.no_kk}='" & Text1.Text & "'"
    .ReportFileName = App.Path & "\keluarga.rpt"
    .WindowState = crptMaximized
    .RetrieveDataFiles
    .Action = 1
End With
End If
End Sub

'WAKTU
Private Sub Timer1_Timer()
'Label77.Caption = Format(Now, "hh : mm : ss")
'Label88.Caption = Format(Now, "dd MMMM yyyy")
End Sub

'FORMAT WAKTU DATAGRID
Sub pormat()
DataGrid1.Columns(4).NumberFormat = ("DD/MM/YYYY")
End Sub
     
'CLEAR FORM
Sub bersih()
Text1 = ""
Text2 = ""
Text3 = ""
Text4 = ""
Option1.Value = False
Option2.Value = False
DTPicker2.Value = Now
Combo1 = ""
Combo2 = ""
Combo3 = ""
Combo5 = ""
Combo6 = ""
Combo4 = ""
combo_agama = ""
Text6 = ""
Text7 = ""
Text8 = ""
Text9 = ""
End Sub

'ENABLE TRUE FORM
Sub tambah()
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Combo4.Enabled = True
Combo5.Enabled = True
Combo6.Enabled = True
Option1.Enabled = True
Option2.Enabled = True
DTPicker2.Enabled = True
combo_agama.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text8.Enabled = True
Text9.Enabled = True
Text1.SetFocus
End Sub

'TAMBAH
Private Sub Command1_Click()
Call bersih
Call tambah
Command2.Enabled = True
End Sub

'CARI DATA
Private Sub Command5_Click()
Adodc1.Recordset.Filter = "nama like '%" + Me.Text5.Text + "%' or no_kk like '%" + Me.Text5.Text + "%'or nik like '%" + Me.Text5.Text + "%'"
End Sub

'MUNCULKAN DATA SAAT PENCARIAN BERAKHIR
Private Sub Text5_Change()
If Text5.Text = "" Then
Adodc1.Refresh
Else
'wkwk
End If
End Sub

'PINDAH DATA DARI DATAGRIDVIEW KE TEXTBOX
Private Sub DataGrid1_Click()
Text1 = Adodc1.Recordset!no_kk
Text2 = Adodc1.Recordset!nik
Text3 = Adodc1.Recordset!nama
Text4 = Adodc1.Recordset!tempat_lahir
DTPicker2 = Adodc1.Recordset!tanggal_lahir
If Adodc1.Recordset!jk = "Laki-Laki" Then
    Option1.Value = True
ElseIf Adodc1.Recordset!jk = "Perempuan" Then
    Option2.Value = True
End If
Combo1 = Adodc1.Recordset!pekerjaan
Combo2 = Adodc1.Recordset!pendidikan
Combo3 = Adodc1.Recordset!alamat
Combo4 = Adodc1.Recordset!kewarganegaraan
combo_agama = Adodc1.Recordset!agama
Combo5 = Adodc1.Recordset!status_keluarga
Combo6 = Adodc1.Recordset!status_perkawinan
Text6 = Adodc1.Recordset!no_paspor
Text7 = Adodc1.Recordset!no_kitas
Text8 = Adodc1.Recordset!nama_ayah
Text9 = Adodc1.Recordset!nama_ibu
Command2.Enabled = False
End Sub

'LOAD
Private Sub Form_Load()
Call bersih
Call tambahcom

'NO TIME
Call pormat
End Sub


'SIMPAN DATA
Private Sub Command2_Click()
If Text1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Or (Option1.Value = False And Option2.Value = False) Or Combo1 = "" Or Combo2 = "" Or Combo3 = "" Or Combo4 = "" Or Combo5 = "" Or Combo6 = "" Or combo_agama = "" Or Text8 = "" Or Text9 = "" Then
MsgBox "LENGKAPI DAHULU DATA YANG AKAN ANDA INPUTKAN !", vbInformation, "PERHATIAN !"
Else
Adodc1.Recordset.AddNew
Adodc1.Recordset!no_kk = Text1.Text
Adodc1.Recordset!nik = Text2.Text
Adodc1.Recordset!nama = Text3.Text
If Option1.Value = True Then
    Adodc1.Recordset!jk = Option1.Caption
ElseIf Option2.Value = True Then
    Adodc1.Recordset!jk = Option2.Caption
End If
Adodc1.Recordset!tempat_lahir = Text4.Text
Adodc1.Recordset!tanggal_lahir = DTPicker2.Value
Adodc1.Recordset!pekerjaan = Combo1.Text
Adodc1.Recordset!pendidikan = Combo2.Text
Adodc1.Recordset!alamat = Combo3.Text
Adodc1.Recordset!kewarganegaraan = Combo4.Text
Adodc1.Recordset!agama = combo_agama.Text
Adodc1.Recordset!status_keluarga = Combo5.Text
Adodc1.Recordset!status_perkawinan = Combo6.Text
Adodc1.Recordset!no_paspor = Text6.Text
Adodc1.Recordset!no_kitas = Text7.Text
Adodc1.Recordset!nama_ayah = Text8.Text
Adodc1.Recordset!nama_ibu = Text9.Text
Adodc1.Recordset.Update
Call bersih
MsgBox "DATA ANDA BERHASIL DISIMPAN !", vbInformation, "INFORMASI !"
Adodc1.Refresh
End If
End Sub

'UBAH
Private Sub Command6_Click()
If Text1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Or (Option1.Value = False And Option2.Value = False) Or Combo1 = "" Or Combo2 = "" Or Combo3 = "" Or Combo4 = "" Or Combo5 = "" Or Combo6 = "" Or combo_agama = "" Or Text8 = "" Or Text9 = "" Then
MsgBox "LENGKAPI DAHULU DATA YANG AKAN ANDA UBAH !", vbInformation, "PERHATIAN !"
Else

Adodc1.Recordset!no_kk = Text1.Text
Adodc1.Recordset!nik = Text2.Text
Adodc1.Recordset!nama = Text3.Text
If Option1.Value = True Then
    Adodc1.Recordset!jk = Option1.Caption
ElseIf Option2.Value = True Then
    Adodc1.Recordset!jk = Option2.Caption
End If
Adodc1.Recordset!tempat_lahir = Text4.Text
Adodc1.Recordset!tanggal_lahir = DTPicker2.Value
Adodc1.Recordset!pekerjaan = Combo1.Text
Adodc1.Recordset!pendidikan = Combo2.Text
Adodc1.Recordset!alamat = Combo3.Text
Adodc1.Recordset!kewarganegaraan = Combo4.Text
Adodc1.Recordset!agama = combo_agama.Text
Adodc1.Recordset!status_keluarga = Combo5.Text
Adodc1.Recordset!status_perkawinan = Combo6.Text
Adodc1.Recordset!no_paspor = Text6.Text
Adodc1.Recordset!no_kitas = Text7.Text
Adodc1.Recordset!nama_ayah = Text8.Text
Adodc1.Recordset!nama_ibu = Text9.Text
Adodc1.Recordset.Update
Call bersih
MsgBox "DATA ANDA BERHASIL DIUBAH !", vbInformation, "INFORMASI !"
Adodc1.Refresh
End If
End Sub

'HAPUS DATA
Private Sub Command7_Click()
If Text1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Or (Option1.Value = False And Option2.Value = False) Or Combo1 = "" Or Combo2 = "" Or Combo3 = "" Or Combo4 = "" Or Combo5 = "" Or Combo6 = "" Or combo_agama = "" Or Text8 = "" Or Text9 = "" Then
MsgBox "LENGKAPI DAHULU DATA YANG AKAN ANDA HAPUS !", vbInformation, "PERHATIAN !"
Else
xx = MsgBox("Apakah Anda yakin akan menghapus data?", vbOKCancel, "Peringatan")
            If xx = vbOK Then
               Adodc1.Recordset.Delete
               Call bersih
MsgBox "DATA ANDA BERHASIL DIHAPUS !", vbInformation, "INFORMASI !"
Adodc1.Refresh
            End If
End If
End Sub



