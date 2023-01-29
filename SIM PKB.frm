VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6750
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9975
   LinkTopic       =   "Form1"
   ScaleHeight     =   6750
   ScaleWidth      =   9975
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2175
      Left            =   0
      TabIndex        =   16
      Top             =   4560
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   3836
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
            LCID            =   1033
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
            LCID            =   1033
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
      Caption         =   "Hapus"
      Height          =   495
      Left            =   7320
      TabIndex        =   15
      Top             =   3600
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Simpan"
      Height          =   495
      Left            =   5280
      TabIndex        =   14
      Top             =   3600
      Width           =   1575
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   7200
      TabIndex        =   13
      Top             =   1560
      Width           =   2055
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   1680
      TabIndex        =   10
      Top             =   2880
      Width           =   2295
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   1680
      TabIndex        =   9
      Top             =   2160
      Width           =   2295
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   1680
      TabIndex        =   8
      Top             =   1680
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1680
      TabIndex        =   7
      Top             =   1200
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1680
      TabIndex        =   6
      Top             =   720
      Width           =   2295
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   615
      Left            =   7560
      Top             =   4920
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1085
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
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
   Begin VB.Label Label8 
      Caption         =   "Biaya Pengujian"
      Height          =   375
      Left            =   5520
      TabIndex        =   12
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "Pembayaran Retribusi"
      Height          =   375
      Left            =   5400
      TabIndex        =   11
      Top             =   840
      Width           =   3855
   End
   Begin VB.Label Label6 
      Caption         =   "Jenis Kendaraan"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "Nomor Kendaraan"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Alamat"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Nama Pemilik"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "No. Uji"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "PERMOHNAN PENDAFTARAN UJI KENDARAAN BERMOTOR"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As New ADODB.Connection
Dim RS As New ADODB.Recordset
Dim koneksi As String
Sub clean()
Text1 = ""
Text2 = ""
Text3 = ""
Text4 = ""
Text5 = ""
Text6 = ""


End Sub
Sub readdata()
Dim result As String
result = "select * from karyawan "
conn.Execute (result)
Adodc1.RecordSource = result
Set DataGrid1.DataSource = Adodc1
Adodc1.Refresh

End Sub

Private Sub Command1_Click()
  result = "select * from karyawan where no_uji = '" & Text1 & "' "
    Set RS = conn.Execute(result)
    If RS.EOF Then
        If MsgBox("Data Akan Disimpan", 36, "Informasi") = vbYes Then
            saves = "insert into karyawan values('" & Text1 & "', '" & Text2 & "', '" & Text3 & "', '" & Text4 & "', '" & Text5 & "', '" & Text6 & "')"
            conn.Execute (saves)
            
        End If
    Call readdata
    Else
        MsgBox ("no_uji Tersedia")
    End If

Call clean
End Sub

Private Sub Command2_Click()
Call clean

End Sub

Private Sub Form_Load()
koneksi = "provider = microsoft.jet.oledb.4.0; data source = " & App.Path & "\pendaftaran.mdb"
Adodc1.ConnectionString = koneksi
conn.Open koneksi
Call readdata
End Sub
