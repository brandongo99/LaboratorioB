VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7995
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5370
   LinkTopic       =   "Form1"
   ScaleHeight     =   7995
   ScaleWidth      =   5370
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command7 
      Caption         =   "Anterior"
      Height          =   495
      Left            =   120
      TabIndex        =   18
      Top             =   6120
      Width           =   735
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Agregar"
      Height          =   495
      Left            =   1080
      TabIndex        =   17
      Top             =   6120
      Width           =   735
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Eliminar"
      Height          =   495
      Left            =   2040
      TabIndex        =   16
      Top             =   6120
      Width           =   735
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Guardar"
      Height          =   495
      Left            =   3000
      TabIndex        =   15
      Top             =   6120
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Siguiente"
      Height          =   495
      Left            =   4200
      TabIndex        =   14
      Top             =   6120
      Width           =   855
   End
   Begin VB.Data Data1 
      BOFAction       =   1  'BOF
      Caption         =   "Estidiantes"
      Connect         =   "Access"
      DatabaseName    =   "C:\Users\estudiante\Desktop\Brandon\SVisualDB\SVisualDB.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      EOFAction       =   1  'EOF
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Estudiantes"
      Top             =   6960
      Width           =   2415
   End
   Begin VB.TextBox Text7 
      DataField       =   "Nombres"
      DataSource      =   "Data1"
      Height          =   405
      Left            =   2520
      TabIndex        =   13
      Top             =   1680
      Width           =   2535
   End
   Begin VB.TextBox Text6 
      DataField       =   "Apellidos"
      DataSource      =   "Data1"
      Height          =   405
      Left            =   2520
      TabIndex        =   12
      Top             =   2280
      Width           =   2535
   End
   Begin VB.TextBox Text5 
      DataField       =   "Edad"
      DataSource      =   "Data1"
      Height          =   405
      Left            =   2520
      TabIndex        =   11
      Top             =   2880
      Width           =   2535
   End
   Begin VB.TextBox Text4 
      DataField       =   "Facultad"
      DataSource      =   "Data1"
      Height          =   405
      Left            =   2520
      TabIndex        =   10
      Top             =   3480
      Width           =   2535
   End
   Begin VB.TextBox Text3 
      DataField       =   "Semestre"
      DataSource      =   "Data1"
      Height          =   405
      Left            =   2520
      TabIndex        =   9
      Top             =   4080
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      DataField       =   "Carné"
      DataSource      =   "Data1"
      Height          =   405
      Left            =   2520
      TabIndex        =   8
      Top             =   1080
      Width           =   2535
   End
   Begin VB.Image Image2 
      Height          =   1455
      Left            =   3120
      Picture         =   "Form1.frx":0000
      Stretch         =   -1  'True
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   1425
      Left            =   3120
      Picture         =   "Form1.frx":1BF64
      Stretch         =   -1  'True
      Top             =   4560
      Width           =   1425
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Nombres"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   240
      TabIndex        =   7
      Top             =   1800
      Width           =   1170
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Apellidos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   240
      TabIndex        =   6
      Top             =   2400
      Width           =   1170
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Edad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   240
      TabIndex        =   5
      Top             =   3000
      Width           =   675
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Facultad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   240
      TabIndex        =   4
      Top             =   3600
      Width           =   1080
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Semestre"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   240
      TabIndex        =   3
      Top             =   4200
      Width           =   1200
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Foto (File)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   240
      TabIndex        =   2
      Top             =   4800
      Width           =   1290
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Carné"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Estudiantes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   1440
      TabIndex        =   0
      Top             =   360
      Width           =   2505
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Data1.Recordset.MoveNext
If Data1.Recordset.EOF Then
Data1.Recordset.MoveFirst
End If
If Data1.Recordset.BOF = True Then
Image1.Visible = True
Image2.Visible = False
End If
If Data1.Recordset.EOF = True Then
Image2.Visible = True
Image1.Visible = False
End If
End Sub

Private Sub Command4_Click()
Data1.Recordset.Update
End Sub

Private Sub Command5_Click()
Data1.Recordset.Delete
End Sub

Private Sub Command6_Click()
Data1.Recordset.AddNew
End Sub

Private Sub Command7_Click()
Data1.Recordset.MovePrevious
If Data1.Recordset.BOF Then
Data1.Recordset.MoveLast
End If
If Data1.Recordset.BOF = False Then
Image1.Visible = True
Image2.Visible = False
End If
If Data1.Recordset.EOF = False Then
Image2.Visible = True
Image1.Visible = False
End If
End Sub

