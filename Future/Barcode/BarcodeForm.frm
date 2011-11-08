VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5235
   ClientLeft      =   1740
   ClientTop       =   2235
   ClientWidth     =   7185
   LinkTopic       =   "Form1"
   ScaleHeight     =   5235
   ScaleWidth      =   7185
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1560
      TabIndex        =   2
      Top             =   2280
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   975
      Left            =   1680
      TabIndex        =   1
      Top             =   960
      Width           =   2775
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "CIA ITF Tall"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   1080
      TabIndex        =   0
      Top             =   3000
      Width           =   5055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  Form1.Label1.Caption = BarcodeITF2from5(Form1.Text1.Text)

End Sub

Private Sub Form_Load()
    ConversionTable(0) = &H28
    ConversionTable(1) = &H202
    ConversionTable(2) = &H82
    ConversionTable(3) = &H280
    ConversionTable(4) = &H22
    ConversionTable(5) = &H220
    ConversionTable(6) = &HA0
    ConversionTable(7) = &HA
    ConversionTable(8) = &H208
    ConversionTable(9) = &H88
End Sub
