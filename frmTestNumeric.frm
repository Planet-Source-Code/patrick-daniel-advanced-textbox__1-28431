VERSION 5.00
Object = "*\ANumericBox.vbp"
Begin VB.Form frmTestNumeric 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Alpha And Numeric Only Textbox"
   ClientHeight    =   4170
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3540
   Icon            =   "frmTestNumeric.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   3540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   375
      Left            =   2145
      TabIndex        =   3
      Top             =   2370
      Width           =   1230
   End
   Begin NumericTextBox.NumericBox NumericBox1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   330
      Width           =   3270
      _ExtentX        =   5768
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   6
      Text            =   ""
      AdvAllowNegative=   -1  'True
      AdvDecimalPlaces=   2
      AdvUCase        =   -1  'True
      CurrencyFormat  =   -1  'True
      DisablePaste    =   -1  'True
   End
   Begin NumericTextBox.NumericBox NumericBox2 
      Height          =   285
      Left            =   135
      TabIndex        =   1
      Top             =   1140
      Width           =   3270
      _ExtentX        =   5768
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Text            =   ""
      AdvMode         =   0
      AdvDecimalPlaces=   2
      CurrencyFormat  =   -1  'True
   End
   Begin NumericTextBox.NumericBox NumericBox3 
      Height          =   285
      Left            =   150
      TabIndex        =   2
      Top             =   1980
      Width           =   3270
      _ExtentX        =   5768
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   ""
      AdvUCase        =   -1  'True
   End
   Begin VB.Label Label5 
      Caption         =   "TRUE NUMERIC VALIDATIONS!!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   90
      TabIndex        =   8
      Top             =   2865
      Width           =   3945
   End
   Begin VB.Label Label4 
      Caption         =   "This is just a small, small sample of what this thing can do.  Download it and check it out!! You will not be dissappointed."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   105
      TabIndex        =   7
      Top             =   3210
      Width           =   3300
   End
   Begin VB.Label Label3 
      Caption         =   "Force Upper Case and Overwrite Mode:"
      Height          =   255
      Left            =   150
      TabIndex        =   6
      Top             =   1680
      Width           =   3360
   End
   Begin VB.Label Label2 
      Caption         =   "Currency Format, 2 Decimals, Right Aligned:"
      Height          =   255
      Left            =   135
      TabIndex        =   5
      Top             =   825
      Width           =   3360
   End
   Begin VB.Label Label1 
      Caption         =   "Allow Negatives and 3 decimals:"
      Height          =   255
      Left            =   135
      TabIndex        =   4
      Top             =   60
      Width           =   3360
   End
End
Attribute VB_Name = "frmTestNumeric"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
   
   Unload Me
   
End Sub

Private Sub Form_Load()

End Sub

Private Sub Form_Unload(Cancel As Integer)
   
   Set frmTestNumeric = Nothing
   
   End
   
End Sub
