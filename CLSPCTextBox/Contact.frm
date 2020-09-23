VERSION 5.00
Begin VB.Form frmContact 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Contact Record"
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6300
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   6300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdMap 
      Height          =   315
      Left            =   5580
      Picture         =   "Contact.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2310
      Width           =   375
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   4110
      TabIndex        =   13
      Top             =   3930
      Width           =   1485
   End
   Begin VB.TextBox txtTelephone 
      Height          =   345
      Left            =   1290
      TabIndex        =   6
      Top             =   2955
      Width           =   2205
   End
   Begin VB.TextBox txtName 
      Height          =   345
      Left            =   1290
      TabIndex        =   0
      Top             =   465
      Width           =   4305
   End
   Begin VB.TextBox txtAdd5 
      Height          =   345
      Left            =   4350
      TabIndex        =   5
      Top             =   2295
      Width           =   1245
   End
   Begin VB.TextBox txtAdd4 
      Height          =   345
      Left            =   1290
      TabIndex        =   4
      Top             =   2295
      Width           =   2205
   End
   Begin VB.TextBox txtAdd3 
      Height          =   315
      Left            =   1290
      TabIndex        =   3
      Top             =   1920
      Width           =   4305
   End
   Begin VB.TextBox txtAdd2 
      Height          =   345
      Left            =   1290
      TabIndex        =   2
      Top             =   1500
      Width           =   4305
   End
   Begin VB.TextBox txtAdd1 
      Height          =   345
      Left            =   1290
      TabIndex        =   1
      Top             =   1095
      Width           =   4305
   End
   Begin VB.Label Label1 
      Caption         =   "Telephone"
      Height          =   195
      Index           =   5
      Left            =   330
      TabIndex        =   12
      Top             =   3030
      Width           =   1005
   End
   Begin VB.Label Label1 
      Caption         =   "Name"
      Height          =   195
      Index           =   4
      Left            =   330
      TabIndex        =   11
      Top             =   540
      Width           =   1005
   End
   Begin VB.Label Label1 
      Caption         =   "Postcode"
      Height          =   195
      Index           =   3
      Left            =   3600
      TabIndex        =   10
      Top             =   2370
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "County"
      Height          =   195
      Index           =   2
      Left            =   330
      TabIndex        =   9
      Top             =   2370
      Width           =   1005
   End
   Begin VB.Label Label1 
      Caption         =   "Town/City"
      Height          =   195
      Index           =   1
      Left            =   330
      TabIndex        =   8
      Top             =   1980
      Width           =   1005
   End
   Begin VB.Label Label1 
      Caption         =   "Address"
      Height          =   195
      Index           =   0
      Left            =   330
      TabIndex        =   7
      Top             =   1170
      Width           =   1005
   End
End
Attribute VB_Name = "frmContact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objText As New clsPCTextBox

Private Sub cmdMap_Click()
    If Len(txtAdd5) Then
        objText.Get_Map_Location
    End If
End Sub

Private Sub Form_Load()
    Set objText = New clsPCTextBox
    objText.Init txtAdd5, True
End Sub

Private Sub cmdClose_Click()
    Set objText = Nothing
    Unload Me
End Sub

