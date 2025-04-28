VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6780
   ClientLeft      =   5160
   ClientTop       =   2685
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   6780
   ScaleWidth      =   6585
   Begin VB.VScrollBar VScroll1 
      Height          =   1095
      Left            =   5640
      TabIndex        =   7
      Top             =   3000
      Width           =   375
   End
   Begin VB.Frame Frame1 
      Caption         =   "Controles"
      Height          =   1215
      Left            =   1920
      TabIndex        =   0
      Top             =   2880
      Width           =   3615
      Begin VB.CommandButton Command3 
         Caption         =   "Command3"
         Height          =   300
         Left            =   1320
         TabIndex        =   6
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   300
         Left            =   120
         TabIndex        =   5
         Text            =   "Text3"
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   300
         Left            =   1320
         TabIndex        =   4
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Height          =   300
         Left            =   120
         TabIndex        =   3
         Text            =   "Text2"
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   300
         Left            =   1320
         TabIndex        =   2
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Left            =   120
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   360
         Width           =   1095
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public xTop As Integer

Private Sub Form_Load()
xTop = 360

    With VScroll1
        .Min = 0
        .Max = 3
        
    End With
End Sub

Private Sub VScroll1_Change()
    Caption = VScroll1.Value
    With VScroll1
        Text1.Top = 360 - (480 * .Value)
        Command1.Top = 360 - (480 * .Value)
        Text1.Visible = Text1.Top >= 360
        Command1.Visible = Command1.Top >= 360
        
        
        Text2.Top = (360 + 480) - (480 * .Value)
        Command2.Top = (360 + 480) - (480 * .Value)
        Text2.Visible = Text2.Top >= 360
        Command2.Visible = Command2.Top >= 360
        
        Text3.Top = (360 + 960) - (480 * .Value)
        Command3.Top = (360 + 960) - (480 * .Value)
        Text3.Visible = Text3.Top >= 360
        Command3.Visible = Command3.Top >= 360
        
        
    End With
    
    
    
    
    
End Sub

Private Sub VScroll1_KeyUp(KeyCode As Integer, Shift As Integer)
'    Caption = KeyCode
End Sub

