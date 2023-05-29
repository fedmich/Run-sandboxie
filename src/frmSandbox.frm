VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   Caption         =   "Run in Sandbox"
   ClientHeight    =   2700
   ClientLeft      =   1095
   ClientTop       =   1425
   ClientWidth     =   6765
   Icon            =   "frmSandbox.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   2700
   ScaleWidth      =   6765
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdGO 
      Caption         =   "Run in Sandbox"
      Height          =   615
      Left            =   2715
      TabIndex        =   4
      Top             =   1680
      Width           =   1335
   End
   Begin VB.ComboBox cboProgram 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   315
      OLEDropMode     =   1  'Manual
      TabIndex        =   3
      Text            =   "C:\Windows\notepad.exe"
      Top             =   1080
      Width           =   6135
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1080
      TabIndex        =   1
      Text            =   "DEV"
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Program:"
      Height          =   195
      Left            =   315
      TabIndex        =   2
      Top             =   840
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sandbox:"
      Height          =   195
      Left            =   315
      TabIndex        =   0
      Top             =   360
      Width           =   675
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    '
    With Combo1
        .AddItem "dev"
        .AddItem "3D"
        .AddItem "games"
        .AddItem "offline"
        
        .ListIndex = 0
        
    End With
End Sub

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Data.Files.Count > 0 Then
        dropped_file Data.Files(1)
    End If
End Sub

Private Sub cboProgram_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Data.Files.Count > 0 Then
        dropped_file Data.Files(1)
    End If
End Sub


Private Sub picked_file()

End Sub

Private Sub dropped_file(file)
    
    cboProgram.Text = file
    add_to_cb cboProgram, CStr(file)
    
End Sub


Private Sub cmdGO_Click()
    Dim box$
    box = Combo1.Text
    If box = "" Then
        Combo1.SetFocus
        Exit Sub
    End If
    
    Dim exe$
    exe = cboProgram.Text
    If exe = "" Then
        cboProgram.SetFocus
        Exit Sub
    End If
    
    Dim args$
    
    'Dim cmd$
    Sandbox_app box, exe, args
    
    add_to_cb cboProgram, exe
End Sub


Private Sub add_to_cb(cb As ComboBox, file$)
    Dim i As Integer
    For i = 0 To cb.ListCount
        If file = cb.List(i) Then
            Exit Sub
        End If
    Next i
    
    cb.AddItem file
End Sub
