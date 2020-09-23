VERSION 5.00
Object = "*\AGandaraControls.vbp"
Begin VB.Form frmClient 
   Caption         =   "Form1"
   ClientHeight    =   4740
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   ScaleHeight     =   4740
   ScaleWidth      =   6975
   StartUpPosition =   3  'Windows-Standard
   Begin GandaraControls.gucScrWin gucScrWin1 
      Height          =   1545
      Left            =   2730
      TabIndex        =   0
      Top             =   2760
      Width           =   3645
      _extentx        =   6429
      _extenty        =   2725
   End
   Begin GandaraControls.gucScrWin gucScrWin2 
      Height          =   3855
      Left            =   330
      TabIndex        =   5
      Top             =   360
      Width           =   2175
      _extentx        =   3836
      _extenty        =   6800
   End
   Begin VB.Label lblValue2 
      Caption         =   "Actual v value 2"
      Height          =   225
      Index           =   1
      Left            =   1920
      TabIndex        =   4
      Top             =   90
      Width           =   1155
   End
   Begin VB.Label lblValue1 
      Caption         =   "Actual v value 1"
      Height          =   225
      Index           =   1
      Left            =   5730
      TabIndex        =   3
      Top             =   2490
      Width           =   1155
   End
   Begin VB.Label lblValue1 
      Caption         =   "Actual h value 1"
      Height          =   225
      Index           =   0
      Left            =   3930
      TabIndex        =   2
      Top             =   4350
      Width           =   1155
   End
   Begin VB.Label lblValue2 
      Caption         =   "Actual h value 2"
      Height          =   225
      Index           =   0
      Left            =   810
      TabIndex        =   1
      Top             =   4230
      Width           =   1155
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mncExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Click()
    gucScrWin1.HideScrollbars
    gucScrWin2.HideScrollbars
End Sub

Private Sub Form_DblClick()
    gucScrWin1.ShowScrollbars
    gucScrWin2.ShowScrollbars
End Sub

Private Sub Form_Load()
    'Initialize a scrolled object.
    With gucScrWin1
        .ActiveScrollbars = egswSBDBoth
        
        .Min(egswSBOVertical) = 25
        .Max(egswSBOVertical) = 99
        .LargeChange(egswSBOVertical) = 25
        .Value(egswSBOVertical) = 50
        
        .Min(egswSBOHorizontal) = 25
        .Max(egswSBOHorizontal) = 99
        .LargeChange(egswSBOHorizontal) = 25
        .Value(egswSBOHorizontal) = 50
        
        .SetScrollbar egswSBOVertical
        .SetScrollbar egswSBOHorizontal
        
        .ShowScrollbars
    End With
    
    'And onother one to prove it's multi-object-capable.
    With gucScrWin2
        .ActiveScrollbars = egswSBDBoth
        
        .Min(egswSBOVertical) = 1
        .Max(egswSBOVertical) = 100
        .LargeChange(egswSBOVertical) = 20
        .Value(egswSBOVertical) = 50
        
        .Min(egswSBOHorizontal) = 1
        .Max(egswSBOHorizontal) = 100
        .LargeChange(egswSBOHorizontal) = 20
        .Value(egswSBOHorizontal) = 50
        
        .SetScrollbar egswSBOVertical
        .SetScrollbar egswSBOHorizontal
        
        .ShowScrollbars
    End With
End Sub

Private Sub gucScrWin1_Change(Scrollbar As GandaraControls.egswSBOrientation, Value As Long)
    lblValue1(Scrollbar).Caption = CStr(Value)
End Sub

Private Sub gucScrWin2_Change(Scrollbar As GandaraControls.egswSBOrientation, Value As Long)
    lblValue2(Scrollbar).Caption = CStr(Value)
End Sub
