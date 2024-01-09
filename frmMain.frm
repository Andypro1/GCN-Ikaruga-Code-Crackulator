VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ikaruga NetRanking Code Calculator"
   ClientHeight    =   3795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9285
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   9285
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkDisable 
      Caption         =   "Disable 33.5m+ hack"
      Height          =   195
      Left            =   1380
      TabIndex        =   30
      Top             =   900
      Width           =   1815
   End
   Begin VB.Frame fraPlayers 
      Caption         =   "Players"
      Height          =   795
      Left            =   6900
      TabIndex        =   27
      Top             =   180
      Width           =   1335
      Begin VB.OptionButton opt2P 
         Caption         =   "2P"
         Height          =   195
         Left            =   180
         TabIndex        =   29
         Top             =   480
         Width           =   795
      End
      Begin VB.OptionButton opt1P 
         Caption         =   "1P"
         Height          =   195
         Left            =   180
         TabIndex        =   28
         Top             =   240
         Width           =   795
      End
   End
   Begin VB.Frame fraRegion 
      Caption         =   "Region"
      Height          =   795
      Left            =   5220
      TabIndex        =   24
      Top             =   180
      Width           =   1515
      Begin VB.OptionButton optPAL 
         Caption         =   "PAL"
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   480
         Width           =   1215
      End
      Begin VB.OptionButton optNTSC 
         Caption         =   "NTSC"
         Height          =   195
         Left            =   240
         TabIndex        =   25
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "&About..."
      Height          =   315
      Left            =   8340
      TabIndex        =   23
      Top             =   0
      Width           =   915
   End
   Begin VB.Frame fraValidity 
      Caption         =   "Validity"
      Height          =   915
      Left            =   5640
      TabIndex        =   21
      Top             =   2820
      Width           =   3495
      Begin VB.Label lblValidity 
         Height          =   615
         Left            =   120
         TabIndex        =   22
         Top             =   180
         Width           =   3315
      End
   End
   Begin VB.TextBox txtResult8Bin 
      Height          =   285
      Left            =   1320
      TabIndex        =   16
      Top             =   2220
      Width           =   7935
   End
   Begin VB.TextBox txtScoreBin 
      Height          =   285
      Left            =   1320
      TabIndex        =   14
      Top             =   3480
      Width           =   3735
   End
   Begin VB.TextBox txtScore 
      Height          =   285
      Left            =   1320
      TabIndex        =   12
      Top             =   3060
      Width           =   1695
   End
   Begin VB.Frame fraMode 
      Caption         =   "Game Mode"
      Height          =   795
      Left            =   3540
      TabIndex        =   8
      Top             =   180
      Width           =   1515
      Begin VB.OptionButton optProto 
         Caption         =   "Prototype"
         Height          =   195
         Left            =   180
         TabIndex        =   10
         Top             =   540
         Width           =   1215
      End
      Begin VB.OptionButton optArcade 
         Caption         =   "Arcade"
         Height          =   255
         Left            =   180
         TabIndex        =   9
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.TextBox txtZeroCodeBin 
      Height          =   285
      Left            =   1320
      TabIndex        =   7
      Top             =   1500
      Width           =   7935
   End
   Begin VB.TextBox txtCodeBin 
      Height          =   285
      Left            =   1320
      TabIndex        =   6
      Top             =   1200
      Width           =   7935
   End
   Begin VB.TextBox txtZeroCode 
      Height          =   285
      Left            =   1380
      TabIndex        =   4
      Top             =   600
      Width           =   1695
   End
   Begin VB.TextBox txtResultBin 
      Height          =   285
      Left            =   1320
      TabIndex        =   2
      Top             =   1920
      Width           =   7935
   End
   Begin VB.TextBox txtCode 
      Height          =   285
      Left            =   1380
      TabIndex        =   0
      Top             =   300
      Width           =   1695
   End
   Begin VB.Label Label10 
      Caption         =   "9 Pos. Result:"
      Height          =   255
      Left            =   300
      TabIndex        =   20
      Top             =   2280
      Width           =   1035
   End
   Begin VB.Label Label9 
      Caption         =   "12 Pos. Result:"
      Height          =   255
      Left            =   180
      TabIndex        =   19
      Top             =   1980
      Width           =   1155
   End
   Begin VB.Label Label8 
      Caption         =   "Binary Zero Code:"
      Height          =   255
      Left            =   0
      TabIndex        =   18
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label7 
      Caption         =   "Binary Code:"
      Height          =   255
      Left            =   360
      TabIndex        =   17
      Top             =   1260
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "In Binary:"
      Height          =   255
      Left            =   600
      TabIndex        =   15
      Top             =   3540
      Width           =   675
   End
   Begin VB.Label Label5 
      Caption         =   "Score:"
      Height          =   255
      Left            =   780
      TabIndex        =   13
      Top             =   3120
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "=============================================================================================================="
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Top             =   2580
      Width           =   9255
   End
   Begin VB.Label Label3 
      Caption         =   "Zero code:"
      Height          =   255
      Left            =   600
      TabIndex        =   5
      Top             =   660
      Width           =   795
   End
   Begin VB.Label Label2 
      Caption         =   $"frmMain.frx":0000
      Height          =   255
      Left            =   1320
      TabIndex        =   3
      Top             =   1740
      Width           =   7935
   End
   Begin VB.Label Label1 
      Caption         =   "Enter code:"
      Height          =   255
      Left            =   540
      TabIndex        =   1
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkDisable_Click()
    GenerateAnalysis
End Sub

Private Sub cmdAbout_Click()
    MsgBox "This is the NetRanking Crackulator" & vbCrLf & _
    "             made by Andypro" & vbCrLf & vbCrLf & _
    "         Version Number: " & App.Major & "." & App.Minor & "." & App.Revision, , "About"
End Sub

Private Sub Form_Load()
    Me.optArcade.Value = True
    Me.optProto.Value = False
    Me.optNTSC.Value = True
    Me.optPAL.Value = False
    Me.opt1P.Value = True
    Me.opt2P.Value = False
    Me.chkDisable.Value = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ExitProgram
End Sub

Private Sub opt1P_Click()
    GenerateAnalysis
End Sub

Private Sub opt2P_Click()
    GenerateAnalysis
End Sub

Private Sub optArcade_Click()
    GenerateAnalysis
End Sub

'  This routine generates the zero code and outputs the codes in
'  binary as well.  If a full code has been entered, it performs the XOR
'  and generates a score as well as checking on the validity of the code.
'  The functions called from here can be found in basCalcs.bas.
Private Sub GenerateAnalysis()
    Dim intValid As Integer
    Me.txtZeroCode.Text = GetZeroCode(Trim(Me.txtCode.Text), IIf(Me.optArcade.Value = True, True, False))
    Me.txtCodeBin.Text = GetCodeBinText(Trim(Me.txtCode.Text))
    Me.txtZeroCodeBin.Text = GetZeroCodeBinText(Trim(Me.txtZeroCode.Text))

    If Len(Trim(Me.txtCode.Text)) = 12 And Len(Trim(Me.txtZeroCode.Text)) = 12 Then
        Me.txtResultBin.Text = GetResultBinText(Trim(Me.txtCodeBin.Text), Trim(Me.txtZeroCodeBin.Text))
        Me.txtResult8Bin.Text = GetResult8BinText(Trim(Me.txtResultBin.Text))
    End If

    If Len(Trim(Me.txtCode.Text)) = 12 Then
        AttemptScoreGuess
        intValid = CheckCodeValidity(Trim(Replace(Me.txtResultBin.Text, " ", "")), Trim(Me.txtScore.Text), Trim(Me.txtScoreBin.Text))

        Select Case intValid
            Case 0
                Me.lblValidity.Caption = "This code appears to be valid." & vbCrLf & _
                "It passes all five tests."
            Case 1
                Me.lblValidity.Caption = "This code is INVALID." & vbCrLf & _
                "It fails the pos. 9 and 10 0's test."
            Case 2
                Me.lblValidity.Caption = "This code is INVALID." & vbCrLf & _
                "It fails the pos. 11 and 12 duplicate test" & vbCrLf & _
                "(or multiple tests)."
            Case 3
                Me.lblValidity.Caption = "This code is INVALID." & vbCrLf & _
                "The score is not divisible by 10" & vbCrLf & _
                "(or fails multiple tests)."
            Case 4
                Me.lblValidity.Caption = "This code is INVALID." & vbCrLf & _
                "The code contains one or more invalid characters" & vbCrLf & _
                "(or fails multiple tests)."
            Case 5
                Me.lblValidity.Caption = "This code is INVALID." & vbCrLf & _
                "The score is too high" & vbCrLf & _
                "(or fails multiple tests)."
            Case Else
        End Select

        Me.txtScore.SetFocus
    Else
        Me.txtScore.Text = ""
        Me.lblValidity.Caption = ""
    End If
End Sub

Private Sub optNTSC_Click()
    GenerateAnalysis
End Sub

Private Sub optPAL_Click()
    GenerateAnalysis
End Sub

Private Sub optProto_Click()
    GenerateAnalysis
End Sub

Private Sub txtCode_Change()
    GenerateAnalysis
End Sub

Private Sub txtScore_Change()
    Me.txtScoreBin.Text = GetScoreBinText(Trim(Me.txtScore.Text))
End Sub
