VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informazioni su..."
   ClientHeight    =   2790
   ClientLeft      =   3285
   ClientTop       =   3150
   ClientWidth     =   6285
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   6285
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   5280
      TabIndex        =   3
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "-"
      Height          =   210
      Left            =   960
      TabIndex        =   2
      Top             =   840
      Width           =   60
   End
   Begin VB.Label lblVer 
      AutoSize        =   -1  'True
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   960
      TabIndex        =   1
      Top             =   600
      Width           =   45
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "LA PIRAMIDE NUMERICA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   345
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   3525
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   510
      Left            =   240
      Picture         =   "Form2.frx":08CA
      Top             =   240
      Width           =   510
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Clicks As Integer

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyH: Image1_Click
        Case vbKeyEscape: Unload Me
    End Select
End Sub

Private Sub Form_Load()
    CF Me
    Clicks = 0
    lblVer.Caption = "Versione " & App.Major & "." & App.Minor & " revisione " & App.Revision
    Label2.Caption = "Autore: -DjGG-" & Chr(13) & _
                     "Data di creazione del progetto: 29/06/2009" & Chr(13) & _
                     "Beta tester: Patty" & Chr(13) & Chr(13) & "FREEWARE" & Chr(13) & Chr(13) & "Dedicato alla mia promessa sposa Samantha"
End Sub

Private Sub Image1_Click()
    If Clicks >= 5 Then
        inputstringa = InputBox("Inserisci il codice per attivare i trucchi", "ATTIVA TRUCCHI")
        If inputstringa = "0000" Then
            frmGame.Command1.Visible = True
            frmGame.Command2.Visible = True
            Clicks = 0
            Unload Me
        ElseIf inputstringa = "9995" Then
            SaveSetting "PNUM", "SCORE", "HISCORE", "0"
            SaveSetting "PNUM", "SCORE", "NSCORE", "<Giocatore>"
            frmGame.LoadScore
            MsgBox ("High Score resettato!"), vbExclamation, "RESET"
        Else
            MsgBox ("Codice non valido"), vbCritical, "TRUCCHI"
            Clicks = 0
            Unload Me
        End If
    Else
        Clicks = Clicks + 1
    End If
End Sub
