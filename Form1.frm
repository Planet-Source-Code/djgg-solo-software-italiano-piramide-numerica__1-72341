VERSION 5.00
Begin VB.Form frmGame 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Piramide numerica  Versione 1.0 rev. 0 -DjGG-"
   ClientHeight    =   6615
   ClientLeft      =   2880
   ClientTop       =   1755
   ClientWidth     =   8880
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   8880
   Begin VB.Timer DEMO 
      Interval        =   150
      Left            =   600
      Top             =   960
   End
   Begin VB.ListBox List1 
      Height          =   5520
      Left            =   6120
      TabIndex        =   26
      Top             =   960
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Soluzione"
      Height          =   615
      Left            =   240
      TabIndex        =   25
      Top             =   2520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ferma il tempo"
      Height          =   615
      Left            =   240
      TabIndex        =   24
      Top             =   1680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer CDOWN 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   120
      Top             =   960
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   120
      Top             =   480
   End
   Begin VB.Timer Timer1 
      Interval        =   150
      Left            =   120
      Top             =   0
   End
   Begin VB.TextBox T15 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   5160
      Locked          =   -1  'True
      TabIndex        =   14
      Text            =   "T15"
      Top             =   5610
      Width           =   735
   End
   Begin VB.TextBox T14 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   3960
      Locked          =   -1  'True
      TabIndex        =   13
      Text            =   "T14"
      Top             =   5610
      Width           =   735
   End
   Begin VB.TextBox T13 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   12
      Text            =   "T13"
      Top             =   5610
      Width           =   735
   End
   Begin VB.TextBox T12 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   11
      Text            =   "T12"
      Top             =   5610
      Width           =   735
   End
   Begin VB.TextBox T11 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   360
      Locked          =   -1  'True
      TabIndex        =   10
      Text            =   "T11"
      Top             =   5610
      Width           =   735
   End
   Begin VB.TextBox T10 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   9
      Text            =   "T10"
      Top             =   4530
      Width           =   735
   End
   Begin VB.TextBox T9 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   "T9"
      Top             =   4530
      Width           =   735
   End
   Begin VB.TextBox T8 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   "T8"
      Top             =   4530
      Width           =   735
   End
   Begin VB.TextBox T7 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   6
      Text            =   "T7"
      Top             =   4530
      Width           =   735
   End
   Begin VB.TextBox T6 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   3960
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "T6"
      Top             =   3450
      Width           =   735
   End
   Begin VB.TextBox T5 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "T5"
      Top             =   3450
      Width           =   735
   End
   Begin VB.TextBox T4 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "T4"
      Top             =   3450
      Width           =   735
   End
   Begin VB.TextBox T3 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "T3"
      Top             =   2370
      Width           =   735
   End
   Begin VB.TextBox T2 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "T2"
      Top             =   2370
      Width           =   735
   End
   Begin VB.TextBox T1 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "T1"
      Top             =   1410
      Width           =   735
   End
   Begin VB.Label HScore 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   7410
      TabIndex        =   28
      Top             =   5760
      Width           =   195
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Punteggio più alto:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   6480
      TabIndex        =   27
      Top             =   5400
      Width           =   2055
   End
   Begin VB.Shape Shape19 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   5
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   1095
      Left            =   6360
      Shape           =   4  'Rounded Rectangle
      Top             =   5280
      Width           =   2295
   End
   Begin VB.Line Line3 
      BorderColor     =   &H0000FFFF&
      X1              =   6600
      X2              =   8400
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Livello"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Left            =   7155
      TabIndex        =   23
      Top             =   4080
      Width           =   675
   End
   Begin VB.Label LivelloLBL 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   435
      Left            =   7395
      TabIndex        =   22
      Top             =   4320
      Width           =   225
   End
   Begin VB.Label TR 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "120"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   435
      Left            =   7200
      TabIndex        =   21
      Top             =   3480
      Width           =   615
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tempo rimanente:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Left            =   6600
      TabIndex        =   20
      Top             =   3240
      Width           =   1785
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0000FFFF&
      X1              =   6600
      X2              =   8400
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Label Punteggio 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   435
      Left            =   7380
      TabIndex        =   19
      Top             =   2520
      Width           =   225
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Punteggio:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Left            =   7005
      TabIndex        =   18
      Top             =   2280
      Width           =   1065
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      X1              =   6600
      X2              =   8400
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Label DT 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   7455
      TabIndex        =   17
      Top             =   1320
      Width           =   105
   End
   Begin VB.Shape Shape18 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   5
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   4095
      Left            =   6360
      Shape           =   4  'Rounded Rectangle
      Top             =   960
      Width           =   2295
   End
   Begin VB.Shape Shape17 
      FillStyle       =   0  'Solid
      Height          =   4095
      Left            =   6480
      Shape           =   4  'Rounded Rectangle
      Top             =   1080
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-DjGG-"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   165
      Left            =   8265
      TabIndex        =   16
      Top             =   600
      Width           =   435
   End
   Begin VB.Label Titolo 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "LA PIRAMIDE NUMERICA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   555
      Left            =   435
      TabIndex        =   15
      Top             =   240
      Width           =   5745
   End
   Begin VB.Shape Shape15 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   975
      Left            =   5040
      Shape           =   2  'Oval
      Top             =   5280
      Width           =   975
   End
   Begin VB.Shape Shape14 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   975
      Left            =   3840
      Shape           =   2  'Oval
      Top             =   5280
      Width           =   975
   End
   Begin VB.Shape Shape13 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   975
      Left            =   2640
      Shape           =   2  'Oval
      Top             =   5280
      Width           =   975
   End
   Begin VB.Shape Shape12 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   975
      Left            =   1440
      Shape           =   2  'Oval
      Top             =   5280
      Width           =   975
   End
   Begin VB.Shape Shape11 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   975
      Left            =   240
      Shape           =   2  'Oval
      Top             =   5280
      Width           =   975
   End
   Begin VB.Shape Shape10 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   975
      Left            =   4440
      Shape           =   2  'Oval
      Top             =   4200
      Width           =   975
   End
   Begin VB.Shape Shape9 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   975
      Left            =   3240
      Shape           =   2  'Oval
      Top             =   4200
      Width           =   975
   End
   Begin VB.Shape Shape8 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   975
      Left            =   2040
      Shape           =   2  'Oval
      Top             =   4200
      Width           =   975
   End
   Begin VB.Shape Shape7 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   975
      Left            =   840
      Shape           =   2  'Oval
      Top             =   4200
      Width           =   975
   End
   Begin VB.Shape Shape6 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   975
      Left            =   3840
      Shape           =   2  'Oval
      Top             =   3120
      Width           =   975
   End
   Begin VB.Shape Shape5 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   975
      Left            =   2640
      Shape           =   2  'Oval
      Top             =   3120
      Width           =   975
   End
   Begin VB.Shape Shape4 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   975
      Left            =   1440
      Shape           =   2  'Oval
      Top             =   3120
      Width           =   975
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   975
      Left            =   3240
      Shape           =   2  'Oval
      Top             =   2040
      Width           =   975
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   975
      Left            =   2040
      Shape           =   2  'Oval
      Top             =   2040
      Width           =   975
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   975
      Left            =   2640
      Shape           =   2  'Oval
      Top             =   1080
      Width           =   975
   End
   Begin VB.Shape Shape16 
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   0
      Top             =   120
      Width           =   9735
   End
   Begin VB.Shape Shape20 
      FillStyle       =   0  'Solid
      Height          =   1095
      Left            =   6480
      Shape           =   4  'Rounded Rectangle
      Top             =   5400
      Width           =   2295
   End
   Begin VB.Menu mnuGame 
      Caption         =   "&Gioco"
      Begin VB.Menu mnuNew 
         Caption         =   "&Nuovo"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnubar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Esci"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&?"
      Begin VB.Menu mnuAbout 
         Caption         =   "&Informazioni su..."
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "frmGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------------------------------
'Nome progetto: Piramide numerica
'Autore: -DjGG-
'BETA Tester: Patty
'Data di creazione: 29/06/2009
'Data di ultimazione: 03/08/2009
'Descrizione di questo programa: Gioco matematico, il cui scopo è quello di indovinare il valore numerico
'                                delle caselle vuote e completare la piramide.
'Regole del gioco: Il numero da inserire nelle palline centrali è dato dalla somma delle due sottostanti
'Esempio:
'
'                       (___)
'
'                    (___)  (___)
'
'                 (___) (___) (___)
'
'               ( 3 ) ( 6 ) ( 9 ) ( 11 )
'
'            ( 1 ) ( 2 ) ( 4 ) ( 5 ) ( 6 )
'
'Note aggiuntive: Dedicato alla mia promessa sposa Samantha.
'-----------------------------------------------------------------------------------------------------------
'
'SCHEMA DELLE CASELLE DI TESTO
'
'              ( T1 ) LINEA5
'
'          ( T2 )  ( T3 ) LINEA4
'
'       ( T4 ) ( T5 ) ( T6 ) LINEA3
'
'   ( T7 ) ( T8 ) ( T9 ) ( T10 ) LINEA2
'
'( T11 )( T12 )( T13 )( T14 )( T15 ) LINEA1
'
'Dichiaro tutte le variabili
Dim Livello, TempoR, OldTempoR, Punti, R1, R2, R3, R4, R5, R6, R7, R8, R9, R10, R11, R12, R13, R14, R15 As Integer
Dim C1, C2, C3, C4, C5, C6, C7, C8, C9, C10, Abilita As Boolean
Dim numROSSI, ShapeC As Integer
Dim mpunt As String
Dim mnome As String
Public Sub NewGame()
Dim x, y As Integer
'Prepara il gioco
    Abilita = False
    RESET
    ATW
    RRESET
    LivelloLBL.Caption = Livello
    numROSSI = 0
    LoadScore
'Genera dei numeri casuali sulle prime 5 caselle della LINEA1
    y = Int(Rnd * 10)
    T11.Text = y
    R11 = T11.Text
    y = Int(Rnd * 10)
    T12.Text = y
    R12 = T12.Text
    y = Int(Rnd * 10)
    T13.Text = y
    R13 = T13.Text
    y = Int(Rnd * 10)
    T14.Text = y
    R14 = T14.Text
    y = Int(Rnd * 10)
    T15.Text = y
    R15 = T15.Text
'Calcola i risultati per la LINEA2
    T7.Text = Val(T11.Text) + Val(T12.Text)
    R7 = T7.Text
    T8.Text = Val(T12.Text) + Val(T13.Text)
    R8 = T8.Text
    T9.Text = Val(T13.Text) + Val(T14.Text)
    R9 = T9.Text
    T10.Text = Val(T14.Text) + Val(T15.Text)
    R10 = T10.Text
'Calcola i risultati per la LINEA3
    T4.Text = Val(T7.Text) + Val(T8.Text)
    R4 = T4.Text
    T5.Text = Val(T8.Text) + Val(T9.Text)
    R5 = T5.Text
    T6.Text = Val(T9.Text) + Val(T10.Text)
    R6 = T6.Text
'Calcola i risultati per la LINEA4
    T2.Text = Val(T4.Text) + Val(T5.Text)
    R2 = T2.Text
    T3.Text = Val(T5.Text) + Val(T6.Text)
    R3 = T3.Text
'Calcola i risultati per la LINEA5
    T1.Text = Val(T2.Text) + Val(T3.Text)
    R1 = T1.Text
'Resetta e imposta il colore blu a tutte le caselle
    RESET
    T1.ForeColor = RGB(0, 0, 255)
    T2.ForeColor = RGB(0, 0, 255)
    T3.ForeColor = RGB(0, 0, 255)
    T4.ForeColor = RGB(0, 0, 255)
    T5.ForeColor = RGB(0, 0, 255)
    T6.ForeColor = RGB(0, 0, 255)
    T7.ForeColor = RGB(0, 0, 255)
    T8.ForeColor = RGB(0, 0, 255)
    T9.ForeColor = RGB(0, 0, 255)
    T10.ForeColor = RGB(0, 0, 255)
    T11.ForeColor = RGB(0, 0, 255)
    T12.ForeColor = RGB(0, 0, 255)
    T13.ForeColor = RGB(0, 0, 255)
    T14.ForeColor = RGB(0, 0, 255)
    T15.ForeColor = RGB(0, 0, 255)
'Blocca alcune caselle per i suggerimenti della LINEA5
    For i = 1 To 10
        x = Int(Rnd * 15)
            If x = 1 Then
                T1.ForeColor = RGB(255, 0, 0)
                T1.Locked = True
                T1.Text = R1
            ElseIf x = 2 Then
                T2.ForeColor = RGB(255, 0, 0)
                T2.Locked = True
                T2.Text = R2
            ElseIf x = 3 Then
                T3.ForeColor = RGB(255, 0, 0)
                T3.Locked = True
                T3.Text = R3
            ElseIf x = 4 Then
                T4.ForeColor = RGB(255, 0, 0)
                T4.Locked = True
                T4.Text = R4
            ElseIf x = 5 Then
                T5.ForeColor = RGB(255, 0, 0)
                T5.Locked = True
                T5.Text = R5
            ElseIf x = 6 Then
                T6.ForeColor = RGB(255, 0, 0)
                T6.Locked = True
                T6.Text = R6
            ElseIf x = 7 Then
                T7.ForeColor = RGB(255, 0, 0)
                T7.Locked = True
                T7.Text = R7
            ElseIf x = 8 Then
                T8.ForeColor = RGB(255, 0, 0)
                T8.Locked = True
                T8.Text = R8
            ElseIf x = 9 Then
                T9.ForeColor = RGB(255, 0, 0)
                T9.Locked = True
                T9.Text = R9
            ElseIf x = 10 Then
                T10.ForeColor = RGB(255, 0, 0)
                T10.Locked = True
                T10.Text = R10
            ElseIf x = 11 Then
                T11.ForeColor = RGB(255, 0, 0)
                T11.Locked = True
                T11.Text = R11
            ElseIf x = 12 Then
                T12.ForeColor = RGB(255, 0, 0)
                T12.Locked = True
                T12.Text = R12
            ElseIf x = 13 Then
                T13.ForeColor = RGB(255, 0, 0)
                T13.Locked = True
                T13.Text = R13
            ElseIf x = 14 Then
                T14.ForeColor = RGB(255, 0, 0)
                T14.Locked = True
                T14.Text = R14
            ElseIf x = 15 Then
                T15.ForeColor = RGB(255, 0, 0)
                T15.Locked = True
                T15.Text = R15
            Else
                i = i - 1
            End If
    Next i
'Conta quanti numeri rossi sono presenti nello schema
    If T1.ForeColor = RGB(255, 0, 0) Then numROSSI = numROSSI + 1
    If T2.ForeColor = RGB(255, 0, 0) Then numROSSI = numROSSI + 1
    If T3.ForeColor = RGB(255, 0, 0) Then numROSSI = numROSSI + 1
    If T4.ForeColor = RGB(255, 0, 0) Then numROSSI = numROSSI + 1
    If T5.ForeColor = RGB(255, 0, 0) Then numROSSI = numROSSI + 1
    If T6.ForeColor = RGB(255, 0, 0) Then numROSSI = numROSSI + 1
    If T7.ForeColor = RGB(255, 0, 0) Then numROSSI = numROSSI + 1
    If T8.ForeColor = RGB(255, 0, 0) Then numROSSI = numROSSI + 1
    If T9.ForeColor = RGB(255, 0, 0) Then numROSSI = numROSSI + 1
    If T10.ForeColor = RGB(255, 0, 0) Then numROSSI = numROSSI + 1
    If T11.ForeColor = RGB(255, 0, 0) Then numROSSI = numROSSI + 1
    If T12.ForeColor = RGB(255, 0, 0) Then numROSSI = numROSSI + 1
    If T13.ForeColor = RGB(255, 0, 0) Then numROSSI = numROSSI + 1
    If T14.ForeColor = RGB(255, 0, 0) Then numROSSI = numROSSI + 1
    If T15.ForeColor = RGB(255, 0, 0) Then numROSSI = numROSSI + 1
    If numROSSI <= 6 Then NewGame
'Attiva il conto alla rovescia
    CDOWN.Enabled = True
    Abilita = True
End Sub

Private Sub CDOWN_Timer()
'Timer di gestione del conto alla rovescia
    If TempoR <= 0 Then
        GameOver
    Else
        TempoR = TempoR - 1
    End If
    TR.Caption = TempoR
End Sub

Private Sub Command1_Click()
'Ferma il tempo
    CDOWN.Enabled = False
End Sub

Private Sub Command2_Click()
'Visualizza la lista della soluzione
    If List1.Visible = True Then
        List1.Visible = False
    Else
        List1.Visible = True
        List1.Clear
        List1.AddItem R1
        List1.AddItem R2
        List1.AddItem R3
        List1.AddItem R4
        List1.AddItem R5
        List1.AddItem R6
        List1.AddItem R7
        List1.AddItem R8
        List1.AddItem R9
        List1.AddItem R10
        List1.AddItem R11
        List1.AddItem R12
        List1.AddItem R13
        List1.AddItem R14
        List1.AddItem R15
    End If
End Sub

Private Sub DEMO_Timer()
'Colora le palline
    If ShapeC = 0 Then
        Colora 255, 255, 0
        ShapeC = ShapeC + 1
    ElseIf ShapeC = 1 Then
        Colora 255, 204, 0
        ShapeC = ShapeC + 1
    ElseIf ShapeC = 2 Then
        Colora 255, 153, 0
        ShapeC = ShapeC + 1
    ElseIf ShapeC = 3 Then
        Colora 255, 102, 0
        ShapeC = ShapeC + 1
    ElseIf ShapeC = 4 Then
        Colora 255, 51, 0
        ShapeC = ShapeC + 1
    ElseIf ShapeC = 5 Then
        Colora 255, 0, 0
        ShapeC = ShapeC + 1
    ElseIf ShapeC = 6 Then
        Colora 255, 51, 0
        ShapeC = ShapeC + 1
    ElseIf ShapeC = 7 Then
        Colora 255, 102, 0
        ShapeC = ShapeC + 1
    ElseIf ShapeC = 8 Then
        Colora 255, 153, 0
        ShapeC = ShapeC + 1
    Else
        Colora 255, 204, 0
        ShapeC = 0
    End If
End Sub

Private Sub Form_Load()
'Resetta le variabili
    RRESET
    Punti = 0
    TempoR = 120
    Livello = 1
    Abilita = False
    ShapeC = 0
'Carica i migliori risultati dal registro di Windows
    LoadScore
'Imposta il titolo, la data, l'ora, e centra la finestra
    DT.Caption = Time & Chr(13) & Date
    Me.Caption = "Piramide numerica Versione " & App.Major & "." & App.Minor & " rev. " & App.Revision & " -DjGG-"
    CF Me
'Pulisce tutte le textbox
    PT
    AWhite
'Inizia a mescolare i numeri
    Randomize
End Sub

Public Sub PT()
'Sub per svuotare le caselle di testo
    T1.Text = ""
    T2.Text = ""
    T3.Text = ""
    T4.Text = ""
    T5.Text = ""
    T6.Text = ""
    T7.Text = ""
    T8.Text = ""
    T9.Text = ""
    T10.Text = ""
    T11.Text = ""
    T12.Text = ""
    T13.Text = ""
    T14.Text = ""
    T15.Text = ""
End Sub
Public Sub LockA()
'Blocca tutte le caselle di testo
    T1.Locked = True
    T2.Locked = True
    T3.Locked = True
    T4.Locked = True
    T5.Locked = True
    T6.Locked = True
    T7.Locked = True
    T8.Locked = True
    T9.Locked = True
    T10.Locked = True
    T11.Locked = True
    T12.Locked = True
    T13.Locked = True
    T14.Locked = True
    T15.Locked = True
End Sub

Public Sub UnlockA()
'Sblocca tutte le caselle di testo
    T1.Locked = False
    T2.Locked = False
    T3.Locked = False
    T4.Locked = False
    T5.Locked = False
    T6.Locked = False
    T7.Locked = False
    T8.Locked = False
    T9.Locked = False
    T10.Locked = False
    T11.Locked = False
    T12.Locked = False
    T13.Locked = False
    T14.Locked = False
    T15.Locked = False
End Sub

Public Sub RESET()
'Esegue pulizia e sblocca le caselle
    PT
    UnlockA
End Sub

Public Sub RRESET()
'Resetta le variabili
    R1 = 0
    R2 = 0
    R3 = 0
    R4 = 0
    R5 = 0
    R6 = 0
    R7 = 0
    R8 = 0
    R9 = 0
    R10 = 0
    R11 = 0
    R12 = 0
    R13 = 0
    R14 = 0
    R15 = 0
    C1 = False
    C2 = False
    C3 = False
    C4 = False
    C5 = False
    C6 = False
    C7 = False
    C8 = False
    C9 = False
    C10 = False
End Sub

Private Sub mnuAbout_Click()
'Informazioni sul gioco
    frmAbout.Show 1
End Sub

Private Sub mnuExit_Click()
'Termina il gioco
    End
End Sub

Private Sub mnuNew_Click()
'Imposta un nuovo gioco
    Abilita = False
    DEMO.Enabled = False
    Colora 255, 255, 255
    TempoR = 120
    OldTempoR = TempoR
    Livello = 1
    If T1.Text <> "" Or _
        T2.Text <> "" Or _
        T3.Text <> "" Or _
        T4.Text <> "" Or _
        T5.Text <> "" Or _
        T6.Text <> "" Or _
        T7.Text <> "" Or _
        T8.Text <> "" Or _
        T9.Text <> "" Or _
        T10.Text <> "" Or _
        T11.Text <> "" Or _
        T12.Text <> "" Or _
        T13.Text <> "" Or _
        T14.Text <> "" Or _
        T15.Text <> "" Then
            response = MsgBox("Terminare la partita corrente per iniziarne una nuova?", vbYesNo, "Nuovo gioco...")
                If response = vbYes Then
                    Punti = 0
                    Punteggio.Caption = Punti
                    NewGame
                Else
                    Exit Sub
                End If
    Else
        NewGame
    End If
End Sub

Public Sub AWhite()
'Imposta il colore bianco a tutti gli Shape e le TextBox
    T1.BackColor = &H80000005
    Shape1.FillColor = &H80000005
    T2.BackColor = &H80000005
    Shape2.FillColor = &H80000005
    T3.BackColor = &H80000005
    Shape3.FillColor = &H80000005
    T4.BackColor = &H80000005
    Shape4.FillColor = &H80000005
    T5.BackColor = &H80000005
    Shape5.FillColor = &H80000005
    T6.BackColor = &H80000005
    Shape6.FillColor = &H80000005
    T7.BackColor = &H80000005
    Shape7.FillColor = &H80000005
    T8.BackColor = &H80000005
    Shape8.FillColor = &H80000005
    T9.BackColor = &H80000005
    Shape9.FillColor = &H80000005
    T10.BackColor = &H80000005
    Shape10.FillColor = &H80000005
    T11.BackColor = &H80000005
    Shape11.FillColor = &H80000005
    T12.BackColor = &H80000005
    Shape12.FillColor = &H80000005
    T13.BackColor = &H80000005
    Shape13.FillColor = &H80000005
    T14.BackColor = &H80000005
    Shape14.FillColor = &H80000005
    T15.BackColor = &H80000005
    Shape15.FillColor = &H80000005
End Sub

'Controllo del testo immesso
'----------------------------------------------------------------------
Private Sub T1_Change()
    CN T1
End Sub

Private Sub T10_Change()
    CN T10
End Sub

Private Sub T11_Change()
    CN T11
End Sub

Private Sub T12_Change()
    CN T12
End Sub

Private Sub T13_Change()
    CN T13
End Sub

Private Sub T14_Change()
    CN T14
End Sub

Private Sub T15_Change()
    CN T15
End Sub

Private Sub T2_Change()
    CN T2
End Sub

Private Sub T3_Change()
    CN T3
End Sub

Private Sub T4_Change()
    CN T4
End Sub

Private Sub T5_Change()
    CN T5
End Sub

Private Sub T6_Change()
    CN T6
End Sub

Private Sub T7_Change()
    CN T7
End Sub

Private Sub T8_Change()
    CN T8
End Sub

Private Sub T9_Change()
    CN T9
End Sub
'----------------------------------------------------------------------

Private Sub Timer1_Timer()
'Colora a caso il titolo del gioco
    Titolo.ForeColor = RGB(255, Rnd * 255, 0)
End Sub

Public Sub ATW()
'Imposta il colore FORECOLOR di tutte le textbox su bianco
    T1.ForeColor = &H80000005
    T2.ForeColor = &H80000005
    T3.ForeColor = &H80000005
    T4.ForeColor = &H80000005
    T5.ForeColor = &H80000005
    T6.ForeColor = &H80000005
    T7.ForeColor = &H80000005
    T8.ForeColor = &H80000005
    T9.ForeColor = &H80000005
    T10.ForeColor = &H80000005
    T11.ForeColor = &H80000005
    T12.ForeColor = &H80000005
    T13.ForeColor = &H80000005
    T14.ForeColor = &H80000005
    T15.ForeColor = &H80000005
End Sub

Public Sub CN(txtbox As TextBox)
'Controlla se è stato inserito un numero nella casella di testo
    If Abilita = True Then
        If IsNumeric(txtbox.Text) = True Then
            'nulla
        Else
            txtbox.Text = ""
            txtbox.SelStart = 0
        End If
        CheckWin
    Else
        'Esce dal controllo
    End If
End Sub

Public Sub CheckWin()
'Controlla se il giocatore ha vinto
'Controlla se le caselle di testo sono vuote
    If T1.Text <> "" And _
    T2.Text <> "" And _
    T3.Text <> "" And _
    T4.Text <> "" And _
    T5.Text <> "" And _
    T6.Text <> "" And _
    T7.Text <> "" And _
    T8.Text <> "" And _
    T9.Text <> "" And _
    T10.Text <> "" And _
    T11.Text <> "" And _
    T12.Text <> "" And _
    T13.Text <> "" And _
    T14.Text <> "" And _
    T15.Text <> "" Then
        FASE2 'compara i valori
    Else
        'esce dal ciclo
    End If
End Sub

Public Sub FASE2()
'Controlla se il risultato finale è quello calcolato dal computer
Dim Result As Integer
Debug.Print "Inizio controllo " & Date & " - " & Time 'Debug
    Result = Val(T11.Text) + Val(T12.Text)
    Debug.Print "T7=" & T7.Text 'Debug
    If Result = Val(T7.Text) Then
        C7 = True
    Else
        C7 = False
    End If
    Result = Val(T12.Text) + Val(T13.Text)
    Debug.Print "T8=" & T8.Text 'Debug
    If Result = Val(T8.Text) Then
        C8 = True
    Else
        C8 = False
    End If
    Result = Val(T13.Text) + Val(T14.Text)
    Debug.Print "T9=" & T9.Text 'Debug
    If Result = Val(T9.Text) Then
        C9 = True
    Else
        C9 = False
    End If
    Result = Val(T14.Text) + Val(T15.Text)
    Debug.Print "T10=" & T10.Text 'Debug
    If Result = Val(T10.Text) Then
        C10 = True
    Else
        C10 = False
    End If
    Result = Val(T7.Text) + Val(T8.Text)
    Debug.Print "T4=" & T4.Text 'Debug
    If Result = Val(T4.Text) Then
        C4 = True
    Else
        C4 = False
    End If
    Result = Val(T8.Text) + Val(T9.Text)
    Debug.Print "T5=" & T5.Text 'Debug
    If Result = Val(T5.Text) Then
        C5 = True
    Else
        C5 = False
    End If
    Result = Val(T9.Text) + Val(T10.Text)
    Debug.Print "T6=" & T6.Text 'Debug
    If Result = Val(T6.Text) Then
        C6 = True
    Else
        C6 = False
    End If
    Result = Val(T4.Text) + Val(T5.Text)
    Debug.Print "T2=" & T2.Text 'Debug
    If Result = Val(T2.Text) Then
        C2 = True
    Else
        C2 = False
    End If
    Result = Val(T5.Text) + Val(T6.Text)
    Debug.Print "T3=" & T3.Text 'Debug
    If Result = Val(T3.Text) Then
        C3 = True
    Else
        C3 = False
    End If
    Result = Val(T2.Text) + Val(T3.Text)
    Debug.Print "T1=" & T1.Text 'Debug
    If Result = Val(T1.Text) Then
        C1 = True
    Else
        C1 = False
    End If
'Debug
'-----------------------------------------------------------------------
Debug.Print "C1=" & C1
Debug.Print "C2=" & C2
Debug.Print "C3=" & C3
Debug.Print "C4=" & C4
Debug.Print "C5=" & C5
Debug.Print "C6=" & C6
Debug.Print "C7=" & C7
Debug.Print "C8=" & C8
Debug.Print "C9=" & C9
Debug.Print "C10=" & C10
'----------------------------------------------------------------------
    Result = Val(T2.Text) + Val(T3.Text)
    If Result = R1 And Val(T1.Text) = R1 Then
    Debug.Print "R1=" & R1 'Debug
'Controllo finale
'----------------------------------------------------------------------
        If C2 = True And C3 = True And C4 = True And _
            C5 = True And C6 = True And C7 = True And _
            C8 = True And C9 = True And C10 = True Then
                CDOWN.Enabled = False
                MsgBox ("HAI VINTO, preparati per il nuovo livello!"), vbExclamation, "GRANDIOSO!"
                Punti = Punti + 15
                OldTempoR = OldTempoR - 10
                TempoR = OldTempoR
                Punteggio.Caption = Punti
                Livello = Livello + 1
                NewGame
            Else
                'esce dal ciclo
            End If
    Else
        'esce dal ciclo
    End If
End Sub

Private Sub Timer2_Timer()
'Aggiorna data e ora
    DT.Caption = Time & Chr(13) & Date
End Sub

Public Sub GameOver()
'Gioco finito
    DEMO.Enabled = True
    CDOWN.Enabled = False
    If mpunt > Punti Then
        MsgBox "HAI PERSO!!!" & Chr(13) & "Purtroppo non hai realizzato un nuovo record", vbCritical, "NOOoooo...."
    Else
        mpunt = Punti
        nome = InputBox("Complimenti hai realizzato un nuovo record, inserisci il tuo nome!", "Gioco finito!", mname)
        SaveSetting "PNUM", "SCORE", "HISCORE", mpunt
        SaveSetting "PNUM", "SCORE", "NSCORE", nome
        LoadScore
    End If
    LockA
End Sub

Public Sub Colora(R As Integer, G As Integer, B As Integer)
'Colorazione delle palline e delle textbox
    Shape1.FillColor = RGB(R, G, B)
    Shape2.FillColor = RGB(R, G, B)
    Shape3.FillColor = RGB(R, G, B)
    Shape4.FillColor = RGB(R, G, B)
    Shape5.FillColor = RGB(R, G, B)
    Shape6.FillColor = RGB(R, G, B)
    Shape7.FillColor = RGB(R, G, B)
    Shape8.FillColor = RGB(R, G, B)
    Shape9.FillColor = RGB(R, G, B)
    Shape10.FillColor = RGB(R, G, B)
    Shape11.FillColor = RGB(R, G, B)
    Shape12.FillColor = RGB(R, G, B)
    Shape13.FillColor = RGB(R, G, B)
    Shape14.FillColor = RGB(R, G, B)
    Shape15.FillColor = RGB(R, G, B)
    T1.BackColor = RGB(R, G, B)
    T2.BackColor = RGB(R, G, B)
    T3.BackColor = RGB(R, G, B)
    T4.BackColor = RGB(R, G, B)
    T5.BackColor = RGB(R, G, B)
    T6.BackColor = RGB(R, G, B)
    T7.BackColor = RGB(R, G, B)
    T8.BackColor = RGB(R, G, B)
    T9.BackColor = RGB(R, G, B)
    T10.BackColor = RGB(R, G, B)
    T11.BackColor = RGB(R, G, B)
    T12.BackColor = RGB(R, G, B)
    T13.BackColor = RGB(R, G, B)
    T14.BackColor = RGB(R, G, B)
    T15.BackColor = RGB(R, G, B)
End Sub

Public Sub LoadScore()
'Carica il punteggio massimo dal registro
    mpunt = GetSetting("PNUM", "SCORE", "HISCORE")
    mname = GetSetting("PNUM", "SCORE", "NSCORE")
    If mpunt = "" Then mpunt = 0 'Se non esiste alcun punteggio registrato imposta 0
    If mname = "" Then mname = "<Giocatore>" 'Se non esiste nessun nome registrato imposta la stringa
    HScore.Caption = mname & Chr(13) & mpunt
End Sub
