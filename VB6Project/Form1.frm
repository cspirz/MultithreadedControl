VERSION 5.00
Object = "{E696C3F8-A01D-4BF8-B645-F179228E4C5F}#1.0#0"; "mscoree.dll"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6840
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   5925
   LinkTopic       =   "Form1"
   ScaleHeight     =   6840
   ScaleWidth      =   5925
   StartUpPosition =   3  'Windows Default
   Begin MultithreadedControlCtl.BackgroundWorker BackgroundWorker1 
      Height          =   1095
      Left            =   120
      TabIndex        =   4
      Top             =   5160
      Width           =   5415
      Object.Visible         =   "True"
      Enabled         =   "True"
      ForegroundColor =   "-2147483630"
      BackgroundColor =   "-2147483633"
      BackColor       =   "Control"
      ForeColor       =   "ControlText"
      Location        =   "8, 344"
      Name            =   "BackgroundWorker"
      Size            =   "361, 73"
      Object.TabIndex        =   "0"
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Process"
      Height          =   375
      Left            =   4200
      TabIndex        =   3
      Top             =   3000
      Width           =   1335
   End
   Begin VB.ListBox List1 
      Height          =   1035
      Left            =   120
      TabIndex        =   1
      Top             =   3960
      Width           =   5415
   End
   Begin VB.TextBox Text1 
      Height          =   2775
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   5415
   End
   Begin VB.Label Label4 
      Caption         =   "Progress..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   6360
      Width           =   5415
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   5160
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Progress..."
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   5520
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Events:"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   3600
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents BackgroundEvents As MultithreadedControl.BackgroundWorker
Attribute BackgroundEvents.VB_VarHelpID = -1

'You will need to
'rebuild the .NET UserControl in order to register it
'on your machine.  Then add the MultiThreadedControl component
'and draw it on your form.
'Finally, add an additional reference to MultiThreadedControl
'using the Project | References... menu button.

Private Sub Command1_Click()
    'Me.Text1.Text = ""
    Me.List1.Clear
    Me.List1.AddItem ("Start processing from VB6: " & DateTime.Now)
    Me.BackgroundWorker1.StartProcessing (Me.Text1.Text)
End Sub

Private Sub Form_Load()
    Set BackgroundEvents = Me.BackgroundWorker1
End Sub

Private Sub BackgroundEvents_StartEvent(ByVal EventText As String)
    Me.List1.AddItem (EventText)
End Sub

Private Sub BackgroundEvents_FinishAsyncEvent(ByVal EventText As String)
    Me.List1.AddItem (EventText)
    Me.List1.AddItem ("Finish processing from VB6:" & DateTime.Now)
End Sub
