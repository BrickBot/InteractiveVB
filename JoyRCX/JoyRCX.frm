VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "JoyRCX-JvK"
   ClientHeight    =   7080
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7845
   LinkTopic       =   "Form1"
   ScaleHeight     =   7080
   ScaleWidth      =   7845
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   1440
      TabIndex        =   29
      Top             =   3360
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6960
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ListBox List1 
      Height          =   1230
      Left            =   720
      Sorted          =   -1  'True
      TabIndex        =   27
      Top             =   5040
      Width           =   2535
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   5280
      TabIndex        =   26
      Top             =   4800
      Width           =   2415
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   5280
      TabIndex        =   22
      Top             =   4320
      Width           =   2415
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   5280
      TabIndex        =   21
      Top             =   3840
      Width           =   2415
   End
   Begin VB.CommandButton CmdStart 
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   600
      TabIndex        =   16
      Top             =   3720
      Width           =   2655
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2160
      TabIndex        =   13
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   600
      TabIndex        =   12
      Top             =   2880
      Width           =   1335
   End
   Begin VB.TextBox TxtDY 
      Height          =   615
      Left            =   2040
      TabIndex        =   11
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox TxtDX 
      Height          =   615
      Left            =   600
      TabIndex        =   10
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox TxtY 
      Height          =   615
      Left            =   2040
      TabIndex        =   6
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox TxtX 
      Height          =   615
      Left            =   600
      TabIndex        =   4
      Top             =   720
      Width           =   1215
   End
   Begin VB.Timer TmrRefresh 
      Enabled         =   0   'False
      Interval        =   75
      Left            =   2880
      Top             =   120
   End
   Begin VB.CheckBox chkHelp 
      Caption         =   "Help"
      Height          =   255
      Left            =   840
      TabIndex        =   3
      Top             =   0
      Value           =   1  'Checked
      Width           =   675
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Help"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   0
      Width           =   495
   End
   Begin VB.TextBox txtResult 
      BackColor       =   &H80000013&
      Height          =   435
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   6480
      Width           =   7575
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "Test"
      Height          =   255
      Left            =   1680
      TabIndex        =   0
      Top             =   0
      Width           =   855
   End
   Begin VB.Label Label12 
      Caption         =   "Lijst Topscores (lager is beter!)"
      Height          =   375
      Left            =   720
      TabIndex        =   28
      Top             =   4680
      Width           =   2535
   End
   Begin VB.Label Label11 
      Caption         =   "Postcode en plaats"
      Height          =   375
      Left            =   3600
      TabIndex        =   25
      Top             =   4800
      Width           =   1575
   End
   Begin VB.Label Label10 
      Caption         =   "Straat en nummer"
      Height          =   375
      Left            =   3600
      TabIndex        =   24
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Label Label8 
      Caption         =   "Naam"
      Height          =   255
      Left            =   3720
      TabIndex        =   23
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Label LblTijd 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3720
      TabIndex        =   20
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label Label9 
      Caption         =   "Aantal seconden"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3720
      TabIndex        =   19
      Top             =   1920
      Width           =   3975
   End
   Begin VB.Label LblPunt 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3720
      TabIndex        =   18
      Top             =   960
      Width           =   3375
   End
   Begin VB.Label Label7 
      Caption         =   "Strafpunten"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3840
      TabIndex        =   17
      Top             =   120
      Width           =   3255
   End
   Begin VB.Label Label6 
      Caption         =   "direction"
      Height          =   375
      Left            =   2160
      TabIndex        =   15
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "motor"
      Height          =   255
      Left            =   600
      TabIndex        =   14
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "dy"
      Height          =   375
      Left            =   2040
      TabIndex        =   9
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "dx"
      Height          =   495
      Left            =   600
      TabIndex        =   8
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Y"
      Height          =   375
      Left            =   2040
      TabIndex        =   7
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "x"
      Height          =   375
      Left            =   600
      TabIndex        =   5
      Top             =   360
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Joypos As Long
Dim ButtonsOud As Integer
Private Type JOYINFO
        wXpos As Long
        wYpos As Long
        wZpos As Long
        wButtons As Long
End Type
Private Declare Function joyGetPos Lib "winmm.dll" (ByVal uJoyID As Long, pji As JOYINFO) As Long
Private WithEvents vpb As LEGOVPBrickLib.VPBrick
Attribute vpb.VB_VarHelpID = -1
Private Port$  'current port
Dim tijd As Long
Dim diroud As Integer
Dim maxlight As Integer
Dim motoroud As Integer
'=========================================================
'NB maxlight instellen (zwarte baan volgen eromheen wit)
'=========================================================
Private Sub CmdStart_Click()
  If Text5.Text = "" Or Text4.Text = "" Or Text3.Text = "" Then
    MsgBox "geef eerst naam en volledig adres"
  Else
    TmrRefresh.Enabled = True
    LblPunt.Caption = "0"
    tijd = 60 * (60 * Hour(Time) + Minute(Time)) + Second(Time)
    CmdStart.Visible = False
  End If
End Sub

Private Sub Text3_Click()
  Text3.Text = ""
End Sub
Private Sub Text4_Click()
  Text4.Text = ""
End Sub
Private Sub Text5_Click()
  Text5.Text = ""
End Sub

Private Sub Text6_Click()
   maxlight = Text6.Text
End Sub

Private Sub TmrRefresh_Timer()
Dim pji As JOYINFO
Dim JOYSTICKID1 As Long
Dim deltaX, DeltaY As Single
Dim commando As String
Dim motor, direct, waarde As Integer
Dim td As Long
'if the timerinterval is to short (sending of commands takes to much time=> error in direction)
'if loose of communication=> reinit through except and cmdtest_click
'To shorten time to communicate this version uses only one send and 1 recieve per cycle
On Error GoTo except
  Joypos = joyGetPos(0, pji)
  TxtX.Text = pji.wXpos
  TxtY.Text = pji.wYpos
  ' TxtButtons.Text = pji.wButtons
  deltaX = (pji.wXpos - 32728) / 100  'hoger => naar links
  DeltaY = (32728 - pji.wYpos) / 100  'hoger => naar boven (omgekeerd)
  TxtDX = deltaX: TxtDY = DeltaY
  direct = 0
  If deltaX > 150 Then
     motor = 2
    Else
       If deltaX < -150 Then motor = 1 Else motor = 3 'beide motoren
  End If
  If DeltaY < -50 Then
     direct = 0
   Else
     If DeltaY > 50 Then direct = 2 Else motor = 4
  End If
  Text1 = motor: Text2 = direct
  td = 60 * (60 * Hour(Time) + Minute(Time)) + Second(Time) - tijd
  LblTijd.Caption = td
  'ook commando sturen als watchdog
  'in rcx als meer dan ....ms na commando niets binnen => motoren omkeren
  waarde = motor + 8 * direct + 32  'var 28 als zonder watchdog joyrcx2
  commando = "setv 31,2," & waarde 'variabele 31 met constavte(2) met waarde (1)
  vpb.Execute commando, LASM  '=> steeds opnieuw watchdog
  Call ContrSens 'ombouwen naar interpr perc+100*switch
 Exit Sub
except: Call cmdTest_Click
End Sub
Private Sub ContrSens()
  Dim nResult As Long
  Dim nPos As Long
  Dim result As Variant
  Dim schak As Boolean
  Dim licht As Integer
  Dim commando As String
  schak = False
   'commando = "poll 9,0"  'waarde senor 0 opvragen schak"
   'nResult = vpb.Execute(commando, LASM, nPos, result)
   commando = "poll 0,21"  'waarde variabele 30 opvragen/licht
   nResult = vpb.Execute(commando, LASM, nPos, result)
   Text6.Text = nResult 'moet later anders
   If nResult > 0 Then
      If nResult > 101 Then
        nResult = nResult - 101
       Call finish
     End If
     If nResult < maxlight Then LblPunt.Caption = LblPunt.Caption + 10 'moet net andersom (zwarte baan volgen
   End If
End Sub
Private Sub finish()
Dim punt As Integer
Dim StrPunt As String
Dim IntRegelnr As Integer
    vpb.Execute "out 1,7", LASM   '1 is off/7=all motoren
   'score berekenen
   punt = Val(LblTijd.Caption) + Val(LblPunt.Caption)
   'op topscorelijst
   StrPunt = punt
   While Len(StrPunt) < 3
       StrPunt = "0" & StrPunt
   Wend
   If StrPunt < Left(List1.List(List1.ListCount - 1), 3) Then MsgBox "gefeliciteerd je bent de nieuwe topscorer"
   List1.AddItem StrPunt & "  " & Text3.Text & " " & Text4.Text & " " & Text5.Text
   TmrRefresh.Enabled = False
   CmdStart.Visible = True
   Open App.Path & "\Topscore.txt" For Output As #1
   For IntRegelnr = 0 To List1.ListCount - 1
     Print #1, List1.List(IntRegelnr)
   Next
   Close
End Sub
Private Sub cmdClose_Click()
  vpb.Close
  txtResult = "Ok"
End Sub

Private Sub FindPort()
  On Error GoTo except
  vpb.FindPort Port$
  txtResult = "Port: " + Port$
  Exit Sub

except:
  txtResult = "Error " + Hex$(Err.Number) + " " + Err.Description
  ErrHelp
End Sub


Private Sub ErrHelp()
  If Err.HelpFile = "" Then
      MsgBox "No Err.HelpFile"
      Exit Sub
    
    With CommonDialog1
      .HelpFile = Err.HelpFile
      .HelpContext = Val(Err.HelpContext)
      .HelpCommand = cdlHelpContext
      .ShowHelp
    End With
  End If
End Sub

Private Sub cmdHelp_Click()
  With CommonDialog1
    .HelpFile = "VPB.hlp"
    .HelpCommand = cdlHelpContents
    .ShowHelp
  End With
End Sub

Private Sub OpenPort()
  On Error GoTo except
  vpb.Open Port$
  txtResult = "Opened " + Port$
  Exit Sub

except:
  txtResult = "Error " + Hex$(Err.Number) + " " + Err.Description
  ErrHelp
End Sub

Private Sub cmdStatus_Click()
  GetStatus
End Sub

Private Sub GetStatus()
  Dim nStatus As StatusResult
  Dim nBrickType As BrickTypes

  On Error GoTo except
  
  nBrickType = vpb.Status(CheckBrickType)
  nStatus = vpb.Status(BrickStatus)
  Select Case nStatus
    Case StatusReady
      txtResult = "Ready"
    Case StatusBusy
      txtResult = "Busy"
    Case Downloading
      txtResult = "Downloading"
    Case NotOpened
      txtResult = "Not opened"
    Case NoTower
      txtResult = "No tower"
    Case BadTower
      txtResult = "Bad tower"
    Case NoBrick
      txtResult = "No brick"
    Case NoFirmware
      txtResult = "No firmware"
    Case BadBrickBattery
      txtResult = " bad battery"
    Case BadComms
      txtResult = "Bad comms"
    Case Else
      txtResult = "Status: " + Str(nStatus)
  End Select

  Exit Sub

except:
  'ignore
End Sub
Private Function SetBrickType(nBrickType As BrickTypes)
  Select Case nBrickType
    Case Scout
      comboBrickType = "Scout"
      vpb.BrickType = Scout
    Case RCX2
      comboBrickType = "RCX2"
      vpb.BrickType = RCX2
    Case MicroScout
      comboBrickType = "MicroScout"
      vpb.BrickType = MicroScout
    Case Else
      comboBrickType.Clear
  End Select
End Function


Private Sub cmdTest_Click()
  On Error GoTo except
  nBrickType = vpb.Status(CheckBrickType)
  SetBrickType (nBrickType)
  nStatus = vpb.Status(BrickStatus)  'unlocks Scout
  If nStatus = StatusReady Then
    If nBrickType = Scout Then
      vpb.Execute "sound 25"
    Else
      vpb.Execute "sound 3"
    End If
    txtResult = "Ok"
  Else
    txtResult = "Not ready"
  End If
  vpb.PortRxBaudRate = 2400 'laagste waardes werken niet 2400=> 1/3 s
  vpb.PortTxBaudRate = 2400 '19200 3s/cyclus (te snel=> teveel interrupts?)
  Exit Sub                  'lagere waarden werken ook niet als waittime omhoog!?

except:
  txtResult = "Error " + Hex$(Err.Number) + " " + Err.Description
   ErrHelp
End Sub
Private Sub initsens()
   Dim commando As String
   commando = "sent 0,1"     'Type sensor instellen sens0,switch"
   vpb.Execute commando, LASM
   commando = "sent 1,3"     'Type sensor instellen sens1,reflection"
   vpb.Execute commando, LASM
   commando = "sent 2,1"     'Type sensor instellen sens2,switch"
   vpb.Execute commando, LASM
   commando = "senm 0,1,0"  'sens0 (boolean/absoluut)
   vpb.Execute commando, LASM
   commando = "senm 1,4,0"  'sens1 (4=perc)
   vpb.Execute commando, LASM
   commando = "senm 2,1,0"  'sens2 (boolen/abs)
   vpb.Execute commando, LASM
   End Sub
Private Sub Form_Activate()
  motoroud = 1
  On Error GoTo except
  FindPort
  OpenPort
  GetStatus
  cmdTest_Click
  Call initsens
  maxlight = 34
  Exit Sub
except:
  txtResult = "Error " + Hex$(Err.Number) + " " + Err.Description
  ErrHelp
End Sub

Private Sub Form_Load()
Dim Hulp As String
 
  Set vpb = New LEGOVPBrickLib.VPBrick
  Port$ = ""
'  ínlezen topscorelijst
  Open App.Path & "\topscore.txt" For Input As #1
   Do While Not EOF(1)
     Line Input #1, Hulp
     List1.AddItem Hulp, 0
   Loop
  Close
End Sub

Private Sub Form_Terminate()
  vpb.Close
  'bewaren list1 in textbestand
End Sub


