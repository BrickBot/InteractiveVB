VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "IVPBrick Test Client"
   ClientHeight    =   8085
   ClientLeft      =   1830
   ClientTop       =   615
   ClientWidth     =   9555
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8085
   ScaleWidth      =   9555
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   5160
      TabIndex        =   44
      Text            =   "?"
      Top             =   7680
      Width           =   1095
   End
   Begin VB.CommandButton CmdInput 
      Caption         =   "Input"
      Height          =   855
      Left            =   3840
      TabIndex        =   43
      Top             =   7200
      Width           =   1215
   End
   Begin VB.Frame Frame4 
      Caption         =   "Type"
      Height          =   615
      Left            =   3840
      TabIndex        =   40
      Top             =   6480
      Width           =   2415
      Begin VB.OptionButton OptType 
         Caption         =   "Reflection"
         Height          =   255
         Index           =   3
         Left            =   1080
         TabIndex        =   42
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton OptType 
         Caption         =   "Switch"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   41
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "MODE"
      Height          =   735
      Left            =   3840
      TabIndex        =   36
      Top             =   5640
      Width           =   2415
      Begin VB.OptionButton OptSMode 
         Caption         =   "Perc"
         Height          =   375
         Index           =   4
         Left            =   1560
         TabIndex        =   39
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton OptSMode 
         Caption         =   "Bool"
         Height          =   375
         Index           =   1
         Left            =   720
         TabIndex        =   38
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton OptSMode 
         Caption         =   "Raw"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   37
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.ComboBox CmbSens 
      Height          =   315
      ItemData        =   "VBTest.frx":0000
      Left            =   3840
      List            =   "VBTest.frx":000D
      TabIndex        =   35
      Text            =   "Sensor"
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "(Re)connect and Test"
      Height          =   495
      Left            =   8040
      TabIndex        =   34
      Top             =   0
      Width           =   1335
   End
   Begin VB.TextBox txtResult 
      BackColor       =   &H80000013&
      Height          =   315
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   4800
      Width           =   9495
   End
   Begin VB.CommandButton cmdStatus 
      Caption         =   "Status"
      Height          =   255
      Left            =   8040
      TabIndex        =   32
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton cmdFirmware 
      Caption         =   "New Firmware"
      Enabled         =   0   'False
      Height          =   255
      Left            =   8040
      TabIndex        =   31
      Top             =   3360
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton CmdPlaySound 
      Caption         =   "PlaySound"
      Height          =   495
      Left            =   0
      TabIndex        =   30
      Top             =   5640
      Width           =   1095
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "VBTest.frx":0026
      Left            =   0
      List            =   "VBTest.frx":0039
      TabIndex        =   29
      Text            =   "Sound"
      Top             =   5160
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "MODE"
      Height          =   735
      Left            =   1320
      TabIndex        =   25
      Top             =   5160
      Width           =   2415
      Begin VB.OptionButton OptMode 
         Caption         =   "On"
         Height          =   255
         Index           =   2
         Left            =   1560
         TabIndex        =   28
         Top             =   360
         Width           =   615
      End
      Begin VB.OptionButton OptMode 
         Caption         =   "OFF"
         Height          =   255
         Index           =   1
         Left            =   840
         TabIndex        =   27
         Top             =   360
         Width           =   735
      End
      Begin VB.OptionButton OptMode 
         Caption         =   "Float"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   26
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "VBTest.frx":006C
      Left            =   1320
      List            =   "VBTest.frx":0085
      TabIndex        =   24
      Text            =   "Outputs"
      Top             =   6720
      Width           =   1215
   End
   Begin VB.CommandButton CmdOutput 
      Caption         =   "Output"
      Height          =   855
      Left            =   1320
      TabIndex        =   23
      Top             =   7200
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Direction"
      Height          =   735
      Left            =   1320
      TabIndex        =   19
      Top             =   5880
      Width           =   2415
      Begin VB.OptionButton DirMode 
         Caption         =   "left"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   22
         Top             =   360
         Width           =   615
      End
      Begin VB.OptionButton DirMode 
         Caption         =   "Change"
         Height          =   255
         Index           =   1
         Left            =   720
         TabIndex        =   21
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton DirMode 
         Caption         =   "Right"
         Height          =   255
         Index           =   2
         Left            =   1560
         TabIndex        =   20
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.CommandButton CmdMakeSound 
      Caption         =   "MakeSound"
      Height          =   375
      Left            =   0
      TabIndex        =   18
      Top             =   7680
      Width           =   1215
   End
   Begin VB.TextBox TxtFreq 
      Height          =   375
      Left            =   0
      TabIndex        =   17
      Text            =   "Freq(30-20000)"
      Top             =   6720
      Width           =   1215
   End
   Begin VB.TextBox TxtTime 
      Height          =   375
      Left            =   0
      TabIndex        =   16
      Text            =   "tijd (1-255x0.1s)"
      Top             =   7200
      Width           =   1215
   End
   Begin VB.CommandButton CmdSetWatch 
      Caption         =   "SetWatch to PC Time"
      Height          =   495
      Left            =   6480
      TabIndex        =   15
      Top             =   5280
      Width           =   1215
   End
   Begin VB.TextBox TxtMonitorIR 
      Height          =   285
      Left            =   8040
      TabIndex        =   14
      Text            =   "IR message"
      Top             =   6000
      Width           =   1335
   End
   Begin VB.CommandButton CmdMonitor 
      Caption         =   "MonitorIR"
      Height          =   375
      Left            =   7920
      TabIndex        =   13
      Top             =   5400
      Width           =   1575
   End
   Begin VB.CheckBox chkHelp 
      Caption         =   "Help"
      Height          =   255
      Left            =   8760
      TabIndex        =   12
      Top             =   1320
      Value           =   1  'Checked
      Width           =   675
   End
   Begin MSComDlg.CommonDialog CdlBestand 
      Left            =   5460
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtSleep 
      Height          =   285
      Left            =   8580
      TabIndex        =   4
      Top             =   3000
      Width           =   435
   End
   Begin VB.ComboBox comboIR 
      Height          =   315
      ItemData        =   "VBTest.frx":00B6
      Left            =   8520
      List            =   "VBTest.frx":00C0
      TabIndex        =   3
      Top             =   2580
      Width           =   915
   End
   Begin VB.ComboBox comboUSB 
      Height          =   315
      ItemData        =   "VBTest.frx":00D1
      Left            =   8520
      List            =   "VBTest.frx":00DE
      TabIndex        =   2
      Top             =   2160
      Width           =   915
   End
   Begin VB.ComboBox comboBrickType 
      Height          =   315
      ItemData        =   "VBTest.frx":00F5
      Left            =   8520
      List            =   "VBTest.frx":0102
      TabIndex        =   5
      Top             =   3840
      Width           =   915
   End
   Begin VB.ComboBox comboProgramSlot 
      Height          =   315
      ItemData        =   "VBTest.frx":011F
      Left            =   8820
      List            =   "VBTest.frx":0132
      TabIndex        =   1
      Top             =   1740
      Width           =   615
   End
   Begin VB.TextBox txtSource 
      Height          =   4635
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   7875
   End
   Begin VB.Label Label7 
      Caption         =   "Waarde sensor"
      Height          =   375
      Left            =   5160
      TabIndex        =   45
      Top             =   7200
      Width           =   1095
   End
   Begin VB.Label Label6 
      Caption         =   "mins."
      Height          =   255
      Left            =   9060
      TabIndex        =   11
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label Label5 
      Caption         =   "Sleep:"
      Height          =   255
      Left            =   8160
      TabIndex        =   10
      Top             =   3000
      Width           =   435
   End
   Begin VB.Label Label4 
      Caption         =   "IR:"
      Height          =   255
      Left            =   8220
      TabIndex        =   9
      Top             =   2640
      Width           =   255
   End
   Begin VB.Label Label3 
      Caption         =   "USB:"
      Height          =   255
      Left            =   8100
      TabIndex        =   8
      Top             =   2220
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "Brick:"
      Height          =   255
      Left            =   8040
      TabIndex        =   7
      Top             =   3840
      Width           =   435
   End
   Begin VB.Label Label1 
      Caption         =   "Slot:"
      Height          =   255
      Left            =   8460
      TabIndex        =   6
      Top             =   1740
      Width           =   315
   End
   Begin VB.Menu File 
      Caption         =   "&File"
      Begin VB.Menu mnuNLasm 
         Caption         =   "New &LASM"
      End
      Begin VB.Menu MnuNMind 
         Caption         =   "New &Mindscript"
      End
      Begin VB.Menu MnuOLASM 
         Caption         =   "&Open LASM"
      End
      Begin VB.Menu MnuOMind 
         Caption         =   "O&pen Mindscript "
      End
      Begin VB.Menu MnuSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu MnuSaveAs 
         Caption         =   "Save &As"
      End
      Begin VB.Menu MnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu MnuConnect 
      Caption         =   "&Connect and Test"
   End
   Begin VB.Menu MnuRunLasm 
      Caption         =   "&Run LASM"
   End
   Begin VB.Menu MnuRMind 
      Caption         =   "R&un Mindscript"
   End
   Begin VB.Menu MnuDLasm 
      Caption         =   "&Download LASM"
   End
   Begin VB.Menu MnuDMind 
      Caption         =   "D&ownload Mindscript"
   End
   Begin VB.Menu MnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu MnuHTest 
         Caption         =   "Help about &Testclient"
      End
      Begin VB.Menu MnuHVPBrick 
         Caption         =   "Help about &VPBrick API"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents vpb As LEGOVPBrickLib.VPBrick
Attribute vpb.VB_VarHelpID = -1
Private Port$  'current port
Dim FNAme As String
Dim BlnLasm As Boolean
Dim OptMod As Integer
Dim OptSmod As Integer
Dim DirMod As Integer
Dim OptTyp As Integer
'The software is still not finished excuses for faults or
'unlogical listings
'Subroutines are given in a logical order
'1. Procedures to initiate the RCX
'2. Procedures connencted to the menu
'3. Subroutines connected to buttons etc
'4. Subroutines connected to textchange etc
'5. Procedures connencted to vpb events
'6. Procedures to close
'================================================================
'1. Procedures to initiate RCX
'================================================================
Private Sub Form_Activate()
  On Error GoTo except
  FindPort
  OpenPort
  GetStatus
  Exit Sub
 OptMod = 0: OptMode(0) = True
   OptSmod = 0: OptSMode(0) = True
 DirMod = 0: DirMode(0) = True
 OptTyp = 1: OptType(1) = True
except:
  txtResult = "Error " + Hex$(Err.Number) + " " + Err.Description
  ErrHelp
End Sub

Private Sub Form_Load()
  Set vpb = New LEGOVPBrickLib.VPBrick
  Port$ = ""
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
Private Sub OpenPort()
  On Error GoTo except
  vpb.Open Port$
  txtResult = "Opened " + Port$
  Exit Sub

except:
  txtResult = "Error " + Hex$(Err.Number) + " " + Err.Description
  ErrHelp
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

Private Function GetBrickType(nBrickType As BrickTypes) As String
  Select Case nBrickType
    Case RCXnoFirmware
      GetBrickType = "RCX (no firmware)"
    Case RCX
      GetBrickType = "RCX"
    Case Scout
      GetBrickType = "Scout"
    Case RCX2
      GetBrickType = "RCX2"
    Case MicroScout
      GetBrickType = "MicroScout"
    Case Else
      GetBrickType = "unknown PBrick"
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
  Exit Sub

except:
  txtResult = "Error " + Hex$(Err.Number) + " " + Err.Description
  ErrHelp
End Sub

Private Sub comboBrickType_Click()
  On Error GoTo except
  
  If comboBrickType = "MicroScout" Then
    vpb.BrickType = MicroScout
  ElseIf comboBrickType = "Scout" Then
    vpb.BrickType = Scout
  ElseIf comboBrickType = "RCX2" Then
    vpb.BrickType = RCX2
  End If
  Exit Sub

except:
  txtResult = "Error " + Hex$(Err.Number) + " " + Err.Description
  ErrHelp
End Sub

Private Sub comboIR_Click()
  On Error GoTo except
  vpb.BrickTxRange = comboIR.ListIndex + ShortRange
  Exit Sub

except:
  txtResult = "Error " + Hex$(Err.Number) + " " + Err.Description
  ErrHelp
End Sub

Private Sub comboProgramSlot_Click()
  On Error GoTo except
  vpb.ProgramSlot = comboProgramSlot
  Exit Sub

except:
  txtResult = "Error " + Hex$(Err.Number) + " " + Err.Description
  ErrHelp
End Sub

Private Sub comboUSB_Click()
  On Error GoTo except
  vpb.PortTxRange = comboUSB.ListIndex + ShortRange
  Exit Sub

except:
  txtResult = "Error " + Hex$(Err.Number) + " " + Err.Description
  ErrHelp
End Sub
'================================================================
'2. Procedures connected to the menu
'================================================================

Private Sub mnuNLasm_Click()
  txtSource = ""
  BlnLasm = True
  txtSource.SetFocus
End Sub

Private Sub MnuNMind_Click()
  txtSource = ""
  BlnLasm = False
  txtSource.SetFocus
End Sub

Private Sub MnuOLASM_Click()
  CdlBestand.InitDir = App.Path '& "\Lasm"
  CdlBestand.Filter = "LASM-file|*.las"
  CdlBestand.DialogTitle = "Open a LASM file"
  CdlBestand.ShowOpen
  FNAme = CdlBestand.FileName
  Fopen (FNAme)
  BlnLasm = True
 End Sub

Private Sub MnuOMind_Click()
  CdlBestand.InitDir = App.Path '& "\Mind\"
  CdlBestand.Filter = "Mindscript-File|*.mnd"
  CdlBestand.DialogTitle = "Open a Mindscript file"
  CdlBestand.ShowOpen
  FNAme = CdlBestand.FileName
  Fopen (FNAme)
  BlnLasm = False
End Sub
Private Sub MnuSave_Click()
  If FNAme = "" Then
     Call MnuSaveAs_Click
  Else
      'If BlnLasm Then
      ChDir App.Path '& "\lasm\" Else ChDir App.Path & "\mind\"
      FSave (FNAme)
      'save as lsm or mnd under a known name
  End If
End Sub

Private Sub MnuSaveAs_Click()
'Dim FNaamV, FNaamN As String
  CdlBestand.InitDir = App.Path
  If BlnLasm Then CdlBestand.Filter = "LASM-File|*.las" Else CdlBestand.Filter = "Mindscript-File|*.mnd"
  CdlBestand.DialogTitle = "Choose a name"
  CdlBestand.ShowSave
  FSave (CdlBestand.FileName)
End Sub
Private Sub Fopen(StrFName As String)
Dim Hulp As String
  txtSource = ""
  Open StrFName For Input As #1
  Do While Not EOF(1)
     Line Input #1, Hulp
     txtSource.Text = txtSource.Text & Hulp & Chr(13) & Chr(10)
  Loop
  Close
  Source$ = StrFName
End Sub
Private Sub FSave(StrFName As String)
   Open StrFName For Output As #1
   Print #1, txtSource
   Close
End Sub

Private Sub MnuExit_Click()
  vpb.Close
  Unload Me
End Sub
Private Sub MnuConnect_Click()
  Call cmdTest_Click
End Sub
Private Sub MnuRunLasm_Click()
 Call Execute(True)
End Sub
Private Sub MnuRunMind_Click()
 Call Execute(False)
End Sub
Private Sub Execute(TLasm As Boolean)
  Dim nResult As Long
  Dim nPos As Long
  Dim result As Variant
  Dim x As String
  
  On Error GoTo except
  Source$ = txtSource
  If TLasm Then nResult = vpb.Execute(Source$, LASM, nPos, result) Else nResult = vpb.Execute(Source$, "", nPos, result)
  txtResult = "Result: " + Str(nResult)
  If VarType(result) = vbByte + vbArray Then
    If UBound(result) > LBound(result) Then
      txtResult = txtResult & ", variant result: "
      For i = LBound(result) To UBound(result)
        txtResult = txtResult + Str(result(i)) + " "
      Next i
    End If
  End If
  Exit Sub

except:
  txtSource.SelStart = nPos
  txtResult = "Error " + Hex$(Err.Number) + " " + Err.Description
  txtSource.SetFocus
  ErrHelp
End Sub


Private Sub MnuDLasm_Click()
  Call Download(True)
End Sub

Private Sub MnuDMind_Click()
  Call Download(False)
End Sub
Private Sub Download(TLasm As Boolean)
  Dim nPos As Long
  
  On Error GoTo except
  
  Source$ = txtSource
  If TLasm Then vpb.Download Source$, LASM, nPos Else vpb.Download Source$, , nPos
  Exit Sub

except:
  txtSource.SelStart = nPos
  txtResult = "Error " + Hex$(Err.Number) + " " + Err.Description
  txtSource.SetFocus
  ErrHelp
End Sub

Private Sub MnuHTest_Click()
  frmBrowser.StartingAddress = App.Path & "\help.htm"
  frmBrowser.Show
End Sub

Private Sub MnuHVPBrick_Click()
  With CdlBestand
    .HelpFile = "VPB.hlp"
    .HelpCommand = cdlHelpContents
    .ShowHelp
  End With
End Sub

'========================================================
'3. Subroutines connected to buttons etc
'========================================================
Private Sub CmdMonitor_Click()
 MsgBox "works with ScriptEd,but not in VB should cause a recieved-event when IR-message is send by RCX"
 If CmdMonitor.Caption = "MonitorIR" Then
   vpb.Monitor (True)
   CmdMonitor.Caption = "No MonitorIR"
 Else
   vpb.Monitor (False)
   CmdMonitor.Caption = "MonitorIR"
  End If
End Sub

Private Sub cmdStatus_Click()
  GetStatus
End Sub

Private Sub GetStatus()
  Dim nStatus As StatusResult
  Dim nBrickType As BrickTypes

  On Error GoTo except
  
  nBrickType = vpb.Status(CheckBrickType)
  SetBrickType (nBrickType)
  nStatus = vpb.Status(BrickStatus)
  Select Case nStatus
    Case StatusReady
      txtResult = GetBrickType(vpb.Status(CheckBrickType)) + " ready"
      If nBrickType = RCX2 Then
        comboProgramSlot = vpb.ProgramSlot
        comboIR.ListIndex = vpb.BrickTxRange - ShortRange
        txtSleep = Str(vpb.PowerDownTime)
      End If
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
      txtResult = GetBrickType(nBrickType) + " bad battery"
    Case BrickMismatch
      txtResult = GetBrickType(nBrickType) + " brick type mismatch, expecting " + GetBrickType(vpb.BrickType)
      vpb.BrickType = nBrickType
    Case BadComms
      txtResult = "Bad comms"
    Case Else
      txtResult = "Status: " + Str(nStatus)
  End Select

  comboUSB.ListIndex = vpb.PortTxRange - ShortRange
  Exit Sub

except:
  'ignore
End Sub


Private Sub CmdPlaySound_Click()
  vpb.Execute "sound " & Combo1.ListIndex + 1
End Sub

Private Sub CmdOutput_Click()
Dim commando As String
  commando = "dir " & DirMod & "," & (Combo2.ListIndex + 1)
  vpb.Execute commando, LASM
  commando = "out " & OptMod & "," & (Combo2.ListIndex + 1)
  vpb.Execute commando, LASM
End Sub

Private Sub CmdInput_Click()
  Dim nResult As Long
  Dim nPos As Long
  Dim result As Variant
  Dim commando As String
 commando = "poll 9," & CmbSens.ListIndex  'waarde senor 0/1/2 opvragen
 nResult = vpb.Execute(commando, LASM, nPos, result)
   Text1.Text = Str(nResult)
End Sub

Private Sub CmdMakeSound_Click()
 Dim commando As String
  commando = "playt " & TxtFreq.Text & "," & TxtTime.Text
  vpb.Execute commando, LASM
End Sub

Private Sub CmdSetWatch_Click()
Dim commando As String
  commando = "setw " & Hour(Now) & "," & Minute(Now)
  vpb.Execute commando, LASM
End Sub

Private Sub DirMode_Click(Index As Integer)
  DirMod = Index
End Sub



Private Sub OptMode_Click(Index As Integer)
  OptMod = Index
End Sub

Private Sub OptSMode_Click(Index As Integer)
 Dim commando As String
 commando = "senm " & CmbSens.ListIndex & "," & Index & ",0" 'Modus sensor NB einde ,1=> incrementeel
 vpb.Execute commando, LASM
    
End Sub

Private Sub OptType_Click(Index As Integer)
 Dim commando As String
  commando = "sent " & CmbSens.ListIndex & "," & Index 'Type sensor instellen
  vpb.Execute commando, LASM
End Sub


Private Sub chkHelp_Click()
 With CdlBestand
    .HelpFile = "VPB.hlp"
    .HelpCommand = cdlHelpContents
    .ShowHelp
  End With
End Sub

Private Sub cmdFirmware_Click()
  On Error GoTo except
  vpb.DownloadFirmware "firm0328.lgo"  'NB was alles weg (path?)
  Exit Sub

except:
  txtResult = "Error " + Hex$(Err.Number) + " " + Err.Description
  ErrHelp
End Sub

Private Sub ErrHelp()
  If chkHelp.Value = Checked Then
    If Err.HelpFile = "" Then
      MsgBox "No Err.HelpFile"
      Exit Sub
    End If
    
    With CdlBestand
      .HelpFile = Err.HelpFile
      .HelpContext = Val(Err.HelpContext)
      .HelpCommand = cdlHelpContext
      .ShowHelp
    End With
  End If
End Sub
'========================================================
'4. Subroutines connected to textchange etc
'========================================================
Private Sub TxtFreq_Click()
  TxtFreq.Text = ""
End Sub

Private Sub TxtTime_Click()
  TxtTime.Text = ""
End Sub


Private Sub txtSleep_Change()
  On Error GoTo except
  vpb.PowerDownTime = Val(txtSleep)
  Exit Sub

except:
  txtResult = "Error " + Hex$(Err.Number) + " " + Err.Description
  ErrHelp
End Sub

Private Sub txtSource_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF1 Then
    help$ = txtSource.SelText
    If Len(help$) = 0 Then
      nStart = txtSource.SelStart
      nEnd = txtSource.SelStart
      While nStart > 0 And Mid$(txtSource, nStart + 1, 1) <> " "
        nStart = nStart - 1
      Wend
      While nEnd < Len(txtSource) And Mid$(txtSource, nEnd + 1, 1) <> " "
        nEnd = nEnd + 1
      Wend
      help$ = Mid$(txtSource, nStart + 1, nEnd - nStart)
    End If
    With CommonDialog1
      .HelpFile = "VPB.hlp"
      If Len(help$) = 0 Then
        .HelpCommand = cdlHelpContents
      Else
        .HelpKey = help$
        .HelpCommand = cdlHelpKey
      End If
      .ShowHelp
    End With
  End If
End Sub
'========================================================
'5. Subroutines connected to vpb events
'========================================================
Private Sub vpb_DownloadDone(ByVal nErrorCode As Long)
  If nErrorCode = 0 Then
    txtResult = "Downloaded ok"
  Else
    txtResult = "Download error: " + Hex$(nErrorCode)
  End If
End Sub

Private Sub vpb_DownloadProgress(ByVal nPercent As Long)
  txtResult = "Download progress: " + Str(nPercent) + "%"
End Sub

Private Sub vpb_Received(ByVal strData As String)
    TxtMonitorIR.Text = strDate
     '1=  0x55,0xff,0x00,0xf7,0x08,0x02,0xfd,0xf9,0x06
     '2= 0x55,0xff,0x00,0xf7,0x08,0x03,0xfc,0xfa,0x05
     '3= 0x55,0xff,0x00,0xf7,0x08,0x04,0xfb,0xfb,0x04
     'Message na insteling message 1/2/3 in RcX, NB 6e =message+1, 7e=15-message
End Sub
'===============================================================
'6. Procedure to close
'===============================================================
Private Sub cmdClose_Click()
  vpb.Close
  txtResult = "Ok"
End Sub

Private Sub Form_Terminate()
  vpb.Close
End Sub

