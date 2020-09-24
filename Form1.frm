VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Graphic VU Meter"
   ClientHeight    =   2865
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6705
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2865
   ScaleWidth      =   6705
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   0
      Top             =   120
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0C0C0&
      Height          =   2850
      Left            =   0
      ScaleHeight     =   2790
      ScaleWidth      =   6585
      TabIndex        =   0
      Top             =   0
      Width           =   6645
      Begin MSComctlLib.ProgressBar ProgressBar 
         Height          =   2580
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   180
         _ExtentX        =   318
         _ExtentY        =   4551
         _Version        =   393216
         Appearance      =   0
         Orientation     =   1
      End
      Begin MSComctlLib.ProgressBar ProgressBar 
         Height          =   2580
         Index           =   1
         Left            =   480
         TabIndex        =   2
         Top             =   120
         Width           =   180
         _ExtentX        =   318
         _ExtentY        =   4551
         _Version        =   393216
         Appearance      =   0
         Orientation     =   1
      End
      Begin MSComctlLib.ProgressBar ProgressBar 
         Height          =   2580
         Index           =   2
         Left            =   840
         TabIndex        =   3
         Top             =   120
         Width           =   180
         _ExtentX        =   318
         _ExtentY        =   4551
         _Version        =   393216
         Appearance      =   0
         Orientation     =   1
      End
      Begin MSComctlLib.ProgressBar ProgressBar 
         Height          =   2580
         Index           =   3
         Left            =   1200
         TabIndex        =   4
         Top             =   120
         Width           =   180
         _ExtentX        =   318
         _ExtentY        =   4551
         _Version        =   393216
         Appearance      =   0
         Orientation     =   1
      End
      Begin MSComctlLib.ProgressBar ProgressBar 
         Height          =   2580
         Index           =   4
         Left            =   1560
         TabIndex        =   5
         Top             =   120
         Width           =   180
         _ExtentX        =   318
         _ExtentY        =   4551
         _Version        =   393216
         Appearance      =   0
         Orientation     =   1
      End
      Begin MSComctlLib.ProgressBar ProgressBar 
         Height          =   2580
         Index           =   5
         Left            =   1920
         TabIndex        =   6
         Top             =   120
         Width           =   180
         _ExtentX        =   318
         _ExtentY        =   4551
         _Version        =   393216
         Appearance      =   0
         Orientation     =   1
      End
      Begin MSComctlLib.ProgressBar ProgressBar 
         Height          =   2580
         Index           =   6
         Left            =   2280
         TabIndex        =   7
         Top             =   120
         Width           =   180
         _ExtentX        =   318
         _ExtentY        =   4551
         _Version        =   393216
         Appearance      =   0
         Orientation     =   1
      End
      Begin MSComctlLib.ProgressBar ProgressBar 
         Height          =   2580
         Index           =   7
         Left            =   2640
         TabIndex        =   8
         Top             =   120
         Width           =   180
         _ExtentX        =   318
         _ExtentY        =   4551
         _Version        =   393216
         Appearance      =   0
         Orientation     =   1
      End
      Begin MSComctlLib.ProgressBar ProgressBar 
         Height          =   2580
         Index           =   8
         Left            =   3000
         TabIndex        =   9
         Top             =   120
         Width           =   180
         _ExtentX        =   318
         _ExtentY        =   4551
         _Version        =   393216
         Appearance      =   0
         Orientation     =   1
      End
      Begin MSComctlLib.ProgressBar ProgressBar 
         Height          =   2580
         Index           =   9
         Left            =   3360
         TabIndex        =   10
         Top             =   120
         Width           =   180
         _ExtentX        =   318
         _ExtentY        =   4551
         _Version        =   393216
         Appearance      =   0
         Orientation     =   1
      End
      Begin MSComctlLib.ProgressBar ProgressBar 
         Height          =   2580
         Index           =   10
         Left            =   3720
         TabIndex        =   11
         Top             =   120
         Width           =   180
         _ExtentX        =   318
         _ExtentY        =   4551
         _Version        =   393216
         Appearance      =   0
         Orientation     =   1
      End
      Begin MSComctlLib.ProgressBar ProgressBar 
         Height          =   2580
         Index           =   11
         Left            =   4080
         TabIndex        =   12
         Top             =   120
         Width           =   180
         _ExtentX        =   318
         _ExtentY        =   4551
         _Version        =   393216
         Appearance      =   0
         Orientation     =   1
      End
      Begin MSComctlLib.ProgressBar ProgressBar 
         Height          =   2580
         Index           =   12
         Left            =   4440
         TabIndex        =   13
         Top             =   120
         Width           =   180
         _ExtentX        =   318
         _ExtentY        =   4551
         _Version        =   393216
         Appearance      =   0
         Orientation     =   1
      End
      Begin MSComctlLib.ProgressBar ProgressBar 
         Height          =   2580
         Index           =   13
         Left            =   4800
         TabIndex        =   14
         Top             =   120
         Width           =   180
         _ExtentX        =   318
         _ExtentY        =   4551
         _Version        =   393216
         Appearance      =   0
         Orientation     =   1
      End
      Begin MSComctlLib.ProgressBar ProgressBar 
         Height          =   2580
         Index           =   14
         Left            =   5160
         TabIndex        =   15
         Top             =   120
         Width           =   180
         _ExtentX        =   318
         _ExtentY        =   4551
         _Version        =   393216
         Appearance      =   0
         Orientation     =   1
      End
      Begin MSComctlLib.ProgressBar ProgressBar 
         Height          =   2580
         Index           =   15
         Left            =   5520
         TabIndex        =   16
         Top             =   120
         Width           =   180
         _ExtentX        =   318
         _ExtentY        =   4551
         _Version        =   393216
         Appearance      =   0
         Orientation     =   1
      End
      Begin MSComctlLib.ProgressBar ProgressBar 
         Height          =   2580
         Index           =   16
         Left            =   5880
         TabIndex        =   17
         Top             =   120
         Width           =   180
         _ExtentX        =   318
         _ExtentY        =   4551
         _Version        =   393216
         Appearance      =   0
         Orientation     =   1
      End
      Begin MSComctlLib.ProgressBar ProgressBar 
         Height          =   2580
         Index           =   17
         Left            =   6240
         TabIndex        =   18
         Top             =   120
         Width           =   180
         _ExtentX        =   318
         _ExtentY        =   4551
         _Version        =   393216
         Appearance      =   0
         Orientation     =   1
      End
      Begin VB.Line Line5 
         Index           =   2
         X1              =   -120
         X2              =   7320
         Y1              =   2640
         Y2              =   2640
      End
      Begin VB.Line Line5 
         Index           =   1
         X1              =   0
         X2              =   7200
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Line Line5 
         Index           =   0
         X1              =   0
         X2              =   7320
         Y1              =   75
         Y2              =   75
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim hmixer As Long                      ' mixer handle
Dim inputVolCtrl As MIXERCONTROL        ' waveout volume control
Dim outputVolCtrl As MIXERCONTROL       ' microphone volume control
Dim rc As Long                          ' return code
Dim OK As Boolean                       ' boolean return code
Dim mxcd As MIXERCONTROLDETAILS         ' control info
Dim vol As MIXERCONTROLDETAILS_SIGNED   ' control's signed value
Dim volume As Long                      ' volume value
Dim volHmem As Long                     ' Volume Buffer
Private VU As VULights                  ' Volume Unit Values
Private FreqNum As Frequency
Private Declare Sub ReleaseCapture Lib "user32" ()
Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2

Private Sub VolVal(VolIs As Long, VolFreq As Double)
For FreqNum = 0 To 9
Next FreqNum
VolIs = volume * 327.67
VolFreq = VU.Freq(FreqNum)
VU.FreqVal = VolIs * VolFreq
End Sub

Private Sub ActivateVolumeUnits()
    For i = 0 To 17
        Lights (i)
    Next i
End Sub

Private Sub Form_Load()
    StayOnTop Me
    
    Timer1.Interval = 6.25
   ' Open the mixer specified by DEVICEID
   rc = mixerOpen(hmixer, DEVICEID, 0, 0, 0)
   If ((MMSYSERR_NOERROR <> rc)) Then
       MsgBox "Couldn't open the mixer."
       Exit Sub
   End If
   ' Get the output volume meter
   OK = GetControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_WAVEOUT, MIXERCONTROL_CONTROLTYPE_PEAKMETER, outputVolCtrl)
   If (OK = True) Then
   
   ' Set frequencies for the Volume Units
    ProgressBar(0).Max = Frequency.Freq31hz + 1
    ProgressBar(0).Min = Frequency.Freq31hz
    ProgressBar(1).Max = Frequency.Freq60hz + 1
    ProgressBar(1).Min = Frequency.Freq60hz
    ProgressBar(2).Max = Frequency.Freq62hz + 1
    ProgressBar(2).Min = Frequency.Freq62hz
    ProgressBar(3).Max = Frequency.Freq125hz + 1
    ProgressBar(3).Min = Frequency.Freq125hz
    ProgressBar(4).Max = Frequency.Freq170hz + 1
    ProgressBar(4).Min = Frequency.Freq170hz
    ProgressBar(5).Max = Frequency.Freq250hz + 1
    ProgressBar(5).Min = Frequency.Freq250hz
    ProgressBar(6).Max = Frequency.Freq310hz + 1
    ProgressBar(6).Min = Frequency.Freq310hz
    ProgressBar(7).Max = Frequency.Freq500hz + 1
    ProgressBar(7).Min = Frequency.Freq500hz
    ProgressBar(8).Max = Frequency.Freq600hz + 1
    ProgressBar(8).Min = Frequency.Freq600hz
    ProgressBar(9).Max = Frequency.Freq1khz + 1
    ProgressBar(9).Min = Frequency.Freq1khz
    ProgressBar(10).Max = Frequency.Freq2khz + 1
    ProgressBar(10).Min = Frequency.Freq2khz
    ProgressBar(11).Max = Frequency.Freq3khz + 1
    ProgressBar(11).Min = Frequency.Freq3khz
    ProgressBar(12).Max = Frequency.Freq4khz + 1
    ProgressBar(12).Min = Frequency.Freq4khz
    ProgressBar(13).Max = Frequency.Freq6khz + 1
    ProgressBar(13).Min = Frequency.Freq6khz
    ProgressBar(14).Max = Frequency.Freq8khz + 1
    ProgressBar(14).Min = Frequency.Freq8khz
    ProgressBar(15).Max = Frequency.Freq12khz + 1
    ProgressBar(15).Min = Frequency.Freq12khz
    ProgressBar(16).Max = Frequency.Freq14khz + 1
    ProgressBar(16).Min = Frequency.Freq14khz
    ProgressBar(17).Max = Frequency.Freq16khz + 1
    ProgressBar(17).Min = Frequency.Freq16khz
   Else
      MsgBox "Couldn't get waveout meter"
   End If
   
   ' Initialize mixercontrol structure
   mxcd.cbStruct = Len(mxcd)
   volHmem = GlobalAlloc(&H0, Len(volume))  ' Allocate a buffer for the volume value
   mxcd.paDetails = GlobalLock(volHmem)
   mxcd.cbDetails = Len(volume)
   mxcd.cChannels = 1
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If (fRecording = True) Then
       StopInput
   End If
   GlobalFree volHmem
   Unload Me
   End
End Sub

Private Sub Form_Resize()
On Error Resume Next
    
    Picture1.Height = ScaleHeight
    
    For i = 0 To 17
        ProgressBar(i).Height = Me.Height - 675
    Next i
    
    Line5(2).Y1 = Me.Height - 675
    Line5(2).Y2 = Me.Height - 675
    Line5(1).Y1 = Picture1.Height / 2
    Line5(1).Y2 = Picture1.Height / 2
    
    Me.Width = 6780
End Sub

Private Sub Form_Terminate()
   If (fRecording = True) Then
       StopInput
   End If
   GlobalFree volHmem
   Unload Me
   End
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If (fRecording = True) Then
       StopInput
   End If
   GlobalFree volHmem
   Unload Me
   End
End Sub

Private Sub Timer1_Timer()
    VU.VolLev = volume / 327.67
    If (volume < 0) Then volume = -volume
    ' Get the current output level
    If (1 = 1) Then
    mxcd.dwControlID = outputVolCtrl.dwControlID
    mxcd.item = outputVolCtrl.cMultipleItems
    rc = mixerGetControlDetails(hmixer, mxcd, MIXER_GETCONTROLDETAILSF_VALUE)
    CopyStructFromPtr volume, mxcd.paDetails, Len(volume)
    If (volume < 0) Then volume = -volume
    End If
    
    If volume = 0 Then
        For i = 0 To 17
            ProgressBar(i).Value = ProgressBar(i).Min
        Next i
        Exit Sub
    End If
    
    ActivateVolumeUnits
End Sub

Private Sub Lights(PBIndex As Integer)

Select Case PBIndex
    Case 0
        FreqNum = Freq31hz
        For VU.InOutLev = CDbl(VU.VolLev * 0.032258064516129) To FreqNum
        Next VU.InOutLev
        ProgressBar(PBIndex).Value = VU.InOutLev

    Case 1
        FreqNum = Freq60hz
        For VU.InOutLev = CDbl(VU.VolLev * 0.161290322580645) To FreqNum
        Next VU.InOutLev
        ProgressBar(PBIndex).Value = VU.InOutLev

    Case 2
        FreqNum = Freq62hz
        For VU.InOutLev = CDbl(VU.VolLev * 1.61290322580645E-02) To FreqNum
        Next VU.InOutLev
        ProgressBar(PBIndex).Value = VU.InOutLev

    Case 3
        FreqNum = Freq125hz
        For VU.InOutLev = CDbl(VU.VolLev * 0.8) To FreqNum
        Next VU.InOutLev
        ProgressBar(PBIndex).Value = VU.InOutLev

    Case 4
        FreqNum = Freq170hz
        For VU.InOutLev = CDbl(VU.VolLev * 0.6) To FreqNum '0.06 not right
        Next VU.InOutLev
        ProgressBar(PBIndex).Value = VU.InOutLev

    Case 5
        FreqNum = Freq250hz
        For VU.InOutLev = CDbl(VU.VolLev * 0.4) To FreqNum
        Next VU.InOutLev
        ProgressBar(PBIndex).Value = VU.InOutLev

    Case 6
        FreqNum = Freq310hz
        For VU.InOutLev = CDbl(VU.VolLev * 0.35) To FreqNum '0.35 not correct
        Next VU.InOutLev
        ProgressBar(PBIndex).Value = VU.InOutLev

    Case 7
        FreqNum = Freq500hz
        For VU.InOutLev = CDbl(VU.VolLev * 0.2) To FreqNum
        Next VU.InOutLev
        ProgressBar(PBIndex).Value = VU.InOutLev

    Case 8
        FreqNum = Freq600hz
        For VU.InOutLev = CDbl(VU.VolLev * 0.1) To FreqNum
        Next VU.InOutLev
        ProgressBar(PBIndex).Value = VU.InOutLev

    Case 9
        FreqNum = Freq1khz
        For VU.InOutLev = CDbl(VU.VolLev * 0.01) To FreqNum
        Next VU.InOutLev
        ProgressBar(PBIndex).Value = VU.InOutLev

    Case 10
        FreqNum = Freq2khz
        For VU.InOutLev = CDbl(VU.VolLev * 0.02) To FreqNum
        Next VU.InOutLev
    ProgressBar(PBIndex).Value = VU.InOutLev

    Case 11
        FreqNum = Freq3khz
        For VU.InOutLev = CDbl(VU.VolLev * 0.03) To FreqNum
        Next VU.InOutLev
        ProgressBar(PBIndex).Value = VU.InOutLev

    Case 12
        FreqNum = Freq4khz
        For VU.InOutLev = CDbl(VU.VolLev * 0.04) To FreqNum
        Next VU.InOutLev
        ProgressBar(PBIndex).Value = VU.InOutLev

    Case 13
        FreqNum = Freq6khz
        For VU.InOutLev = CDbl(VU.VolLev * 0.06) To FreqNum
        Next VU.InOutLev
        ProgressBar(PBIndex).Value = VU.InOutLev

    Case 14
        FreqNum = Freq8khz
        For VU.InOutLev = CDbl(VU.VolLev * 0.08) To FreqNum
        Next VU.InOutLev
        ProgressBar(PBIndex).Value = VU.InOutLev

    Case 15
        FreqNum = Freq12khz
        For VU.InOutLev = CDbl(VU.VolLev * 0.12) To FreqNum
        Next VU.InOutLev
        ProgressBar(PBIndex).Value = VU.InOutLev

    Case 16
        FreqNum = Freq14khz
        For VU.InOutLev = CDbl(VU.VolLev * 0.14) To FreqNum
        Next VU.InOutLev
        ProgressBar(PBIndex).Value = VU.InOutLev

    Case 17
        FreqNum = Freq16khz
        For VU.InOutLev = CDbl(VU.VolLev * 0.16) To FreqNum
        Next VU.InOutLev
        ProgressBar(PBIndex).Value = VU.InOutLev

End Select
End Sub
