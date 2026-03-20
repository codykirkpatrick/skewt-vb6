VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Skew-T Diagram"
   ClientHeight    =   7755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11010
   Icon            =   "SkewT.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7755
   ScaleWidth      =   11010
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPolling 
      Caption         =   "Polling"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   12
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   2160
      Top             =   0
   End
   Begin VB.Frame Frame3 
      Caption         =   "Cursor Location"
      Height          =   735
      Left            =   120
      TabIndex        =   10
      Top             =   6960
      Width           =   2175
      Begin VB.Label lblReadout 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "Hover inside diagram..."
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Data"
      Height          =   2535
      Left            =   120
      TabIndex        =   8
      Top             =   3120
      Width           =   2175
      Begin VB.Label lblStats 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Caption         =   "Awaiting data..."
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2175
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.CommandButton cmdReplot 
      Caption         =   "Replot Same File"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   1080
      Width           =   1935
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2640
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save Image"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   2040
      Width           =   1935
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset Plot"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1560
      Width           =   1935
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   7215
      Left            =   2400
      MousePointer    =   2  'Cross
      ScaleHeight     =   7185
      ScaleWidth      =   8505
      TabIndex        =   1
      Top             =   480
      Width           =   8535
   End
   Begin VB.CommandButton cmdPlotFile 
      Caption         =   "Plot File"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Caption         =   "Actions"
      Height          =   2655
      Left            =   120
      TabIndex        =   7
      Top             =   360
      Width           =   2175
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Status messages appear here..."
      Height          =   1095
      Left            =   240
      TabIndex        =   4
      Top             =   5760
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Skew-T Diagram"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   2
      Top             =   120
      Width           =   8175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const b As Double = 17.625
Private Const c As Double = 243.04
Private Const skewFactor As Single = 3#

Dim CurrentFilePath As String
Private LastFileTimeStamp As Variant

Private Sub cmdPlotFile_Click()
  On Error GoTo CancelHandler
  With CommonDialog1
    .Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
    .DialogTitle = "Select Sounding Data File"
    .InitDir = App.Path
    .CancelError = True
    .ShowOpen
    CurrentFilePath = .fileName
  End With
  
  PlotData CurrentFilePath
  
  Timer1.Enabled = True
  Label2.Caption = Label2.Caption & " Auto-polling enabled."
  cmdPolling.Caption = "Polling is active"
  
  Exit Sub

CancelHandler:
  If Err.Number <> cdlCancel Then
    MsgBox "Error opening file: " & Err.Description
  End If

End Sub

Private Function PlotData(fileName As String)
  Dim T As Single, z As Single, Dwpt As Single, RH As Single
  Dim TRAW As Single, ZRaw As Single, RRaw As Single, PRaw As Single
  Dim lastT As Single, lastZ As Single, lastD As Single
  Dim isFirstPoint As Boolean, isTSPOT As Boolean
  Dim pValues As Variant
  
  Dim surfZ As Single, surfT As Single, surfTd As Single
  
  Dim i As Integer
  Dim x1 As Single, y1 As Single
  Dim x2 As Single, y2 As Single
  Dim fileNum As Integer
  Dim fileType As Integer
  Dim shortName As String
  Dim lastSlashIndex As Integer
    
  ' Pressure levels that will plot at right when they appear
  pValues = Array(850, 700, 500, 300, 200, 100)
  
  ' Set up file input and read the three header lines
  fileNum = FreeFile
  Open fileName For Input Shared As #fileNum
  For i = 1 To 3
    Line Input #fileNum, Junk
  Next i
  
  ' Get shortName to identify file type
  lastSlashIndex = InStrRev(fileName, "\")
  If lastSlashIndex > 0 Then
    shortName = Mid(fileName, lastSlashIndex + 1)
  Else
    shortName = fileName
  End If
  fileType = InStr(1, fileName, "TSPOTINT", vbTextCompare)
  
  ' Plot the actual data
  isFirstPoint = True
  Do While Not EOF(fileNum)
    Line Input #fileNum, Row
    If fileType > 0 Then
      PRaw = Val(Mid$(Row, 16, 15))
      TRAW = Val(Mid$(Row, 31, 15))
      RRaw = Val(Mid$(Row, 46, 15))
      ZRaw = Val(Mid$(Row, 91, 15))
    Else
      PRaw = Val(Mid$(Row, 31, 15))
      TRAW = Val(Mid$(Row, 46, 15))
      RRaw = Val(Mid$(Row, 61, 15))
      ZRaw = Val(Mid$(Row, 1, 15))
    End If
    z = ZRaw / 1000#
    If z > 16.01 Then
      Exit Do
    End If
    T = TRAW
    RH = RRaw
    TD = GetDewpoint(T, RH)
    
    ' If a new P level is hit, plot that tick mark on the right
    For i = LBound(pValues) To UBound(pValues)
      If lastP >= pValues(i) And PRaw < pValues(i) Then
        Picture1.ForeColor = &HCC8822
        Picture1.DrawWidth = 1
        Picture1.Line (38, z)-(43, z)
        Picture1.CurrentX = 43.5
        Picture1.CurrentY = z + 0.25
        Picture1.Print pValues(i) & "mb"
      End If
    Next i
        
    ' Format surface data so it can be displayed in the data label
    If isFirstPoint Then
      surfZ = z: surfT = T: surfTd = TD: surfP = PRaw
      strsurfT = Format(surfT, "0.0")
      strsurfTD = Format(surfTd, "0.0")
      strsurfP = Format(surfP, "0.0")
      strsurfT = Right(Space(5) & strsurfT, 5)
      strsurfTD = Right(Space(5) & strsurfTD, 5)
      strsurfP = Right(Space(6) & strsurfP, 6)
      isFirst = False
    Else
      Picture1.DrawWidth = 2
      Picture1.Line (lastT + (lastZ * skewFactor), lastZ)-(T + (z * skewFactor), z), vbRed
      Picture1.Line (lastTD + (lastZ * skewFactor), lastZ)-(TD + (z * skewFactor), z), vbGreen
    End If
    
    ' Format last-line-read data so it can be displayed in the data label
    lastT = T
    lastTD = TD
    lastZ = z
    lastP = PRaw
    strlastT = Format(lastT, "0.0")
    strlastTD = Format(lastTD, "0.0")
    strlastZ = Format(lastZ, "0.00")
    strPressure = Format(PRaw, "00.0")
    strlastT = Right(Space(5) & strlastT, 5)
    strlastTD = Right(Space(5) & strlastTD, 5)
    strlastZ = Right(Space(5) & strlastZ, 5)
    strPressure = Right(Space(6) & strPressure, 6)

    isFirstPoint = False
    
  Loop

  lblStats.Caption = "--- CONDITIONS ---" & vbCrLf & _
    vbCrLf & _
    "      Surface" & vbCrLf & _
    "   TMPC " & strsurfT & " C" & vbCrLf & _
    "   DWPC " & strsurfTD & " C" & vbCrLf & _
    "   PRES " & strsurfP & " mb" & vbCrLf & _
    vbCrLf & _
    "     Last Line" & vbCrLf & _
    "   TMPC " & strlastT & " C" & vbCrLf & _
    "   DWPC " & strlastTD & " C" & vbCrLf & _
    "   HGHT " & strlastZ & " km" & vbCrLf & _
    "   PRES " & strPressure & " mb"

  Close #fileNum
  
  Picture1.CurrentX = 10
  Picture1.CurrentY = 15.8
  Picture1.ForeColor = &HFFFFFF
  Picture1.Print shortName
  
  Label1.Caption = fileName
  Label2.Caption = "Skew-T plotted!"
  
  Beep
  
End Function

Private Sub cmdSave_Click()
  On Error GoTo SaveError
  
  Dim filePath As String
  Dim shortName As String
  Dim lastSlashIndex As Integer
  
  ' Get name of Skew-T file
  lastSlashIndex = InStrRev(CurrentFilePath, "\")
  If lastSlashIndex > 0 Then
    shortName = Mid(CurrentFilePath, lastSlashIndex + 1)
  Else
    shortName = CurrentFilePath
  End If
  
  ' Set storage location
  filePath = App.Path & "\" & shortName & ".bmp"
  
  ' Save image and update status window
  SavePicture Picture1.Image, filePath
  Label2.Caption = "Image saved!  Check current folder."
  Beep
  
  Exit Sub
  
SaveError:
  MsgBox "Error saving image", vbCritical, "File Error"
  Label2.Caption = "Image could not be saved."
  
End Sub

Private Function DrawBasics(sFactor As Single)
  Picture1.ForeColor = vbRed
  Picture1.DrawStyle = vbDot
  Picture1.DrawWidth = 1
  
  ' Temperature lines -- this range is the range at the bottom of the plot
  For T = -60 To 40 Step 20
    If T <= 0 Then
      Picture1.ForeColor = &HC87F09
    Else
      Picture1.ForeColor = vbRed
    End If
    
    Picture1.Line (T, 0)-(T + (16 * sFactor), 16)
    Picture1.CurrentX = T - 2
    Picture1.CurrentY = 0.01
    Picture1.Print T
  Next T

  ' Height lines
  Picture1.ForeColor = vbWhite
  Picture1.DrawStyle = vbSolid
  Picture1.DrawWidth = 1
  For i = 0 To 15 Step 3
    Picture1.Line (-44.1, i)-(50, i)
    Picture1.CurrentX = -50
    Picture1.CurrentY = i + 0.3
    HgtValue = i
    HgtLabel = Format(HgtValue, "0 km")
    HgtLabel = Right(Space(5) & HgtLabel, 5)
    Picture1.Print HgtLabel
  Next i
  
  Label2.Caption = "Finished drawing base plot."

End Function

Private Function GetDewpoint(T As Single, RH As Single) As Single
  Dim gamma As Double
  If RH < 0.1 Then RH = 0.1
  gamma = Log(RH / 100) + ((b * T) / (c + T))
  GetDewpoint = (c * gamma) / (b - gamma)

  ' I think the program moves so fast that this message is not helpful
  Label2.Caption = "Calculating a dewpoint from RH..."

End Function

Private Function GetMixingRatioTemp(w As Single, z As Single) As Single
  Dim P As Double
  Dim e As Double
  Dim gamma As Double
  P = 1013.25 * Exp(-z / 8.25)
  e = (w * P) / (621.97 + w)
  gamma = Log(e / 6.112)
  GetMixingRatioTemp = (243.04 * gamma) / (17.625 - gamma)

  ' I think the program moves so fast that this message is not helpful
  Label2.Caption = "Calculating a mixing ratio..."

End Function

Private Sub DrawMixingRatios(sFactor As Single)
  Dim wValues As Variant
  Dim w As Variant
  Dim z As Single, stepZ As Single
  Dim T1 As Single, T2 As Single
  Dim x1 As Single, x2 As Single
  
  ' I think the program moves so fast that this message is not helpful
  Label2.Caption = "Drawing mixing ratio lines..."
  
  ' Which mixing ratio lines to plot (g/kg, obv.)
  wValues = Array(0.5, 1, 5, 10, 15, 20)
  
  ' Height interval for recalculating the line slope
  stepZ = 0.25
  
  Picture1.DrawWidth = 1
  Picture1.DrawStyle = 1
  Picture1.ForeColor = &H2B2BA5
  
  For Each w In wValues
    For z = 0 To 5.75 Step stepZ
      T1 = GetMixingRatioTemp(CSng(w), z)
      x1 = T1 + (z * sFactor)
      
      T2 = GetMixingRatioTemp(CSng(w), z + stepZ)
      x2 = T2 + ((z + stepZ) * sFactor)
      
      Picture1.Line (x1, z)-(x2, z + stepZ)
      Picture1.CurrentX = GetMixingRatioTemp(CSng(w), 5.7) + (5.7 * sFactor)
      Picture1.CurrentY = 6.6
      Picture1.Print w
    Next z
  Next w
  
  ' I think the program moves so fast that this message is not helpful
  Label2.Caption = "Drawing mixing ratio lines...finished."

  Picture1.DrawStyle = 0
End Sub

Private Sub DrawCompleteGrid()
  Picture1.Cls
  Picture1.Scale (-50, 16)-(50, -0.5)
  
  DrawMixingRatios skewFactor
  DrawBasics skewFactor

End Sub

Private Sub cmdReset_Click()
  Picture1.Cls
  DrawCompleteGrid
  Label2.Caption = "Skew-T cleared!"
  Label1.Caption = "Skew-T Diagram"
  lblStats.Caption = "Awaiting data..."
End Sub

Private Sub cmdReplot_Click()
  If CurrentFilePath = "" Then
    MsgBox "Please open a file first using the Plot button.", vbExclamation
    Label2.Caption = "No file selected!"
    Exit Sub
  End If
  
  Label2.Caption = "Plotting file " & CurrentFilePath
  DrawCompleteGrid
  PlotData CurrentFilePath
  
End Sub

Private Sub cmdPolling_Click()
    If CurrentFilePath = "" Then
        MsgBox "Please plot a file first.": Exit Sub
    End If
    If Timer1.Enabled = True Then
        Timer1.Enabled = False
        cmdPolling.Caption = "Start polling..."
        cmdPolling.BackColor = &H8000000F
        Label2.Caption = Label2.Caption & " Polling disabled."
    Else
        Timer1.Enabled = True
        cmdPolling.Caption = "Polling is active"
        cmdPolling.BackColor = &HC0FFC0
        Label2.Caption = Label2.Caption & " Polling enabled."
        LastFileTimeStamp = FileDateTime(CurrentFilePath)
    End If
End Sub

Private Sub Form_Load()
  Picture1.AutoRedraw = True
End Sub

Private Sub Form_Activate()
  DrawCompleteGrid
  Form1.Caption = "Skew-T v" & App.Major & "." _
    & App.Minor & "." & App.Revision
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  lblReadout.Caption = "Hover inside diagram..."
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' The skewFactor should probably be a universal constant since
  ' it is used elsewhere. Fix in a future version.
  Dim mouseT As Single
  Dim mouseZ As Single
  Dim skewFactor As Single: skewFactor = 3
  
  mouseZ = Y
  mouseT = X - (mouseZ * skewFactor)
  
  strZ = Format(mouseZ, "0.0")
  strT = Format(mouseT, "0.0")
  
  strZ = Right(Space(4) & strZ, 4)
  strT = Right(Space(5) & strT, 5)
  
  lblReadout.Caption = "HGHT " & strZ & " km" & vbCrLf & "TEMP " & strT & " C"
End Sub

Public Sub StartMonitoring(strPath As String)
    CurrentFilePath = strPath
    Call PlotData(CurrentFilePath)
    
    LastFileTimeStamp = FileDateTime(CurrentFilePath)
    Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
    Static IsProcessing As Boolean
    If IsProcessing Then Exit Sub
    
    On Error GoTo CleanUp
    IsProcessing = True
    
    If Dir(CurrentFilePath) = "" Then
        Timer1.Enabled = False
        cmdPolling.Caption = "Start polling..."
        Exit Sub
    End If
    
    Dim CurrentTimestamp As Variant
    CurrentTimestamp = FileDateTime(CurrentFilePath)
    
    If CurrentTimestamp <> LastFileTimeStamp Then
        LastFileTimeStamp = CurrentTimestamp
        Call PlotData(CurrentFilePath)
    End If

CleanUp:
    IsProcessing = False

End Sub

