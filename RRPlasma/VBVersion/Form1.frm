VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000000&
   Caption         =   "Form1"
   ClientHeight    =   6345
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11145
   Icon            =   "FORM1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   423
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   743
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   5595
      Left            =   8010
      TabIndex        =   4
      Top             =   45
      Width           =   3105
      Begin VB.CommandButton cmdChangePalette 
         Caption         =   "Grey palette"
         Height          =   330
         Index           =   2
         Left            =   1710
         TabIndex        =   28
         Top             =   5160
         Width           =   1290
      End
      Begin VB.CommandButton cmdChangePalette 
         Caption         =   "Reverse palette"
         Height          =   330
         Index           =   1
         Left            =   1710
         TabIndex        =   27
         Top             =   4770
         Width           =   1290
      End
      Begin VB.CommandButton cmdChangePalette 
         Caption         =   "Sort palette"
         Height          =   330
         Index           =   0
         Left            =   1710
         TabIndex        =   24
         Top             =   4365
         Width           =   1290
      End
      Begin VB.ListBox List1 
         Height          =   3570
         Left            =   1680
         Sorted          =   -1  'True
         TabIndex        =   22
         Top             =   375
         Width           =   1290
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H80000000&
         Caption         =   " Effects"
         Height          =   1860
         Left            =   165
         TabIndex        =   16
         Top             =   2505
         Width           =   1425
         Begin VB.CommandButton cmdEffects 
            Caption         =   "Darken"
            Height          =   225
            Index           =   4
            Left            =   195
            TabIndex        =   26
            Top             =   1515
            Width           =   1035
         End
         Begin VB.CommandButton cmdEffects 
            Caption         =   "Brighten"
            Height          =   225
            Index           =   3
            Left            =   195
            TabIndex        =   25
            Top             =   1230
            Width           =   1035
         End
         Begin VB.CommandButton cmdEffects 
            Caption         =   " H-smooth"
            Height          =   225
            Index           =   0
            Left            =   195
            TabIndex        =   19
            Top             =   285
            Width           =   1035
         End
         Begin VB.CommandButton cmdEffects 
            Caption         =   "V-smooth"
            Height          =   225
            Index           =   1
            Left            =   195
            TabIndex        =   18
            Top             =   570
            Width           =   1035
         End
         Begin VB.CommandButton cmdEffects 
            Caption         =   "H&&V-smooth"
            Height          =   225
            Index           =   2
            Left            =   195
            TabIndex        =   17
            Top             =   855
            Width           =   1035
         End
      End
      Begin VB.CommandButton cmdGO 
         BackColor       =   &H00C0FFC0&
         Caption         =   "GO"
         Height          =   375
         Left            =   180
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   4515
         Width           =   1410
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save ""Plasma.bmp"""
         Height          =   465
         Left            =   210
         TabIndex        =   14
         Top             =   5025
         Width           =   1395
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H80000000&
         Height          =   2205
         Left            =   150
         TabIndex        =   5
         Top             =   150
         Width           =   1425
         Begin VB.HScrollBar HScroll1 
            Height          =   210
            Index           =   0
            LargeChange     =   8
            Left            =   180
            Max             =   128
            TabIndex        =   9
            Top             =   465
            Width           =   1110
         End
         Begin VB.HScrollBar HScroll1 
            Height          =   210
            Index           =   1
            LargeChange     =   16
            Left            =   195
            Max             =   128
            Min             =   4
            SmallChange     =   4
            TabIndex        =   8
            Top             =   1020
            Value           =   4
            Width           =   1110
         End
         Begin VB.CheckBox chkWrap 
            Caption         =   "Wrap X"
            Height          =   240
            Index           =   0
            Left            =   300
            TabIndex        =   7
            Top             =   1410
            Width           =   960
         End
         Begin VB.CheckBox chkWrap 
            Caption         =   "Wrap Y"
            Height          =   240
            Index           =   1
            Left            =   300
            TabIndex        =   6
            Top             =   1725
            Width           =   990
         End
         Begin VB.Label LabNum 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Index           =   0
            Left            =   1050
            TabIndex        =   13
            Top             =   210
            Width           =   315
         End
         Begin VB.Label Label1 
            BackColor       =   &H80000000&
            Caption         =   "Graininess = "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   12
            Top             =   225
            Width           =   810
         End
         Begin VB.Label LabNum 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Index           =   1
            Left            =   900
            TabIndex        =   11
            Top             =   735
            Width           =   315
         End
         Begin VB.Label Label1 
            BackColor       =   &H80000000&
            Caption         =   "Scale = "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   300
            TabIndex        =   10
            Top             =   750
            Width           =   675
         End
      End
      Begin VB.Label LabPalName 
         Caption         =   "PalName"
         Height          =   210
         Left            =   1680
         TabIndex        =   23
         Top             =   4050
         Width           =   1155
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000000&
         Caption         =   "Palettes"
         Height          =   240
         Left            =   1740
         TabIndex        =   20
         Top             =   150
         Width           =   690
      End
   End
   Begin VB.PictureBox PicHandle 
      BackColor       =   &H000000FF&
      Height          =   180
      Left            =   3675
      ScaleHeight     =   8
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   7
      TabIndex        =   2
      Top             =   3960
      Width           =   165
   End
   Begin VB.PictureBox picPal 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   90
      ScaleHeight     =   14
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   512
      TabIndex        =   1
      Top             =   6000
      Width           =   7740
   End
   Begin VB.PictureBox PIC 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3435
      Left            =   180
      ScaleHeight     =   229
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   257
      TabIndex        =   0
      Top             =   240
      Width           =   3855
   End
   Begin VB.Label LabRGB 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "RGB"
      Height          =   300
      Left            =   7980
      TabIndex        =   21
      Top             =   6000
      Width           =   1500
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label3"
      Height          =   285
      Left            =   9555
      TabIndex        =   3
      Top             =   6000
      Width           =   1590
   End
   Begin VB.Shape Shape1 
      Height          =   3675
      Left            =   120
      Top             =   180
      Width           =   3975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Form1.frm

' Plasmas Galore VB version:  by Robert Rayment

' 21/5/02 - 23/5/02

Option Base 1
DefLng A-W  ' All variables Long unless
DefSng X-Z  ' Singles or otherwise defined.

Dim sizex As Integer, sizey As Integer ' PIC size

' Plasma parameters
Dim StartNoise
Dim StartStepsize
Dim WrapX As Boolean
Dim WrapY As Boolean

' Flag if a plasma done
Dim GoDone As Boolean

' RGB components
Dim red As Byte
Dim green As Byte
Dim blue As Byte

Dim PaletteMaxIndex           ' Max Palette size
Dim IndexArray() As Integer   ' To hold palette indexes
Dim ColArray() As Long        ' To hold colors for dosplaying
Dim Colors()   ' The palette colors

Dim zInitX, zInitY   ' To resize PIC
Dim PathSpec$, PalDir$


Private Sub Form_Load()

Caption = " Plasmas Galore VB version: by Robert Rayment"

' Initial PIC size
sizex = 320
sizey = 256

With PIC
   .Width = sizex
   .Height = sizey
   .Top = 12
   .Left = 12
End With

' Show initial PIC sizes
Label3.Caption = "HxW = " & Str$(PIC.Height) & Str$(PIC.Width) ' -4 to avoid label change

' PIC border
With Shape1
   .Top = PIC.Top - 2
   .Left = PIC.Left - 2
   .Width = sizex + 4
   .Height = sizey + 4
End With

' Pic to show palette
picPal.Left = PIC.Left - 2
picPal.Width = 512 + 4

' PicHandle init coords
zInitX = 0
zInitY = 0
PicHandle.BackColor = vbGreen
'PicHandle position
PicHandle.Left = PIC.Left + PIC.Width + 1
PicHandle.Top = PIC.Top + PIC.Height - PicHandle.Height + 3

' Initial plasma parameters

StartNoise = 0
StartStepsize = 32

HScroll1(0).Value = 0
LabNum(0).Caption = "0"
HScroll1(1).Value = 32
LabNum(1).Caption = "32"

chkWrap(0).Value = Checked
chkWrap(1).Value = Checked
WrapX = True
WrapY = True

cmdGO.BackColor = RGB(192, 255, 192)   ' Light Green

Show

ReDim IndexArray(1 To sizex, 1 To sizey)
ReDim ColArray(1 To sizex, 1 To sizey)
PaletteMaxIndex = 511
ReDim Colors(0 To PaletteMaxIndex)


'Get app path
PathSpec$ = App.Path
If Right$(PathSpec$, 1) <> "\" Then PathSpec$ = PathSpec$ & "\"

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' NEED TO BACK UP TO GET PAL FILES
' If Pals folder not in parallel with App Folder
' or compiled exe not in App Folder will
' need to alter this code here
' ...\VBVersion\
' want ...\Pal\  ie back up 1 level
p = InStrRev(PathSpec$, "\", Len(PathSpec$) - 1)   ' NB VB6 only

' To pick up *.pal files in app folder - 1
'------------------
'Get FIRST entry
PalDir$ = Left$(PathSpec$, p) & "Pals\"
Spec$ = PalDir$ & "*.pal"
a$ = Dir$(Spec$)
'------------------

If a$ = "" Then
   MsgBox "Pal files not found"
   Unload Me
   End
End If

List1.AddItem UCase$(Left$(a$, 1)) & LCase$(Mid$(a$, 2))
Do
  a$ = Trim$(Dir)    'Gets NEXT entry
  If a$ <> "" Then
     List1.AddItem UCase$(Left$(a$, 1)) & LCase$(Mid$(a$, 2))
  Else   'a$="" indicates no more entries
     Exit Do
  End If
Loop

' Starting palette
PalSpec$ = PalDir$ & "Moon.pal"
LabPalName.Caption = "Moon.pal"
'PalSpec$ = PalDir$ & "Blues.pal"

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

ReadPAL PalSpec$

GoDone = False

End Sub

'###### GO GO GO #############################################

Private Sub cmdGO_Click()

GoDone = True

MousePointer = vbHourglass

cmdGO.BackColor = RGB(255, 192, 192)

DisableControls

Plasma

DisplayArray

EnableControls

cmdGO.BackColor = RGB(192, 255, 192)

MousePointer = vbDefault

End Sub

'###### PLASMA ################################################

Private Sub Plasma()

' CORE PLASMA ROUTINE

' Adapted & extended from a QB prog by Alan King 1996

ReDim IndexArray(1 To sizex, 1 To sizey)
ReDim ColArray(1 To sizex, 1 To sizey)

Randomize Timer

' PaletteMaxIndex = 511
' Graininess & scale also depends on palette
' Rapidly changing palettes will be grainier
' StartNoise = 0   ' graininess
' StartStepsize = 32  ' scale

Noise = StartNoise
Stepsize = StartStepsize

RndNoise = 0

Do
   
   For iy = 1 To sizey Step Stepsize
   For ix = 1 To sizex Step Stepsize
      
      If Noise > 0 Then RndNoise = (Rnd * (2 * Noise) - Noise) * Stepsize
       
      If iy + Stepsize >= sizey Then
         If WrapY Then
           Col = IndexArray(ix, 1) + RndNoise ' Gives vertical wrapping
         Else
           Col = IndexArray(ix, sizey) + RndNoise
         End If
      ElseIf ix + Stepsize >= sizex Then
         If WrapX Then
            Col = IndexArray(1, iy) + RndNoise  ' Gives horizontal wrapping
         Else
            Col = IndexArray(sizex, iy) + RndNoise
         End If
      ElseIf Stepsize = StartStepsize Then
         Col = Rnd * (2 * PaletteMaxIndex) - PaletteMaxIndex
      Else
          Col = IndexArray(ix, iy)
          Col = Col + IndexArray(ix + Stepsize, iy)
          Col = Col + IndexArray(ix, iy + Stepsize)
          Col = Col + IndexArray(ix + Stepsize, iy + Stepsize)
          Col = Col \ 4 + RndNoise
      End If
      
      If Stepsize > 1 Then 'BoxLine
          
          For ky = iy To iy + Stepsize
              If ky <= sizey Then
                  For kx = ix To ix + Stepsize
                      If kx <= sizex Then IndexArray(kx, ky) = Col
                  Next kx
              End If
          Next ky
      
      End If
      If Stepsize <= 1 Then IndexArray(ix, iy) = Col
   
   Next ix
   Next iy

Stepsize = Stepsize \ 2

Loop Until Stepsize <= 1


' Find Col max/min
smax = -10000
smin = 10000
For iy = sizey To 1 Step -1
For ix = sizex To 1 Step -1
   Col = IndexArray(ix, iy)
   If Col < smin Then smin = Col
   If Col > smax Then smax = Col
Next ix
Next iy

' Scale colors indices to range
' & put full RGB() palette into ColArray()
sdiv = (smax - smin)
If sdiv <= 0 Then sdiv = 1
zmul = PaletteMaxIndex / sdiv

For iy = sizey To 1 Step -1
For ix = sizex To 1 Step -1
   
   Index = (IndexArray(ix, iy) - smin) * zmul
   ' Check Index in change
   If Index < 0 Then
      Index = 0
   End If
   If Index > PaletteMaxIndex Then
      Index = PaletteMaxIndex
   End If
   ' Put in new Index
   IndexArray(ix, iy) = Index
   
   ColArray(ix, iy) = Colors(Index)

Next ix
Next iy

End Sub

'###### WRAP SWITCH ##########################################

Private Sub chkWrap_Click(Index As Integer)

Select Case Index
Case 0: WrapX = Not WrapX
Case 1: WrapY = Not WrapY
End Select

End Sub

'#####   SMOOTHING #########################################

Private Sub cmdEffects_Click(Index As Integer)

If GoDone = False Then Exit Sub

MousePointer = vbHourglass

cmdGO.BackColor = RGB(255, 192, 192)

DisableControls

Select Case Index
Case 0: SmoothX
Case 1: SmoothY
Case 2: SmoothX: SmoothY
Case 3: Brighten
Case 4: Darken
End Select

DisplayArray

EnableControls

cmdGO.BackColor = RGB(192, 255, 192)

MousePointer = vbDefault

End Sub

Private Sub SmoothX()  ' Two point wrapped smoothing in X direc
' ColArray(ix=1 to sizex, iy=1 to sizey)

For iy = 1 To sizey
For ix = 1 To sizex
    
    ixL = (ix - 1)   ' If ix = 1 ixL = 0
    If ix = 1 Then ixL = sizex
    ixR = (ix + 1)   ' If ix = sizex ixR = sizex+1
    If ix = sizex Then ixR = 1
    
    LNGtoRGB ColArray(ixL, iy)
    savred = red: savgreen = green: savblue = blue
    
    LNGtoRGB ColArray(ixR, iy)
    savred = savred + red: savgreen = savgreen + green: savblue = savblue + blue
    
    red = savred \ 2
    green = savgreen \ 2
    blue = savblue \ 2
    
    ColArray(ix, iy) = RGB(red, green, blue)

Next ix
Next iy

End Sub

Private Sub SmoothY()  ' Two point wrapped smoothing in Y direc
' ColArray(ix=1 to sizex, iy=1 to sizey)

For ix = 1 To sizex
For iy = 1 To sizey
    
    iyA = (iy - 1)      ' If iy = 1 iyA = 0
    If iy = 1 Then iyA = sizey
    iyB = (iy + 1)      ' If iy = sizey iyB = sizey+1
    If iy = sizey Then iyB = 1
    
    LNGtoRGB ColArray(ix, iyA)
    savred = red: savgreen = green: savblue = blue
    
    LNGtoRGB ColArray(ix, iyB)
    savred = savred + red: savgreen = savgreen + green: savblue = savblue + blue
    
    red = savred \ 2
    green = savgreen \ 2
    blue = savblue \ 2
    
    ColArray(ix, iy) = RGB(red, green, blue)

Next iy
Next ix

End Sub


Private Sub Brighten()
' ColArray(ix=1 to sizex, iy=1 to sizey)

Increment = 4

For ix = 1 To sizex
For iy = 1 To sizey
    
   LNGtoRGB ColArray(ix, iy)
   If red <= 255 - Increment Then red = red + Increment
   If green <= 255 - Increment Then green = green + Increment
   If blue <= 255 - Increment Then blue = blue + Increment
   
   ColArray(ix, iy) = RGB(red, green, blue)

Next iy
Next ix

End Sub

Private Sub Darken()
' ColArray(ix=1 to sizex, iy=1 to sizey)

Increment = 4

For ix = 1 To sizex
For iy = 1 To sizey
    
   LNGtoRGB ColArray(ix, iy)
   If red >= Increment Then red = red - Increment
   If green >= Increment Then green = green - Increment
   If blue >= Increment Then blue = blue - Increment
   
   ColArray(ix, iy) = RGB(red, green, blue)

Next iy
Next ix

End Sub


'######  DISPLAY ColArray(ix, iy) TO PIC #####################

Private Sub DisplayArray()

PIC.Cls

PicHandle.BackColor = vbRed

For ix = 1 To sizex
   For iy = 1 To sizey
         
      PIC.PSet (ix - 1, iy - 1), ColArray(ix, iy) 'Colors(Index)
   
   Next iy
Next ix

PicHandle.BackColor = vbGreen

End Sub


'##### PALETTE CHANGERS #################################

Private Sub cmdChangePalette_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

MousePointer = vbHourglass

cmdGO.BackColor = RGB(255, 192, 192)

DoEvents

Select Case Index
Case 0
   Quicksort Colors()
Case 1
   ReversePalette
Case 2
   GreyPalette
End Select

' Show palette
For N = 0 To 511
   picPal.Line (N, 0)-(N, 18), Colors(N)
Next N
picPal.Refresh

If GoDone Then

   For iy = 1 To sizey
   For ix = 1 To sizex
       ColArray(ix, iy) = Colors(IndexArray(ix, iy))
   Next ix
   Next iy

   DisplayArray

End If

cmdGO.BackColor = RGB(192, 255, 192)

MousePointer = vbDefault

End Sub

Sub Quicksort(lngArray() As Long)  ' Used  forSort palette
'1 dimensional long array sorted in ascending order from k to max
'All variables Long
Max = UBound(lngArray)
If Max = 1 Then Exit Sub
k = LBound(lngArray)
If k = Max Then Exit Sub
m = Max \ 2: ReDim sortl(m), sortr(m)
S = 1: sortl(1) = k: sortr(1) = Max
Do While S <> 0
      ll = sortl(S): mm = sortr(S): S = S - 1
   Do While ll < mm
      i = ll: j = mm
      p = (ll + mm) \ 2
      VX& = lngArray(p)
      Do While i <= j
         Do While lngArray(i) < VX&: i = i + 1: Loop
         Do While VX& < lngArray(j): j = j - 1: Loop
         If i <= j Then
          'SWAP zarr(i), zarr(j)
          VY& = lngArray(i): lngArray(i) = lngArray(j): lngArray(j) = VY&
          i = i + 1: j = j - 1
         End If
      Loop
      If i < mm Then
       S = S + 1: sortl(S) = i: sortr(S) = mm
      End If
      mm = j
   Loop
Loop
Erase sortl, sortr
End Sub

Private Sub ReversePalette()
   
   For i = 0 To 255
      Col = Colors(i)
      Colors(i) = Colors(511 - i)
      Colors(511 - i) = Col
   Next i

End Sub

Private Sub GreyPalette()

   For i = 0 To 511
      
    LNGtoRGB Colors(i)
    
    'Col = (1& * red + 1& * green + 1& * blue) \ 3
    Col = (1& * red * 0.3 + green * 0.6 + blue * 0.1)
    Col = Col And 255
    Colors(i) = RGB(Col, Col, Col)
   
   Next i

End Sub


'##### PALETTE INPUT #################################

Public Sub ReadPAL(PalSpec$)
' Read JASC-PAL palette file
' Any error shown by PalSpec$ = ""
' Else RGB into Colors(i) Long

'Dim red As Byte, green As Byte, blue As Byte
On Error GoTo palerror
Open PalSpec$ For Input As #1
Line Input #1, a$
p = InStr(1, a$, "JASC")
If p = 0 Then PalSpec$ = "": Exit Sub
   
   'JASC-PAL
   '0100
   '256
   Line Input #1, a$
   Line Input #1, a$

   For N = 0 To 255
      If EOF(1) Then Exit For
      Line Input #1, a$
      ParsePAL a$ ', red, green, blue
      Colors(N) = RGB(red, green, blue)
   Next N
   Close #1
   
   ' Extend palette to 512 colors
   ReDim Preserve Colors(0 To 511)
   
   For N = 255 To 1 Step -1
      Colors(2 * N - 1) = Colors(N)
   Next N
   Colors(510) = Colors(509)
   Colors(511) = Colors(510)

   For N = 2 To 510 Step 2
      LNGtoRGB Colors(N - 1)
      savred = red: savgreen = green: savblue = blue
      R1 = RGB(red, green, blue)
      LNGtoRGB Colors(N + 1)
      R2 = RGB(red, green, blue)
      savr = (savred + red) \ 2
      savg = (savgreen + green) \ 2
      savb = (savblue + blue) \ 2
      Colors(N) = RGB(savr, savg, savb)
   Next N
   Colors(511) = Colors(510)
   
   ' Show palette
   For N = 0 To 511
      picPal.Line (N, 0)-(N, 18), Colors(N)
   Next N
   picPal.Refresh

Exit Sub
'===========
palerror:
PalSpec$ = ""
   
   MsgBox "Palette file error or not there"
   'Resume Next
   Unload Me
   End

Exit Sub
End Sub

Public Sub ParsePAL(ain$) ', red As Byte, green As Byte, blue As Byte)
'Input string ain$, with 3 numbers(R G B) with
'space separators and then any text
ain$ = LTrim(ain$)
lena = Len(ain$)
R$ = ""
g$ = ""
b$ = ""
num = 0 'R
nt = 0
For i = 1 To lena
   c$ = Mid$(ain$, i, 1)
   
   If c$ <> " " Then
      If nt = 0 Then num = num + 1
      nt = 1
      If num = 4 Then Exit For
      If Asc(c$) < 48 Or Asc(c$) > 57 Then Exit For
      If num = 1 Then R$ = R$ + c$
      If num = 2 Then g$ = g$ + c$
      If num = 3 Then b$ = b$ + c$
   Else
      nt = 0
   End If
Next i
red = Val(R$): green = Val(g$): blue = Val(b$)
End Sub


Private Sub List1_Click()

PalName$ = List1.List(List1.ListIndex)
LabPalName.Caption = PalName$
PalSpec$ = PalDir$ & PalName$

ReadPAL PalSpec$

MousePointer = vbHourglass

cmdGO.BackColor = RGB(255, 192, 192)

DoEvents

For iy = 1 To sizey
For ix = 1 To sizex
    ColArray(ix, iy) = Colors(IndexArray(ix, iy))
Next ix
Next iy

If GoDone Then DisplayArray

cmdGO.BackColor = RGB(192, 255, 192)

MousePointer = vbDefault

End Sub

Private Sub picPal_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

LNGtoRGB picPal.Point(X, Y)
LabRGB.Caption = "RGB " & Str$(red) & Str$(green) & Str$(blue)

End Sub

'######  CONTROLS SWITCH ########################

Private Sub DisableControls()
Frame2.Enabled = False
Frame1.Enabled = False
List1.Enabled = False
PicHandle.Enabled = False
DoEvents
End Sub
Private Sub EnableControls()
Frame2.Enabled = True
Frame1.Enabled = True
List1.Enabled = True
PicHandle.Enabled = True
DoEvents
End Sub

'###### CONVERT LOng Color to RGB components ################

Private Sub LNGtoRGB(ByVal LongCul As Long)
'Dim red As Byte
'Dim green As Byte
'Dim blue As Byte
    red = LongCul And &HFF
    green = (LongCul \ &H100) And &HFF
    blue = (LongCul \ &H10000) And &HFF
    R = RGB(red, green, blue)
    
End Sub

'###### CHANGE StartNoise (Graininess) & StartStepsize (Scale) ##############

Private Sub HScroll1_Change(Index As Integer)

Select Case Index
Case 0
   StartNoise = HScroll1(0).Value
   LabNum(0).Caption = Str$(StartNoise)
Case 1
   StartStepsize = HScroll1(1).Value
   LabNum(1).Caption = Str$(StartStepsize)
End Select

End Sub

'######  Show PIC RGB Color ########################################

Private Sub PIC_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

LNGtoRGB PIC.Point(X, Y)
LabRGB.Caption = "RGB " & Str$(red) & Str$(green) & Str$(blue)

End Sub

'#######  RESIZE PIC ##############################################

Private Sub PicHandle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
zInitX = X
zInitY = Y
End Sub

Private Sub PicHandle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo LabHError


PicHandle.MousePointer = vbSizeAll


If Button = 1 Then
   
   ' Test PicHandle's new position
   
   LabHLeft = PicHandle.Left + (X - zInitX)
   LabHTop = PicHandle.Top + (Y - zInitY)
   
   PICWidth = (LabHLeft - PIC.Left) - 1
   PICHeight = (LabHTop - PIC.Top) + PicHandle.Height - 2
   
   ' Limit lower & upper sizes
   If PICWidth < 128 Or PICWidth > 512 + 4 Then Exit Sub
   If PICHeight < 128 Or PICHeight > 384 + 4 Then Exit Sub
   
   ' Re-position PicHandle
   PicHandle.Left = LabHLeft
   PicHandle.Top = LabHTop
   
   ' Resize PIC after using PicHandle
   sizex = (PicHandle.Left - PIC.Left) - 1
   sizey = (PicHandle.Top - PIC.Top) + PicHandle.Height - 2

   ' Force multiples of 8
   remx = sizex Mod 8
   If remx <> 0 Then sizex = sizex - remx
   remy = sizey Mod 8
   If remy <> 0 Then sizey = sizey - remy
   ' and resize again
   PIC.Width = sizex
   PIC.Height = sizey

   'Re-position PicHandle to new PIC size
   PicHandle.Left = PIC.Left + PIC.Width + 1
   PicHandle.Top = PIC.Top + PIC.Height - PicHandle.Height + 3
   
   'GoDone = False

   ReDim IndexArray(1 To sizex, 1 To sizey) As Integer
   ReDim ColArray(1 To sizex, 1 To sizey)

   With Shape1
      .Top = PIC.Top - 2
      .Left = PIC.Left - 2
      .Width = sizex + 4
      .Height = sizey + 4
   End With

   Label3.Caption = "HxW = " & Str$(PIC.Height) & Str$(PIC.Width) ' -4 to avoid label change

   DoEvents
   
End If

Exit Sub
'========
LabHError:
    PicHandle.Top = PIC.Top + PIC.Height
    PicHandle.Left = PIC.Left + PIC.Width
End Sub

Private Sub PicHandle_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

   PicHandle.MousePointer = vbDefault

   'cmdGO_Click

End Sub


'###### SAVE BMP to "Plasma.BMP" #################################

Private Sub cmdSave_Click()

SavePicture PIC.Image, "Plasma.bmp"

End Sub

