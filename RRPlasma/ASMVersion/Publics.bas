Attribute VB_Name = "Module1"
'Module1: Publics.bas

' Plasma Worlds by Trebor Tnemyar

' June 2002

' Mainly to hold Publics

Option Base 1
DefLng A-W
DefSng X-Z

' As suggested by John Galanopoulos
' Used in Sub Tunnel
Public Const QS_KEY = &H1
Public Const QS_PAINT = &H20
Public Const QS_MOUSEBUTTON = &H4
Public Declare Function GetQueueStatus Lib "user32" _
(ByVal qsFlags As Long) As Long

'--------------------------------------------------------------------------
' --------------------------------------------------------------
' Windows API - For timing (not used yet)

Public Declare Function timeGetTime& Lib "winmm.dll" ()

'------------------------------------------------------------------------------

' Copy one array to another of same number of bytes
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
(Destination As Any, Source As Any, ByVal Length As Long)

'To fill BITMAP structure
Public Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" _
(ByVal hObject As Long, ByVal Lenbmp As Long, Publicbmp As Any) As Long

Public Type BITMAP
   bmType As Long              ' Type of bitmap
   bmWidth As Long             ' Pixel width
   bmHeight As Long            ' Pixel height
   bmWidthBytes As Long        ' Byte width = 3 x Pixel width
   bmPlanes As Integer         ' Color depth of bitmap
   bmBitsPixel As Integer      ' Bits per pixel, must be 16 or 24
   bmBits As Long              ' This is the pointer to the bitmap data  !!!
End Type

'NB PICTURE STORED IN MEMORY UPSIDE DOWN
'WITH INCREASING MEMORY GOING UP THE PICTURE
'bmp.bmBits points to the bottom left of the picture

Public bmp As BITMAP
'------------------------------------------------------------------------------

' Structures for StretchDIBits
Public Type BITMAPINFOHEADER ' 40 bytes
   biSize As Long
   biwidth As Long
   biheight As Long
   biPlanes As Integer
   biBitCount As Integer
   biCompression As Long
   biSizeImage As Long
   biXPelsPerMeter As Long
   biYPelsPerMeter As Long
   biClrUsed As Long
   biClrImportant As Long
End Type

Public Type BITMAPINFO
   bmiH As BITMAPINFOHEADER
   'bmiH As RGBTRIPLE         'NB Palette NOT NEEDED for 16,24 & 32-bit
End Type
Public bm As BITMAPINFO

' For transferring drawing in an integer array to Form or PicBox
Public Declare Function StretchDIBits Lib "gdi32" (ByVal hdc As Long, _
ByVal DesXOffset As Long, ByVal DesYOffset As Long, _
ByVal DesWidth As Long, ByVal DesHeight As Long, _
ByVal SrcXOffset As Long, ByVal SrcYOffset As Long, _
ByVal SrcWidth As Long, ByVal SrcHeight As Long, _
lpBits As Any, lpBitsInfo As BITMAPINFO, _
ByVal wUsage As Long, ByVal dwRop As Long) As Long
'------------------------------------------------------------------------------

' For calling machine code
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" _
(ByVal lpMCode As Long, _
ByVal Long1 As Long, ByVal Long2 As Long, _
ByVal Long3 As Long, ByVal Long4 As Long) As Long

Public ptrMC, ptrStruc    ' Ptrs to Machine Code & Structure

'-----------------------------------------------------------------
'Plasma MCode Structure
Public Type PlasmaStruc
   sizex As Long           ' 128 - 512
   sizey As Long           ' 128 - 384
   PtrIndexArray As Long
   PtrColArray As Long
   RandSeed As Long        ' (0-.99999) * 511
   Noise As Long           ' 0 - 128
   Stepsize As Long        ' 0 - 128
   WrapX As Long           ' -1 True, 0 False
   WrapY As Long           ' -1 True, 0 False
   PaletteMaxIndex As Long ' 511
   PtrColors As Long
End Type
Public PlasmaMCODE As PlasmaStruc
Public PlasmaArray() As Byte  ' Array to hold machine code
'res = CallWindowProc(ptrMC, ptrStruc, 0&, 0&, 0&)
' Plasma parameters
Public StartNoise
Public StartStepsize
Public WrapX As Boolean
Public WrapY As Boolean
Public RandSeed
'-----------------------------------------------------------------

'Effects MCode Structure
Public Type EffectsStruc
   sizex As Long           ' 128 - 512
   sizey As Long           ' 128 - 384
   PtrColArray As Long
End Type
'   Opcode As Long  ' 0 x, 1 y, 2 x & y - smoothing
                    ' 3 brighten, 4 darken (Increment=4)
Public EffectsMCODE As EffectsStruc
Public EffectsArray() As Byte  ' Array to hold machine code
'res = CallWindowProc(ptrMC, ptrStruc, 0pcode, 0&, 0&)

'-----------------------------------------------------------------------
'Scrolling MCode Structure
Public Type ScrollStruc
   sizex As Long           ' 128 - 512
   sizey As Long           ' 128 - 384
   PtrColArray As Long
   PtrLineStore As Long
   HScrollValue As Long    ' -4 to +4
   VScrollValue As Long    ' -4 to +4
End Type
Public ScrollMCODE As ScrollStruc
Public ScrollArray() As Byte  ' Array to hold machine code
'res = CallWindowProc(ptrMC, ptrStruc, 0, 0&, 0&)
' Scrolling values
Public HScrollValue, VScrollValue
Public ScrollingDone As Boolean
Public Linestore() As Long

'-----------------------------------------------------------------------
'Color cycling MCode Structure
Public Type CycleStruc
   sizex As Long           ' 128 - 512
   sizey As Long           ' 128 - 384
   PtrColArray As Long
   CycleValue As Long      ' -4 to +4
End Type
Public CycleMCODE As CycleStruc
Public CycleArray() As Byte  ' Array to hold machine code
'res = CallWindowProc(ptrMC, ptrStruc, 0, 0&, 0&)
Public CyclingDone As Boolean
' Color cycling value
Public CycleValue
'-----------------------------------------------------------------------

' Transform Mcode Structure
Public Type TransformStruc
   sizex As Long           ' 128 - 512
   sizey As Long           ' 128 - 384
   PtrColArray As Long
   dimx As Long
   dimy As Long
   PtrSmallArray As Long
   Ptrixdent As Long
   Ptrzrdsq As Long        ' Single data
   ixa As Long
   ixsc As Long
   zScale As Single
   zWScale As Single
   zHscale As Single
   ShadeNumber As Long
   PtrzSinTab As Long      ' Single data Sin(0-181 deg)
End Type

Public TransformMCODE As TransformStruc
Public TransformArray() As Byte  ' Array to hold machine code
' Transform parameters
Public zSinTab()
Public TransformType    ' 0 Cylinder, 1 Spheroid, 2 Landscape
Public SmallArray()
Public ixdent()
Public zrdsq()
Public RotateDone As Boolean
Public Speed            '+/- 10
Public zWScale, zHscale, zScale
Public dimx, dimy
Public ixa, ixsc
Public TransformSize
Public ShadeNumber   ' 0 none, 1 half, 2 center

'-----------------------------------------------------------------------
' Tunnel

Public Type TunnelStructure
   sizex As Long           ' 128 - 512
   sizey As Long           ' 128 - 384
   PtrCoordsX As Long      ' (1-SqSize)Integer array
   PtrCoordsY As Long      ' (1-SqSize)Integer array
   PtrColArray As Long     ' (1-SqSize,1-SqSize)Long array
   PtrColors As Long       ' (1-511)Long array
   PtrIndexArray As Long   ' (1-SqSize,1-SqSize)Integer array
   kystart As Long         ' for advancing tunnel
   kxstart As Long         ' for rotating tunnel
   zTrimmer As Single      ' .005
   NumSmooths As Long
End Type
Public TunnelStruc As TunnelStructure
Public TunnelArray() As Byte  ' Array to hold machine code
'ptrMC = VarPtr(TunnelArray(1))
'ptrStruc = VarPtr(TunnelStruc.sizex)
'res = CallWindowProc(ptrMC, ptrStruc, CylinProj, 0&, 0&) ;0
'res = CallWindowProc(ptrMC, ptrStruc, LUTProj, 0&, 0&) ;2
Public Const CylinProj = 0
Public Const LUTProj = 1
Public Const SmoothTunnel = 2

Public CoordsX() As Integer
Public CoordsY() As Integer
Public NumSmooths
Public kystart
Public kxstart
Public zTrimmer
Public TunnelSpeed
Public Spin As Boolean

Public TunnelDone As Boolean

'-----------------------------------------------------------------------
' Landscape Mcode Structure
Public Type LandscapeStruc
   sizex As Long           ' 128 - 512
   sizey As Long           ' 128 - 384
   PtrColArray As Long
   dimx As Long
   dimy As Long
   PtrSmallArray As Long
   PtrIndexArray As Long
   LAH As Long
   TANG As Long
   xeye As Single
   yeye As Single
   zvp As Single
   zdslope As Single
   zslopemult As Single
   zhtmult  As Single
   DISH As Long
   Reflect As Long
End Type

Public LandscapeMCODE As LandscapeStruc
Public LandscapeArray() As Byte  ' Array to hold machine code

' Additional Landscape variables
Public LandscapeDone As Boolean
Public VSBF, HSANG, VSHT   ' Values for for/back move, Angle & Height
Public DISH As Boolean
Public Reflect
Public ZeroArray()         ' Landscape background
Public SmoothXCount, SmoothYCount, SmoothXYCount
Public BrightenCount, DarkenCount
Public zdslope ' Starting slope of ray for each column scanned.  Larger
               ' values flatten & smaller values exaggerate heights.
Public zhtmult     ' Voxel ht multiplier.
Public zslopemult  ' Slope multiplier.  Larger -ve values look
                   ' more vertically down.
Public zvp         ' Starting height of viewer.  Higher values lift
                   ' the viewer & give a more pointed perspective.

Public LAH         ' Look ahead distance - the larger it is the more
                   ' of the FloorSurf is scanned and the slower
''
Public TANG
Public xeye, yeye, zEL
Public FirstIn
Public SA      ' flag Speed (1), Angle (2) scrollbars

'-----------------------------------------------------------------------


Public sizex, sizey  ' PIC size

' Flag if a plasma done
Public PlasmaDone As Boolean

' RGB components
Public red As Byte, green As Byte, blue As Byte

Public PaletteMaxIndex           ' Max Palette size = 511
Public IndexArray() As Integer   ' To hold palette indexes
Public ColArray() As Long        ' To hold colors for displaying
Public Colors() As Long          ' The palette colors
Public zInitX, zInitY            ' To resize PIC
Public PathSpec$, PalDir$        ' App & Pal paths
Public MapDone As Boolean

'---------------------------------

Public Const pi# = 3.1415926535898
Public Const d2r# = pi# / 180
Public Const r2d# = 180 / pi#


Public Sub Loadmcode(InFile$, MCCode() As Byte)
'Load machine code into MCCode() byte array
On Error GoTo InFileErr
If Dir$(InFile$) = "" Then
   MsgBox (InFile$ & " missing")
   DoEvents
   Unload Form1
   End
End If
Open InFile$ For Binary As #1
MCSize& = LOF(1)
If MCSize& = 0 Then
InFileErr:
   MsgBox (InFile$ & " missing")
   DoEvents
   Unload Form1
   End
End If
ReDim MCCode(MCSize&)
Get #1, , MCCode
Close #1
On Error GoTo 0
End Sub

Public Sub FillBMPStruc(bwidth, bheight)
 
 With bm.bmiH
   .biSize = 40
   .biwidth = bwidth
   .biheight = bheight
   .biPlanes = 1
   .biBitCount = 32           ' Sets up BGRA pixels
   .biCompression = 0
   BytesPerScanLine = (.biwidth * .biBitCount * 4)
   .biSizeImage = BytesPerScanLine * Abs(.biheight)
   
   .biXPelsPerMeter = 0
   .biYPelsPerMeter = 0
   .biClrUsed = 0
   .biClrImportant = 0
 End With

End Sub


'##### PALETTE INPUT #####################################

Public Sub ReadPAL(PalSpec$)
' Read JASC-PAL palette file
' Any error shown by PalSpec$ = ""
' Else RGB into Colors(i) Long

'Dim red As Byte, green As Byte, blue As Byte
On Error GoTo palerror
Open PalSpec$ For Input As #1
Line Input #1, A$
p = InStr(1, A$, "JASC")
If p = 0 Then PalSpec$ = "": Exit Sub
   
   'JASC-PAL
   '0100
   '256
   Line Input #1, A$
   Line Input #1, A$

   For N = 0 To 255
      If EOF(1) Then Exit For
      Line Input #1, A$
      ParsePAL A$ ', red, green, blue
      'Colors(N) = RGB(red, green, blue)
      Colors(N) = RGB(blue, green, red) ' Reverse R & B
   Next N
   Close #1
   
   ' Extend palette to 512 colors
   
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
Exit Sub
'===========
palerror:
PalSpec$ = ""
   
   MsgBox "Palette file error or not there"
   Err.Clear

End Sub

Public Sub ParsePAL(ain$) ', red As Byte, green As Byte, blue As Byte)
'Input string ain$, with 3 numbers(R G B) with
'space separators and then any text
On Error GoTo parserror
ain$ = LTrim(ain$)
lena = Len(ain$)
R$ = ""
G$ = ""
B$ = ""
num = 0 'R
nt = 0
For i = 1 To lena
   C$ = Mid$(ain$, i, 1)
   
   If C$ <> " " Then
      If nt = 0 Then num = num + 1
      nt = 1
      If num = 4 Then Exit For
      If Asc(C$) < 48 Or Asc(C$) > 57 Then Exit For
      If num = 1 Then R$ = R$ + C$
      If num = 2 Then G$ = G$ + C$
      If num = 3 Then B$ = B$ + C$
   Else
      nt = 0
   End If
Next i
red = Val(R$): green = Val(G$): blue = Val(B$)
Exit Sub
'=====
parserror:
Resume Next
End Sub

'###### CONVERT Long Color to RGB components ################

Public Sub LNGtoRGB(ByVal LongCul As Long)
' Public red As Byte, green As Byte, blue As Byte

red = LongCul And &HFF
green = (LongCul \ &H100) And &HFF
blue = (LongCul \ &H10000) And &HFF
'R = RGB(red, green, blue)
    
End Sub

Public Function zATan2(ByVal zy As Single, ByVal zx As Single)
' Find angle Atan from -pi#/2 to +pi#/2
' Public pi#
If zx <> 0 Then
   zATan2 = Atn(zy / zx)
   If (zx < 0) Then
      If (zy < 0) Then zATan2 = zATan2 - pi# Else zATan2 = zATan2 + pi#
   End If
Else  ' zx=0
   If Abs(zy) > Abs(zx) Then   'Must be an overflow
      If zy > 0 Then zATan2 = pi# / 2 Else zATan2 = -pi# / 2
   Else
      zATan2 = 0   'Must be an underflow
   End If
End If
End Function

Public Sub Quicksort(lngArray() As Long)
'1 dimensional long array sorted in ascending order from k to max
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
      VX = lngArray(p)
      Do While i <= j
         Do While lngArray(i) < VX: i = i + 1: Loop
         Do While VX < lngArray(j): j = j - 1: Loop
         If i <= j Then
          'SWAP lngArray(i), lngArray(j)
          VY = lngArray(i): lngArray(i) = lngArray(j): lngArray(j) = VY
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

