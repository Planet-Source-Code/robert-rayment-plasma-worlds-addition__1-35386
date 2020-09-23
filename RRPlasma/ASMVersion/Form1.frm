VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000000&
   Caption         =   "Form1"
   ClientHeight    =   7815
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11145
   Icon            =   "FORM1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   521
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   743
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox PIC2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000007&
      FillStyle       =   0  'Solid
      Height          =   1440
      Left            =   5400
      ScaleHeight     =   92
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   124
      TabIndex        =   78
      Top             =   4095
      Width           =   1920
   End
   Begin VB.Frame frmPlusMinus 
      Height          =   615
      Left            =   10365
      TabIndex        =   68
      ToolTipText     =   "Resizer"
      Top             =   5640
      Width           =   690
      Begin VB.CommandButton cmdResize 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   270
         TabIndex        =   72
         Top             =   375
         Width           =   165
      End
      Begin VB.CommandButton cmdResize 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   270
         TabIndex        =   71
         Top             =   150
         Width           =   165
      End
      Begin VB.CommandButton cmdResize 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   465
         TabIndex        =   70
         Top             =   255
         Width           =   165
      End
      Begin VB.CommandButton cmdResize 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   75
         TabIndex        =   69
         Top             =   255
         Width           =   165
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "?"
      Height          =   225
      Left            =   240
      TabIndex        =   66
      Top             =   7470
      Width           =   240
   End
   Begin VB.Frame frmLandscape 
      Caption         =   "Landscape"
      Height          =   1350
      Left            =   7515
      TabIndex        =   41
      Top             =   6345
      Width           =   3555
      Begin VB.CheckBox chkReflected 
         Caption         =   "Reflect"
         Height          =   210
         Left            =   135
         TabIndex        =   67
         Top             =   975
         Width           =   825
      End
      Begin VB.CommandButton cmdLandscape 
         Caption         =   "DISHED"
         Height          =   270
         Index           =   1
         Left            =   150
         TabIndex        =   65
         Top             =   570
         Width           =   765
      End
      Begin VB.Frame frmLControls 
         Height          =   1125
         Left            =   1005
         TabIndex        =   54
         Top             =   105
         Width           =   2430
         Begin VB.CommandButton cmdMap 
            Caption         =   " Map"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1890
            TabIndex        =   79
            Top             =   720
            Width           =   435
         End
         Begin VB.TextBox txtHT 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   240
            Left            =   1425
            TabIndex        =   62
            Text            =   "1"
            Top             =   465
            Width           =   285
         End
         Begin VB.CheckBox chkLandscapeStop 
            Caption         =   " S"
            BeginProperty Font 
               Name            =   "Wingdings"
               Size            =   8.25
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   1965
            Style           =   1  'Graphical
            TabIndex        =   60
            Top             =   375
            Width           =   255
         End
         Begin VB.VScrollBar VSHeight 
            Height          =   765
            Left            =   1425
            Max             =   32000
            Min             =   -32000
            TabIndex        =   59
            Top             =   195
            Width           =   315
         End
         Begin VB.TextBox txtTANG 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   255
            Left            =   720
            TabIndex        =   58
            Text            =   "0"
            Top             =   480
            Width           =   390
         End
         Begin VB.HScrollBar HSTANG 
            Height          =   270
            Left            =   450
            Max             =   32000
            Min             =   -32000
            TabIndex        =   57
            Top             =   480
            Width           =   915
         End
         Begin VB.TextBox txtForwardBackward 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   255
            Left            =   90
            TabIndex        =   56
            Text            =   "0"
            Top             =   465
            Width           =   300
         End
         Begin VB.VScrollBar VSForwardBackward 
            Height          =   780
            Left            =   90
            Max             =   32000
            Min             =   -3200
            TabIndex        =   55
            Top             =   195
            Width           =   315
         End
         Begin VB.Label Label8 
            Caption         =   "Height"
            Height          =   180
            Left            =   1770
            TabIndex        =   64
            Top             =   120
            Width           =   525
         End
         Begin VB.Label Label7 
            Caption         =   "Speed"
            Height          =   180
            Left            =   435
            TabIndex        =   63
            Top             =   120
            Width           =   705
         End
         Begin VB.Label Label4 
            Caption         =   "Angle"
            Height          =   225
            Left            =   720
            TabIndex        =   61
            Top             =   765
            Width           =   465
         End
      End
      Begin VB.CommandButton cmdLandscape 
         Caption         =   "FLAT"
         Height          =   270
         Index           =   0
         Left            =   150
         TabIndex        =   42
         Top             =   255
         Width           =   765
      End
   End
   Begin VB.Frame frmTransform 
      Caption         =   "Transform"
      Height          =   1350
      Left            =   2520
      TabIndex        =   39
      Top             =   6345
      Width           =   3480
      Begin VB.CommandButton cmdTransform 
         Caption         =   "Hyperboloid"
         Height          =   285
         Index           =   2
         Left            =   195
         TabIndex        =   76
         Top             =   915
         Width           =   975
      End
      Begin VB.CommandButton cmdTransform 
         Caption         =   "Ellipsoid"
         Height          =   270
         Index           =   1
         Left            =   195
         TabIndex        =   75
         Top             =   585
         Width           =   975
      End
      Begin VB.Frame frmTControls 
         Height          =   1035
         Left            =   1200
         TabIndex        =   43
         Top             =   150
         Width           =   2145
         Begin VB.CheckBox chkRotationStop 
            Caption         =   " S"
            BeginProperty Font 
               Name            =   "Wingdings"
               Size            =   8.25
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   1755
            Style           =   1  'Graphical
            TabIndex        =   53
            Top             =   645
            Width           =   255
         End
         Begin VB.TextBox txtShade 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   255
            Left            =   690
            TabIndex        =   50
            Text            =   "0"
            Top             =   690
            Width           =   195
         End
         Begin VB.HScrollBar SBShade 
            Height          =   270
            Left            =   420
            Max             =   4
            TabIndex        =   49
            Top             =   690
            Width           =   735
         End
         Begin VB.OptionButton optRotationStart 
            Caption         =   "S"
            BeginProperty Font 
               Name            =   "Wingdings"
               Size            =   8.25
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080FF80&
            Height          =   240
            Left            =   1755
            Style           =   1  'Graphical
            TabIndex        =   48
            Top             =   285
            Width           =   270
         End
         Begin VB.TextBox txtSpeed 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   255
            Left            =   675
            TabIndex        =   47
            Text            =   "0"
            Top             =   165
            Width           =   240
         End
         Begin VB.HScrollBar SBSpeed 
            Height          =   270
            Left            =   420
            Max             =   10
            Min             =   -10
            TabIndex        =   46
            Top             =   165
            Width           =   750
         End
         Begin VB.TextBox txtVSB 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   285
            Left            =   120
            TabIndex        =   45
            Text            =   "10"
            Top             =   405
            Width           =   270
         End
         Begin VB.VScrollBar VSBTransform 
            Height          =   795
            LargeChange     =   2
            Left            =   120
            Max             =   40
            Min             =   2
            TabIndex        =   44
            Top             =   165
            Value           =   10
            Width           =   270
         End
         Begin VB.Label Label9 
            Caption         =   "Distance"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   135
            Left            =   435
            TabIndex        =   77
            Top             =   480
            Width           =   600
         End
         Begin VB.Label Label6 
            Caption         =   "Shade"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Left            =   1170
            TabIndex        =   52
            Top             =   795
            Width           =   480
         End
         Begin VB.Label Label5 
            Caption         =   "Speed"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Left            =   1185
            TabIndex        =   51
            Top             =   165
            Width           =   480
         End
      End
      Begin VB.CommandButton cmdTransform 
         Caption         =   "Cylinder"
         Height          =   270
         Index           =   0
         Left            =   195
         TabIndex        =   40
         Top             =   255
         Width           =   960
      End
   End
   Begin VB.Frame frmCycle 
      Caption         =   "Cycle"
      Height          =   1050
      Left            =   1560
      TabIndex        =   34
      Top             =   6345
      Width           =   915
      Begin VB.OptionButton optCycleStop 
         Caption         =   "S"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   330
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   780
         Width           =   285
      End
      Begin VB.OptionButton optCycleStart 
         Caption         =   "S"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   225
         Left            =   315
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   240
         Width           =   255
      End
      Begin VB.TextBox txtCycle 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   255
         Left            =   315
         TabIndex        =   36
         Text            =   "0"
         Top             =   495
         Width           =   240
      End
      Begin VB.HScrollBar SBCycle 
         Height          =   255
         Left            =   75
         Max             =   4
         Min             =   -4
         TabIndex        =   35
         Top             =   495
         Width           =   750
      End
   End
   Begin VB.Frame frmScrolling 
      Caption         =   "Scrolling"
      Height          =   1050
      Left            =   255
      TabIndex        =   27
      Top             =   6345
      Width           =   1260
      Begin VB.OptionButton optScrollStop 
         Caption         =   "S"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   900
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   765
         Width           =   270
      End
      Begin VB.OptionButton optScrollStart 
         Caption         =   "S"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   225
         Left            =   900
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   225
         Width           =   255
      End
      Begin VB.TextBox txtVScroll 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   255
         Left            =   105
         TabIndex        =   31
         Text            =   "3"
         Top             =   450
         Width           =   255
      End
      Begin VB.TextBox txtHScroll 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   270
         Left            =   660
         TabIndex        =   30
         Text            =   "1"
         Top             =   465
         Width           =   225
      End
      Begin VB.HScrollBar HScroll2 
         Height          =   270
         Left            =   405
         Max             =   4
         Min             =   -4
         TabIndex        =   28
         Top             =   465
         Width           =   750
      End
      Begin VB.VScrollBar VScroll2 
         Height          =   735
         Left            =   105
         Max             =   4
         Min             =   -4
         TabIndex        =   29
         Top             =   225
         Value           =   4
         Width           =   270
      End
   End
   Begin VB.Frame frmDesign 
      Caption         =   "Design"
      Height          =   5595
      Left            =   8025
      TabIndex        =   4
      Top             =   45
      Width           =   3105
      Begin VB.CommandButton cmdChangePalette 
         Caption         =   "Grey palette"
         Height          =   330
         Index           =   2
         Left            =   1710
         TabIndex        =   74
         Top             =   5160
         Width           =   1290
      End
      Begin VB.CommandButton cmdChangePalette 
         Caption         =   "Reverse palette"
         Height          =   330
         Index           =   1
         Left            =   1710
         TabIndex        =   73
         Top             =   4725
         Width           =   1290
      End
      Begin VB.CommandButton cmdChangePalette 
         Caption         =   "Sort palette"
         Height          =   330
         Index           =   0
         Left            =   1695
         TabIndex        =   24
         Top             =   4305
         Width           =   1290
      End
      Begin VB.ListBox List1 
         Height          =   3570
         ItemData        =   "FORM1.frx":0442
         Left            =   1680
         List            =   "FORM1.frx":0444
         Sorted          =   -1  'True
         TabIndex        =   22
         Top             =   375
         Width           =   1290
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H80000000&
         Caption         =   "  Effects"
         Height          =   1875
         Left            =   180
         TabIndex        =   16
         Top             =   2460
         Width           =   1425
         Begin VB.CommandButton cmdEffects 
            Caption         =   " Darken"
            Height          =   225
            Index           =   4
            Left            =   195
            TabIndex        =   26
            Top             =   1545
            Width           =   1035
         End
         Begin VB.CommandButton cmdEffects 
            Caption         =   " Brighten"
            Height          =   225
            Index           =   3
            Left            =   195
            TabIndex        =   25
            Top             =   1245
            Width           =   1035
         End
         Begin VB.CommandButton cmdEffects 
            Caption         =   " H-smooth"
            Height          =   225
            Index           =   0
            Left            =   195
            TabIndex        =   19
            Top             =   285
            Width           =   1020
         End
         Begin VB.CommandButton cmdEffects 
            Caption         =   "V-smooth"
            Height          =   225
            Index           =   1
            Left            =   195
            TabIndex        =   18
            Top             =   585
            Width           =   1035
         End
         Begin VB.CommandButton cmdEffects 
            Caption         =   "H&&V-smooth"
            Height          =   225
            Index           =   2
            Left            =   195
            TabIndex        =   17
            Top             =   885
            Width           =   1035
         End
      End
      Begin VB.CommandButton cmdGO 
         BackColor       =   &H00C0FFC0&
         Caption         =   "GO"
         Height          =   375
         Left            =   195
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   4500
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
         Top             =   195
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
         Caption         =   "PalNmae"
         Height          =   210
         Left            =   1740
         TabIndex        =   23
         Top             =   3990
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
      Left            =   4110
      ScaleHeight     =   8
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   7
      TabIndex        =   2
      Top             =   3690
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
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3435
      Left            =   150
      ScaleHeight     =   229
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   257
      TabIndex        =   0
      Top             =   270
      Width           =   3855
   End
   Begin VB.Frame frmTunnel 
      Caption         =   "Tunnel"
      Height          =   1350
      Left            =   6030
      TabIndex        =   80
      Top             =   6345
      Width           =   1455
      Begin VB.CheckBox chkSpin 
         Caption         =   "Spin"
         Height          =   210
         Left            =   210
         TabIndex        =   85
         Top             =   915
         Width           =   660
      End
      Begin VB.TextBox txtTunnelSpeed 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   255
         Left            =   435
         TabIndex        =   84
         Text            =   "0"
         Top             =   435
         Width           =   240
      End
      Begin VB.HScrollBar HSTunnelSpeed 
         Height          =   270
         Left            =   180
         Max             =   16
         Min             =   1
         TabIndex        =   83
         Top             =   420
         Value           =   1
         Width           =   750
      End
      Begin VB.CheckBox chkTunnelStop 
         Caption         =   " S"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   1065
         Style           =   1  'Graphical
         TabIndex        =   82
         Top             =   900
         Width           =   255
      End
      Begin VB.OptionButton optTunnelStart 
         Caption         =   "S"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   240
         Left            =   1065
         Style           =   1  'Graphical
         TabIndex        =   81
         Top             =   435
         Width           =   270
      End
      Begin VB.Label Label10 
         Caption         =   "Speed"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   330
         TabIndex        =   86
         Top             =   225
         Width           =   480
      End
   End
   Begin VB.Shape Shape2 
      Height          =   1485
      Left            =   135
      Top             =   6270
      Width           =   10950
   End
   Begin VB.Label LabRGB 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "RGB"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   7920
      TabIndex        =   21
      Top             =   5985
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   9210
      TabIndex        =   3
      Top             =   5985
      Width           =   1080
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

' Plasma Worlds ASM API version:  by Robert Rayment

' June 2002

' Addition 10/6/02
' Cylindrical projection Tunnel plus rotation

' Additions 5/6/02
' Toggle Small Map of plasma
' Track shown on map for Landscape
' Allow -ve heights for Landscape
' Left half shading added for 3D shapes

Option Base 1
DefLng A-W  ' All variables Long unless
DefSng X-Z  ' Singles or otherwise defined.



Private Sub cmdHelp_Click()

MsgBox " Press GO to get a new plasma pattern" & vbCrLf & vbCrLf & _
" Green/Red dot are Start/Stop buttons." & vbCrLf & vbCrLf & _
" Remember you can use the arrow" & vbCrLf & _
" keys on an active scroll bar" & vbCrLf & vbCrLf & _
" NB Plasma.bmp is saved in the app. folder.", , "Plasma Worlds"

End Sub

Private Sub Form_Load()

Caption = " Plasma Worlds: by Robert Rayment"

'Get app path
PathSpec$ = App.Path
If Right$(PathSpec$, 1) <> "\" Then PathSpec$ = PathSpec$ & "\"

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

With PIC2   ' Map pic
   .ScaleMode = vbPixels
   .Width = 128
   .Height = 96
   .Top = 421 '1 'PIC.Top
   .Left = 370 '1 'PIC.Left
   .DrawMode = vbXorPen
End With
PIC2.Visible = False

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

' PicMoveHandle init coords
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

' Scrolling Parameters
' Line store
If sizex > sizey Then
   ReDim Linestore(sizex)
Else
   ReDim Linestore(sizey)
End If
optScrollStart.Caption = Chr$(108)
optScrollStop.Caption = Chr$(108)
optScrollStop.Value = True
HScroll2.Value = 1
HScrollValue = 1
VScroll2.Value = 0
VScrollValue = 0

' Color cycling Parameters
optCycleStart.Caption = Chr$(108)   ' Windings font
optCycleStop.Caption = Chr$(108)
optCycleStop.Value = True
SBCycle.Value = 1
CycleValue = 1

' Transform
optRotationStart.Caption = Chr$(108)
chkRotationStop.Caption = Chr$(108)
Speed = 1
RotateDone = True ' -1
SBSpeed.Value = Speed
frmTControls.Visible = False

' Tunnel
optTunnelStart.Caption = Chr$(108)
chkTunnelStop.Caption = Chr$(108)
TunnelSpeed = 4
HSTunnelSpeed.Value = 4
Spin = False
chkSpin.Value = Unchecked

' Landscape
chkLandscapeStop.Caption = Chr$(108)
frmLControls.Visible = False
chkReflected.Value = Unchecked

' Map
MapDone = False

Show

ReDim IndexArray(1 To sizex, 1 To sizey) As Integer
ReDim ColArray(1 To sizex, 1 To sizey) As Long
PaletteMaxIndex = 511
ReDim Colors(0 To PaletteMaxIndex)

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' NEED TO BACK UP TO GET PAL FILES (Hence FileListBox not used)
' If Pals folder not in parallel with App Folder
' or compiled exe not in App Folder will
' need to alter this code here
' ...\ASMVersion\
' want ...\Pal\  ie back up 1 level
p = InStrRev(PathSpec$, "\", Len(PathSpec$) - 1)  ' NB VB6 only

' To pick up *.pal files in app folder - 1
'------------------
'Get FIRST entry
PalDir$ = Left$(PathSpec$, p) & "Pals\"
Spec$ = PalDir$ & "*.pal"
A$ = Dir$(Spec$)
'------------------

If A$ = "" Then
   MsgBox "Pal files not found"
   Unload Me
   End
End If

' Fill list box, capitalise first letter, lower case the rest,
' alphabetic property set at design

List1.AddItem UCase$(Left$(A$, 1)) & LCase$(Mid$(A$, 2))
Do
  A$ = Trim$(Dir)    'Gets NEXT entry
  If A$ <> "" Then
     List1.AddItem UCase$(Left$(A$, 1)) & LCase$(Mid$(A$, 2))
  Else   'a$="" indicates no more entries
     Exit Do
  End If
Loop


' Starting palette
PalSpec$ = PalDir$ & "Blues.pal"
LabPalName.Caption = "Blues.pal"
'PalSpec$ = PalDir$ & "Blues.pal"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

ReadPAL PalSpec$

ShowPalette

PlasmaDone = False

' Load machine code
Loadmcode PathSpec$ & "Plasma.bin", PlasmaArray()
Loadmcode PathSpec$ & "Effects.bin", EffectsArray()
Loadmcode PathSpec$ & "Scroll.bin", ScrollArray()
Loadmcode PathSpec$ & "Cycle.bin", CycleArray()
Loadmcode PathSpec$ & "Transform.bin", TransformArray()
Loadmcode PathSpec$ & "Landscape.bin", LandscapeArray()
Loadmcode PathSpec$ & "PTunnel.bin", TunnelArray()

KeyPreview = True

End Sub

'###### GO GO GO #############################################

Private Sub cmdGO_Click()

PlasmaDone = True

MousePointer = vbHourglass
cmdGO.BackColor = RGB(255, 192, 192)
DisableControls

ResizePIC sizex, sizey

Plasma

DisplayArray

' Start Counters to reproduce plasma
' at large scale for Landscape
SmoothXCount = 0
SmoothYCount = 0
SmoothXYCount = 0
BrightenCount = 0
DarkenCount = 0

EnableControls
cmdGO.BackColor = RGB(192, 255, 192)
MousePointer = vbDefault

TunnelDone = True
optTunnelStart.Value = False

End Sub

'###### MAP ###############################################

Private Sub cmdMap_Click()

' Map large display to small display PIC2

If Not PlasmaDone Then Exit Sub
'If RotateDone = False Then Exit Sub

PIC2.Visible = False

MapDone = Not MapDone

If MapDone Then
   
   FillBMPStruc sizex, sizey
   
   PIC2.Picture = LoadPicture
   PIC2.DrawMode = 13
   PIC2.Visible = True
   
   If StretchDIBits(PIC2.hdc, _
      0, 0, PIC2.Width, PIC2.Height, _
      0, 0, sizex, sizey, _
      ColArray(1, 1), bm, _
      1, vbSrcCopy) = 0 Then
         MsgBox ("Blit Error - Map")
         Unload Me
         End
   End If
   
   PIC2.Refresh
   PIC2.DrawMode = 7

End If

End Sub



Private Sub PIC2_Click()
   PIC2.Visible = False
   MapDone = False
End Sub

'###### PLASMA ################################################

Private Sub Plasma()

ReDim IndexArray(1 To sizex, 1 To sizey) As Integer
ReDim ColArray(1 To sizex, 1 To sizey) As Long

' Ensure same random sequence each time,
' this will always give the same pattern
' each time GO is pressed
'Rnd -1
'Randomize 1

' This gives a different pattern each time
' GO is pressed
Randomize Timer

RandSeed = Rnd * 511

' PaletteMaxIndex = 511
' Graininess & scale also depends on palette
' Rapidly changing palettes will be grainier
' StartNoise = 0      (0 - 128) ' graininess
' StartStepsize = 32  (0 - 128) ' scale

Noise = StartNoise   ' Graininess
Stepsize = StartStepsize   ' Scale

' Fill Plasma structure
PlasmaMCODE.sizex = sizex
PlasmaMCODE.sizey = sizey
PlasmaMCODE.PtrIndexArray = VarPtr(IndexArray(1, 1))
PlasmaMCODE.PtrColArray = VarPtr(ColArray(1, 1))
PlasmaMCODE.RandSeed = RandSeed
PlasmaMCODE.Noise = Noise
PlasmaMCODE.Stepsize = Stepsize
PlasmaMCODE.WrapX = WrapX
PlasmaMCODE.WrapY = WrapY
PlasmaMCODE.PaletteMaxIndex = PaletteMaxIndex
PlasmaMCODE.PtrColors = VarPtr(Colors(0))

ptrMC = VarPtr(PlasmaArray(1))
ptrStruc = VarPtr(PlasmaMCODE.sizex)

res = CallWindowProc(ptrMC, ptrStruc, 0&, 0&, 0&)

' Solid block on to plasma
'Cul = RGB(255, 0, 0)
'For iy = 1 To 20
'For ix = 1 To 20
'   ColArray(ix, iy) = Cul
'Next ix
'Next iy


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

If Not PlasmaDone Then Exit Sub

MousePointer = vbHourglass
cmdGO.BackColor = RGB(255, 192, 192)
DisableControls

ResizePIC sizex, sizey

'   Opcode As Long  ' 0 x, 1 y, 2 x & y - smoothing
'                   ' 3 brighten, 4 darken (Increment=4)

' Effects structure
EffectsMCODE.sizex = sizex
EffectsMCODE.sizey = sizey
EffectsMCODE.PtrColArray = VarPtr(ColArray(1, 1))
ptrMC = VarPtr(EffectsArray(1))
ptrStruc = VarPtr(EffectsMCODE.sizex)

' NB Counters for reproducing plasma
' at large scale for Landscape

Select Case Index
Case 0: SmoothX
        SmoothXCount = SmoothXCount + 1
Case 1: SmoothY
        SmoothYCount = SmoothYCount + 1
Case 2: SmoothXY
        SmoothXYCount = SmoothXYCount + 1
Case 3: Brighten
        BrightenCount = BrightenCount + 1
Case 4: Darken
        DarkenCount = DarkenCount + 1
End Select

EnableControls
cmdGO.BackColor = RGB(192, 255, 192)
MousePointer = vbDefault

End Sub

Private Sub SmoothX()  ' Two point wrapped smoothing in X direc
' ColArray(ix=1 to sizex, iy=1 to sizey)
Opcode = 0
res = CallWindowProc(ptrMC, ptrStruc, Opcode, 0&, 0&)
DisplayArray

End Sub

Private Sub SmoothY()  ' Two point wrapped smoothing in Y direc
' ColArray(ix=1 to sizex, iy=1 to sizey)
Opcode = 1
res = CallWindowProc(ptrMC, ptrStruc, Opcode, 0&, 0&)
DisplayArray

End Sub

Private Sub SmoothXY()  ' 8 point wrapped smoothing in Y direc
' ColArray(ix=1 to sizex, iy=1 to sizey)
Opcode = 2
res = CallWindowProc(ptrMC, ptrStruc, Opcode, 0&, 0&)
DisplayArray

End Sub


Private Sub Brighten()
' ColArray(ix=1 to sizex, iy=1 to sizey)
Opcode = 3
res = CallWindowProc(ptrMC, ptrStruc, Opcode, 0&, 0&)
DisplayArray

End Sub

Private Sub Darken()
' ColArray(ix=1 to sizex, iy=1 to sizey)
Opcode = 4
res = CallWindowProc(ptrMC, ptrStruc, Opcode, 0&, 0&)
DisplayArray

End Sub

'######  DISPLAY ColArray(ix, iy) TO PIC #####################

Private Sub DisplayArray()

PIC.Cls

If ScrollingDone = False And CyclingDone = False Then
   PicHandle.BackColor = vbRed
End If

FillBMPStruc sizex, sizey

If StretchDIBits(PIC.hdc, _
   0, 0, sizex, sizey, _
   0, 0, sizex, sizey, _
   ColArray(1, 1), bm, _
   1, vbSrcCopy) = 0 Then
      MsgBox ("Blit Error - DisplayArray")
      Unload Me
      End
End If
   
PIC.Refresh

PicHandle.BackColor = vbGreen

End Sub

'###### SHOW PALETTE ##################################

Private Sub ShowPalette()
   For N = 0 To 511
      LNGtoRGB Colors(N)   ' Need to reverse R & B
      picPal.Line (N, 0)-(N, 18), RGB(blue, green, red)
   Next N
   picPal.Refresh
End Sub

'##### PALETTE CHANGERS #################################

Private Sub cmdChangePalette_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

' SORT, REVERSE & GREY PALETTE

Select Case Index
Case 0   'SORT
   Quicksort Colors()   ' Sort palette on Long colors
Case 1   ' REVERSE
   ReversePalette
Case 2   'GREY
   GreyPalette
End Select

ShowPalette

If PlasmaDone Then ' ie plasma drawn

   ResizePIC sizex, sizey  ' get screen back after Cyl/Sphere display
   
   ' Refill ColArray
   For iy = 1 To sizey
   For ix = 1 To sizex
       ColArray(ix, iy) = Colors(IndexArray(ix, iy))
   Next ix
   Next iy
   
   ' Redraw plasma
   DisplayArray

End If

End Sub

Private Sub ReversePalette()
For i = 0 To 255
   Col = Colors(i)
   Colors(i) = Colors(511 - i)
   Colors(511 - i) = Col
Next i
End Sub

Private Sub GreyPalette()

ColMin = 100000
ColMax = -10000

For i = 0 To 511
      
   LNGtoRGB Colors(i)
    
   'Col = (1& * red + 1& * green + 1& * blue) \ 3     ' Even
   Col = (1& * red * 0.3 + green * 0.6 + blue * 0.1)  ' Visual
   Col = Col And 255
   Colors(i) = RGB(Col, Col, Col)
   If Col > ColMax Then ColMax = Col
   If Col < ColMin Then ColMin = Col
      
Next i

'Rescale  newCol = (Col - ColMin)*255/(ColMax-ColMin)

zColMul = 255 / (ColMax - ColMin)
For i = 0 To 511
      
   LNGtoRGB Colors(i)
    
   Col = ((1& * red - ColMin) * zColMul) And 255
   Colors(i) = RGB(Col, Col, Col)
      
Next i
   
End Sub


Private Sub Form_Unload(Cancel As Integer)

ScrollingDone = True
CyclingDone = True
RotationDone = True
TunnelDone = True
LandscapeDone = True
DoEvents

Unload Me

End

End Sub


'###### SELECT PALETTE FROM LISTBOX #########################

Private Sub List1_Click()

PalName$ = List1.List(List1.ListIndex)
LabPalName.Caption = PalName$
PalSpec$ = PalDir$ & PalName$

ReadPAL PalSpec$
   
ShowPalette

If Not PlasmaDone Then Exit Sub

ResizePIC sizex, sizey

MousePointer = vbHourglass
cmdGO.BackColor = RGB(255, 192, 192)
DoEvents

For iy = 1 To sizey
For ix = 1 To sizex
    ColArray(ix, iy) = Colors(IndexArray(ix, iy))
Next ix
Next iy

DisplayArray

cmdGO.BackColor = RGB(192, 255, 192)
MousePointer = vbDefault

End Sub

Private Sub picPal_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
LNGtoRGB picPal.Point(x, y)
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


'###### CHANGE StartNoise (Graininess) & StartStepsize (Scale) ##############

Private Sub HScroll1_Change(Index As Integer)

Select Case Index
Case 0   ' Graininess
   StartNoise = HScroll1(0).Value
   LabNum(0).Caption = Str$(StartNoise)
Case 1   ' Scale
   StartStepsize = HScroll1(1).Value
   LabNum(1).Caption = Str$(StartStepsize)
End Select

End Sub

'######  Show PIC RGB Color ########################################

Private Sub PIC_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
LNGtoRGB PIC.Point(x, y)
LabRGB.Caption = "RGB " & Str$(red) & Str$(green) & Str$(blue)
End Sub

'#######  RESIZE PIC ##############################################

Private Sub PicHandle_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
zInitX = x
zInitY = y
End Sub

Private Sub PicHandle_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

On Error GoTo LabHError

PicHandle.MousePointer = vbSizeAll

If Button = 1 Then
   
   ' Test PicHandle's new position
   
   LabHLeft = PicHandle.Left + (x - zInitX)
   LabHTop = PicHandle.Top + (y - zInitY)
   
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
   
   PlasmaDone = False

   If sizex > sizey Then
      ReDim Linestore(sizex)
   Else
      ReDim Linestore(sizey)
   End If
   ReDim IndexArray(1 To sizex, 1 To sizey) As Integer
   ReDim ColArray(1 To sizex, 1 To sizey) As Long

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

Private Sub PicHandle_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   PicHandle.MousePointer = vbDefault
End Sub


'###### SAVE BMP to "Plasma.BMP" only  ############################

Private Sub cmdSave_Click()

SavePicture PIC.Image, "Plasma.bmp"

End Sub

'###### SCROLLING ##############################################

Private Sub HScroll2_Change()
HScrollValue = HScroll2.Value
txtHScroll.Text = Trim$(HScrollValue)
End Sub

Private Sub VScroll2_Change()
VScrollValue = VScroll2.Value
txtVScroll.Text = Trim$(-VScrollValue)
End Sub

Private Sub optScrollStart_Click()
If Not PlasmaDone Then
   optScrollStop.Value = True
   Exit Sub
End If

CyclingDone = True
optCycleStop.Value = True

ResizePIC sizex, sizey

PicHandle.Enabled = False
frmDesign.Enabled = False
frmTransform.Enabled = False
frmLandscape.Enabled = False
frmPlusMinus.Enabled = False
frmTunnel.Enabled = False
DoEvents

ScrollMCODE.sizex = sizex
ScrollMCODE.sizey = sizey
ScrollMCODE.PtrColArray = VarPtr(ColArray(1, 1))
ScrollMCODE.PtrLineStore = VarPtr(Linestore(1))

ptrMC = VarPtr(ScrollArray(1))
ptrStruc = VarPtr(ScrollMCODE.sizex)

ScrollingDone = False

Do
   
   ScrollMCODE.HScrollValue = HScrollValue
   ScrollMCODE.VScrollValue = VScrollValue

   res = CallWindowProc(ptrMC, ptrStruc, 0&, 0&, 0&)

   DisplayArray

   DoEvents

Loop Until ScrollingDone

End Sub

Private Sub optScrollStop_Click()

ScrollingDone = True

frmDesign.Enabled = True
frmTransform.Enabled = True
frmLandscape.Enabled = True
frmPlusMinus.Enabled = True
PicHandle.Enabled = True
frmTunnel.Enabled = True

DoEvents

End Sub

'###### COLOR CYCLING ###########################################

Private Sub SBCycle_Change()
CycleValue = SBCycle.Value
txtCycle.Text = Trim$(CycleValue)
End Sub

Private Sub optCycleStart_Click()

If Not PlasmaDone Then
   optCycleStop.Value = True
   Exit Sub
End If

ScrollingDone = True
optScrollStop.Value = True

ResizePIC sizex, sizey

PicHandle.Enabled = False
frmDesign.Enabled = False
frmTransform.Enabled = False
frmLandscape.Enabled = False
frmPlusMinus.Enabled = False
frmTunnel.Enabled = False

DoEvents

CycleMCODE.sizex = sizex
CycleMCODE.sizey = sizey
CycleMCODE.PtrColArray = VarPtr(ColArray(1, 1))
CycleMCODE.CycleValue = CycleValue

ptrMC = VarPtr(CycleArray(1))
ptrStruc = VarPtr(CycleMCODE.sizex)

CyclingDone = False

Do

   CycleMCODE.CycleValue = CycleValue
   
   res = CallWindowProc(ptrMC, ptrStruc, 0, 0&, 0&)

   DisplayArray

   DoEvents

Loop Until CyclingDone

End Sub

Private Sub optCycleStop_Click()

CyclingDone = True

PicHandle.Enabled = True
frmDesign.Enabled = True
frmTransform.Enabled = True
frmLandscape.Enabled = True
frmPlusMinus.Enabled = True
frmTunnel.Enabled = True

DoEvents

End Sub

'########## TRANSFORMS ###################################################
'########## Cylinder Sphere Hyperboloid  #################################

Private Sub cmdTransform_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
' Cylinder Sphere Hyperboloid

If Not PlasmaDone Then Exit Sub

PIC2_Click  ' Clear map

RotationDone = True

PicHandle.Enabled = False
frmDesign.Enabled = False
frmScrolling.Enabled = False
frmCycle.Enabled = False
frmLandscape.Enabled = False
frmPlusMinus.Enabled = False
frmTControls.Visible = True
frmTunnel.Enabled = False

TransformType = Index

ResizePIC 512, 384

TransformSize = 10
VSBTransform.Value = TransformSize
txtVSB.Text = "10"

ShadeNumber = 0
SBShade.Value = ShadeNumber
txtShade.Text = "0"

TransformPreCalc

ShowTransformOnce

End Sub

Private Sub TransformPreCalc()
' Cylinder Sphere Hyperboloid

' From cmdTransform

'Public zSinTab()
ReDim zSinTab(0 To 271)
For i = 0 To 271
   zSinTab(i) = Sin(i * d2r#)
Next i

dimx = sizex
dimy = sizey

If sizex > 256 Then

   dimx = 256
   dimy = 256 * (sizey / sizex)
   If dimy > 384 Then
      dimy = 384
      dimx = dimy / (sizey / sizex)
   End If

End If

ReDim SmallArray(dimx, dimy)

zWScale = dimx / sizex
zHscale = dimy / sizey

ReDim ixdent(sizey)
ReDim zrdsq(sizey)

' Parameters
ixa = sizex \ 2
ixsc = ixa
yb = 0.5 * sizey
xc = 0.5 * sizex + 1
yc = 0.5 * sizey '+ 1

' Pre-calculate radius squared & slice indentation
For iy = 1 To sizey
   Select Case TransformType
   Case 0   ' Cylinder
      zrdsq(iy) = ixa ^ 2
      ixdent(iy) = xc - ixa
   Case 1   ' Ellipsoid
      zrdsq(iy) = ixa ^ 2 * (1 - ((iy - yc) / yb) ^ 2)
      ixdent(iy) = xc - Sqr(zrdsq(iy))
   Case 2   ' Hyperboloid
      zrdsq(iy) = (2 * ixa / 3) ^ 2 * (1 + ((iy - yc) / yb) ^ 2)
      ixdent(iy) = xc - Sqr(zrdsq(iy))
   End Select
   
   If ixdent(iy) < 1 Then ixdent(iy) = 1
   If ixdent(iy) > sizex Then ixdent(iy) = sizex

Next iy

' Proportionality factor
zScale = sizex / (2 * pi#)

' Fill TransformMCODE struc
PtrColArray = VarPtr(ColArray(1, 1))
PtrSmallArray = VarPtr(SmallArray(1, 1))
Ptrixdent = VarPtr(ixdent(1))
Ptrzrdsq = VarPtr(zrdsq(1))
PtrzSinTab = VarPtr(zSinTab(0))

TransformMCODE.sizex = sizex
TransformMCODE.sizey = sizey
TransformMCODE.PtrColArray = PtrColArray
TransformMCODE.dimx = dimx
TransformMCODE.dimy = dimy
TransformMCODE.PtrSmallArray = PtrSmallArray
TransformMCODE.Ptrixdent = Ptrixdent
TransformMCODE.Ptrzrdsq = Ptrzrdsq
TransformMCODE.ixa = ixa
TransformMCODE.ixsc = ixsc
TransformMCODE.zScale = zScale
TransformMCODE.zWScale = zWScale
TransformMCODE.zHscale = zHscale
TransformMCODE.ShadeNumber = ShadeNumber
TransformMCODE.PtrzSinTab = PtrzSinTab

ptrMC = VarPtr(TransformArray(1))
ptrStruc = VarPtr(TransformMCODE.sizex)

End Sub


Private Sub optRotationStart_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
' Rotation of Cylinder Sphere Hyperboloid

RotateDone = False

Do

   FillBMPStruc dimx, dimy
   
   'FillSmallArray  ' for VB only
   
   TransformMCODE.ixsc = ixsc  ' Changed by Speed
   TransformMCODE.ShadeNumber = ShadeNumber

   res = CallWindowProc(ptrMC, ptrStruc, 0&, 0&, 0&)  ' for ASM only
   
   PW = dimx * (10 / TransformSize)    ' dimx*(5 to 1/4)
   PH = dimy * (10 / TransformSize)    ' dimy*(5 to 1/4)
   PL = 256 - PW \ 2
   PT = 192 - PH \ 2
   PIC.Cls
   If StretchDIBits(PIC.hdc, _
      PL, PT, PW, PH, _
      0, 0, dimx, dimy, _
      SmallArray(1, 1), bm, _
      1, vbSrcCopy) = 0 Then
         MsgBox ("Blit Error - Rotation Start")
         Unload Me
         End
   End If
   
   PIC.Refresh

   ' Rotate Left/Right
   ixsc = ixsc + Speed
   If ixsc > sizex Then ixsc = 1
   If ixsc <= 0 Then ixsc = sizex

   DoEvents

Loop Until RotateDone

optRotationStart.Value = False

End Sub

Private Sub chkRotationStop_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
' Cylinder Sphere Hyperboloid

chkRotationStop.Value = Unchecked

RotateDone = True

PicHandle.Enabled = True
frmDesign.Enabled = True
frmScrolling.Enabled = True
frmCycle.Enabled = True
frmLandscape.Enabled = True
frmPlusMinus.Enabled = True
frmTControls.Visible = False
frmTunnel.Enabled = True

Erase SmallArray, ixdent, zrdsq

End Sub

Private Sub SBSpeed_Change()
' Cylinder Sphere Hyperboloid Speed
Speed = SBSpeed.Value
txtSpeed.Text = Trim$(Str$(Speed))

End Sub

Private Sub SBShade_Change()
' Cylinder Sphere Hyperboloid Shade
ShadeNumber = SBShade.Value
If ShadeNumber > 3 Then
   SBShade.Value = 0
   ShadeNumber = 0
End If
txtShade.Text = Trim$(Str$(ShadeNumber))

End Sub

Private Sub VSBTransform_Change()
' Cylinder Sphere Hyperboloid Size
TransformSize = VSBTransform.Value
txtVSB.Text = Trim$(Str$(TransformSize))

End Sub

Private Sub ShowTransformOnce()
' Cylinder Sphere Hyperboloid
' After TransformPreCalc

'FillSmallArray  'For VB only

res = CallWindowProc(ptrMC, ptrStruc, 0&, 0&, 0&)  ' For ASM only

 ' Set bm structure for SmallArray
FillBMPStruc dimx, dimy

' Center on WxH=512x384 picbox
PLEFT = (512 - dimx) \ 2
PTOP = (384 - dimy) \ 2

If StretchDIBits(PIC.hdc, _
   PLEFT, PTOP, dimx, dimy, _
   0, 0, dimx, dimy, _
   SmallArray(1, 1), bm, _
   1, vbSrcCopy) = 0 Then
      MsgBox ("Blit Error")
      Unload Me
      End
End If

PIC.Refresh

End Sub

Private Sub FillSmallArray()
' Cylinder Sphere Hyperboloid

' From ShowTransformOnce & optRotationStart

'####### This code is ASM-ified in main loop  ########
'####### but left here for reference          ########
 
 zpi = pi# / dimx
 
 For iy = sizey To 1 Step -1

   zrd = zrdsq(iy) + 1
   
   iyyd = 1 + (iy - 1) * zHscale
   
   If iyyd <= dimy Then
   
      For ixs = ixdent(iy) To sizex - ixdent(iy)
   
         zx = ixa - ixs
         zy = Sqr(zrd - (zx * zx))
         ztheta = zATan2(zy, zx)
   
         ixt = ixsc + zScale * ztheta
   
         If ixt < 1 Then
            ixt = sizex + ixt
         End If
         If ixt > sizex Then
            ixt = ixt - sizex
            If ixt > sizex Then ixt = ixt - sizex
         End If
   
         ixxd = 1 + (ixs - 1) * zWScale
         If ixxd > dimx Then Exit For
         
         Col = ColArray(ixt, iy)
         
         For kyy = iyyd To iyyd + 1
            'If kyy >= 1 And kyy <= dimy Then
            If kyy <= dimy Then
               For kxx = ixxd To ixxd + 1
                  If kxx <= dimx Then
                     
                     If kxx = ixxd Then  ' Need only calc Col once here
                        
                        If ShadeNumber > 0 Then
                           ' zpi = pi# / dimx
                           'LNGtoRGB Col
                           red = Col And &HFF
                           green = (Col \ &H100) And &HFF
                           blue = (Col \ &H10000) And &HFF
                           If ShadeNumber = 1 Then ' right half shading
                              deg = (zpi * kxx / 2) * r2d#
                              zdivor = zSinTab(deg)
                              'zdivor = Sin(zpi * kxx / 2)
                           ElseIf ShadeNumber = 2 Then   ' center shading
                              deg = (zpi * kxx) * r2d#
                              zdivor = zSinTab(deg)
                              'zdivor = Sin(zpi * kxx)
                           ElseIf ShadeNumber = 3 Then   ' left half shading
                              deg = (pi# / 2 + (zpi * kxx / 2)) * r2d#
                              zdivor = zSinTab(deg)
                              'zdivor = Sin(pi# / 2 + (zpi * kxx / 2))
                           End If
                           
                           R = red * zdivor
                           G = green * zdivor
                           B = blue * zdivor
                           Col = RGB(R, G, B)
                        End If
                     
                     End If   ' kxx = ixxd
                     
                     SmallArray(kxx, kyy) = Col
                  
                  Else
                     Exit For   ' kxx out of range
                  End If
               Next kxx
            Else  ' kyy out of range
               Exit For
            End If
         Next kyy
   
      Next ixs
   
   End If   ' iyyd <= dimy

Next iy

'############################################################

End Sub

'###### PIC RESIZERS #####################################

Private Sub ResizePIC(ByVal PWIDTH, ByVal PHEIGHT)

PIC.Width = PWIDTH
PIC.Height = PHEIGHT
PIC.Cls
DoEvents

'Re-position PicHandle to new PIC size
PicHandle.Left = PIC.Left + PIC.Width + 1
PicHandle.Top = PIC.Top + PIC.Height - PicHandle.Height + 3
   
With Shape1
   .Top = PIC.Top - 2
   .Left = PIC.Left - 2
   .Width = PWIDTH + 4
   .Height = PHEIGHT + 4
End With

Label3.Caption = "HxW = " & Str$(PIC.Height) & Str$(PIC.Width)
PIC.BackColor = 0
DoEvents

End Sub

Private Sub cmdResize_Click(Index As Integer)

ChangeSize = 0
Select Case Index
Case 0   ' - dx
   If sizex > 128 Then sizex = sizex - 8: ChangeSize = 1
Case 1   ' + dx
   If sizex < 512 Then sizex = sizex + 8: ChangeSize = 1
Case 2   ' - dy
   If sizey > 128 Then sizey = sizey - 8: ChangeSize = 1
Case 3   ' + dy
   If sizey < 384 Then sizey = sizey + 8: ChangeSize = 1
End Select

If ChangeSize = 1 Then
   ResizePIC sizex, sizey
   PlasmaDone = False

   If sizex > sizey Then
      ReDim Linestore(sizex)
   Else
      ReDim Linestore(sizey)
   End If
   ReDim IndexArray(1 To sizex, 1 To sizey) As Integer
   ReDim ColArray(1 To sizex, 1 To sizey) As Long
End If

End Sub

'###### TUNNEL ######################################

Private Sub optTunnelStart_Click()

If Not PlasmaDone Or TunnelDone = False Then
   optTunnelStart.Value = False
   Exit Sub
End If

TunnelDone = True
chkTunnelStop.Value = Checked

PIC2_Click  ' Clear any map
PIC.Cls

NumSmooths = SmoothXCount + SmoothYCount + SmoothXYCount

MousePointer = vbHourglass
cmdGO.BackColor = RGB(255, 192, 192)
DisableControls

Plasma

EnableControls
cmdGO.BackColor = RGB(192, 255, 192)
MousePointer = vbDefault
DoEvents

ReDim IndexLinestore(sizex)

PicHandle.Enabled = False
frmDesign.Enabled = False
frmScrolling.Enabled = False
frmCycle.Enabled = False
frmTransform.Enabled = False
frmLandscape.Enabled = False
frmPlusMinus.Enabled = False

ReDim CoordsX(sizex, sizey)
ReDim CoordsY(sizex, sizey)

zTrimmer = 0.005

' Fill ASM Tunnel Structure
TunnelStruc.sizex = sizex
TunnelStruc.sizey = sizey
TunnelStruc.PtrCoordsX = VarPtr(CoordsX(1, 1))
TunnelStruc.PtrCoordsY = VarPtr(CoordsY(1, 1))
TunnelStruc.PtrColArray = VarPtr(ColArray(1, 1))
TunnelStruc.PtrColors = VarPtr(Colors(0))
TunnelStruc.PtrIndexArray = VarPtr(IndexArray(1, 1))
TunnelStruc.kystart = kystart
TunnelStruc.kxstart = kxstart
TunnelStruc.zTrimmer = zTrimmer
TunnelStruc.NumSmooths = NumSmooths

ptrMC = VarPtr(TunnelArray(1))
ptrStruc = VarPtr(TunnelStruc.sizex)

' ASM  Cylindrical Projection   CylinProj Opcode = 0
res = CallWindowProc(ptrMC, ptrStruc, CylinProj, 0&, 0&)
   
'--------------------------------
' VB VB VB Cylindrical Projection
'For iy = 1 To sizey
'For ix = 1 To sizex
'   CoordsX(ix, iy) = 1
'   CoordsY(ix, iy) = 1
'Next ix
'Next iy

'ixdc = sizex \ 2
'iydc = sizey \ 2
'
'For ky = sizey To 1 Step -1
'For kx = sizex To 1 Step -1
'   zdx = kx - ixdc
'   zdy = ky - iydc
'   zRadius = Sqr(zdx ^ 2 + zdy ^ 2)
'   If sizex > sizey Then
'      iys = (sizey / sizex) * zRadius - 2   ' -2 gives hole at center
'   Else
'      iys = (sizex / sizey) * zRadius - 2
'   End If
'   If iys >= 1 And iys <= sizey Then
'      zAngle = zATan2(zdy, zdx) + pi# + 0.005    ' +.005 better integerization
'      ixs = zAngle * sizex / (2 * pi#)
'      If ixs >= 1 And ixs <= sizex Then
'         ColArray(kx, ky) = Colors(IndexArray(ixs, iys))
'         ' Make a LUT, nb integer arrays
'         CoordsX(kx, ky) = ixs
'         CoordsY(kx, ky) = iys
'      End If
'      ' so each kx,ky has an associated ixs,iys
'   End If
'Next kx
'Next ky
'--------------------------------

If NumSmooths > 0 Then     ' Smooth palette directly
   'ASM SmoothTunnel=2
   res = CallWindowProc(ptrMC, ptrStruc, SmoothTunnel, 0&, 0&)
   '---------------
   ' VB VB VB
'   For N = 1 To NumSmooths ' for tunnel, palette can
'      SmoothColors         ' easily be revcovered by
'      IndexArraySmoothXY   ' clicking palette in list
'   Next N                  ' box again.
   
   ShowPalette
End If

TunnelStruc.PtrCoordsX = VarPtr(CoordsX(1, 1))
TunnelStruc.PtrCoordsY = VarPtr(CoordsY(1, 1))

TunnelDone = False


kystart = 0
kxstart = 0

FillBMPStruc sizex, sizey

Do

   TunnelStruc.kystart = kystart
   TunnelStruc.kxstart = kxstart
   
   ' Advance & Spin   ASM  LUT Projection  LUTProj Opcode=1
   res = CallWindowProc(ptrMC, ptrStruc, LUTProj, 0&, 0&)
   
   '-----------------------------
   ' VB VB VB LUT Projection
'   For ky = sizey To 1 Step -1
'   For kx = sizex To 1 Step -1
'      iy = CoordsY(kx, ky) - kystart
'      If iy < 1 Then iy = sizey + iy
'      ix = CoordsX(kx, ky) - kxstart
'      If ix < 1 Then ix = sizex + ix
'      ColArray(kx, ky) = Colors(IndexArray(ix, iy))
'   Next kx
'   Next ky
   '-----------------------------
   
   ' Advance Y
   kystart = kystart + TunnelSpeed
   If kystart > sizey Then kystart = 1
   
   If Spin Then ' Advance X
      kxstart = kxstart + TunnelSpeed
      If kxstart > sizex Then kxstart = 1
   End If
   
   If StretchDIBits(PIC.hdc, _
      0, 0, sizex, sizey, _
      0, 0, sizex, sizey, _
      ColArray(1, 1), bm, _
      1, vbSrcCopy) = 0 Then
         MsgBox ("Blit Error - Tunnel")
         Unload Me
         End
   End If
      
   PIC.Refresh
   
   If GetQueueStatus(QS_MOUSEBUTTON Or QS_KEY) <> 0 Then DoEvents

Loop Until TunnelDone


End Sub

Private Sub chkTunnelStop_Click()

chkTunnelStop.Value = Unchecked
If TunnelDone = True Then Exit Sub
optTunnelStart.Value = False

TunnelDone = True
DoEvents

PicHandle.Enabled = True
frmDesign.Enabled = True
frmScrolling.Enabled = True
frmCycle.Enabled = True
frmTransform.Enabled = True
frmLandscape.Enabled = True
frmPlusMinus.Enabled = True

End Sub

Private Sub HSTunnelSpeed_Change()
TunnelSpeed = HSTunnelSpeed.Value
txtTunnelSpeed.Text = Trim$(Str$(TunnelSpeed))
End Sub

Private Sub chkSpin_Click()
Spin = Not Spin
End Sub

Private Sub SmoothColors()
' Smooth palette directly, 2-point runner
For i = 0 To 511

   im1 = i - 1: If im1 < 0 Then im1 = 511 + im1
   ip1 = i + 1: If ip1 > 511 Then ip1 = ip1 - 511
   
   LNGtoRGB Colors(im1)
   savr = red
   savb = blue
   savg = green
   LNGtoRGB Colors(ip1)
   savr = savr + red
   savb = savb + blue
   savg = savg + green
   
   Colors(i) = RGB(savr \ 2, savg \ 2, savb \ 2)

Next i

End Sub


Private Sub IndexArraySmoothXY()  ' 4 point wrapped smoothing in
' X & Y direc for IndexArray(ix=1 to sizex, iy=1 to sizey)
' ie Color number avearging

For ix = 1 To sizex
For iy = 1 To sizey
    
    iyA = (iy - 1)      ' If iy = 1 iyA = 0
    If iy = 1 Then iyA = sizey
    iyB = (iy + 1)      ' If iy = sizey iyB = sizey+1
    If iy = sizey Then iyB = 1
    
    savc = IndexArray(ix, iyA)
    savc = savc + IndexArray(ix, iyB)
    
    ixL = (ix - 1)   ' If ix = 1 ixL = 0
    If ix = 1 Then ixL = sizex
    ixR = (ix + 1)   ' If ix = sizex ixR = sizex+1
    If ix = sizex Then ixR = 1
    
    savc = savc + IndexArray(ixL, iy)
    savc = savc + IndexArray(ixR, iy)
    
    savc = savc \ 4
    
    IndexArray(ix, iy) = savc

Next iy
Next ix


End Sub

'###### LANDSCAPE ###################################

Private Sub cmdLandscape_Click(Index As Integer)

If Not PlasmaDone Then Exit Sub

If Index = 0 Then DISH = False Else DISH = True

TransformType = 2

PicHandle.Enabled = False
frmDesign.Enabled = False
frmTunnel.Enabled = False
frmScrolling.Enabled = False
frmCycle.Enabled = False
frmTransform.Enabled = False
frmPlusMinus.Enabled = False
frmLControls.Visible = True
frmTunnel.Enabled = False

' Resize to max picture
If sizex <> 512 Or sizey <> 384 Then
   
   ResizePIC 512, 384
   
   sizex = 512
   sizey = 384
   
   ' Save counts coz cmdGO zeroes them
   ' & cmdEffects_Click Increments them
   svSmoothXCount = SmoothXCount
   svSmoothYCount = SmoothYCount
   svSmoothXYCount = SmoothXYCount
   svBrightenCount = BrightenCount
   svDarkenCount = DarkenCount
   
   cmdGO_Click

   If svSmoothXCount > 0 Then
      For i = 1 To svSmoothXCount
         cmdEffects_Click 0
      Next i
   End If

   If svSmoothYCount > 0 Then
      For i = 1 To svSmoothYCount
         cmdEffects_Click 1
      Next i
   End If

   If svSmoothXYCount > 0 Then
      For i = 1 To svSmoothXYCount
         cmdEffects_Click 2
      Next i
   End If
   
   If svBrightenCount > 0 Then
      For i = 1 To svBrightenCount
         cmdEffects_Click 3
      Next i
   End If

   If svDarkenCount > 0 Then
      For i = 1 To svDarkenCount
         cmdEffects_Click 4
      Next i
   End If

   ' Put counts back to zero
   SmoothXCount = 0
   SmoothYCount = 0
   SmoothXYCount = 0
   BrightenCount = 0
   DarkenCount = 0

End If

'cmdMap_Click
   
   ' ReDim for other uses
   If sizex > sizey Then
      ReDim Linestore(sizex)
   Else
      ReDim Linestore(sizey)
   End If


PIC.Cls

' So always dealing with a ColArray(512,384) ie a floor surface

' Make height map from ColArray(sizex,sizey)
ReDim IndexArray(1 To sizex, 1 To sizey) As Integer
For iy = 1 To sizey
For ix = 1 To sizex
   LNGtoRGB ColArray(ix, iy)
   IndexArray(ix, iy) = (1 * red + green + blue) \ 2
   ' the 1 * forces integer calculation of bytes
Next ix
Next iy
' Heights 0 to 382

' Projection surface
dimx = 512
dimy = 384
ReDim SmallArray(dimx, dimy)

' Background color
ReDim ZeroArray(dimx, dimy)
For iy = 1 To dimy
For ix = 1 To dimx
   ZeroArray(ix, iy) = ColArray(128, 128) 'RGB(255, 255, 255)
Next ix
Next iy

' Floorcast from ColArray & IndexArray on to SmallArray
' and blit SmallArray to PIC

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
zdslope = 0.03  ' Starting slope of ray for each column
                ' scanned. Larger values flatten & smaller
                ' values exaggerate heights.

zhtmult = 0 '0.5  ' Starting Voxel ht multiplier.
txtHT.Text = Round(zhtmult, 1)

zslopemult = -150 ' Slope multiplier.  Larger -ve values
                  ' look more vertically down.

zvp = 256  ' Starting height of viewer.  Higher values lift
           ' the viewer & give a more pointed perspective.

LAH = 192 '128  ' Look ahead distance - the larger it is the more
           ' of the FloorSurf is scanned and the slower it
           ' will be.

'-----------------
' Starting eye position
xeye = sizex \ 2
yeye = 1

' Initial speed
zEL = 1
txtForwardBackward.Text = Trim$(Str$(zEL))

FirstIn = 0
VSForwardBackward.Value = -1
FirstIn = 1

' Initial angle
TANG = 33 '55
txtTANG.Text = Trim$(Str$(TANG))

' Starting values for scrollers
VSBF = 0
HANG = 0
VSHT = 0

' Fill Landscape struc

LandscapeMCODE.sizex = sizex
LandscapeMCODE.sizey = sizey
LandscapeMCODE.PtrColArray = VarPtr(ColArray(1, 1))
LandscapeMCODE.dimx = dimx
LandscapeMCODE.dimy = dimy
LandscapeMCODE.PtrSmallArray = VarPtr(SmallArray(1, 1))
LandscapeMCODE.PtrIndexArray = VarPtr(IndexArray(1, 1))
LandscapeMCODE.LAH = LAH
LandscapeMCODE.TANG = TANG
LandscapeMCODE.xeye = xeye
LandscapeMCODE.yeye = yeye
LandscapeMCODE.zvp = zvp
LandscapeMCODE.zdslope = zdslope
LandscapeMCODE.zslopemult = zslopemult
LandscapeMCODE.zhtmult = zhtmult
LandscapeMCODE.DISH = DISH
LandscapeMCODE.Reflect = Reflect

'Public LandscapeArray() As Byte  ' Array to hold machine code
'Public LandscapeMCODE As LandscapeStruc
ptrMC = VarPtr(LandscapeArray(1))
ptrStruc = VarPtr(LandscapeMCODE.sizex)

 
 ' Set bm structure for SmallArray
FillBMPStruc dimx, dimy

LandscapeDone = False
HSTANG.SetFocus

Do

xeye = xeye + zEL * Sin(TANG * d2r#)
yeye = yeye + zEL * Cos(TANG * d2r#)

If xeye > sizex Then xeye = 1 ': yeye = yeye - 1.41 * Cos(TANG * dtr#)
If xeye < 1 Then xeye = sizex ': yeye = yeye - 1.41 * Cos(TANG * dtr#)

If yeye > sizey Then yeye = 1 ': xeye = xeye - 1.41 * Sin(TANG * dtr#)
If yeye < 1 Then yeye = sizey ': xeye = xeye - 1.41 * Sin(TANG * dtr#)

If MapDone Then   ' Show position marker on map

   znfx = 2 * Sin(TANG * dtr#)
   znfy = 2 * Cos(TANG * dtr#)
   
   ixm = xeye + znfx
   iym = yeye + znfy
   
   If ixm > sizex Then
      ixm = ixm - sizex
   End If
   If ixm < 1 Then
      ixm = sizex + ixm
   End If
   If iym < 1 Then
      iym = sizey + iym
   End If
   If iym > sizey Then
      iym = iym - sizey
   End If
   
   ixm = ixm / 4
   iym = PIC2.Height - iym / 4
   
   PIC2.PSet (ixm, iym), RGB(255, 255, 255)
   PIC2.Line (ixm - 2, iym - 2)-(ixm + 2, iym + 2), RGB(255, 255, 255), BF
   PIC2.Refresh
   DoEvents

End If

' Background color
CopyMemory SmallArray(1, 1), ZeroArray(1, 1), dimx * dimy * 4

' Changers within Loop
LandscapeMCODE.xeye = xeye
LandscapeMCODE.yeye = yeye
LandscapeMCODE.TANG = TANG
LandscapeMCODE.zhtmult = zhtmult
LandscapeMCODE.Reflect = Reflect

'FloorCASTER  ' for VB only

res = CallWindowProc(ptrMC, ptrStruc, 0&, 0&, 0&) ' for ASM only

'Project SmallArray to PIC

 If StretchDIBits(PIC.hdc, _
   0, 0, 512, 384, _
   0, 0, dimx, dimy, _
   SmallArray(1, 1), bm, _
   1, vbSrcCopy) = 0 Then
      MsgBox ("Blit Error - Landscape")
      Unload Me
      End
End If

PIC.Refresh

PIC2.PSet (ixm, iym), RGB(255, 255, 255)
PIC2.Line (ixm - 2, iym - 2)-(ixm + 2, iym + 2), RGB(255, 255, 255), BF
PIC2.Refresh

DoEvents

Loop Until LandscapeDone

End Sub

Private Sub FloorCASTER()

' Left in for reference

' TANG values picked up from User
zrayang = (TANG * 2 * pi#) - dimx / 2  ' TANG = angle in degrees
xvp = xeye
yvp = yeye
'-----------------

For C = 0 To dimx - 1
   xray = xvp:  yray = yvp:  zray = zvp
   ' ie start at same viewpoint for each column
   
   ' x & y components to move a unit distance along ray
   zdy = Cos((zrayang + C) / 360)
   zdx = Sin((zrayang + C) / 360)
   
   ' Starting ray height below floor
   If DISH Then   ' DISH
   
      zdz = 2 * zdslope * zslopemult / (Cos((C - dimx / 2) / 360))
   
   Else  ' FLAT
   
      zdz = 2 * zdslope * zslopemult
   
   End If
   
   ' zslopemult - larger -ve values look more vertically down
   
   zvoxscale = 0  ' Moves zray up as each pixel is drawn so that
                  ' only higher pixels are drawn thereafter.
                  ' zray is also modified according to the distance
                  ' away from the viewer by the line zdz = zdz + zdslope
   
   row = 1        ' First screen row in SmallArray
   
   For cstep = 0 To LAH
      
      'Ensure rows and columns stay on FloorSurf
      ' & wrap
      ixR = xray
      If ixR < 1 Then ixR = sizex + ixR
      If ixR > sizex Then ixR = ixR - sizex
      
      iyr = yray
      If iyr < 1 Then iyr = sizey + iyr
      If iyr > sizey Then iyr = iyr - sizey
      
      FSHt = IndexArray(ixR, iyr)   ' Height array
      
      zHt = FSHt * zhtmult
      ' * zhtmult simply scales heights without distortion
      
      If zHt > zray Then   ' NB integer comp OK
         
         Do
            ' Copy color from floor to screen array
            
'            Put in some random dots
'            LNGtoRGB ColArray(ixr, iyr)
'            red = (1 * red + Rnd * (dimy - row) / 64) And 255
'            green = (1 * green + Rnd * (dimy - row) / 64) And 255
'            blue = (1 * blue + Rnd * (dimy - row) / 64) And 255
'            Cul = RGB(red, green, blue)
'            SmallArray(C + 1, row) = Cul 'ColArray(ixr, iyr)
'            If Reflect = 1 Then  ' REFLECT
'               SmallArray(C + 1, dimy - row) = Cul 'ColArray(ixr, iyr)
'            End If
            
            SmallArray(C + 1, row) = ColArray(ixR, iyr)
            If Reflect = 1 Then  ' REFLECT
               SmallArray(C + 1, dimy - row) = ColArray(ixR, iyr)
            End If
            
            zdz = zdz + zdslope
            zray = zray + zvoxscale
            row = row + 1
            If row > dimy Then
               Exit Sub
               'cstep = LAH
               'C = dimx - 1
               'Exit Do
            End If
        
            ' Collision detection OMITTED
        
        Loop Until zray > zHt ' NB integer comp OK
         
      End If
      xray = xray + zdx
      yray = yray + zdy
      zray = zray + zdz
      zvoxscale = zvoxscale + zdslope
   
   Next cstep   ' 0 -> LAH
   
Next C   ' 1 -> dimx

End Sub


Private Sub chkLandscapeStop_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

chkLandscapeStop.Value = Unchecked

LandscapeDone = True

If MapDone Then
   PIC2.Visible = False
   MapDone = Not MapDone
End If

DoEvents

PicHandle.Enabled = True
frmDesign.Enabled = True
frmScrolling.Enabled = True
frmCycle.Enabled = True
frmTunnel.Enabled = True
frmTransform.Enabled = True
frmPlusMinus.Enabled = True
frmTunnel.Enabled = True

frmLControls.Visible = False

Erase SmallArray, ZeroArray

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If LandscapeDone = True Then Exit Sub

' SA=0 Ht, SA=1 Speed, SA=2 Angle
'HSTANG.SetFocus
FirstIn = 0
Select Case KeyCode
' Angle
Case vbKeyLeft: HSTANG.SetFocus
Case vbKeyRight: HSTANG.SetFocus
' Speed
Case vbKeyUp:
VSForwardBackward.SetFocus
Case vbKeyDown:
VSForwardBackward.SetFocus
' HALT
Case vbKeySpace
   zEL = 0
   SA = 0
   txtForwardBackward.Text = Round(zEL, 1)
End Select

FirstIn = 1

End Sub


Private Sub HSTANG_Change()
' Landscape angle
' SA=0 Ht, SA=1 Speed, SA=2 Angle

If FirstIn = 0 Then Exit Sub

If HSTANG.Value < HSANG Then
   incr = -3
   'If SA = 1 Then zEL = zEL - 0.2
Else
   incr = 3
   'If SA = 1 Then zEL = zEL + 0.2
End If

SA = 2

HSANG = HSTANG.Value
TANG = TANG + incr
If TANG > 360 Then TANG = 0
If TANG < 0 Then TANG = 360
txtTANG.Text = Trim$(Str$(TANG))

If zEL > 10 Then zEL = 10
If zEL < -10 Then zEL = -10
txtForwardBackward.Text = Round(zEL, 1)

End Sub

Private Sub VSForwardBackward_Change()
' Landscape speed
' SA=0 Ht, SA=1 Speed, SA=2 Angle

If FirstIn = 0 Then Exit Sub

If VSForwardBackward.Value < VSBF Then
   zincr = -0.2
   'If SA = 2 Then TANG = TANG + 3
Else
   zincr = 0.2
   'If SA = 2 Then TANG = TANG - 3
End If

SA = 1

VSBF = VSForwardBackward.Value
zEL = zEL - zincr
If zEL > 10 Then zEL = 10
If zEL < -10 Then zEL = -10
txtForwardBackward.Text = Round(zEL, 1)

If TANG > 360 Then TANG = 0
If TANG < 0 Then TANG = 360
txtTANG.Text = Trim$(Str$(TANG))

End Sub

Private Sub VSHeight_Change()
' Landscape height
' SA=0 Ht, SA=1 Speed, SA=2 Angle
SA = 0
If VSHeight.Value < VSHT Then zincr = -0.1 Else zincr = 0.1
VSHT = VSHeight.Value
zhtmult = zhtmult - zincr
If zhtmult > 2.5 Then zhtmult = 2.5
If zhtmult < -3.5 Then zhtmult = -3.5
txtHT.Text = Round(zhtmult, 1)

End Sub

Private Sub chkReflected_Click()
' Landscape reflector

If chkReflected.Value = Unchecked Then
   Reflect = 0
Else
   Reflect = 1
End If

End Sub

