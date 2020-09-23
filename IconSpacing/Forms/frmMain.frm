VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Change listview's thumbnail spacing (demo)"
   ClientHeight    =   7950
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8715
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7950
   ScaleWidth      =   8715
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkReset 
      Caption         =   "Reset thumbnail space to default"
      Height          =   615
      Left            =   6720
      TabIndex        =   5
      Top             =   4920
      Width           =   1815
   End
   Begin VB.ComboBox cboArrMode 
      Height          =   315
      Left            =   6720
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   4440
      Width           =   1695
   End
   Begin MSComctlLib.ImageList iml 
      Left            =   4560
      Top             =   4200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Slider sldX 
      Height          =   375
      Left            =   840
      TabIndex        =   6
      Top             =   5880
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      Max             =   1000
   End
   Begin MSComctlLib.ListView lvw 
      Height          =   3855
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   6800
      Arrange         =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      OLEDragMode     =   1
      _Version        =   393217
      Icons           =   "iml"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDragMode     =   1
      NumItems        =   0
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   495
      Left            =   6960
      TabIndex        =   8
      Top             =   7200
      Width           =   1455
   End
   Begin MSComctlLib.Slider sldY 
      Height          =   1935
      Left            =   480
      TabIndex        =   1
      Top             =   4200
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   3413
      _Version        =   393216
      Orientation     =   1
      Max             =   1000
   End
   Begin VB.Label lblXOldPixel 
      Caption         =   "Thumbnail Space - X  (Twips)"
      Height          =   375
      Left            =   1080
      TabIndex        =   12
      Top             =   7200
      Width           =   5295
   End
   Begin VB.Label lblYOldPixel 
      Caption         =   "Thumbnail Space - Y (Twips)"
      Height          =   375
      Left            =   1080
      TabIndex        =   11
      Top             =   5040
      Width           =   5415
   End
   Begin VB.Label lblYInPixel 
      Caption         =   "Thumbnail Space - Y (Twips)"
      Height          =   375
      Left            =   1080
      TabIndex        =   10
      Top             =   4680
      Width           =   3375
   End
   Begin VB.Label lblXInPixel 
      Caption         =   "Thumbnail Space - X  (Twips)"
      Height          =   375
      Left            =   1080
      TabIndex        =   9
      Top             =   6840
      Width           =   3255
   End
   Begin VB.Label lblArrMode 
      Caption         =   "Arrange listview:"
      Height          =   375
      Left            =   6720
      TabIndex        =   2
      Top             =   4200
      Width           =   1815
   End
   Begin VB.Label lblXInTwips 
      Caption         =   "Thumbnail Space - X  (Twips)"
      Height          =   375
      Left            =   1080
      TabIndex        =   7
      Top             =   6480
      Width           =   3255
   End
   Begin VB.Label lblYInTwips 
      Caption         =   "Thumbnail Space - Y (Twips)"
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      Top             =   4320
      Width           =   3375
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Project:       IconSpacing
' Last revision: 2007.05.20
' Version:       1.0.0
'----------------------------------------------------------------------------------------
' This is a demo which shows how to set the icon spacing in a listview's
' icon view.
' Copyright 2007 Ulrik Gustafsson.
' This demo is released "as is" without warranty or guaranty of any kind.
'

'********************************************************************************
' Procedure : Sub Form_Load
' DateTime  : 2007-05-20
' Author    : Ulrik Gustafsson
' Purpose   : Populates listview and combo.
'           :
' Method    : None
'           :
' Remarks   : None
'           :
' Returns   :
' Arguments : None
'********************************************************************************
Private Sub Form_Load()

    Dim i As Long

    '   Add some test items to the listview.
    For i = 1 To 10
        lvw.ListItems.Add , , "Test " & CStr(i), 1
    Next i

    '   Fill the combo and set the arrange mode of
    '   the listview.
    With cboArrMode
        .AddItem "lvwAutoNone"
        .AddItem "lvwAutoLeft"
        .AddItem "lvwAutoTop"
        .ListIndex = 1
        lvw.Arrange = cboArrMode.ListIndex
    End With

    '   Execute the new thumbnail space and
    '   print out x y values on labels.
    '
    SetThumbnailSpace lvw, sldX.Value, sldY.Value
    
    pvPrintPixelLabels sldX.Value, sldY.Value
    pvPrintTwipsLabels sldX.Value, sldY.Value
    pvPrintOldPixelLabels 0, 0
    
End Sub


'********************************************************************************
' Procedure : Sub pvEnableDisableControls
' DateTime  : 2007-05-20
' Author    : Ulrik Gustafsson
' Purpose   : Enable/Disable controls
'           : depedning of the blnDisabled parameter.
' Method    : None
'           :
' Remarks   : None
'           :
' Returns   :
' Arguments : blnReset - If controls should be enabled or not
'********************************************************************************
Private Sub pvEnableDisableControls(ByVal blnDisabled As Boolean)

    Dim ctl As Control
    
    For Each ctl In Me.Controls

        If TypeName(ctl) = "Label" Or TypeName(ctl) = "Slider" Then
            ctl.Enabled = Not blnDisabled
        End If

    Next

    lblArrMode.Enabled = True
    
End Sub
Private Sub cboArrMode_Click()

    Dim blnDisabled As Boolean

    '   Change enabled/disabled properties
    '   dependning on the listview's
    '   "arrange" mode
    lvw.Arrange = cboArrMode.ListIndex
    blnDisabled = cboArrMode.ListIndex = 0
    
    pvEnableDisableControls blnDisabled
    
End Sub

'********************************************************************************
' Procedure : Sub pvPrintTwipsLabels
' DateTime  : 2007-05-20
' Author    : Ulrik Gustafsson
' Purpose   : Prints x and y values of the thumbnails space
'           : on labels (twips).

' Method    : None
'           :
' Remarks   : None
'           :
' Returns   :
' Arguments :
'           : lngSpaceTwipsX        - The Thumbnail space (x) space specified in twips
'           : lngSpaceTwipsY        - The Thumbnail space (y) space specified in twips
'********************************************************************************
Private Sub pvPrintTwipsLabels(ByVal lngSpaceTwipsX As Long, _
                               ByVal lngSpaceTwipsY As Long)
                               
    lblXInTwips.Caption = "Thumbnail space - X  (Twips) " & CStr(lngSpaceTwipsX)
    lblYInTwips.Caption = "Thumbnail space - Y  (Twips) " & CStr(lngSpaceTwipsY)
End Sub

'********************************************************************************
' Procedure : Sub pvPrintPixelLabels
' DateTime  : 2007-05-20
' Author    : Ulrik Gustafsson
' Purpose   : Prints x and y values of the thumbnails space
'           : on labels (pixels).

' Method    : None
'           :
' Remarks   : None
'           :
' Returns   :
' Arguments :
'           : lngSpaceTwipsX        - The Thumbnail space (x) space specified in twips
'           : lngSpaceTwipsY        - The Thumbnail space (y) space specified in twips
'********************************************************************************
Private Sub pvPrintPixelLabels(ByVal lngSpaceTwipsX As Long, _
                               ByVal lngSpaceTwipsY As Long)
                               
    lblXInPixel.Caption = "Thumbnail space - X  (Pixels) " & CStr(TwipsToPixelsX(lngSpaceTwipsX))
    lblYInPixel.Caption = "Thumbnail space - Y  (Pixels) " & CStr(TwipsToPixelsY(lngSpaceTwipsY))
End Sub


'********************************************************************************
' Procedure : Sub pvPrintOldPixelLabels
' DateTime  : 2007-05-20
' Author    : Ulrik Gustafsson
' Purpose   : Prints x and y values of the thumbnails space
'           : on labels (previous pixels values).

' Method    : None
'           :
' Remarks   : None
'           :
' Returns   :
' Arguments :
'           : lngOldPixelsX        - The previous Thumbnail space (x) space specified in twips
'           : lngOldPixelsY        - The previous Thumbnail space (y) space specified in twips
'********************************************************************************
Private Sub pvPrintOldPixelLabels(ByVal lngOldPixelsX As Long, _
                                  ByVal lngOldPixelsY As Long)
                               
    lblXOldPixel.Caption = "Thumbnail space - X  (Previous thumbnail spacing - pixels) " & CStr(lngOldPixelsX)
    lblYOldPixel.Caption = "Thumbnail space - Y  (Previous thumbnail spacing - pixels) " & CStr(lngOldPixelsY)
End Sub
'********************************************************************************
' Procedure : Sub chkReset
' DateTime  : 2007-05-20
' Author    : Ulrik Gustafsson
' Purpose   : Resets icon spacing to default if checked.

' Method    : None
'           :
' Remarks   : None
'           :
' Returns   :
'********************************************************************************
Private Sub chkReset_Click()

    Dim blnReset As Boolean
    Dim lngOldSpacePixelX As Long
    Dim lngOldSpacePixelY As Long
    
    blnReset = chkReset.Value
    pvEnableDisableControls blnReset
    
    '   Execute the new thumbnail space and
    '   print out x y values on labels.
    SetThumbnailSpace lvw, sldX.Value, sldY.Value, lngOldSpacePixelX, lngOldSpacePixelY, blnReset
    
    pvPrintPixelLabels sldX.Value, sldY.Value
    pvPrintTwipsLabels sldX.Value, sldY.Value
    pvPrintOldPixelLabels lngOldSpacePixelX, lngOldSpacePixelY
    
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub


Private Sub sldX_Scroll()

    Dim lngOldSpacePixelX As Long
    Dim lngOldSpacePixelY As Long
    
    '   Execute the new thumbnail space and
    '   print out x y values on labels.
    SetThumbnailSpace lvw, sldX.Value, sldY.Value, lngOldSpacePixelX, lngOldSpacePixelY
    
    pvPrintPixelLabels sldX.Value, sldY.Value
    pvPrintTwipsLabels sldX.Value, sldY.Value
    pvPrintOldPixelLabels lngOldSpacePixelX, lngOldSpacePixelY
    
End Sub

Private Sub sldY_Scroll()

    
    Dim lngOldSpacePixelX As Long
    Dim lngOldSpacePixelY As Long
    
    '   Execute the new thumbnail space and
    '   print out x y values on labels.
    SetThumbnailSpace lvw, sldX.Value, sldY.Value, lngOldSpacePixelX, lngOldSpacePixelY
    
    pvPrintPixelLabels sldX.Value, sldY.Value
    pvPrintTwipsLabels sldX.Value, sldY.Value
    pvPrintOldPixelLabels lngOldSpacePixelX, lngOldSpacePixelY
    
End Sub
