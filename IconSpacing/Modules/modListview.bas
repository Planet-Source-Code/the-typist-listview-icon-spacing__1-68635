Attribute VB_Name = "modListview"
' The LVM_SETICONSPACING message adjusts the spacing between the icons in a listview
' that has the ICON style style set.

' Syntax
'            lResult = SendMessage(               returns LRESULT in lResult
'            (HWND) hWndControl,                  handle to destination control
'            (UINT) LVM_SETICONSPACING,           message ID
'            (WPARAM) wParam,                     = 0; not used, must be zero
'            (LPARAM) lParam                      = (LPARAM) MAKELONG(cx, cy)
'            );

' The parameters in the LPARAM are the following:
'    cx  - which is the distance between icons on the x-axis
'    cy  - which is the distance between icons on the y-axis

' To set the spacing between icons he cx or cy values must *include* the size of the icon and the amount of empty space
' desired between the icons else the *icons will overlap each other*.
'
' The message returns a DWORD value that contains the previous cx in the low word and the previous cy in the high word.
' To reset the thumbnail spaicing to default the LPARAM parameter should be -1

Option Explicit
'
'   Listview api messages.
Private Const LVM_FIRST As Long = &H1000
Private Const LVM_SETICONSPACING = LVM_FIRST + 53

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

'********************************************************************************
' Procedure : Sub SetThumbnailSpace
' DateTime  : 2007-05-20
' Author    : Ulrik Gustafsson
' Purpose   : Sets the x and y thumbnail space
'           : on a listview in icon view.
'           :
' Method    : Uses the LVM_SETICONSPACING api message.
'           :
' Remarks   : You only have to bother about the thumbnail spaces in twips.
'           : The procedure takes care of both including the size of the icons
'           : and converting the twips values to pixels.
'           :

' Arguments : lvw                   - The listview to process
'           : lngSpaceTwipsX        - The Thumbnail space (x) space specified in twips
'           : lngSpaceTwipsY        - The Thumbnail space (y) space specified in twips
'           : lngOldSpacePixelX     - The previous Thumbnail space (x) space specified in twips
'           :                         Optional parameter, default = -1
'           : lngOldSpacePixelY     - The previousThumbnail space (y) space specified in twips
'           :                         Optional parameter, default = -1
'           : blnReset              - Flag to mark if we want to reset thumbnail space
'           :                         to default. (value -1) Optional parameter, default
'           :                         is false.
'********************************************************************************
Public Sub SetThumbnailSpace(ByVal lvw As ListView, _
                             ByVal lngSpaceTwipsX As Long, _
                             ByVal lngSpaceTwipsY As Long, _
                             Optional ByRef lngOldSpacePixelX As Long = -1, _
                             Optional ByRef lngOldSpacePixelY As Long = -1, _
                             Optional blnReset As Boolean = False)

    Dim cxcy           As Long
    Dim lngSpacePixelX As Long
    Dim lngSpacePixelY As Long
    Dim lngRet         As Long
    Dim lngWidthPixelX As Long
    Dim lngHeightPixelY As Long
    
    Const c_ResetSpace As Long = -1
    
    '   Convert the submitted twips to pixels.
    lngSpacePixelX = TwipsToPixelsX(lngSpaceTwipsX)
    lngSpacePixelY = TwipsToPixelsY(lngSpaceTwipsY)
    
    '   Convert the listitem's coords to pixels.
    lngWidthPixelX = TwipsToPixelsX(lvw.ListItems(1).Width)
    lngHeightPixelY = TwipsToPixelsY(lvw.ListItems(1).Height)

    '   The thumbnail space is equal to the list item's width or height +
    '   the thumbnail space from the submitted parameters.    '
    
    lngSpacePixelX = lngWidthPixelX + lngSpacePixelX
    lngSpacePixelY = lngHeightPixelY + lngSpacePixelY

    '   Convert the pixel x, y values with MAKELONG function (macro) which
    '   converts the cxcy value to a long integer.
    '
    If Not blnReset Then
        cxcy = MakeLong(lngSpacePixelX, lngSpacePixelY)
    Else
    
        '   We are resetting the thumbnail space.
        cxcy = c_ResetSpace
    End If

    '   Apply the thumbnail space on the listview.
    lngRet = SendMessage(lvw.hwnd, LVM_SETICONSPACING, 0, ByVal cxcy)
    
    '   Retrieve the previous thumbnail spaces if specified.
    If lngOldSpacePixelX > -1 Then
    
        lngOldSpacePixelX = LoWord(lngRet)
        lngOldSpacePixelY = HiWord(lngRet)
        
        '  Subtract the values - the listitem coords.
        lngOldSpacePixelX = lngOldSpacePixelX - lngWidthPixelX
        lngOldSpacePixelY = lngOldSpacePixelY - lngHeightPixelY
              
    End If
    
End Sub
Public Function TwipsToPixelsX(ByRef lngPixelsX As Long) As Long
    TwipsToPixelsX = lngPixelsX \ Screen.TwipsPerPixelX
End Function
Public Function TwipsToPixelsY(ByRef lngPixelsY As Long) As Long
    TwipsToPixelsY = lngPixelsY \ Screen.TwipsPerPixelY
End Function

'********************************************************************************
' Procedure : Function MakeLong
' DateTime  : 2007-05-20
' Author    : Ulrik Gustafsson
' Purpose   : Combines two integers into a long integer.
'           : A vb translation of the C MAKELONG macro.
'           :
' Returns   : The "long integer"
'           :
' Arguments : LoWord -  The low-order word of the new value.
'             HiWord -  The high-order word of the new value.
'********************************************************************************
Private Function MakeLong(ByVal LoWord As Integer, ByVal HiWord As Integer) As Long
    MakeLong = ((HiWord * &H10000) + LoWord)
End Function

Private Function LoWord(ByVal LongIn As Long) As Integer
   Call CopyMemory(LoWord, LongIn, 2)
End Function

Private Function HiWord(ByVal LongIn As Long) As Integer
   Call CopyMemory(HiWord, ByVal (VarPtr(LongIn) + 2), 2)
End Function
