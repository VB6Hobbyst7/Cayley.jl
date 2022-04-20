Attribute VB_Name = "modPastePicture"
'This module copied from the DVD that comes with "Professional Excel Development"
'Used in method ShowCommandBarPopup so that we can store custom icons on the sheet
'"Custom Icons" rather than having to manage separate .bmp files. Philip Swannell 14/11/2013

'An alternative approach that might work and would not require Windows API calls would be to use
'SavePicture and LoadPicture to save the pictures to temporary files.

' Description:  Creates a standard Picture object from whatever is on the clipboard.
'               This object can then be assigned to (for example) an Image control
'               on a userform. The PastePicture function takes an optional argument
'               of the picture type - xlBitmap or xlPicture.
'
'               This is a standard drop-in module, not changed at all for the book.
'               It can be treated as a 'black-box' to get a picture from the clipboard.
'
' Authors:      Rob Bovey, www.appspro.com
'               Stephen Bullen, www.oaltd.co.uk
'
' Chapter Change Overview
' Ch#   Comment
' --------------------------------------------------------------
' 08    Initial version, added to support the table-driven commandbar builder.
'       Copied from a public example on the bmsltd.ie web site.

Option Explicit
Option Private Module

' **************************************************************************
' Module Constant Declarations Follow
' **************************************************************************

' API Format types.
Private Const CF_BITMAP As Long = 2
'Private Const CF_PALETTE As Long = 9
Private Const CF_ENHMETAFILE As Long = 14
Private Const IMAGE_BITMAP As Long = 0
Private Const LR_COPYRETURNORG As Long = &H4

' **************************************************************************
' Module Type Structure Declarations Follow
' **************************************************************************
' Declare a UDT to store a GUID for the IPicture OLE interface.
Private Type guid
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

' Declare a UDT to store the bitmap information.
Private Type uPicDesc
    Size As Long
    Type As Long
    hPic As Long
    hPal As Long
End Type

' **************************************************************************
' Module DLL Declarations Follow
' **************************************************************************
'PGS 24/11/2015 used Microsoft Office 2010 Code Compatibility Inspector to make changes for 64-bit compatibility

' Does the clipboard contain a bitmap/metafile?
Private Declare PtrSafe Function IsClipboardFormatAvailable Lib "USER32" (ByVal wFormat As Long) As Long
' Open the clipboard to read
Private Declare PtrSafe Function OpenClipboard Lib "USER32" (ByVal hWnd As LongPtr) As Long
' Get a pointer to the bitmap/metafile. PGS: using "Alias "GetClipboardDataA"" does not work, removing it does work...
'Private Declare PtrSafe Function GetClipboardData Lib "USER32" Alias "GetClipboardDataA" (ByVal wFormat As Long) As LongPtr
Private Declare PtrSafe Function GetClipboardData Lib "USER32" (ByVal wFormat As Long) As LongPtr
' Close the clipboard
Private Declare PtrSafe Function CloseClipboard Lib "USER32" () As Long
' Convert the handle into an OLE IPicture interface.
'Private Declare PtrSafe Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As uPicDesc, RefIID As GUID, ByVal fPictureOwnsHandle As LongPtr, IPic As IPicture) As LongPtr
'Declaration below taken from http://www.vbaexpress.com/forum/showthread.php?25275-Saving-Clipboard-data-as-picture/page2
Private Declare PtrSafe Function OleCreatePictureIndirect Lib "oleaut32.dll" (PicDesc As uPicDesc, RefIID As guid, ByVal fPictureOwnsHandle As LongPtr, IPic As IPicture) As Long
' Create our own copy of the metafile, so it doesn't get wiped out by subsequent clipboard updates.
Declare PtrSafe Function CopyEnhMetaFile Lib "gdi32" Alias "CopyEnhMetaFileA" (ByVal hemfSrc As LongPtr, ByVal lpszFile As String) As LongPtr
' Create our own copy of the bitmap, so it doesn't get wiped out by subsequent clipboard updates.
Declare PtrSafe Function CopyImage Lib "USER32" (ByVal handle As LongPtr, ByVal un1 As Long, ByVal N1 As Long, ByVal N2 As Long, ByVal un2 As Long) As LongPtr

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Comments:   Get a Picture object showing whatever's on the clipboard.
'
' Arguments:  lXlPicType  The type of picture to create. Can be one of:
'                         - xlPicture to create a metafile (default)
'                         - xlBitmap to create a bitmap
'
' Date        Developer           Action
' ---------------------------------------------------------------------------------------------------------------------------------
' 30 Oct 98   Stephen Bullen      Created
' 15 Nov 98   Stephen Bullen      Updated to create our own copies of the clipboard images
'
Public Function PastePicture(Optional lXlPicType As Long = xlPicture) As IPicture

          Dim hPicAvail As Long
          Dim lReturn As Long
          Dim hPtr As Variant    'Long or LongPtr
          Dim lPicType As Long
          Dim hCopy As Variant    'Long or LongPtr

          ' Convert the type of picture requested from the xl constant to the API constant.
1         On Error GoTo ErrHandler
2         lPicType = IIf(lXlPicType = xlBitmap, CF_BITMAP, CF_ENHMETAFILE)

          ' Check if the clipboard contains the required Format.
3         hPicAvail = IsClipboardFormatAvailable(lPicType)

4         If hPicAvail <> 0 Then

              ' Get access to the clipboard.
5             lReturn = OpenClipboard(0&)

6             If lReturn > 0 Then

                  ' Get a handle to the image data.
7                 hPtr = GetClipboardData(lPicType)

                  ' Create our own copy of the image on the clipboard, in the appropriate Format.
8                 If lPicType = CF_BITMAP Then
9                     hCopy = CopyImage(hPtr, IMAGE_BITMAP, 0, 0, LR_COPYRETURNORG)
10                Else
11                    hCopy = CopyEnhMetaFile(hPtr, vbNullString)
12                End If

                  ' Release the clipboard to other programs.
13                lReturn = CloseClipboard()

                  ' If we got a handle to the image, convert it into a Picture object and return it
14                If hPtr <> 0 Then Set PastePicture = CreatePicture(hCopy, 0, lPicType)

15            End If
16        End If
17        Exit Function
ErrHandler:
18        Throw "#PastePicture (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Comments:   Converts a image (and palette) handle into a Picture object.
'
' Date        Developer           Action
' ---------------------------------------------------------------------------------------------------------------------------------
' 30 Oct 98   Stephen Bullen      Created
'
Private Function CreatePicture(ByVal hPic As Long, ByVal hPal As Long, ByVal lPicType) As IPicture

          ' OLE Picture types.
          Const PICTYPE_BITMAP As Long = 1
          Const PICTYPE_ENHMETAFILE As Long = 4

          Dim IID_IDispatch As guid
          Dim IPic As IPicture
          Dim lReturn As Variant    'Long or LongPtr
          Dim uPicInfo As uPicDesc

          ' Create the Interface GUID (for the IPicture interface).
1         On Error GoTo ErrHandler
2         With IID_IDispatch
3             .Data1 = &H7BF80980
4             .Data2 = &HBF32
5             .Data3 = &H101A
6             .Data4(0) = &H8B
7             .Data4(1) = &HBB
8             .Data4(2) = &H0
9             .Data4(3) = &HAA
10            .Data4(4) = &H0
11            .Data4(5) = &H30
12            .Data4(6) = &HC
13            .Data4(7) = &HAB
14        End With

          ' Fill uPicInfo with necessary parts.
15        With uPicInfo
16            .Size = Len(uPicInfo)        ' Length of structure.
17            .Type = IIf(lPicType = CF_BITMAP, PICTYPE_BITMAP, PICTYPE_ENHMETAFILE)        ' Type of Picture
18            .hPic = hPic        ' Handle to image.
19            .hPal = IIf(lPicType = CF_BITMAP, hPal, 0)        ' Handle to palette (if bitmap).
20        End With

          ' Create the Picture object.
21        lReturn = OleCreatePictureIndirect(uPicInfo, IID_IDispatch, True, IPic)

          ' If an error occured, show the description.
22        If lReturn <> 0 Then Debug.Print "Create Picture: " & OLEError(lReturn)

          ' Return the new Picture object.
23        Set CreatePicture = IPic
24        Exit Function
ErrHandler:
25        Throw "#CreatePicture (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Comments:   Gets the message text for standard OLE errors
'
' Arguments:  lErrNum     The error number to retrieve the message text for.
'
' Date        Developer           Action
' ---------------------------------------------------------------------------------------------------------------------------------
' 30 Oct 98   Stephen Bullen      Created
'
Private Function OLEError(ByVal lErrNum As Long) As String

          ' OLECreatePictureIndirect return values.
          Const E_ABORT As Long = &H80004004
          Const E_ACCESSDENIED As Long = &H80070005
          Const E_FAIL As Long = &H80004005
          Const E_HANDLE As Long = &H80070006
          Const E_INVALIDARG As Long = &H80070057
          Const E_NOINTERFACE As Long = &H80004002
          Const E_NOTIMPL As Long = &H80004001
          Const E_OUTOFMEMORY As Long = &H8007000E
          Const E_POINTER As Long = &H80004003
          Const E_UNEXPECTED As Long = &H8000FFFF
          Const S_OK As Long = &H0

1         Select Case lErrNum
              Case E_ABORT
2                 OLEError = " Aborted"
3             Case E_ACCESSDENIED
4                 OLEError = " Access Denied"
5             Case E_FAIL
6                 OLEError = " General Failure"
7             Case E_HANDLE
8                 OLEError = " Bad/Missing Handle"
9             Case E_INVALIDARG
10                OLEError = " Invalid Argument"
11            Case E_NOINTERFACE
12                OLEError = " No Interface"
13            Case E_NOTIMPL
14                OLEError = " Not Implemented"
15            Case E_OUTOFMEMORY
16                OLEError = " Out of Memory"
17            Case E_POINTER
18                OLEError = " Invalid Pointer"
19            Case E_UNEXPECTED
20                OLEError = " Unknown Error"
21            Case S_OK
22                OLEError = " Success!"
23        End Select
End Function
