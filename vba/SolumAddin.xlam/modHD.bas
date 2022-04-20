Attribute VB_Name = "modHD"
Option Explicit
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : CONVERT_PICS
' Author    : Philip Swannell
' Date      : 10-Nov-2018
' Purpose   : Prior versions of method PastePictureToActiveCell inserted the pictures as links rather than embedding them
'             This method loops through the pictures inserted using such early version and converts them to compressed
'             embedded pictures, gets the location of the file to insert by parsing the comments that the old code
'             inserted, and copes with the possibility that the images inserted via the old code have been moved on the
'             sheet so that they are no long in the same cell as the associated comment.
' -----------------------------------------------------------------------------------------------------------------------
Sub CONVERT_PICS()

          Dim ws As Worksheet
1         On Error GoTo ErrHandler
2         Set ws = ActiveSheet
          Dim CommentText As String
          Dim FileName As String
          Dim MD5 As String
          Dim Path As String
          Dim STK As clsStacker
3         Set STK = CreateStacker()
          Dim STK2 As clsStringAppend
4         Set STK2 = New clsStringAppend
          Dim LookupArray
          Dim Message As String
          Dim NumConverted As Long
          Dim NumFound As Long
          Dim S As Shape
          Dim s2 As Shape
          Dim tmp As String

5         Message = "Convert ""linked"" pictures in the active worksheet to embedded pictures?"
6         If MsgBoxPlus(Message, vbOKCancel + vbDefaultButton2 + vbQuestion, "Convert Pictures", "Yes, convert", "No, do nothing") <> vbOK Then Exit Sub

          'Loop through all the comments in the sheet, parsing them to build up a look up table of file name to file path
7         For Each S In ws.Shapes
8             If S.Type = msoComment Then
9                 CommentText = S.TextFrame.Characters.caption
10                Path = sStringBetweenStrings(CommentText, "Path: ", vbLf)
11                FileName = sSplitPath(Path)
12                MD5 = sStringBetweenStrings(CommentText, "MD5: ", vbLf)
13                STK.Stack2D sArrayRange(FileName, Path, MD5)
14            End If
15        Next S

16        LookupArray = STK.Report

17        For Each S In ws.Shapes
18            If S.Type = msoLinkedPicture Then
19                NumFound = NumFound + 1
20                FileName = S.Name
21                FileName = sRegExReplace(FileName, " \(\d+\)$", vbNullString, False)
22                Path = sVLookup(FileName, LookupArray)
23                MD5 = sVLookup(FileName, LookupArray, 3)
24                If sIsErrorString(Path) Then
25                    Message = "Cannot find image file to replace LinkedPicture '" + S.Name + "' at cell " + S.TopLeftCell.address
26                ElseIf Not sFileExists(Path) Then
27                    Message = "File '" + Path + "' needed to replace replace LinkedPicture '" + S.Name + "' at cell " + S.TopLeftCell.address + " cannot be found."
28                ElseIf sFileCheckSum(Path) <> MD5 Then
29                    Message = "File '" + Path + "' needed to replace replace LinkedPicture '" + S.Name + "' at cell " + S.TopLeftCell.address + " is found but its MD5 is '" + sFileCheckSum(Path) + "' rather than '" + MD5 + "'."
30                Else
31                    Message = "Processing Linked Picture '" + S.Name + "' at cell " + S.TopLeftCell.address
32                    Set s2 = ws.Shapes.AddPicture2(FileName:=Path, linktofile:=msoFalse, _
                          savewithdocument:=msoCTrue, Left:=S.Left, Top:=S.Top, Width:=S.Width, Height:=S.Height, compress:=msoPictureCompressTrue)
33                    tmp = S.Name
34                    S.Delete
35                    s2.Name = tmp
36                    s2.AlternativeText = Path
37                    NumConverted = NumConverted + 1
38                End If
39                Debug.Print Message
40                STK2.Append Message + vbLf
41            End If
42        Next

43        Message = "All done" + vbLf + _
              "Num linked images found: " + CStr(NumFound) + vbLf + _
              "Num linked images converted: " + CStr(NumConverted) + vbLf

44        If NumConverted > 0 Then
45            Message = Message + vbLf + vbLf + "Details:" + vbLf + STK2.Report
46        End If

47        MsgBoxPlus Message, vbOKCancel + IIf(NumFound = NumConverted, vbInformation, vbExclamation)

48        Exit Sub
ErrHandler:
49        SomethingWentWrong "#CONVERT_PICS (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : PastePictureToActiveCell
' Author    : Philip Swannell
' Date      : 08-Oct-2018
' Purpose   : Allows user to select image file from disk and insert into the sheet. Picture will be placed inside the active
'             cell. Within limits, that cell is made bigger to accommodate. The name of the file is pasted to the cell and a
'             cell comment is added giving the path to the file, file size and MD5.
'             Assigned to Ctrl + 2
' -----------------------------------------------------------------------------------------------------------------------
Sub PastePictureToActiveCell()

          Const MinColumnWidth = 30.18    'if smaller than this the column width is set to it
          Const MaxCellHeight = 150
          Const WidthRatio = 0.9
          Const topGap = 10
          Const BottomGap = 15
          Const CommentFont = "Segoe UI"
          Const CommentFontSize = 9
          Const CommentHeight = 40
          Dim pName As String
          Dim Target As Range
          Const Title = "Paste Picture"

          Dim CommentText As String
          Dim CommentWidth
          Dim FileFilter As String
          Dim FileName As Variant
          Dim NeedToOverwrite As Boolean
          Dim SUH As clsScreenUpdateHandler

1         On Error GoTo ErrHandler

2         If ActiveCell Is Nothing Then
3             Exit Sub
4         End If
5         If Not UnprotectAsk(ActiveSheet, Title) Then Exit Sub

6         Set Target = ActiveCell

7         FileFilter = "Image Files (*.jpg;*.jpeg;*.gif;*.png;*.svg;*.tif;*.bmp),*.jpg;*.jpeg;*.gif;*.png;*.svg;*.tif;*.bmp"

8         FileName = GetOpenFilenameWrap("PastePictureToTarget", FileFilter, , _
              Title & ": Select file", "Paste", False, True, Target)
9         If VarType(FileName) <> vbString Then Exit Sub

10        Set SUH = CreateScreenUpdateHandler()

11        If Target.ColumnWidth < MinColumnWidth Then
12            Target.ColumnWidth = MinColumnWidth
13        End If

14        Target.Value = "'" + sSplitPath(FileName)
15        Target.VerticalAlignment = xlVAlignBottom
16        Target.HorizontalAlignment = xlHAlignCenter
          Dim Factor As Double
          Dim Factor1 As Double
          Dim Factor2 As Double
          Dim p As Object
          Dim q As Picture

17        For Each q In Target.Parent.Pictures
18            If q.TopLeftCell.address = Target.address Then
19                NeedToOverwrite = True
20                Exit For
21            End If
22        Next

23        If NeedToOverwrite Then
              Dim Prompt As String
24            Prompt = "There is already a picture in cell " & Replace(Target.address, "$", vbNullString) + vbLf + vbLf + _
                  "Do you want to replace it?"
25            If MsgBoxPlus(Prompt, vbYesNo + vbQuestion, "Paste picture", "Yes, Replace", "No, Do Nothing") <> vbYes Then Exit Sub
26            For Each p In Target.Parent.Pictures
27                If p.TopLeftCell.address = Target.address Then
28                    p.Delete
29                End If
30            Next
31        End If

          Dim ws As Worksheet
32        Set ws = Target.Parent

33        Set p = ws.Shapes.AddPicture(FileName, msoFalse, msoCTrue, ActiveCell.Left, ActiveCell.Top, -1, -1)

34        pName = sSplitPath(FileName)
35        If IsInCollection(Target.Parent.Shapes, pName) Then
              Dim j As Long
36            j = 2
37            Do While IsInCollection(Target.Parent.Shapes, sSplitPath(FileName) & " (" & CStr(j) & ")")
38                j = j + 1
39            Loop
40            pName = sSplitPath(FileName) & " (" & CStr(j) & ")"
41        End If

42        p.Name = pName

43        Factor1 = Target.Width * WidthRatio / p.Width
44        Factor2 = (MaxCellHeight - topGap - BottomGap) / p.Height
45        If Factor1 < Factor2 Then
46            Factor = Factor1
47        Else
48            Factor = Factor2
49        End If

50        p.ScaleWidth Factor, msoFalse, msoScaleFromTopLeft
51        p.AlternativeText = FileName
52        If WidthRatio < 1 Then
53            p.Left = Target.Left + Target.Width * (1 - WidthRatio) / 2
54            p.Top = Target.Top + topGap
55        End If

56        If Target.Height < p.Height + topGap + BottomGap Then
57            Target.RowHeight = p.Height + topGap + BottomGap
58        End If

59        CommentText = "Path: " & FileName & vbLf & _
              "Size: " & Format$(sFileInfo(FileName, "Size"), "###,###") & vbLf & _
              "MD5: " & sFileInfo(FileName, "MD5")

60        CommentWidth = sColumnMax(sStringWidth(sTokeniseString(CommentText, vbLf), CommentFont, CommentFontSize))(1, 1) + 10

61        With Target
62            .ClearComments
63            .AddComment
64            .Comment.Visible = False
65            .Comment.text text:=CommentText
66            With .Comment.Shape.TextFrame
67                .Characters.Font.Name = CommentFont
68                .Characters.Font.Size = CommentFontSize
69            End With
70            .Comment.Shape.Width = CommentWidth
71            .Comment.Shape.Height = CommentHeight
72        End With
73        Application.OnRepeat "Repeat Insert Picture", "PastePictureToActiveCell"

74        Exit Sub
ErrHandler:
75        SomethingWentWrong "#PastePictureToTarget (line " & CStr(Erl) + "): " & Err.Description & "!", vbExclamation, Title
End Sub

