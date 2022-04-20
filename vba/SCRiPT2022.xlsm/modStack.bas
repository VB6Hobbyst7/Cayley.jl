Attribute VB_Name = "modStack"
Option Explicit

'---------------------------------------------------------------------------------------------------
' Procedure : HStack
' Purpose   : Places arrays horizontally side by side. If the arrays are of unequal height then they will be padded
'             underneath with #NA! values.
'  Notes   1) Input arrays to range can have 0,1, or 2 dimensions
'          2) output array has lower bound 1, whatever the lower bounds of the inputs
'          3) input arrays of 1 dimension are treated as if they were columns, different from SAI equivalent fn.
' -----------------------------------------------------------------------------------------------------------------------
Public Function HStack(ParamArray Arrays()) As Variant

    Dim AllC As Long
    Dim AllR As Long
    Dim c As Long
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim R As Long
    Dim ReturnArray() As Variant
    Dim Y0 As Long

    On Error GoTo ErrHandler

    Static NA As Variant
    If IsEmpty(NA) Then NA = CVErr(xlErrNA)

    If IsMissing(Arrays) Then
        HStack = CreateMissing()
    Else
        For i = LBound(Arrays) To UBound(Arrays)
            If TypeName(Arrays(i)) = "Range" Then Arrays(i) = Arrays(i).Value
            If IsMissing(Arrays(i)) Then
                R = 0: c = 0
            Else
                Select Case NumDimensions(Arrays(i))
                    Case 0
                        R = 1: c = 1
                    Case 1
                        R = UBound(Arrays(i)) - LBound(Arrays(i)) + 1
                        c = 1
                    Case 2
                        R = UBound(Arrays(i), 1) - LBound(Arrays(i), 1) + 1
                        c = UBound(Arrays(i), 2) - LBound(Arrays(i), 2) + 1
                End Select
            End If
            If R > AllR Then AllR = R
            AllC = AllC + c
        Next i

        If AllR = 0 Then
            HStack = CreateMissing()
            Exit Function
        End If

        ReDim ReturnArray(1 To AllR, 1 To AllC)

        Y0 = 1
        For i = LBound(Arrays) To UBound(Arrays)
            If Not IsMissing(Arrays(i)) Then
                Select Case NumDimensions(Arrays(i))
                    Case 0
                        R = 1: c = 1
                        ReturnArray(1, Y0) = Arrays(i)
                    Case 1
                        R = UBound(Arrays(i)) - LBound(Arrays(i)) + 1
                        c = 1
                        For j = 1 To R
                            ReturnArray(j, Y0) = Arrays(i)(j + LBound(Arrays(i)) - 1)
                        Next j
                    Case 2
                        R = UBound(Arrays(i), 1) - LBound(Arrays(i), 1) + 1
                        c = UBound(Arrays(i), 2) - LBound(Arrays(i), 2) + 1

                        For j = 1 To R
                            For k = 1 To c
                                ReturnArray(j, Y0 + k - 1) = Arrays(i)(j + LBound(Arrays(i), 1) - 1, k + LBound(Arrays(i), 2) - 1)
                            Next k
                        Next j

                End Select
                If R < AllR Then        'Pad with #NA! values
                    For j = R + 1 To AllR
                        For k = 1 To c
                            ReturnArray(j, Y0 + k - 1) = NA
                        Next k
                    Next j
                End If

                Y0 = Y0 + c
            End If
        Next i
        HStack = ReturnArray
    End If

    Exit Function
ErrHandler:
    HStack = "#HStack (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : VStack
' Purpose   : Places arrays vertically on top of one another. If the arrays are of unequal width then they will be
'             padded to the right with #NA! values.
'  Notes   1) Input arrays to range can have 0, 1, or 2 dimensions.
'          2) output array has lower bound 1, whatever the lower bounds of the inputs.
'          3) input arrays of 1 dimension are treated as if they were rows, same as SAI equivalent fn.
' -----------------------------------------------------------------------------------------------------------------------
Function VStack(ParamArray Arrays()) As Variant
    Dim AllC As Long
    Dim AllR As Long
    Dim c As Long
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim R As Long
    Dim R0 As Long
    Dim ReturnArray() As Variant
    On Error GoTo ErrHandler

    Static NA As Variant
    If IsMissing(Arrays) Then
        VStack = CreateMissing()
    Else
        If IsEmpty(NA) Then NA = CVErr(xlErrNA)

        For i = LBound(Arrays) To UBound(Arrays)
            If TypeName(Arrays(i)) = "Range" Then Arrays(i) = Arrays(i).Value
            If IsMissing(Arrays(i)) Then
                R = 0: c = 0
            Else
                Select Case NumDimensions(Arrays(i))
                    Case 0
                        R = 1: c = 1
                    Case 1
                        R = 1
                        c = UBound(Arrays(i)) - LBound(Arrays(i)) + 1
                    Case 2
                        R = UBound(Arrays(i), 1) - LBound(Arrays(i), 1) + 1
                        c = UBound(Arrays(i), 2) - LBound(Arrays(i), 2) + 1
                End Select
            End If
            If c > AllC Then AllC = c
            AllR = AllR + R
        Next i

        If AllR = 0 Then
            VStack = CreateMissing
            Exit Function
        End If

        ReDim ReturnArray(1 To AllR, 1 To AllC)

        R0 = 1
        For i = LBound(Arrays) To UBound(Arrays)
            If Not IsMissing(Arrays(i)) Then
                Select Case NumDimensions(Arrays(i))
                    Case 0
                        R = 1: c = 1
                        ReturnArray(R0, 1) = Arrays(i)
                    Case 1
                        R = 1
                        c = UBound(Arrays(i)) - LBound(Arrays(i)) + 1
                        For j = 1 To c
                            ReturnArray(R0, j) = Arrays(i)(j + LBound(Arrays(i)) - 1)
                        Next j
                    Case 2
                        R = UBound(Arrays(i), 1) - LBound(Arrays(i), 1) + 1
                        c = UBound(Arrays(i), 2) - LBound(Arrays(i), 2) + 1

                        For j = 1 To R
                            For k = 1 To c
                                ReturnArray(R0 + j - 1, k) = Arrays(i)(j + LBound(Arrays(i), 1) - 1, k + LBound(Arrays(i), 2) - 1)
                            Next k
                        Next j

                End Select
                If c < AllC Then
                    For j = 1 To R
                        For k = c + 1 To AllC
                            ReturnArray(R0 + j - 1, k) = NA
                        Next k
                    Next j
                End If
                R0 = R0 + R
            End If
        Next i

        VStack = ReturnArray
    End If
    Exit Function
ErrHandler:
    VStack = "#VStack (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Sub FlipVectorSign(ByRef V)
          Dim i As Long

1         On Error GoTo ErrHandler
2         For i = LBound(V) To UBound(V)
3             V(i) = -V(i)
4         Next

5         Exit Sub
ErrHandler:
6         Throw "#FlipVectorSign (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
