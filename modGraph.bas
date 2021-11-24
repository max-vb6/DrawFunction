Attribute VB_Name = "modGraph"
Option Explicit

Public Enum enuReturns
    lpSucceed = 0
    lpSyntaxError = 1
    lpRunError = 2
End Enum

Public bpDeclareX As Boolean
Public bpDeclareY As Boolean
Public lpRX As Long
Public lpRY As Long

Public Sub DrawLine(ByVal picObject As PictureBox, ByVal lpX1 As Long, ByVal lpY1 As Long, ByVal lpX2 As Long, ByVal lpY2 As Long, ByVal lpBaseX As Long, ByVal lpBaseY As Long, Optional clrColor As ColorConstants = vbBlack)
    On Error Resume Next
    picObject.Line (lpBaseX + lpX1, lpBaseY - lpY1)-(lpBaseX + lpX2, lpBaseY - lpY2), clrColor
End Sub

Public Function DrawFunctionX(ByVal picObject As PictureBox, ByVal strFormula As String, Optional clrColor As ColorConstants = vbBlack, Optional bpJumpZero As Boolean) As Boolean
    On Error Resume Next
    Dim lpX As Long
    Dim lpY As Long
    Dim lpPreviousX As Long
    Dim lpPreviousY As Long
    Dim lpXPos As Long
    Dim enuReturn As enuReturns
    bpDeclareX = False
    bpDeclareY = True
    lpPreviousY = picObject.ScaleHeight / 2
    lpRY = lpPreviousY
    lpXPos = Val(CalculateString(strFormula, enuReturn))
    If enuReturn = lpSyntaxError Then
        DrawFunctionX = False
        Exit Function
    End If
    lpPreviousX = lpXPos
    For lpY = picObject.ScaleHeight / 2 To -picObject.ScaleHeight / 2 Step -1
        If bpJumpZero = True Then
            If lpY = 0 Then
                lpPreviousY = -1
                lpRY = lpPreviousY
                lpPreviousX = Val(CalculateString(strFormula, enuReturn))
                If enuReturn = lpSyntaxError Then
                    DrawFunctionX = False
                    Exit Function
                End If
                GoTo JumpEX
            End If
        End If
        lpRY = lpY
        lpXPos = Val(CalculateString(strFormula, enuReturn))
        If enuReturn = lpSyntaxError Then
            DrawFunctionX = False
            Exit Function
        End If
        lpX = lpXPos
        DrawLine picObject, lpX, lpY, lpPreviousX, lpPreviousY, picObject.ScaleWidth / 2, picObject.ScaleHeight / 2, clrColor
        lpPreviousX = lpX
        lpPreviousY = lpY
        DoEvents
JumpEX:
    Next lpY
    bpDeclareX = False
    bpDeclareY = False
    DrawFunctionX = True
End Function

Public Function DrawFunctionY(ByVal picObject As PictureBox, ByVal strFormula As String, Optional clrColor As ColorConstants = vbBlack, Optional bpJumpZero As Boolean) As Boolean
    On Error Resume Next
    Dim lpX As Long
    Dim lpY As Long
    Dim lpPreviousX As Long
    Dim lpPreviousY As Long
    Dim lpYPos As Long
    Dim enuReturn As enuReturns
    bpDeclareX = True
    bpDeclareY = False
    lpPreviousX = -picObject.ScaleWidth / 2
    lpRX = lpPreviousX
    lpYPos = Val(CalculateString(strFormula, enuReturn))
    If enuReturn = lpSyntaxError Then
        DrawFunctionY = False
        Exit Function
    End If
    lpPreviousY = lpYPos
    For lpX = -picObject.ScaleWidth / 2 To picObject.ScaleWidth / 2
        If bpJumpZero = True Then
            If lpX = 0 Then
                lpPreviousX = 1
                lpRX = lpPreviousX
                lpPreviousY = Val(CalculateString(strFormula, enuReturn))
                If enuReturn = lpSyntaxError Then
                    DrawFunctionY = False
                    Exit Function
                End If
                GoTo JumpEX
            End If
        End If
        lpRX = lpX
        lpYPos = Val(CalculateString(strFormula, enuReturn))
        If enuReturn = lpSyntaxError Then
            DrawFunctionY = False
            Exit Function
        End If
        lpY = lpYPos
        DrawLine picObject, lpX, lpY, lpPreviousX, lpPreviousY, picObject.ScaleWidth / 2, picObject.ScaleHeight / 2, clrColor
        lpPreviousX = lpX
        lpPreviousY = lpY
        DoEvents
JumpEX:
    Next lpX
    bpDeclareX = False
    bpDeclareY = False
    DrawFunctionY = True
End Function

Private Function CalculateString(ByVal strFormula As String, ByRef lpReturn As enuReturns) As String
On Error GoTo errH
Dim lngTmp As Long
lngTmp = frmMain.SC.Eval(strFormula)
CalculateString = Trim(Str(frmMain.SC.Eval(strFormula)))
lpReturn = lpSucceed
Exit Function
errH:
If Err.Number = 1002 Then
lpReturn = lpSyntaxError
Else
lpReturn = lpRunError
End If
End Function
