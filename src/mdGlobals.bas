Attribute VB_Name = "mdGlobals"
Option Explicit

Public Enum UcsBarCodeTypeEnum
    ucsBctAuto = 0
    ucsBctEan13
    ucsBctEan8
    ucsBctEan128
    ucsBctUpcA
    ucsBctUpcE
End Enum

Public Sub PushError()

End Sub

Public Sub PopPrintError(sFunction As String, sModule As String)
    Debug.Print sModule & "." & sFunction & ": " & Error
End Sub

Public Sub PopRaiseError(sFunction As String, sModule As String)
    Err.Raise Err.Number, sModule & "." & sFunction & vbCrLf & Err.Source, Err.Description
End Sub

Public Function C_Lng(v As Variant) As Long
    On Error Resume Next
    C_Lng = CLng(v)
    On Error GoTo 0
End Function

Public Function C_Str(v As Variant) As String
    On Error Resume Next
    C_Str = CStr(v)
    On Error GoTo 0
End Function

Public Function InitStdFont() As StdFont
    Set InitStdFont = New StdFont
End Function

Public Function GetModuleInstance(sModule As String, sInstance As String, Optional sDebugID As String) As String

End Function

Public Function IsOnlyDigits(sText As String) As Boolean
    If LenB(sText) <> 0 Then
        IsOnlyDigits = Not (sText Like "*[!0-9]*")
    End If
End Function

Public Property Get EmptyVariantArray() As Variant
    EmptyVariantArray = Array()
End Property


Public Function At(Data As Variant, ByVal Index As Long, Optional Default As String) As String
    Const FUNC_NAME     As String = "At"
    
    On Error GoTo EH
    At = Default
    If IsArray(Data) Then
        If LBound(Data) <= Index And Index <= UBound(Data) Then
            At = C_Str(Data(Index))
        End If
    End If
    Exit Function
EH:
End Function
