VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBarCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=========================================================================
' $Header: /Dreem/1.5/Src/Common/cBarCode.cls 15    14.11.14 16:55 Wqw $
'
'   Dreem Enterprise Project
'   Copyright (c) 2003-2014 Unicontsoft
'
'   BarCode generator
'
' $Log: /Dreem/1.5/Src/Common/cBarCode.cls $
'
' 15    14.11.14 16:55 Wqw
' REF: polzwa empty variant array
'
' 14    12.11.11 19:23 Wqw
' REF: chars i back color na ean 128
'
' 13    4.07.11 0:08 Wqw
' REF: kalkulatsiq na laDigitsLength w GetEan128BarCode
'
' 12    2.07.11 17:13 Wqw
' REF: err handling dump-i params
'
' 11    31.05.10 14:51 Wqw
' ADD: Property HangSeparators, ShowDigits, HeightPercent, Function
' GetBarCode. REF: bql background, wdiga m_clrInvalid ako check digit-a e
' razlichen
'
' 10    4.02.10 19:54 Wqw
' REF: re-dim s expliciten type
'
' 9     13.01.10 14:52 Wqw
' REF: polzwa InitStdFont
'
' 8     3.11.09 17:19 Wqw
' REF: w debug mode MODULE_NAME sydyrja debug id
'
' 7     21.07.09 20:00 Wqw
' REF: impl instance name
'
' 6     17.03.09 17:51 Wqw
' DEL: dead code
'
' 5     7.02.09 1:28 Wqw
' ADD: Property InvalidColor, CurrentColor. REF: params na GetEanBarCode,
' GetUpcBarCode
'
' 4     27.11.08 12:36 Wqw
' ADD: GetEan128BarCode
'
' 3     5.11.08 10:55 Wqw
' ADD: impl na pvUpcEBarCode
'
' 2     4.09.08 12:28 Wqw
' REF: awtomatichno si izchislqwa bounding box
'
' 1     29.08.08 17:42 Wqw
' Initial implementation
'
'=========================================================================
Option Explicit
Private Const STR_MODULE_NAME As String = "cBarCode"

'=========================================================================
' API
'=========================================================================

Private Const WHITE_BRUSH                   As Long = 0
Private Const DT_CENTER                     As Long = &H1
Private Const DT_VCENTER                    As Long = &H4
Private Const DT_SINGLELINE                 As Long = &H20
Private Const TRANSPARENT                   As Long = 1

Private Declare Function CreateEnhMetaFile Lib "gdi32" Alias "CreateEnhMetaFileA" (ByVal hdcRef As Long, ByVal lpFileName As String, ByVal lpRect As Long, ByVal lpDescription As String) As Long
Private Declare Function CloseEnhMetaFile Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (lpPictDesc As PICTDESC, riid As Any, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Type RECT
    Left                As Long
    Top                 As Long
    Right               As Long
    Bottom              As Long
End Type

Private Type PICTDESC
    Size                As Long
    Type                As Long
    hBmpOrIcon          As Long
    hPal                As Long
End Type

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const STR_EAN_CHARACTER_SET     As String = "0001101|0011001|0010011|0111101|0100011|0110001|0101111|0111011|0110111|0001011"
Private Const STR_EAN13_SYSTEM          As String = "111111|110100|110010|110001|101100|100110|100011|101010|101001|100101"
Private Const STR_SUF2_SYSTEM           As String = "11|10|01|00"
Private Const STR_SUF5_SYSTEM           As String = "00111|01011|01101|01110|10011|11001|11100|10101|10110|11010"
Private Const STR_UPCE_SYSTEM           As String = "000111001011001101001110010011011001011100010101010110011010"
Private Const STR_EAN128_B_CHARACTERS   As String = "!""#$%&'()*+,-./0123456789:;<=>?@ABCDEFGHIJKLMNOPQRSTUVWXYZ[\]^_`abcdefghijklmnopqrstuvwxyz{|}~"
Private Const STR_EAN128_SYSTEM         As String = "1101100110011001101100110011001101001001100010010001100100010011001001100100010011000100100011001001100100100011001000100110001001001011001110010011011100100110011101011100110010011101100100111001101100111001011001011100110010011101101110010011001110100111011011101110100110011100101100111001001101110110010011100110100111001100101101101100011011000110110001101101010001100010001011000100010001101011000100010001101000100011000101101000100011000101000110001000101011011100010110001110100011011101011101100010111000110100011101101110111011011010001110110001011101101110100011011100010" & _
                                                    "110111011101110101100011101000110111000101101110110100011101100010111000110101110111101011001000010111100010101010011000010100001100100101100001001000011010000101100100001001101011001000010110000100100110100001001100001010000110100100001100101100001001011001010000111101110101100001010010001111010101001111001001011110010010011110101111001001001111010010011110010111101001001111001010011110010010110110111101101111011011110110110101011110001010001111010001011110101111010001011110001011110101000111101000101011101111010111101110111010111101111010111011010000100110100100001101001110011000111010"

Private Enum UcsEanDigitTypeEnum
    ucsEdtLeftOdd
    ucsEdtLeftEven
    ucsEdtRight
End Enum

Private Enum UcsEan128SystemEnum
    ucsEseSystemB
    ucsEseSystemC
End Enum

Private Enum UcsVertAlignEnum
    ucsValBottom = 1
    ucsValTop = 2
    ucsValCenter = 3
End Enum

Private m_hDC                   As Long
Private m_rcBound               As RECT
Private m_lBleed                As Long
Private m_bHangSeparators       As Boolean
Private m_bShowDigits           As Boolean
Private m_dblHeightPercent      As Double
Private m_oFont                 As IFont
Private m_clrInvalid            As OLE_COLOR
Private m_clrCurrent            As OLE_COLOR
Private m_sInstanceName         As String
#If DebugMode Then
    Private m_sDebugID          As String
#End If

'=========================================================================
' Error management
'=========================================================================

Friend Property Get frInstanceName() As String
    frInstanceName = m_sInstanceName
End Property

Private Property Get MODULE_NAME() As String
    #If DebugMode Then
        MODULE_NAME = GetModuleInstance(STR_MODULE_NAME, frInstanceName, m_sDebugID)
    #Else
        MODULE_NAME = GetModuleInstance(STR_MODULE_NAME, frInstanceName)
    #End If
End Property

Private Sub PrintError(sFunction As String)
    PushError
    PopPrintError sFunction, MODULE_NAME
End Sub

'=========================================================================
' Methods
'=========================================================================

Property Get InvalidColor() As OLE_COLOR
    InvalidColor = m_clrInvalid
End Property

Property Let InvalidColor(ByVal clrValue As OLE_COLOR)
    m_clrInvalid = clrValue
End Property

Property Get CurrentColor() As OLE_COLOR
    CurrentColor = m_clrCurrent
End Property

Property Get HangSeparators() As Boolean
    HangSeparators = m_bHangSeparators
End Property

Property Let HangSeparators(ByVal bValue As Boolean)
    m_bHangSeparators = bValue
End Property

Property Get ShowDigits() As Boolean
    ShowDigits = m_bShowDigits
End Property

Property Let ShowDigits(ByVal bValue As Boolean)
    m_bShowDigits = bValue
End Property

Property Get HeightPercent() As Double
    HeightPercent = m_dblHeightPercent
End Property

Property Let HeightPercent(dblValue As Double)
    m_dblHeightPercent = dblValue
End Property

'=========================================================================
' Methods
'=========================================================================

Public Function Init(ByVal lBleed As Long, ByVal bHangSeparators As Boolean, ByVal bShowDigits As Boolean) As Boolean
    Const FUNC_NAME     As String = "Init"
    Const STR_FONTS     As String = "OCR-B 10 BT:520|ocrb10:480|Arial Narrow:600"
    Dim vElem           As Variant
    
    On Error GoTo EH
    m_lBleed = lBleed
    m_bHangSeparators = bHangSeparators
    m_bShowDigits = bShowDigits
    m_clrInvalid = vbBlack
    m_dblHeightPercent = 100
    Set m_oFont = InitStdFont()
    m_oFont.Bold = True
    m_oFont.Size = 600
    For Each vElem In Split(STR_FONTS, "|")
        vElem = Split(vElem, ":")
        m_oFont.Name = vElem(0)
        If m_oFont.Name = vElem(0) Then
            m_oFont.Size = vElem(1)
            Exit For
        End If
    Next
    '--- success
    Init = True
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Public Function GetBarCode(sCode As String, eType As UcsBarCodeTypeEnum) As StdPicture
    Const FUNC_NAME     As String = "GetBarCode"
    
    On Error GoTo EH
    Select Case eType
    Case ucsBctEan13, ucsBctEan8
        Set GetBarCode = GetEanBarCode(sCode, eType)
    Case ucsBctUpcA, ucsBctUpcE
        Set GetBarCode = GetUpcBarCode(sCode, eType)
    Case ucsBctEan128
        Set GetBarCode = GetEan128BarCode(sCode)
    Case Else '--- auto
        If IsOnlyDigits(sCode) Then
            If Len(sCode) > 8 + 2 Then
                Set GetBarCode = GetEanBarCode(sCode, ucsBctEan13)
            Else
                Set GetBarCode = GetEanBarCode(sCode, ucsBctEan8)
            End If
        Else
            Set GetBarCode = GetEan128BarCode(sCode)
        End If
    End Select
    Exit Function
EH:
    PrintError FUNC_NAME & "(sCode=" & sCode & ", eType=" & eType & ")"
    Resume Next
End Function
                    
Public Function GetEanBarCode(sCode As String, Optional eType As UcsBarCodeTypeEnum = ucsBctAuto) As StdPicture
    Const FUNC_NAME     As String = "GetEanBarCode"
    Dim sBarCode        As String
    Dim sSuffix         As String
    Dim lSize           As Long
    
    On Error GoTo EH
    lSize = Len(sCode)
    If eType = ucsBctEan13 Or (eType = ucsBctAuto And lSize >= 13 And lSize <= 13 + 5) Then
        sBarCode = Left$(sCode, 13)
        sSuffix = Mid$(sCode, 14)
        eType = ucsBctEan13
    ElseIf eType = ucsBctEan8 Or (eType = ucsBctAuto And lSize >= 8 And lSize <= 8 + 5) Then
        sBarCode = Left$(sCode, 8)
        sSuffix = Mid$(sCode, 9)
        eType = ucsBctEan8
    Else
        sBarCode = sCode
        eType = IIf(Len(sBarCode) > 8, ucsBctEan13, ucsBctEan8)
    End If
    If eType = ucsBctEan8 Then
        Set GetEanBarCode = pvEan8BarCode(sBarCode, sSuffix)
    ElseIf eType = ucsBctEan13 Then
        Set GetEanBarCode = pvEan13BarCode(sBarCode, sSuffix)
    End If
    sCode = sBarCode & sSuffix
    Exit Function
EH:
    PrintError FUNC_NAME & "(sCode=" & sCode & ", eType=" & eType & ")"
    Resume Next
End Function

Public Function GetUpcBarCode(sCode As String, Optional eType As UcsBarCodeTypeEnum = ucsBctAuto) As StdPicture
    Const FUNC_NAME     As String = "GetUpcBarCode"
    Dim sBarCode        As String
    Dim sSuffix         As String
    Dim lSize           As Long
    
    On Error GoTo EH
    lSize = Len(sCode)
    If eType = ucsBctUpcA Or (eType = ucsBctAuto And lSize = 12 Or lSize = 12 + 2 Or lSize = 12 + 5) Then
        sBarCode = Left$(sCode, 12)
        sSuffix = Mid$(sCode, 13)
        eType = ucsBctUpcA
    ElseIf eType = ucsBctUpcE Or (eType = ucsBctAuto And lSize = 8 Or lSize = 8 + 2 Or lSize = 8 + 5) Then
        sBarCode = Left$(sCode, 8)
        sSuffix = Mid$(sCode, 9)
        eType = ucsBctUpcE
    ElseIf eType = ucsBctUpcA Or (eType = ucsBctAuto And lSize >= 12 And lSize <= 12 + 5) Then
        sBarCode = Left$(sCode, 12)
        sSuffix = Mid$(sCode, 13)
        eType = ucsBctUpcA
    ElseIf eType = ucsBctUpcE Or (eType = ucsBctAuto And lSize >= 8 And lSize <= 8 + 5) Then
        sBarCode = Left$(sCode, 8)
        sSuffix = Mid$(sCode, 9)
        eType = ucsBctUpcE
    Else
        sBarCode = sCode
        eType = IIf(Len(sBarCode) > 8, ucsBctUpcA, ucsBctUpcE)
    End If
    If eType = ucsBctUpcE Then
        Set GetUpcBarCode = pvUpcEBarCode(sBarCode, sSuffix)
    ElseIf eType = ucsBctUpcA Then
        Set GetUpcBarCode = pvUpcABarCode(sBarCode, sSuffix)
    End If
    sCode = sBarCode & sSuffix
    Exit Function
EH:
    PrintError FUNC_NAME & "(sCode=" & sBarCode & ", eType=" & eType & ")"
    Resume Next
End Function

Public Function GetEan128BarCode(sBarCode As String) As StdPicture
    Const FUNC_NAME     As String = "GetEan128BarCode"
    Dim lIdx            As Long
    Dim lPos            As Long
    Dim lStep           As Long
    Dim sCheckSum       As String
    Dim lRunningSum     As Long
    Dim laDigitsLength() As Long
    Dim eSystem         As UcsEan128SystemEnum
    Dim lOutIdx         As Long
    Dim lSize           As Long
    Dim saEncoding()    As String
    Dim rcLetter        As RECT
    Dim hPrevFont       As Long
    
    On Error GoTo EH
    '--- init local vars
    ReDim laDigitsLength(Len(sBarCode) + 2) As Long
    ReDim saEncoding(Len(sBarCode) + 1) As String
    m_clrCurrent = vbBlack
    lStep = 100
    lPos = 9 * lStep
    '--- prepare digits counter
    For lIdx = Len(sBarCode) To 1 Step -1
        If Asc(Mid$(sBarCode, lIdx, 1)) >= &H30 And Asc(Mid$(sBarCode, lIdx, 1)) <= &H39 Then
            lOutIdx = lOutIdx + 1
        Else
            lOutIdx = 0
        End If
        laDigitsLength(lIdx) = lOutIdx
        If InStr(1, STR_EAN128_B_CHARACTERS & " ", Mid$(sBarCode, lIdx, 1)) = 0 Then
            m_clrCurrent = m_clrInvalid
        End If
    Next
    lIdx = 1
    lOutIdx = 1
    eSystem = ucsEseSystemB
    Do While lIdx <= Len(sBarCode)
        If laDigitsLength(lIdx) >= 4 And lIdx = 1 Then
            ''' започваме със С
            eSystem = ucsEseSystemC
            lRunningSum = 105
            Do While laDigitsLength(lIdx) > 1
                lRunningSum = lRunningSum + lOutIdx * Mid$(sBarCode, lIdx, 2)
                saEncoding(lOutIdx) = "C"
                lOutIdx = lOutIdx + 1
                lIdx = lIdx + 2
            Loop
        ElseIf laDigitsLength(lIdx) >= 6 Then
            ''' смяна на В със С
            eSystem = ucsEseSystemC
            lRunningSum = lRunningSum + lOutIdx * 99
            saEncoding(lOutIdx) = ">"
            lOutIdx = lOutIdx + 1
            Do While laDigitsLength(lIdx) > 1
                lRunningSum = lRunningSum + lOutIdx * Mid$(sBarCode, lIdx, 2)
                saEncoding(lOutIdx) = "C"
                lOutIdx = lOutIdx + 1
                lIdx = lIdx + 2
            Loop
        ElseIf laDigitsLength(lIdx) >= 4 And laDigitsLength(lIdx) + lIdx - 1 = Len(sBarCode) Then
            '''завършваме със С
            If laDigitsLength(lIdx) Mod 2 Then
                lRunningSum = lRunningSum + lOutIdx * InStr(1, STR_EAN128_B_CHARACTERS, Mid$(sBarCode, lIdx, 1))
                saEncoding(lOutIdx) = "B"
                lOutIdx = lOutIdx + 1
                lIdx = lIdx + 1
            End If
            eSystem = ucsEseSystemC
            saEncoding(lOutIdx) = ">"
            lRunningSum = lRunningSum + lOutIdx * 99
            lOutIdx = lOutIdx + 1
            Do While laDigitsLength(lIdx) > 1
                lRunningSum = lRunningSum + lOutIdx * Mid$(sBarCode, lIdx, 2)
                saEncoding(lOutIdx) = "C"
                lOutIdx = lOutIdx + 1
                lIdx = lIdx + 2
            Loop
        Else
            If lIdx = 1 Then
                lRunningSum = 104
            End If
            If eSystem = ucsEseSystemC Then
                eSystem = ucsEseSystemB
                saEncoding(lOutIdx) = "<"
                lRunningSum = lRunningSum + lOutIdx * 100
                lOutIdx = lOutIdx + 1
            End If
            saEncoding(lOutIdx) = "B"
            lRunningSum = lRunningSum + lOutIdx * InStr(1, STR_EAN128_B_CHARACTERS, Mid$(sBarCode, lIdx, 1))
            lOutIdx = lOutIdx + 1
            lIdx = lIdx + 1
        End If
    Loop
    sCheckSum = lRunningSum Mod 103
    lSize = lOutIdx
    lIdx = 1
    lOutIdx = 1
    pvCreateMeta (53 + 11 * (lSize - 1)) * lStep, 55 * lStep
    If saEncoding(lOutIdx) = "B" Then
        pvOutputBars pvGetEan128Bars("STARTB"), lPos, lStep, " " ' , "•"
    Else
        pvOutputBars pvGetEan128Bars("STARTC"), lPos, lStep, " " ' , "·"
    End If
    For lOutIdx = 1 To lSize
        If saEncoding(lOutIdx) = "B" Then
            pvOutputBars pvGetEan128Bars(Mid$(sBarCode, lIdx, 1), ucsEseSystemB), lPos, lStep, " " ' , Mid$(sBarCode, lIdx, 1)
            lIdx = lIdx + 1
        ElseIf saEncoding(lOutIdx) = "C" Then
            pvOutputBars pvGetEan128Bars(Mid$(sBarCode, lIdx, 2), ucsEseSystemC), lPos, lStep, " " ' , Mid$(sBarCode, lIdx, 2)
            lIdx = lIdx + 2
        ElseIf saEncoding(lOutIdx) = ">" Then
            pvOutputBars pvGetEan128Bars("CODEC", ucsEseSystemB), lPos, lStep, " " ' , "·"
        ElseIf saEncoding(lOutIdx) = "<" Then
            pvOutputBars pvGetEan128Bars("CODEB", ucsEseSystemC), lPos, lStep, " " ' , "•"
        End If
    Next
    pvOutputBars pvGetEan128Bars(sCheckSum, ucsEseSystemC), lPos, lStep, " " ' , sCheckSum
    pvOutputBars pvGetEan128Bars("STOP", ucsEseSystemB), lPos, lStep, " "
    pvOutputBars pvGetEan128Bars("TERMINATOR", ucsEseSystemB), lPos, lStep, " "
    If m_bShowDigits Then
        hPrevFont = SelectObject(m_hDC, m_oFont.hFont)
        rcLetter.Left = ((m_rcBound.Right - m_rcBound.Left) - 7 * lStep * Len(sBarCode)) \ 2
        rcLetter.Top = m_rcBound.Bottom - 10 * lStep
        rcLetter.Bottom = m_rcBound.Bottom
        For lIdx = 1 To Len(sBarCode)
            rcLetter.Right = rcLetter.Left + 7 * lStep
            Call DrawText(m_hDC, Mid$(sBarCode, lIdx, 1), 1, rcLetter, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE)
            rcLetter.Left = rcLetter.Right + 0 * lStep
        Next
        Call SelectObject(m_hDC, hPrevFont)
    End If
    
    Set GetEan128BarCode = pvMetaToPicture()
    Exit Function
EH:
    PrintError FUNC_NAME & "(sBarCode=" & sBarCode & ")"
    Resume Next
End Function

'= private ===============================================================

Private Function pvUpcEBarCode(sBarCode As String, sSuffix As String) As StdPicture
    Const FUNC_NAME     As String = "pvUpcEBarCode"
    Dim sSystem         As String
    Dim lStep           As Long
    Dim lPos            As Long
    Dim lIdx            As Long
    Dim vBars           As Variant
    Dim sOrigCode       As String
    Dim sOrigSuffix     As String
    
    On Error GoTo EH
    '--- init loca vars
    sOrigCode = sBarCode
    sOrigSuffix = sSuffix
    sBarCode = pvLeftDigits(sBarCode, 8)
    If LenB(sSuffix) <> 0 Then
        If Len(sSuffix) > 2 Then
            sSuffix = pvLeftDigits(sSuffix, 5)
        Else
            sSuffix = pvLeftDigits(sSuffix, 2)
        End If
    End If
    m_clrCurrent = IIf(sOrigCode = sBarCode And sOrigSuffix = sSuffix, vbBlack, m_clrInvalid)
    lStep = 100
    lPos = 9 * lStep
    sSystem = Mid$(STR_UPCE_SYSTEM, 6 * Mid$(sBarCode, 2, 1) + 1, 6)
    If Mid$(sBarCode, 1, 1) <> "0" Then
        sSystem = Replace(Replace(Replace(sSystem, "0", "2"), "1", "0"), "2", "1")
    End If
    '--- output
    pvCreateMeta (70 + IIf(LenB(sSuffix) <> 0, 11 + Len(sSuffix) * 9, 0)) * lStep, 55 * lStep
    pvOutputBars Empty, 2 * lStep, lStep, Left$(sBarCode, 1), ucsValBottom
    pvOutputBars pvGetBars("!"), lPos, lStep
    For lIdx = 2 To 7
        vBars = pvGetBars(Mid$(sBarCode, lIdx, 1), IIf(Mid$(sSystem, lIdx - 1, 1) = "1", ucsEdtLeftOdd, ucsEdtLeftEven))
        pvOutputBars vBars, lPos, lStep, IIf(lIdx > 1, Mid$(sBarCode, lIdx, 1), vbNullString)
    Next
    pvOutputBars pvGetBars("#"), lPos, lStep
    pvOutputBars pvGetBars("$"), lPos, lStep
    pvOutputBars Empty, lPos, lStep, Mid$(sBarCode, 8, 1), ucsValBottom
    '--- suffix
    lPos = lPos + 3 * lStep
    pvOutputSuffix sSuffix, lPos, lStep
    '--- success
    Set pvUpcEBarCode = pvMetaToPicture()
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Private Function pvUpcABarCode(sBarCode As String, sSuffix As String) As StdPicture
    Const FUNC_NAME     As String = "pvUpcABarCode"
    Dim sSystem         As String
    Dim lIdx            As Long
    Dim lCheck          As Long
    Dim lPos            As Long
    Dim lStep           As Long
    Dim vBars           As Variant
    Dim sOrigCode       As String
    Dim sOrigSuffix     As String
    
    On Error GoTo EH
    '--- init local vars
    sOrigCode = sBarCode
    sOrigSuffix = sSuffix
    sSystem = Split(STR_EAN13_SYSTEM, "|")(0)
    sBarCode = pvLeftDigits(sBarCode, 12)
    If LenB(sSuffix) <> 0 Then
        If Len(sSuffix) > 2 Then
            sSuffix = pvLeftDigits(sSuffix, 5)
        Else
            sSuffix = pvLeftDigits(sSuffix, 2)
        End If
    End If
    m_clrCurrent = IIf(sOrigCode = sBarCode And sOrigSuffix = sSuffix, vbBlack, m_clrInvalid)
    lStep = 100
    lPos = 9 * lStep
    '--- calc check digit
    For lIdx = 1 To 11
        lCheck = lCheck + C_Lng(Mid$(sBarCode, lIdx, 1)) * IIf(lIdx Mod 2 = 1, 3, 1)
    Next
    Mid$(sBarCode, 12, 1) = C_Str((10 - (lCheck Mod 10)) Mod 10)
    '--- output
    pvCreateMeta (113 + IIf(LenB(sSuffix) <> 0, 11 + Len(sSuffix) * 9, 0)) * lStep, 55 * lStep
    pvOutputBars Empty, 2 * lStep, lStep, Left$(sBarCode, 1), ucsValBottom
    pvOutputBars pvGetBars("!"), lPos, lStep
    For lIdx = 1 To 6
        vBars = pvGetBars(Mid$(sBarCode, lIdx, 1), IIf(Mid$(sSystem, lIdx, 1) = "1", ucsEdtLeftOdd, ucsEdtLeftEven))
        pvOutputBars vBars, lPos, lStep, IIf(lIdx > 1, Mid$(sBarCode, lIdx, 1), vbNullString)
    Next
    pvOutputBars pvGetBars("*"), lPos, lStep
    For lIdx = 1 To 6
        vBars = pvGetBars(Mid$(sBarCode, lIdx + 6, 1), ucsEdtRight)
        pvOutputBars vBars, lPos, lStep, IIf(lIdx < 6, Mid$(sBarCode, lIdx + 6, 1), vbNullString)
    Next
    pvOutputBars pvGetBars("!"), lPos, lStep
    pvOutputBars Empty, lPos, lStep, Mid$(sBarCode, 12, 1), ucsValBottom
    '--- suffix
    lPos = lPos + 3 * lStep
    pvOutputSuffix sSuffix, lPos, lStep
    '--- success
    Set pvUpcABarCode = pvMetaToPicture()
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Private Function pvEan8BarCode(sBarCode As String, sSuffix As String) As StdPicture
    Const FUNC_NAME     As String = "pvEan8BarCode"
    Dim lIdx            As Long
    Dim lCheck          As Long
    Dim lPos            As Long
    Dim lStep           As Long
    Dim vBars           As Variant
    Dim sOrigCode       As String
    Dim sOrigSuffix     As String
    
    On Error GoTo EH
    '--- init local vars
    sOrigCode = sBarCode
    sOrigSuffix = sSuffix
    sBarCode = pvLeftDigits(sBarCode, 8)
    If LenB(sSuffix) <> 0 Then
        If Len(sSuffix) > 2 Then
            sSuffix = pvLeftDigits(sSuffix, 5)
        Else
            sSuffix = pvLeftDigits(sSuffix, 2)
        End If
    End If
    lStep = 100
    lPos = 7 * lStep
    '--- calc check digit
    For lIdx = 1 To 7
        lCheck = lCheck + C_Lng(Mid$(sBarCode, lIdx, 1)) * IIf(lIdx Mod 2 = 1, 3, 1)
    Next
    Mid$(sBarCode, 8, 1) = C_Str((10 - (lCheck Mod 10)) Mod 10)
    m_clrCurrent = IIf(sOrigCode = sBarCode And sOrigSuffix = sSuffix, vbBlack, m_clrInvalid)
    '--- output
    pvCreateMeta (81 + IIf(LenB(sSuffix) <> 0, 11 + Len(sSuffix) * 9, 0)) * lStep, 55 * lStep
    pvOutputBars pvGetBars("!"), lPos, lStep
    For lIdx = 1 To 4
        vBars = pvGetBars(Mid$(sBarCode, lIdx, 1), ucsEdtLeftOdd)
        pvOutputBars vBars, lPos, lStep, Mid$(sBarCode, lIdx, 1)
    Next
    pvOutputBars pvGetBars("*"), lPos, lStep
    For lIdx = 1 To 4
        vBars = pvGetBars(Mid$(sBarCode, lIdx + 4, 1), ucsEdtRight)
        pvOutputBars vBars, lPos, lStep, Mid$(sBarCode, lIdx + 4, 1)
    Next
    pvOutputBars pvGetBars("!"), lPos, lStep
    '--- suffix
    lPos = lPos + 9 * lStep
    pvOutputSuffix sSuffix, lPos, lStep
    '--- success
    Set pvEan8BarCode = pvMetaToPicture()
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Private Function pvEan13BarCode(sBarCode As String, sSuffix As String) As StdPicture
    Const FUNC_NAME     As String = "pvEan13BarCode"
    Dim sSystem         As String
    Dim lIdx            As Long
    Dim lCheck          As Long
    Dim lPos            As Long
    Dim lStep           As Long
    Dim vBars           As Variant
    Dim sOrigCode       As String
    Dim sOrigSuffix     As String
    
    On Error GoTo EH
    '--- init local vars
    sOrigCode = sBarCode
    sOrigSuffix = sSuffix
    sSystem = Split(STR_EAN13_SYSTEM, "|")(C_Lng(Left$(sBarCode, 1)))
    sBarCode = pvLeftDigits(sBarCode, 13)
    If LenB(sSuffix) <> 0 Then
        If Len(sSuffix) > 2 Then
            sSuffix = pvLeftDigits(sSuffix, 5)
        Else
            sSuffix = pvLeftDigits(sSuffix, 2)
        End If
    End If
    lStep = 100
    lPos = 11 * lStep
    '--- calc check digit
    For lIdx = 1 To 12
        lCheck = lCheck + C_Lng(Mid$(sBarCode, lIdx, 1)) * IIf(lIdx Mod 2 = 1, 1, 3)
    Next
    Mid$(sBarCode, 13, 1) = C_Str((10 - (lCheck Mod 10)) Mod 10)
    m_clrCurrent = IIf(sOrigCode = sBarCode And sOrigSuffix = sSuffix, vbBlack, m_clrInvalid)
    '--- output
    pvCreateMeta (113 + IIf(LenB(sSuffix) <> 0, 11 + Len(sSuffix) * 9, 0)) * lStep, 55 * lStep
    pvOutputBars Empty, 1 * lStep, lStep, Left$(sBarCode, 1)
    pvOutputBars pvGetBars("!"), lPos, lStep
    For lIdx = 1 To 6
        vBars = pvGetBars(Mid$(sBarCode, lIdx + 1, 1), IIf(Mid$(sSystem, lIdx, 1) = "1", ucsEdtLeftOdd, ucsEdtLeftEven))
        pvOutputBars vBars, lPos, lStep, Mid$(sBarCode, lIdx + 1, 1)
    Next
    pvOutputBars pvGetBars("*"), lPos, lStep
    For lIdx = 1 To 6
        vBars = pvGetBars(Mid$(sBarCode, lIdx + 7, 1), ucsEdtRight)
        pvOutputBars vBars, lPos, lStep, Mid$(sBarCode, lIdx + 7, 1)
    Next
    pvOutputBars pvGetBars("!"), lPos, lStep
    pvOutputBars Empty, lPos, lStep, ">"
    '--- suffix
    lPos = lPos + 3 * lStep
    pvOutputSuffix sSuffix, lPos, lStep
    '--- success
    Set pvEan13BarCode = pvMetaToPicture()
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Private Sub pvOutputSuffix(sSuffix As String, lPos As Long, ByVal lStep As Long)
    Const FUNC_NAME     As String = "pvOutputSuffix"
    Dim sSystem         As String
    Dim vBars           As Variant
    Dim lIdx            As Long
    Dim lCheck          As Long
    
    On Error GoTo EH
    If LenB(sSuffix) = 0 Then
        Exit Sub
    End If
    If Len(sSuffix) = 2 Then
        sSystem = Split(STR_SUF2_SYSTEM, "|")(C_Lng(sSuffix) Mod 4)
    Else
        For lIdx = 1 To 5
            lCheck = lCheck + C_Lng(Mid$(sSuffix, lIdx, 1)) * IIf(lIdx Mod 2 = 1, 3, 9)
        Next
        sSystem = Split(STR_SUF5_SYSTEM, "|")(lCheck Mod 10)
    End If
    For lIdx = 1 To Len(sSuffix)
        pvOutputBars pvGetBars(IIf(lIdx = 1, "?", "/")), lPos, lStep, , ucsValTop
        vBars = pvGetBars(Mid$(sSuffix, lIdx, 1), IIf(Mid$(sSystem, lIdx, 1) = "1", ucsEdtLeftOdd, ucsEdtLeftEven))
        pvOutputBars vBars, lPos, lStep, Mid$(sSuffix, lIdx, 1), ucsValTop
    Next
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Function pvGetEan128Bars(sDigit As String, Optional ByVal eType As UcsEan128SystemEnum = ucsEseSystemB) As Variant
    Dim sChar           As String
    Dim vResult         As Variant
    Dim lPos            As Long
    Dim lIdx            As Long
    Dim lCount          As Long
    
    If sDigit = "STOP" Then
        sChar = "11000111010"
    ElseIf sDigit = "TERMINATOR" Then
        sChar = "11"
    ElseIf sDigit = "STARTB" Then
        sChar = "11010010000"
    ElseIf sDigit = "CODEC" Then
        sChar = "10111011110"
    ElseIf sDigit = "STARTC" Then
        sChar = "11010011100"
    ElseIf sDigit = "CODEB" Then
        sChar = "10111101110"
    ElseIf eType = ucsEseSystemB Then
        lCount = InStr(1, STR_EAN128_B_CHARACTERS, sDigit)
        sChar = Mid$(STR_EAN128_SYSTEM, 1 + 11 * lCount, 11)
    Else ' If eType = ucsEseSystemC Then
        sChar = Mid$(STR_EAN128_SYSTEM, 1 + 11 * C_Lng(sDigit), 11)
    End If
    vResult = EmptyVariantArray
    lPos = 1
    Do While lPos <= Len(sChar)
        For lIdx = lPos + 1 To Len(sChar)
            If Mid$(sChar, lPos, 1) <> Mid$(sChar, lIdx, 1) Then
                Exit For
            End If
        Next
        ReDim Preserve vResult(0 To UBound(vResult) + 1) As Variant
        vResult(UBound(vResult)) = IIf(Mid$(sChar, lPos, 1) = "0", lPos - lIdx, lIdx - lPos)
        lPos = lIdx
    Loop
    pvGetEan128Bars = vResult
End Function

Private Function pvGetBars(sDigit As String, Optional ByVal eType As UcsEanDigitTypeEnum = ucsEdtLeftOdd) As Variant
    Const FUNC_NAME     As String = "pvGetBars"
    Dim sChar           As String
    Dim vResult         As Variant
    Dim lPos            As Long
    Dim lIdx            As Long
    
    On Error GoTo EH
    If sDigit = "!" Then '--- guard bars
        sChar = "101"
    ElseIf sDigit = "*" Then '-- separator
        sChar = "01010"
    ElseIf sDigit = "?" Then '-- suffix guard bars
        sChar = "1011"
    ElseIf sDigit = "/" Then '-- suffix separator
        sChar = "01"
    ElseIf sDigit = "#" Then '--- centar guard bars
        sChar = "01010"
    ElseIf sDigit = "$" Then '--- right guard bar
        sChar = "1"
    Else
        If Not IsNumeric(sDigit) Then
            sDigit = 0
        End If
        sChar = Split(STR_EAN_CHARACTER_SET, "|")(C_Lng(sDigit))
        If eType = ucsEdtRight Then
            sChar = Replace(Replace(sChar, "0", "2"), "1", "0")
        ElseIf eType = ucsEdtLeftEven Then
            sChar = StrReverse(Replace(Replace(sChar, "0", "2"), "1", "0"))
        End If
    End If
    vResult = EmptyVariantArray
    lPos = 1
    Do While lPos <= Len(sChar)
        For lIdx = lPos + 1 To Len(sChar)
            If Mid$(sChar, lPos, 1) <> Mid$(sChar, lIdx, 1) Then
                Exit For
            End If
        Next
        ReDim Preserve vResult(0 To UBound(vResult) + 1) As Variant
        vResult(UBound(vResult)) = IIf(Mid$(sChar, lPos, 1) = "0", lPos - lIdx, lIdx - lPos)
        lPos = lIdx
    Loop
    pvGetBars = vResult
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Private Sub pvCreateMeta(ByVal lWidth As Long, ByVal lHeight As Long)
    m_hDC = CreateEnhMetaFile(0, vbNullString, 0, vbNullString)
    Call SetTextColor(m_hDC, m_clrCurrent)
    Call SetBkMode(m_hDC, TRANSPARENT)
    m_rcBound.Right = lWidth
    m_rcBound.Bottom = Round(lHeight * m_dblHeightPercent / 100#)
    Call FillRect(m_hDC, m_rcBound, GetStockObject(WHITE_BRUSH))
End Sub

Private Sub pvOutputBars( _
            vBars As Variant, _
            lStart As Long, _
            ByVal lStep As Long, _
            Optional sDigit As String, _
            Optional lVertAlign As UcsVertAlignEnum = ucsValBottom)
    Const FUNC_NAME     As String = "pvOutputBars"
    Dim vElem           As Variant
    Dim rc              As RECT
    Dim rcLetter        As RECT
    Dim hPrevFont       As Long
    Dim lBottom         As Long
    Dim hBrush          As Long
    
    On Error GoTo EH
    rc.Top = 1 * lStep
    If lVertAlign = ucsValCenter Then
        lBottom = m_rcBound.Bottom - 20 * lStep
    ElseIf lVertAlign = ucsValTop Then
        If m_bShowDigits Then
            lBottom = m_rcBound.Bottom - 5 * lStep
        Else
            lBottom = m_rcBound.Bottom
        End If
        rc.Top = 9 * lStep
    ElseIf m_bShowDigits Then
        lBottom = m_rcBound.Bottom - 10 * lStep
    ElseIf m_bHangSeparators Then
        lBottom = m_rcBound.Bottom - 5 * lStep
    Else
        lBottom = m_rcBound.Bottom
    End If
    rcLetter.Left = lStart
    hBrush = CreateSolidBrush(m_clrCurrent)
    If IsArray(vBars) Then
        For Each vElem In vBars
            If vElem > 0 Then
                If m_bHangSeparators And LenB(sDigit) = 0 And lVertAlign = ucsValBottom Then
                    rc.Bottom = lBottom + 5 * lStep
                Else
                    rc.Bottom = lBottom
                End If
                rc.Left = lStart + m_lBleed
                rc.Right = lStart + vElem * lStep - m_lBleed
                Call FillRect(m_hDC, rc, hBrush)
            End If
            lStart = lStart + Abs(vElem) * lStep
        Next
    Else
        lStart = lStart + 7 * lStep
    End If
    If m_bShowDigits And LenB(sDigit) <> 0 Then
        rcLetter.Right = lStart
        If lVertAlign = ucsValTop Then
            rcLetter.Bottom = rc.Top
        Else
            rcLetter.Top = lBottom - 1 * lStep
            rcLetter.Bottom = m_rcBound.Bottom
        End If
        hPrevFont = SelectObject(m_hDC, m_oFont.hFont)
        Call DrawText(m_hDC, sDigit, Len(sDigit), rcLetter, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE)
        Call SelectObject(m_hDC, hPrevFont)
    End If
    Call DeleteObject(hBrush)
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Function pvMetaToPicture() As IPicture
    Dim uDesc           As PICTDESC
    Dim aGuid(0 To 3)   As Long
    
    '--- fill struct
    With uDesc
        .Size = Len(uDesc)
        .Type = vbPicTypeEMetafile
        .hBmpOrIcon = CloseEnhMetaFile(m_hDC)
    End With
    m_hDC = 0
    '--- Fill in magic IPicture GUID {7BF80980-BF32-101A-8BBB-00AA00300CAB}
    aGuid(0) = &H7BF80980
    aGuid(1) = &H101ABF32
    aGuid(2) = &HAA00BB8B
    aGuid(3) = &HAB0C3000
    '--- Create picture from bitmap handle
    Call OleCreatePictureIndirect(uDesc, aGuid(0), True, pvMetaToPicture)
End Function

Private Function pvLeftDigits(sText As String, ByVal lSize As Long) As String
    Dim lIdx            As Long
    Dim lAsc            As Long
    
    pvLeftDigits = Left$(sText & String$(lSize, "0"), lSize)
    For lIdx = 1 To Len(pvLeftDigits)
        lAsc = Asc(Mid$(pvLeftDigits, lIdx, 1))
        If lAsc < &H30 Or lAsc > &H39 Then
            Mid$(pvLeftDigits, lIdx, 1) = "0"
        End If
    Next
End Function

#If DebugMode Then
    Private Sub Class_Initialize()
        DebugInstanceInit MODULE_NAME, m_sDebugID, Me
    End Sub

    Private Sub Class_Terminate()
        DebugInstanceTerm MODULE_NAME, m_sDebugID
    End Sub
#End If
