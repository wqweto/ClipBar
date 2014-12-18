VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "ClipBar"
   ClientHeight    =   4520
   ClientLeft      =   180
   ClientTop       =   820
   ClientWidth     =   5680
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4520
   ScaleWidth      =   5680
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBarCode 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   3708
      Left            =   0
      ScaleHeight     =   3710
      ScaleWidth      =   5560
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   504
      Width           =   5556
   End
   Begin VB.TextBox txtCode 
      Height          =   288
      Left            =   84
      TabIndex        =   0
      Text            =   "1234567890128"
      Top             =   84
      Width           =   2952
   End
   Begin VB.Label labAbout 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ClipBar © 2011-2014 Unicontsoft"
      Height          =   200
      Left            =   3190
      TabIndex        =   2
      Top             =   80
      Width           =   3520
      WordWrap        =   -1  'True
   End
   Begin VB.Menu mnuMain 
      Caption         =   "File"
      Index           =   0
      Begin VB.Menu mnuFile 
         Caption         =   "Save"
         Index           =   0
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFile 
         Caption         =   "Copy"
         Index           =   1
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuFile 
         Caption         =   "Exit"
         Index           =   3
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Settings"
      Index           =   1
      Begin VB.Menu mnuSettings 
         Caption         =   "EAN-8/13"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu mnuSettings 
         Caption         =   "UPC-A/E"
         Index           =   1
      End
      Begin VB.Menu mnuSettings 
         Caption         =   "EAN-128"
         Index           =   2
      End
      Begin VB.Menu mnuSettings 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuSettings 
         Caption         =   "Digits"
         Index           =   4
      End
      Begin VB.Menu mnuSettings 
         Caption         =   "Separators"
         Index           =   5
      End
      Begin VB.Menu mnuSettings 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuSettings 
         Caption         =   "Bleed"
         Index           =   7
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'=========================================================================
' Constants and member variables
'=========================================================================

Private m_oBarCode          As cBarCode
Private m_lBleed            As Long
Private m_sFileName         As String

Private Enum UcsMenuIdx
    '--- file
    ucsMnuSave = 0
    ucsMnuCopy = 1
    ucsMnuExit = 3
    '--- settings
    ucsMnuEAN = 0
    ucsMnuUPC = 1
    ucsMnuEAN128 = 2
    ucsMnuDigits = 4
    ucsMnuSeparators = 5
    ucsMnuBleed = 7
    [_ucsMnuSettingsMin] = ucsMnuEAN
    [_ucsMnuSettingsMax] = ucsMnuSeparators
End Enum

'=========================================================================
' Control events
'=========================================================================

Private Sub Form_Load()
    Const FUNC_NAME     As String = "Form_Load"
    Dim lIdx            As Long
    Dim vSettings       As Variant
    
    On Error GoTo EH
    m_lBleed = 10
    vSettings = Split(GetSetting(App.Title, "Default", "Settings", "1 0 0 0 1 1"))
    For lIdx = [_ucsMnuSettingsMin] To [_ucsMnuSettingsMax]
        If At(vSettings, lIdx - [_ucsMnuSettingsMin]) = "1" Then
            mnuSettings_Click CInt(lIdx)
        End If
    Next
    mnuSettings_Click -1
    Exit Sub
EH:
    MsgBox Error, vbCritical
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    picBarCode.Width = ScaleWidth - 2 * picBarCode.Left
    picBarCode.Height = ScaleHeight - picBarCode.Top
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Const FUNC_NAME     As String = "Form_Unload"
    Dim lIdx            As Long
    Dim vSettings       As Variant
    
    On Error GoTo EH
    ReDim vSettings(0 To [_ucsMnuSettingsMax] - [_ucsMnuSettingsMin]) As String
    For lIdx = [_ucsMnuSettingsMin] To [_ucsMnuSettingsMax]
        vSettings(lIdx - [_ucsMnuSettingsMin]) = -mnuSettings(lIdx).Checked
    Next
    SaveSetting App.Title, "Default", "Settings", Join(vSettings)
    Exit Sub
EH:
    MsgBox Error, vbCritical
End Sub

Private Sub mnuFile_Click(Index As Integer)
    On Error GoTo EH
    Select Case Index
    Case ucsMnuSave
        If ShowOpenSaveDialog(m_sFileName, "WMF - Windows Metafile|*.wmf|All files (*.*)|*.*", , hWnd, "WMF", "Barcode image", ucsOsdSave) Then
            SavePicture picBarCode.Picture, m_sFileName
            MsgBox "Barcode successfully saved to " & m_sFileName, vbExclamation
        End If
    Case ucsMnuCopy
        Clipboard.Clear
        Clipboard.SetData picBarCode.Picture
        MsgBox "Barcode successfully copied to clipboard", vbExclamation
    Case ucsMnuExit
        Unload Me
    End Select
    Exit Sub
EH:
    MsgBox Error, vbCritical
End Sub

Private Sub mnuSettings_Click(Index As Integer)
    Dim sText           As String
    
    On Error GoTo EH
    Select Case Index
    Case ucsMnuEAN, ucsMnuUPC, ucsMnuEAN128
        mnuSettings(ucsMnuEAN).Checked = (Index = ucsMnuEAN)
        mnuSettings(ucsMnuUPC).Checked = (Index = ucsMnuUPC)
        mnuSettings(ucsMnuEAN128).Checked = (Index = ucsMnuEAN128)
    Case ucsMnuDigits, ucsMnuSeparators
        mnuSettings(Index).Checked = Not mnuSettings(Index).Checked
    Case ucsMnuBleed
        sText = InputBox("Enter ink bleed (in percents):", , m_lBleed)
        If StrPtr(sText) = 0 Then
            Exit Sub
        End If
        m_lBleed = C_Lng(sText)
        If m_lBleed > 99 Then
            m_lBleed = 99
        ElseIf m_lBleed < 0 Then
            m_lBleed = 0
        End If
    End Select
    Set m_oBarCode = New cBarCode
    m_oBarCode.Init m_lBleed, mnuSettings(ucsMnuSeparators).Checked, mnuSettings(ucsMnuDigits).Checked
    txtCode_Change
    Exit Sub
EH:
    MsgBox Error, vbCritical
End Sub

Private Sub txtCode_Change()
    On Error GoTo EH
    If m_oBarCode Is Nothing Then
        Exit Sub
    End If
    If mnuSettings(ucsMnuUPC).Checked Then
        Set picBarCode.Picture = m_oBarCode.GetUpcBarCode(txtCode.Text)
    ElseIf mnuSettings(ucsMnuEAN128).Checked Then
        Set picBarCode.Picture = m_oBarCode.GetEan128BarCode(txtCode.Text)
    Else
        Set picBarCode.Picture = m_oBarCode.GetEanBarCode(txtCode.Text)
    End If
    Exit Sub
EH:
    MsgBox Error, vbCritical
End Sub

