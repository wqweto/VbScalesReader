VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2316
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   3624
   LinkTopic       =   "Form1"
   ScaleHeight     =   2316
   ScaleWidth      =   3624
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   600
      Left            =   924
      TabIndex        =   0
      Top             =   1428
      Width           =   1608
   End
   Begin MSCommLib.MSComm MSComm1 
      Index           =   0
      Left            =   3024
      Top             =   0
      _ExtentX        =   974
      _ExtentY        =   974
      _Version        =   393216
      DTREnable       =   -1  'True
      BaudRate        =   2400
      InputMode       =   1
   End
   Begin VB.Label labInfo 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Waiting for device. . ."
      Height          =   264
      Left            =   504
      TabIndex        =   1
      Top             =   504
      Width           =   2532
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'=========================================================================
' API declares
'=========================================================================

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function IsBadReadPtr Lib "kernel32" (ByVal lp As Long, ByVal ucb As Long) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long

'=========================================================================
' Constants and member variables
'=========================================================================

'--- errors
Private Const ERR_CRITICAL          As String = "Неочаквана грешка: "
Private Const ERR_USER_CANCEL       As String = "Прекъсване от потребителя"
Private Const ERR_DEVICE_TIMEOUT    As String = "Изтекло време за отговор от устройство"
Private Const ERR_ZERO_WEIGHT       As String = "Липсва измерване"
'--- strings
Private Const STR_WAITING           As String = "Получава тегло от везната..."
Private Const STR_OVERLOAD          As String = "Претоварване"
Private Const STR_UNSTABLE          As String = "Текущо"
Private Const STR_STABLE            As String = "Измерено"
'--- numeric
Private Const MAX_RETRY             As Long = 3
Private Const DBL_COM_TIMEOUT       As Double = 3 '--- in seconds
Private Const DBL_MAX_WEIGHT        As Double = 100
Private Const DBL_EPSILON           As Double = 0.000001


Private m_sLastError            As String
Private m_uData()               As UcsScaleDataType
Private m_bCancel               As Boolean

Private Type UcsScaleDataType
    Protocol            As UcsScaleProtocolEnum
    Received            As String
    Request             As String
    Response            As String
    Status              As UcsScaleStatusEnum
    Weight              As Double
End Type

Private Enum UcsScaleProtocolEnum
    ucsScaProtocolCas
    ucsScaProtocolElicom
    ucsScaProtocolDibal
    ucsScaProtocolMettler
    ucsScaProtocolDelmac
    ucsScaProtocolBimco
End Enum

Private Enum UcsScaleStatusEnum
    ucsScaStatusStable
    ucsScaStatusUnstable
    ucsScaStatusOverload
    ucsScaStatusUnderload
End Enum

Private Enum UcsControlSymbols
    SOH = 1
    STX = 2
    ETX = 3
    EOT = 4
    ENQ = 5
    ACK = 6
    NAK = &H15
    DC1 = &H11
    BMK_REQ = &H36
    ELI_REQ = &HAA
    ELI_UNS = &HBB
End Enum

Private Enum UcsParseResultEnum
    ucsScaResultContinue = 0     '--- must be 0
    ucsScaResultHasResult
    ucsScaResultRetrySend
    ucsScaResultRetryZero
End Enum

'=========================================================================
' Properties
'=========================================================================

Private Property Get pvInfo() As String
    pvInfo = labInfo.Caption
End Property

Private Property Let pvInfo(sValue As String)
    labInfo.Caption = sValue
    Refresh
End Property

'=========================================================================
' Control events
'=========================================================================

Private Sub Command1_Click()
    m_bCancel = True
End Sub

Private Sub Form_Load()
    ReDim m_uData(0 To 1) As UcsScaleDataType
    m_uData(0).Protocol = ucsScaProtocolBimco
    If Not pvOpenPort(0, 1, 9600, m_sLastError) Then
        MsgBox m_sLastError, vbExclamation
    End If
    m_uData(1).Protocol = ucsScaProtocolDelmac
    If Not pvOpenPort(1, 2, 1200, m_sLastError) Then
        MsgBox m_sLastError, vbExclamation
    End If
End Sub

Private Sub Form_Activate()
    Dim sError          As String
    Dim lResult         As Long
    
    pvInfo = STR_WAITING
    Do
        If Not pvReadWeight(DBL_COM_TIMEOUT, MAX_RETRY, False, lResult, sError) Then
            pvInfo = sError
            If Not Visible Then
                Unload Me
            End If
            Exit Sub
        End If
        With m_uData(lResult)
            Select Case .Status
            Case ucsScaStatusOverload
                pvInfo = STR_OVERLOAD
            Case ucsScaStatusUnstable, ucsScaStatusUnderload
                pvInfo = STR_UNSTABLE & " " & Format(.Weight, "0.000")
            Case ucsScaStatusStable
                pvInfo = STR_STABLE & " " & Format(.Weight, "0.000")
                If .Weight > 0 And .Weight < DBL_MAX_WEIGHT Then
                    Exit Do
                End If
            End Select
        End With
    Loop
End Sub

Private Sub Form_Click()
    Form_Activate
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    m_bCancel = True
End Sub

'=========================================================================
' Methods
'=========================================================================

Private Function pvOpenPort(ByVal lIndex As Long, ByVal lPort As Long, ByVal lSpeed As Long, sError As String) As Boolean
    On Error GoTo EH
    If MSComm1.UBound < lIndex Then
        Load MSComm1(lIndex)
    End If
    With MSComm1(lIndex)
        If .PortOpen Then
            .PortOpen = False
        End If
        .CommPort = lPort
        .Settings = lSpeed & ",n,8,1"
        .RThreshold = 1
        .PortOpen = True
        If .PortOpen Then
            '--- success
            pvOpenPort = True
        End If
    End With
QH:
    Exit Function
EH:
    sError = ERR_CRITICAL & Err.Description
    Resume QH
End Function

Private Function pvReadWeight( _
            ByVal dblTimeout As Double, _
            ByVal lMaxRetry As Long, _
            ByVal bAllowZero As Boolean, _
            lResult As Long, _
            sError As String) As Boolean
    Dim lRetry          As Long
    Dim dblTimerEx      As Double
    Dim lIdx            As Long
    Dim eResult         As UcsParseResultEnum
    Dim bHasZero        As Boolean
    
    On Error GoTo EH
    m_bCancel = False
    For lRetry = 1 To lMaxRetry
        dblTimerEx = TimerEx
        For lIdx = 0 To UBound(m_uData)
            Select Case m_uData(lIdx).Protocol
            Case ucsScaProtocolCas
                m_uData(lIdx).Request = Chr$(ENQ)
            Case ucsScaProtocolElicom
                m_uData(lIdx).Request = Chr$(ELI_REQ)
            Case ucsScaProtocolDibal
                m_uData(lIdx).Request = Chr$(STX)
            Case ucsScaProtocolMettler
                m_uData(lIdx).Request = "SI" & vbCrLf
            Case ucsScaProtocolDelmac
                m_uData(lIdx).Request = vbNullString
            Case ucsScaProtocolBimco
                m_uData(lIdx).Request = Chr$(BMK_REQ)
            End Select
            m_uData(lIdx).Received = vbNullString
            If LenB(m_uData(lIdx).Request) <> 0 Then
                pvLogDataDump "Out(" & MSComm1(lIdx).CommPort & ")", m_uData(lIdx).Request
                MSComm1(lIdx).Output = m_uData(lIdx).Request
            End If
        Next
        Do While TimerEx < dblTimerEx + dblTimeout
            DoEvents
            If m_bCancel Then
                sError = ERR_USER_CANCEL
                GoTo QH
            End If
            For lIdx = 0 To UBound(m_uData)
                Select Case m_uData(lIdx).Protocol
                Case ucsScaProtocolCas
                    eResult = pvParseCasResponse(m_uData(lIdx), bAllowZero)
                Case ucsScaProtocolElicom
                    eResult = pvParseElicomResponse(m_uData(lIdx), bAllowZero)
                Case ucsScaProtocolDibal
                    eResult = pvParseDibalResponse(m_uData(lIdx), bAllowZero)
                Case ucsScaProtocolMettler
                    eResult = pvParseMettlerResponse(m_uData(lIdx), bAllowZero)
                Case ucsScaProtocolDelmac
                    eResult = pvParseDelmacResponse(m_uData(lIdx), bAllowZero)
                Case ucsScaProtocolBimco
                    eResult = pvParseBimcoResponse(m_uData(lIdx), bAllowZero)
                Case Else
                    eResult = ucsScaResultContinue
                End Select
                Select Case eResult
                Case ucsScaResultHasResult
                    lResult = lIdx
                    '--- success
                    pvReadWeight = True
                    GoTo QH
                Case ucsScaResultRetrySend, ucsScaResultRetryZero
                    If eResult = ucsScaResultRetryZero Then
                        bHasZero = True
                    End If
                    m_uData(lIdx).Received = vbNullString
                    If LenB(m_uData(lIdx).Request) <> 0 Then
                        pvLogDataDump "Out(" & MSComm1(lIdx).CommPort & ")", m_uData(lIdx).Request
                        MSComm1(lIdx).Output = m_uData(lIdx).Request
                    End If
                End Select
            Next
        Loop
    Next
    sError = IIf(bHasZero, ERR_ZERO_WEIGHT, ERR_DEVICE_TIMEOUT)
QH:
    Exit Function
EH:
    sError = ERR_CRITICAL & Err.Description
    Resume QH
End Function

Private Function pvParseCasResponse(uData As UcsScaleDataType, ByVal bAllowZero As Boolean) As UcsParseResultEnum
    Dim lStart          As Long
    Dim lEnd            As Long
    Dim lSign           As Long
    
    With uData
        If InStr(.Received, Chr$(NAK)) > 0 Then
            .Request = Chr$(ENQ)
            pvParseCasResponse = ucsScaResultRetrySend
        ElseIf InStr(.Received, Chr$(ACK)) > 0 And .Request = Chr$(ENQ) Then
            .Request = Chr$(DC1)
            pvParseCasResponse = ucsScaResultRetrySend
        Else
            lStart = InStr(.Received, Chr$(SOH) & Chr$(STX))
            lEnd = InStr(.Received, Chr$(ETX) & Chr$(EOT))
            If lStart > 0 And lEnd > lStart + 2 Then
                .Response = Replace(Mid$(.Received, lStart + 2, lEnd - lStart - 2), Chr$(ACK), vbNullString)
                .Status = ucsScaStatusStable
                .Weight = 0
                Select Case Left$(.Response, 1)
                Case "F"
                    .Status = ucsScaStatusOverload
                Case "U"
                    .Status = ucsScaStatusUnstable
                End Select
                lSign = IIf(Mid$(.Response, 2, 1) = "-", -1, 1)
                .Weight = lSign * Val(Mid$(.Response, 3))
                If .Status = ucsScaStatusStable And Abs(.Weight) < DBL_EPSILON And Not bAllowZero Then
                    .Request = Chr$(ENQ)
                    pvParseCasResponse = ucsScaResultRetryZero
                Else
                    If lSign < 0 Then
                        .Status = ucsScaStatusUnderload
                    End If
                    '--- success
                    pvParseCasResponse = ucsScaResultHasResult
                End If
            ElseIf lEnd > 0 Then
                .Received = Mid$(.Received, lEnd + 2)
            End If
        End If
    End With
End Function

Private Function pvParseElicomResponse(uData As UcsScaleDataType, ByVal bAllowZero As Boolean) As UcsParseResultEnum
    Dim lIdx            As Long
    Dim lChar           As Long
    Dim lSum            As Long
    
    With uData
        If InStr(.Received, Chr$(ELI_UNS)) > 0 Then
            If Not bAllowZero Then
                pvParseElicomResponse = ucsScaResultRetryZero
            Else
                .Status = ucsScaStatusUnstable
                '--- success
                pvParseElicomResponse = ucsScaResultHasResult
            End If
        ElseIf Len(.Received) >= 4 Then
            .Response = vbNullString
            For lIdx = 1 To 3
                lChar = Asc(Mid$(.Received, lIdx, 1))
                lSum = (lSum + lChar) And &HFF&
                .Response = .Response & Right$("0" & Hex$(lChar), 2)
            Next
            If lSum <> Asc(Mid$(.Received, 4, 1)) Then
                pvParseElicomResponse = ucsScaResultRetrySend
            Else
                .Status = ucsScaStatusStable
                .Weight = Val(.Response) / 1000#
                '--- success
                pvParseElicomResponse = ucsScaResultHasResult
            End If
        End If
    End With
End Function

Private Function pvParseDibalResponse(uData As UcsScaleDataType, ByVal bAllowZero As Boolean) As UcsParseResultEnum
    Dim lEnd            As Long

    With uData
        If InStr(.Received, Chr$(ACK)) > 0 Then
            .Request = "10" & vbCrLf
            pvParseDibalResponse = ucsScaResultRetrySend
        ElseIf InStr(.Received, Chr$(STX)) > 0 Then
            .Request = Chr$(ACK)
            pvParseDibalResponse = ucsScaResultRetrySend
        Else
            lEnd = InStr(.Received, " ")
            If lEnd > 0 Then
                .Response = Mid$(.Received, 1, lEnd - 1)
                If Not IsNumeric(.Response) And Not bAllowZero Then
                    .Request = Chr$(STX)
                    pvParseDibalResponse = ucsScaResultRetryZero
                Else
                    .Status = IIf(IsNumeric(.Response), ucsScaStatusStable, ucsScaStatusUnstable)
                    .Weight = Val(.Response) / 1000#
                    If .Status = ucsScaStatusStable And .Weight < -DBL_EPSILON Then
                        .Status = ucsScaStatusUnderload
                    End If
                    '--- success
                    pvParseDibalResponse = ucsScaResultHasResult
                End If
            End If
        End If
    End With
End Function

Private Function pvParseMettlerResponse(uData As UcsScaleDataType, ByVal bAllowZero As Boolean) As UcsParseResultEnum
    Dim lStart          As Long
    Dim lEnd            As Long
    Dim lSign           As Long

    With uData
        lStart = InStr(.Received, "S ")
        lEnd = InStr(.Received, " Kg")
        If lStart > 0 And lEnd > lStart + 2 Then
            .Response = Mid$(.Received, lStart + 2, lEnd - lStart - 2)
            .Status = ucsScaStatusStable
            .Weight = 0
            Select Case Left$(.Response, 1)
            Case "X"
                .Status = ucsScaStatusOverload
            Case "D"
                .Status = ucsScaStatusUnstable
            End Select
            lSign = IIf(Mid$(.Response, 3, 1) = "-", -1, 1)
            .Weight = lSign * Val(Mid$(.Response, 4))
            If .Status = ucsScaStatusStable And Abs(.Weight) < DBL_EPSILON And Not bAllowZero Then
                pvParseMettlerResponse = ucsScaResultRetryZero
            Else
                If lSign < 0 Then
                    .Status = ucsScaStatusUnderload
                End If
                '--- success
                pvParseMettlerResponse = ucsScaResultHasResult
            End If
        ElseIf lEnd > 0 Then
            .Received = Mid$(.Received, lEnd + 3)
        End If
    End With
End Function

Private Function pvParseDelmacResponse(uData As UcsScaleDataType, ByVal bAllowZero As Boolean) As UcsParseResultEnum
    Const FLAG_UNDRLOAD As Long = &H80
    Const FLAG_OVERLOAD As Long = &H40
    Const FLAG_STABLE   As Long = &H10
    Const FLAG_ZERO     As Long = 8
    Const MASK_SCALE    As Long = 7
    Const NUM_DIGITS    As Long = 5
    Dim lStart          As Long
    Dim lEnd            As Long
    Dim lControl        As Long
    
    With uData
        lStart = InStr(.Received, Chr$(STX))
        lEnd = InStr(.Received, Chr$(ETX))
        If lStart > 0 And lEnd > lStart + 1 Then
            .Response = Mid$(.Received, lStart + 1, lEnd - lStart - 1)
            lControl = Asc(Right$(.Response, 1))
            If (lControl And FLAG_ZERO) <> 0 And Not bAllowZero Then
                pvParseDelmacResponse = ucsScaResultRetryZero
            Else
                If (lControl And FLAG_UNDRLOAD) <> 0 Then
                    .Status = ucsScaStatusUnderload
                ElseIf (lControl And FLAG_OVERLOAD) <> 0 Then
                    .Status = ucsScaStatusOverload
                ElseIf (lControl And FLAG_STABLE) <> 0 Then
                    .Status = ucsScaStatusStable
                Else
                    .Status = ucsScaStatusUnstable
                End If
                .Weight = Val(.Response) / (10 ^ (NUM_DIGITS - (lControl And MASK_SCALE)))
                '--- success
                pvParseDelmacResponse = ucsScaResultHasResult
            End If
        ElseIf lEnd > 0 Then
            .Received = Mid$(.Received, lEnd + 1)
        End If
    End With
End Function

Private Function pvParseBimcoResponse(uData As UcsScaleDataType, ByVal bAllowZero As Boolean) As UcsParseResultEnum
    Const IDX_WIEGHT    As Long = 0
    Const IDX_POWER10   As Long = 3
    Const IDX_STATUS    As Long = 4
    Const IDX_ERROR     As Long = 5
    Const IDX_CRC       As Long = 6
    Const FLAG_STABLE   As Long = 1
    Const FLAG_UNDRLOAD As Long = 4
    Dim lStart          As Long
    Dim lIdx            As Long
    Dim lSum            As Long
    Dim baRecv()        As Byte
    
    With uData
        lStart = Len(.Received) - 7
        If lStart > 0 Then
            .Response = Mid$(.Received, lStart)
            baRecv = StrConv(.Response, vbFromUnicode)
            For lIdx = 0 To IDX_CRC - 1
                lSum = lSum + baRecv(lIdx)
            Next
            If lSum <> (baRecv(IDX_CRC) * &H100& Or baRecv(IDX_CRC + 1)) Or baRecv(IDX_POWER10) > 10 Then
                pvParseBimcoResponse = ucsScaResultRetrySend
            Else
                If baRecv(IDX_ERROR) <> 0 Then
                    .Status = ucsScaStatusOverload
                ElseIf (baRecv(IDX_STATUS) And FLAG_STABLE) <> 0 Then
                    .Status = ucsScaStatusStable
                ElseIf (baRecv(IDX_STATUS) And FLAG_UNDRLOAD) = 0 Then
                    .Status = ucsScaStatusUnderload
                Else
                    .Status = ucsScaStatusUnstable
                End If
                .Weight = (IIf(baRecv(IDX_WIEGHT) And &H80, &HFF000000, 0) _
                    Or baRecv(IDX_WIEGHT) * &H10000 _
                    Or baRecv(IDX_WIEGHT + 1) * &H100& _
                    Or baRecv(IDX_WIEGHT + 2)) / (10 ^ baRecv(IDX_POWER10))
                If .Status = ucsScaStatusStable And Abs(.Weight) < DBL_EPSILON And Not bAllowZero Then
                    pvParseBimcoResponse = ucsScaResultRetryZero
                Else
                    '--- success
                    pvParseBimcoResponse = ucsScaResultHasResult
                End If
            End If
        End If
    End With
End Function

Private Sub MSComm1_OnComm(Index As Integer)
    Dim sInput          As String
    
    With MSComm1(Index)
        If .CommEvent = comEvReceive Then
            If .InputMode = comInputModeText Then
                sInput = .Input
            Else
                sInput = StrConv(.Input, vbUnicode)
            End If
            pvLogDataDump "In(" & .CommPort & ")", sInput
            m_uData(Index).Received = m_uData(Index).Received & sInput
        End If
    End With
End Sub

Private Sub pvLogDataDump(sType As String, sData As String)
    Dim baData()        As Byte
    
'    Debug.Print sType & ":", "[" & vData & "]", Timer
    Debug.Print sType & ":", Timer
    baData = StrConv(sData, vbFromUnicode)
    Debug.Print DesignDumpArray(baData);
End Sub

'== shared helpers ========================================================

Public Property Get TimerEx() As Double
    Dim cFreq           As Currency
    Dim cValue          As Currency
    
    Call QueryPerformanceFrequency(cFreq)
    Call QueryPerformanceCounter(cValue)
    TimerEx = cValue / cFreq
End Property

Private Function DesignDumpArray(baData() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1) As String
    If Size < 0 Then
        Size = UBound(baData) + 1 - Pos
    End If
    If Size > 0 Then
        DesignDumpArray = DesignDumpMemory(VarPtr(baData(Pos)), Size)
    End If
End Function

Private Function DesignDumpMemory(ByVal lPtr As Long, ByVal lSize As Long) As String
    Dim lIdx            As Long
    Dim sHex            As String
    Dim sChar           As String
    Dim lValue          As Long
    Dim aResult()       As String
    
    ReDim aResult(0 To (lSize + 15) \ 16) As String
    For lIdx = 0 To ((lSize + 15) \ 16) * 16
        If lIdx < lSize Then
            If IsBadReadPtr(lPtr, 1) = 0 Then
                Call CopyMemory(lValue, ByVal lPtr, 1)
                sHex = sHex & Right$("0" & Hex$(lValue), 2) & " "
                If lValue >= 32 Then
                    sChar = sChar & Chr$(lValue)
                Else
                    sChar = sChar & "."
                End If
            Else
                sHex = sHex & "?? "
                sChar = sChar & "."
            End If
        Else
            sHex = sHex & "   "
        End If
        If ((lIdx + 1) Mod 4) = 0 Then
            sHex = sHex & " "
        End If
        If ((lIdx + 1) Mod 16) = 0 Then
            aResult(lIdx \ 16) = Right$("000" & Hex$(lIdx - 15), 4) & " - " & sHex & sChar
            sHex = vbNullString
            sChar = vbNullString
        End If
        lPtr = (lPtr Xor &H80000000) + 1 Xor &H80000000
    Next
    DesignDumpMemory = Join(aResult, vbCrLf)
End Function

