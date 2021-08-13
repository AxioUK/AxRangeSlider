VERSION 5.00
Begin VB.UserControl AxRangeSlider 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   ClientHeight    =   960
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5160
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForwardFocus    =   -1  'True
   ScaleHeight     =   64
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   344
   ToolboxBitmap   =   "AxRangeSlider.ctx":0000
End
Attribute VB_Name = "AxRangeSlider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'-UC-VB6-----------------------------
'UC Name  : AxRangeSlider
'Version  : 0.08.01
'Editor   : David Rojas [AxioUK]
'Date     : 07/08/2021
'------------------------------------
Option Explicit

Private Declare Function MulDiv Lib "kernel32.dll" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function SetRECL Lib "user32" Alias "SetRect" (lpRect As RECTL, ByVal X As Long, ByVal Y As Long, ByVal W As Long, ByVal H As Long) As Long

Private Declare Function GetSysColor Lib "user32.dll" (ByVal nIndex As Long) As Long
Private Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hwnd As Long, ByVal hDC As Long) As Long

Private Declare Function GdiplusStartup Lib "GdiPlus.dll" (Token As Long, inputbuf As GDIPlusStartupInput, Optional ByVal outputbuf As Long = 0) As Long
Private Declare Sub GdiplusShutdown Lib "GdiPlus.dll" (ByVal Token As Long)
Private Declare Function GdipCreateLineBrushFromRectWithAngleI Lib "GdiPlus.dll" (ByRef mRect As RECTL, ByVal mColor1 As Long, ByVal mColor2 As Long, ByVal mAngle As Single, ByVal mIsAngleScalable As Long, ByVal mWrapMode As Long, ByRef mLineGradient As Long) As Long
Private Declare Function GdipCreateFromHDC Lib "GdiPlus.dll" (ByVal mhDC As Long, ByRef mGraphics As Long) As Long
Private Declare Function GdipCreatePen1 Lib "GdiPlus.dll" (ByVal mColor As Long, ByVal mWidth As Single, ByVal mUnit As Long, ByRef mPen As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "GdiPlus.dll" (ByVal mGraphics As Long) As Long
Private Declare Function GdipDeleteBrush Lib "GdiPlus.dll" (ByVal Brush As Long) As Long
Private Declare Function GdipDeletePen Lib "GdiPlus.dll" (ByVal mPen As Long) As Long
Private Declare Function GdipCreatePath Lib "GdiPlus.dll" (ByRef mBrushMode As Long, ByRef mPath As Long) As Long
Private Declare Function GdipAddPathLineI Lib "GdiPlus.dll" (ByVal mPath As Long, ByVal mX1 As Long, ByVal mY1 As Long, ByVal mX2 As Long, ByVal mY2 As Long) As Long
Private Declare Function GdipAddPathArcI Lib "GdiPlus.dll" (ByVal mPath As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long, ByVal mStartAngle As Single, ByVal mSweepAngle As Single) As Long
Private Declare Function GdipClosePathFigures Lib "GdiPlus.dll" (ByVal mPath As Long) As Long
Private Declare Function GdipDeletePath Lib "GdiPlus.dll" (ByVal mPath As Long) As Long
Private Declare Function GdipDrawPath Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mPen As Long, ByVal mPath As Long) As Long
Private Declare Function GdipFillPath Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mBrush As Long, ByVal mPath As Long) As Long
Private Declare Function GdipSetSmoothingMode Lib "GdiPlus.dll" (ByVal Graphics As Long, ByVal SmoothingMd As Long) As Long
Private Declare Function GdipCreateSolidFill Lib "gdiplus" (ByVal RGBA As Long, ByRef Brush As Long) As Long
Private Declare Function GdipDrawLine Lib "GdiPlus.dll" (ByVal Graphics As Long, ByVal Pen As Long, ByVal X1 As Single, ByVal Y1 As Single, ByVal X2 As Single, ByVal Y2 As Single) As Long
'---
Private Declare Function DrawTextA Lib "user32.dll" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
'Private Declare Function DrawTextW Lib "user32.dll" (ByVal hDC As Long, lpStr As Long, ByVal nCount As Long, ByRef lpRect As RECT, ByVal wFormat As Long) As Long
'---
'Private Declare Function GetCursorPos Lib "user32.dll" (ByRef lpPoint As POINTL) As Long
'Private Declare Function WindowFromPoint Lib "user32.dll" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Private Type RECTL
    Left As Long
    Top As Long
    Width As Long
    Height As Long
End Type

Private Type RECTS
    Left As Single
    Top As Single
    Width As Single
    Height As Single
End Type

Private Type Points
   X As Single
   Y As Single
End Type

Private Type GDIPlusStartupInput
    GdiPlusVersion           As Long
    DebugEventCallback       As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs   As Long
End Type

Private Type pPoints
  X1 As Long
  X2 As Long
  Valor As String
End Type

Public Enum pStyle
  pVertical
  pHorizontal
End Enum

Public Enum CallOutPosition
  coLeft
  coTop
  coRight
  coBottom
End Enum

Public Enum coMark
  cmNothing
  cmBar
  cmLeft
  cmRight
End Enum

Public Enum eTypeValue
  eDateValue
  eNumValue
  eLetterValue
End Enum

Public Enum eDateValueI
   byDay
   byMonth
   byYear
End Enum

'Constants
'Private Const CombineModeExclude As Long = &H4
Private Const WrapModeTileFlipXY = &H3
'Private Const SmoothingModeHighQuality As Long = &H2
Private Const SmoothingModeAntiAlias As Long = &H4
Private Const LOGPIXELSX As Long = 88
Private Const LOGPIXELSY As Long = 90
Private Const TLS_MINIMUM_AVAILABLE As Long = 64
'Private Const IDC_HAND As Long = 32649
Private Const UnitPixel As Long = &H2&
Private Const DT_CENTER = &H1

'Define EVENTS-------------------
Public Event Click()
Public Event DblClick()
Public Event ChangeMarks(vLeftMark As String, vRightMark As String)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

'Property Variables:
'Private hFontCollection As Long
Private GdipToken As Long
Private nScale    As Single
Private hGraphics As Long

Private m_Enabled       As Boolean
Private m_BackColor     As OLE_COLOR
Private m_BorderColor   As OLE_COLOR
Private m_ForeColor1    As OLE_COLOR
Private m_ForeColor2    As OLE_COLOR
Private m_GradientColor1 As OLE_COLOR
Private m_GradientColor2 As OLE_COLOR
Private m_ColorLeftMark As OLE_COLOR
Private m_ColorRightMark  As OLE_COLOR
Private m_ValuesLineColor As OLE_COLOR
Private m_ColorSelector As OLE_COLOR
Private m_BorderWidth   As Long
Private m_CornerCurve   As Long

Private m_ValueType As eTypeValue
Private m_DateValueIntervalBy As eDateValueI

Private m_Font1 As StdFont
Private m_Font2 As StdFont
Private m_Style As pStyle
Private mActive As coMark
Private iPts()  As Points
Private pPts()  As pPoints
Private lMark   As RECTL
Private rMark   As RECTL
Private Bar     As RECTL

Private m_MarkLValue As String
Private m_MarkRValue As String

Private m_Min As String
Private m_Max As String
Private m_Interval As Long
Private sLMark As String
Private sRMark As String
Private BarPos As Long
Private lPos   As Long
Private rPos   As Long
Private lBar   As Long
Private rBar   As Long


Private Function AddString(mCaption As String, RCT As RECT, mFont As Font, TextColor As OLE_COLOR)
Dim pFont       As IFont
Dim lFontOld    As Long
    
On Error GoTo ErrF
With UserControl
  Set pFont = mFont
  lFontOld = SelectObject(.hDC, pFont.hFont)

  .ForeColor = TextColor
    
  DrawTextA .hDC, mCaption, -1, RCT, DT_CENTER
  
  Call SelectObject(.hDC, lFontOld)
  
ErrF:
  Set pFont = Nothing
End With
End Function

Public Sub Refresh()
Draw
End Sub

Private Sub SetMarkValues()
Dim I As Integer

For I = 0 To UBound(pPts)
    If pPts(I).Valor = m_MarkLValue Then
      lPos = pPts(I).X1
      mActive = cmLeft
      Draw
    End If
    If pPts(I).Valor = m_MarkRValue Then
      rPos = pPts(I).X1
      mActive = cmRight
      Draw
    End If
Next I
  
  mActive = cmNothing
End Sub

Private Sub Draw()
Dim I As Long, lPoints As Long, tPoints As Long
Dim REC As RECTL
Dim cpLeft As RECT, cpRight As RECT

On Error GoTo ErrDraw
With UserControl
  .Cls
  GdipCreateFromHDC .hDC, hGraphics
  GdipSetSmoothingMode hGraphics, SmoothingModeAntiAlias

  'Control BackColor
  .BackColor = m_BackColor
  
  'draw Background Bar
  SetRECL REC, 10, .ScaleHeight / 3, .ScaleWidth - 20, .ScaleHeight / 3
  DrawRoundRect hGraphics, REC, RGBA(m_GradientColor1, 100), RGBA(m_GradientColor2, 100), RGBA(m_BorderColor, 100), m_BorderWidth, m_CornerCurve
  'Get Points Values
  tPoints = fSetPoints(REC, m_Min, m_Max)

  SetRECL REC, 10, (.ScaleHeight / 2), (.ScaleWidth - 20), (.ScaleHeight / 3)
  DrawPoints hGraphics, REC, m_ValuesLineColor, m_BorderWidth, m_ForeColor1, tPoints, pHorizontal
  
  'Draw Slider Bar
  If lPos < rPos Then
    SetRECL Bar, lPos, (.ScaleHeight / 3) + 1, rPos - lPos, (.ScaleHeight / 3) - 2
  Else
    SetRECL Bar, rPos, (.ScaleHeight / 3) + 1, lPos - rPos, (.ScaleHeight / 3) - 2
  End If
  
  lBar = Bar.Left
  rBar = Bar.Width
  If mActive = cmBar Then SetRECL Bar, BarPos - (rBar / 2), (.ScaleHeight / 3) + 1, rBar, (.ScaleHeight / 3) - 2
  DrawRoundRect hGraphics, Bar, RGBA(m_ColorSelector, 50), RGBA(m_ColorSelector, 50), RGBA(m_BorderColor, 60), 1, m_CornerCurve
  'Refresh values for move Marks
  lBar = Bar.Left
  rBar = Bar.Width

  'Draw Marks
  Dim bW As Long
  bW = IIf(m_ValueType = eDateValue, UserControl.TextWidth((pPts(0).Valor) & "000"), UserControl.TextWidth("0000"))
  
  If mActive = cmBar Then
    SetRECL lMark, (lBar - (bW / 2)), (.ScaleHeight / 3) - 20, bW, 20
    SetRECL rMark, (lBar + rBar - (bW / 2)), (.ScaleHeight / 3) - 20, bW, 20
  Else
    SetRECL lMark, (lPos - (bW / 2)), (.ScaleHeight / 3) - 20, bW, 20
    SetRECL rMark, (rPos - (bW / 2)), (.ScaleHeight / 3) - 20, bW, 20
  End If
  DrawBubble hGraphics, lMark, RGBA(m_BorderColor, 100), 1, RGBA(m_ColorLeftMark, 100), 2, 5, 5, coBottom
  DrawBubble hGraphics, rMark, RGBA(m_BorderColor, 100), 1, RGBA(m_ColorRightMark, 100), 2, 5, 5, coBottom

  'Set Marks Rects
  If mActive = cmBar Then
    SetREC2 cpLeft, (lBar - (bW / 2)), (.ScaleHeight / 3) - 20, (lBar - (bW / 2)) + (bW), 20
    SetREC2 cpRight, (lBar + rBar - (bW / 2)), (.ScaleHeight / 3) - 20, (lBar + rBar - (bW / 2)) + (bW), 20
  Else
    SetREC2 cpLeft, (lPos - (bW / 2)), (.ScaleHeight / 3) - 20, (lPos - (bW / 2)) + (bW), 20
    SetREC2 cpRight, (rPos - (bW / 2)), (.ScaleHeight / 3) - 20, (rPos - (bW / 2)) + (bW), 20

  End If
  
  'Draw Marks Captions
  For I = 0 To UBound(pPts)
    Select Case mActive
      Case Is = cmLeft
          If lPos >= pPts(I).X1 And lPos <= pPts(I).X2 Then
            sLMark = pPts(I).Valor
          End If
      Case Is = cmRight
          If rPos >= pPts(I).X1 And rPos <= pPts(I).X2 Then
            sRMark = pPts(I).Valor
          End If
      Case Is = cmBar
          lPos = lBar:  rPos = lBar + rBar
          If lPos >= pPts(I).X1 And lPos <= pPts(I).X2 Then sLMark = pPts(I).Valor
          If rPos >= pPts(I).X1 And rPos <= pPts(I).X2 Then sRMark = pPts(I).Valor
    End Select
  Next I
  
  AddString sLMark, cpLeft, m_Font2, m_ForeColor2
  AddString sRMark, cpRight, m_Font2, m_ForeColor2
  
  RaiseEvent ChangeMarks(sLMark, sRMark)
End With

ErrDraw:
 GdipDeleteGraphics hGraphics
End Sub

Private Sub DrawPoints(ByVal iGraphics As Long, RCT As RECTL, ColorLine As OLE_COLOR, LineWidth As Long, _
                      ColorText As OLE_COLOR, vPoints As Long, lStyle As pStyle)
Dim I As Integer
Dim hPen As Long
Dim X As Single, Y As Single
Dim W As Single, H As Single
Dim sREC As RECT
Dim rW As Single, stMark As String
Dim sSep As String, pY2 As Single
Dim pSpace As Single, vMark As Long
Dim sPar As Boolean

X = RCT.Left:  Y = RCT.Top
W = RCT.Width: H = RCT.Height

On Error GoTo zEnd
sPar = False
pSpace = (W / vPoints) * nScale
'Debug.Print "pSpace:" & pSpace

ReDim iPts(vPoints) As Points
  
rW = IIf(m_ValueType = eDateValue, UserControl.TextWidth((pPts(0).Valor) & "000"), UserControl.TextWidth("000"))
        
For I = 0 To vPoints Step m_Interval

  iPts(I).X = X + (pSpace * I)
  iPts(I).Y = Y + (H / 3)
  
  If rW * (vPoints / Interval) > W Then
    pY2 = IIf(sPar = False, iPts(I).Y + 20, iPts(I).Y + 10)
  Else
    pY2 = iPts(I).Y + 10
  End If
  
  UserControl.Line (iPts(I).X, iPts(I).Y)-(iPts(I).X, pY2), ColorLine
  
  sREC.Left = iPts(I).X - 2 - (rW / 2)
  If rW * (vPoints / Interval) > W Then
    sREC.Top = IIf(sPar = False, iPts(I).Y + 20, iPts(I).Y + 10)
  Else
    sREC.Top = iPts(I).Y + 10
  End If
  sREC.Right = sREC.Left + rW + 5
  sREC.Bottom = sREC.Top + UserControl.TextHeight("00") + 2
  
  Select Case m_ValueType
    Case Is = eLetterValue
        stMark = Chr$(Asc(m_Min) + (I))
    Case Is = eNumValue
        stMark = CLng(m_Min) + (I)
    Case Is = eDateValue
      sSep = IIf(InStr(m_Min, "-") <> 0, "-", "/")
      'Debug.Print sSep
      If m_DateValueIntervalBy = byDay Then
          stMark = DateAdd("d", CDbl(I), CDate(m_Min))
      ElseIf m_DateValueIntervalBy = byMonth Then
          stMark = Format$(DateAdd("m", CDbl(I), CDate(m_Min)), "MM-yyyy")
      ElseIf m_DateValueIntervalBy = byYear Then
          stMark = Year(DateAdd("yyyy", CDbl(I), CDate(m_Min)))
      End If
  End Select
  'Debug.Print stMark
  AddString stMark, sREC, m_Font1, ColorText
  sPar = Not sPar
Next I
  
zEnd:
  GdipDeletePen hPen

End Sub

Private Function DrawRoundRect(ByVal hGraphics As Long, RECT As RECTL, ByVal Color1 As Long, Color2 As Long, _
                               ByVal BorderColor As Long, ByVal BorderWidth As Long, ByVal Round As Long) As Long
    Dim hPen As Long
    Dim hBrush As Long
    Dim mPath As Long
    Dim mRound As Long
    
    If m_BorderWidth <> 0 Then GdipCreatePen1 BorderColor, BorderWidth * nScale, &H2, hPen
    GdipCreateLineBrushFromRectWithAngleI RECT, Color1, Color2, 0, 0, WrapModeTileFlipXY, hBrush
    
    GdipCreatePath &H0, mPath   '&H0
    
    With RECT
        mRound = GetSafeRound((Round * nScale), .Width, .Height)
        If mRound = 0 Then mRound = 1
            GdipAddPathArcI mPath, .Left, .Top, mRound, mRound, 180, 90
            GdipAddPathArcI mPath, (.Left + .Width) - mRound, .Top, mRound, mRound, 270, 90
            GdipAddPathArcI mPath, (.Left + .Width) - mRound, (.Top + .Height) - mRound, mRound, mRound, 0, 90
            GdipAddPathArcI mPath, .Left, (.Top + .Height) - mRound, mRound, mRound, 90, 90
    End With
    
    GdipClosePathFigures mPath
    GdipFillPath hGraphics, hBrush, mPath
    GdipDrawPath hGraphics, hPen, mPath
    
    Call GdipDeletePath(mPath)
    Call GdipDeleteBrush(hBrush)
    Call GdipDeletePen(hPen)
End Function

Private Function DrawBubble(ByVal hGraphics As Long, RCT As RECTL, BorderColor As Long, BorderWidth As Long, BackColor As Long, CornerCurve As Long, coWidth As Long, coLen As Long, COPos As CallOutPosition) As Long
    Dim mPath As Long
    Dim hPen As Long
    Dim hBrush As Long
    Dim mRound As Long
    Dim Xx As Long, Yy As Long
'    Dim MidBorder As Long
    Dim lMax As Long
    Dim coAngle  As Long

With RCT
        
    coAngle = coWidth / 2

    mRound = GetSafeRound(CornerCurve * nScale, .Width, .Height)
    
    Select Case COPos
        Case coLeft
            .Left = .Left + coLen
            .Width = .Width - coLen
            lMax = .Height - (mRound * 2)
            If coWidth > lMax Then coWidth = lMax
        Case coTop
            .Top = .Top + coLen
            .Height = .Height - coLen
            lMax = .Width - (mRound * 2)
            If coWidth > lMax Then coWidth = lMax
        Case coRight
            .Width = .Width - coLen
            lMax = .Height - (mRound * 2)
            If coWidth > lMax Then coWidth = lMax
        Case coBottom
            .Height = .Height - coLen
            lMax = .Width - (mRound * 2)
            If coWidth > lMax Then coWidth = lMax
    End Select

    GdipCreatePen1 BorderColor, BorderWidth, UnitPixel, hPen
    GdipCreateSolidFill BackColor, hBrush
    Call GdipCreatePath(&H0, mPath)
                    
    GdipAddPathArcI mPath, .Left, .Top, mRound * 2, mRound * 2, 180, 90

    If COPos = coTop Then
        Xx = .Left + (.Width - coWidth) / 2
        If mRound = 0 Then GdipAddPathLineI mPath, .Left, .Top, .Left, .Top
        GdipAddPathLineI mPath, Xx, .Top, Xx + coAngle, .Top - coLen
        GdipAddPathLineI mPath, Xx + coAngle, .Top - coLen, Xx + coWidth, .Top
    End If

    GdipAddPathArcI mPath, .Left + .Width - mRound * 2, .Top, mRound * 2, mRound * 2, 270, 90

    If COPos = coRight Then
        Yy = .Top + (.Height - coWidth) / 2
        Xx = .Left + .Width
        If mRound = 0 Then GdipAddPathLineI mPath, .Left + .Width, .Top, .Left + .Width, .Top
        GdipAddPathLineI mPath, Xx, Yy, Xx + coLen, Yy + coAngle
        GdipAddPathLineI mPath, Xx + coLen, Yy + coAngle, Xx, Yy + coWidth
    End If

    GdipAddPathArcI mPath, .Left + .Width - mRound * 2, .Top + .Height - mRound * 2, mRound * 2, mRound * 2, 0, 90

    If COPos = coBottom Then
        Xx = .Left + (.Width - coWidth) / 2
        Yy = .Top + .Height
        If mRound = 0 Then GdipAddPathLineI mPath, .Left + .Width, .Top + .Height, .Left + .Width, .Top + .Height
        GdipAddPathLineI mPath, Xx + coWidth, Yy, Xx + coAngle, Yy + coLen
        GdipAddPathLineI mPath, Xx + coAngle, Yy + coLen, Xx, Yy
    End If

    GdipAddPathArcI mPath, .Left, .Top + .Height - mRound * 2, mRound * 2, mRound * 2, 90, 90
    
    If COPos = coLeft Then
        Yy = .Top + (.Height - coWidth) / 2
        If mRound = 0 Then GdipAddPathLineI mPath, .Left, .Top + .Height, .Left, .Top + .Height
        GdipAddPathLineI mPath, .Left, Yy + coWidth, .Left - coLen, Yy + coAngle
        GdipAddPathLineI mPath, .Left - coLen, Yy + coAngle, .Left, Yy
    End If
End With
        
    GdipClosePathFigures mPath
    GdipFillPath hGraphics, hBrush, mPath
    GdipDrawPath hGraphics, hPen, mPath
    
    Call GdipDeletePath(mPath)
    Call GdipDeleteBrush(hBrush)
    Call GdipDeletePen(hPen)

End Function

Private Function fSetPoints(RCT As RECTL, minVal As String, maxVal As String) As Long
Dim p As Integer, iSpace As Single, I As Integer

On Error GoTo ErrPoints

Select Case m_ValueType
  Case Is = eNumValue
      p = CLng(maxVal) - CLng(minVal)
  Case Is = eLetterValue
      p = Asc(maxVal) - Asc(minVal)
  Case Is = eDateValue
      If m_DateValueIntervalBy = byDay Then
          p = DateDiff("d", CDate(minVal), CDate(maxVal))
      ElseIf m_DateValueIntervalBy = byMonth Then
          p = DateDiff("m", CDate(minVal), CDate(maxVal))
      ElseIf m_DateValueIntervalBy = byYear Then
          p = DateDiff("yyyy", CDate(minVal), CDate(maxVal))
      End If
End Select

iSpace = (RCT.Width / p) * nScale

ReDim pPts(p) As pPoints

fSetPoints = p

For I = 0 To UBound(pPts)

  pPts(I).X1 = RCT.Left + (iSpace * I)
  pPts(I).X2 = pPts(I).X1 + iSpace
  
  Select Case m_ValueType
    Case Is = eLetterValue
          pPts(I).Valor = Chr$(Asc(minVal) + I)
          
    Case Is = eNumValue
          pPts(I).Valor = CLng(minVal) + I
          
    Case Is = eDateValue
      If m_DateValueIntervalBy = byDay Then
          pPts(I).Valor = DateAdd("d", CDbl(I), CDate(minVal))
      ElseIf m_DateValueIntervalBy = byMonth Then
          pPts(I).Valor = Format$(DateAdd("m", CDbl(I), CDate(minVal)), "MM-yyyy")
      ElseIf m_DateValueIntervalBy = byYear Then
          pPts(I).Valor = Format$(DateAdd("yyyy", CDbl(I), CDate(minVal)), "yyyy")
      End If
      
  End Select
  
Next I

Exit Function
ErrPoints:
  Debug.Print "Error setting points"
End Function

Private Function GetSafeRound(Angle As Integer, Width As Long, Height As Long) As Integer
    Dim lRet As Integer
    lRet = Angle
    If lRet * 2 > Height Then lRet = Height \ 2
    If lRet * 2 > Width Then lRet = Width \ 2
    GetSafeRound = lRet
End Function

Private Function GetWindowsDPI() As Double
    Dim hDC As Long, LPX  As Double ', LPY As Double
    hDC = GetDC(0)
    LPX = CDbl(GetDeviceCaps(hDC, LOGPIXELSX))
    'LPY = CDbl(GetDeviceCaps(hDC, LOGPIXELSY))
    ReleaseDC 0, hDC

    If (LPX = 0) Then
        GetWindowsDPI = 1#
    Else
        GetWindowsDPI = LPX / 96#
    End If
End Function

'Private Function IconCharCode(ByVal New_IconCharCode As String) As Long
'    New_IconCharCode = UCase(Replace(New_IconCharCode, Space(1), vbNullString))
'    New_IconCharCode = UCase(Replace(New_IconCharCode, "U+", "&H"))
'    If Not VBA.Left$(New_IconCharCode, 2) = "&H" And Not IsNumeric(New_IconCharCode) Then
'        IconCharCode = "&H" & New_IconCharCode
'    Else
'        IconCharCode = New_IconCharCode
'    End If
'End Function

Private Sub InitGDI()
    Dim GdipStartupInput As GDIPlusStartupInput
    GdipStartupInput.GdiPlusVersion = 1&
    Call GdiplusStartup(GdipToken, GdipStartupInput, ByVal 0)
End Sub

'Private Function ReadValue(ByVal lProp As Long, Optional Default As Long) As Long
'    Dim I       As Long
'    For I = 0 To TLS_MINIMUM_AVAILABLE - 1
'        If TlsGetValue(I) = lProp Then
'            ReadValue = TlsGetValue(I + 1)
'            Exit Function
'        End If
'    Next
'    ReadValue = Default
'End Function

Private Function RGBA(ByVal RGBColor As Long, ByVal Opacity As Long) As Long
  If (RGBColor And &H80000000) Then RGBColor = GetSysColor(RGBColor And &HFF&)
  RGBA = (RGBColor And &HFF00&) Or (RGBColor And &HFF0000) \ &H10000 Or (RGBColor And &HFF) * &H10000
  Opacity = CByte((Abs(Opacity) / 100) * 255)
  If Opacity < 128 Then
      If Opacity < 0& Then Opacity = 0&
      RGBA = RGBA Or Opacity * &H1000000
  Else
      If Opacity > 255& Then Opacity = 255&
      RGBA = RGBA Or (Opacity - 128&) * &H1000000 Or &H80000000
  End If
End Function

'Private Sub SafeRange(Value, Min, Max)
'    If Value < Min Then Value = Min
'    If Value > Max Then Value = Max
'End Sub

Private Function SetREC2(lpRect As RECT, ByVal X As Long, ByVal Y As Long, ByVal R As Long, ByVal B As Long) As Long
  lpRect.Left = X:    lpRect.Top = Y
  lpRect.Right = R:   lpRect.Bottom = B
End Function

Private Function SetRECS(lpRect As RECTS, ByVal X As Long, ByVal Y As Long, ByVal W As Long, ByVal H As Long) As Long
  lpRect.Left = X:    lpRect.Top = Y
  lpRect.Width = W:   lpRect.Height = H
End Function

Private Sub TerminateGDI()
    Call GdiplusShutdown(GdipToken)
End Sub

Private Sub UserControl_Initialize()
InitGDI
nScale = GetWindowsDPI

End Sub

Private Sub UserControl_InitProperties()
'hFontCollection = ReadValue(&HFC)

  m_Enabled = True
  m_Style = pHorizontal
  m_BorderColor = &HFF8080
  m_BackColor = &H8000000F
  m_GradientColor1 = &HFF&
  m_GradientColor2 = &HC000&
  m_BorderWidth = 1
  m_CornerCurve = 10
  m_ForeColor1 = &H8D4214
  m_ForeColor2 = &HFFFFFF
  Set m_Font1 = UserControl.Font
  Set m_Font2 = UserControl.Font
  m_ColorRightMark = vbRed
  m_ColorLeftMark = &HFFC0C0
  m_ValuesLineColor = &H8D4214
  m_Min = "0"
  m_Max = "100"
  m_Interval = 10
  m_ValueType = eNumValue
  m_DateValueIntervalBy = byDay
  m_ColorSelector = &H0&
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim hB As Long

hB = (UserControl.ScaleHeight / 3)
If Button = vbLeftButton Then
  If X > lMark.Left And X < (lMark.Left + lMark.Width) And Y > lMark.Top And Y < (lMark.Top + lMark.Height) Then
    mActive = cmLeft
  End If
  If X > rMark.Left And X < (rMark.Left + rMark.Width) And Y > rMark.Top And Y < (rMark.Top + rMark.Height) Then
    mActive = cmRight
  End If
  If X > Bar.Left + 5 And X < (Bar.Left + Bar.Width) - 5 And Y > Bar.Top And Y < (Bar.Top + Bar.Height) Then
    mActive = cmBar
  End If
  '--Border ActiveBar
  If X > Bar.Left - 2 And X < Bar.Left + 2 And Y > Bar.Top And Y < (Bar.Top + Bar.Height) Then
    mActive = cmLeft
  ElseIf X > (Bar.Left + Bar.Width) - 2 And X < (Bar.Left + Bar.Width) + 2 And Y > Bar.Top And Y < (Bar.Top + Bar.Height) Then
    mActive = cmRight
  End If
End If

RaiseEvent MouseDown(Button, Shift, X, Y)
RaiseEvent Click
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If X > Bar.Left - 2 And X < Bar.Left + 2 And Y > Bar.Top And Y < (Bar.Top + Bar.Height) Then
  UserControl.MousePointer = vbSizeWE
ElseIf X > (Bar.Left + Bar.Width) - 2 And X < (Bar.Left + Bar.Width) + 2 And Y > Bar.Top And Y < (Bar.Top + Bar.Height) Then
  UserControl.MousePointer = vbSizeWE
Else
  UserControl.MousePointer = vbDefault
End If

If Button = vbLeftButton Then
  Select Case mActive
    Case Is = cmLeft
      If X >= 10 And X <= (UserControl.ScaleWidth - 10) Then lPos = X
      UserControl.MousePointer = vbSizeWE
    Case Is = cmRight
      If X >= 10 And X <= (UserControl.ScaleWidth - 10) Then rPos = X
      UserControl.MousePointer = vbSizeWE
    Case Is = cmBar
      If X - (rBar / 2) >= 10 And X + (rBar / 2) <= (UserControl.ScaleWidth - 10) Then BarPos = X
      UserControl.MousePointer = vbSizeWE
  End Select
  Refresh
End If
RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
mActive = cmNothing
UserControl.MousePointer = vbDefault
RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_Paint()
Refresh
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
With PropBag
  m_Enabled = .ReadProperty("Enabled", True)
  m_Style = .ReadProperty("Style", 0)
  m_BorderColor = .ReadProperty("BorderColor", &HFF8080)
  m_BackColor = .ReadProperty("BackColor", &H8000000F)
  m_GradientColor1 = .ReadProperty("GradientColor1", &HFF&)
  m_GradientColor2 = .ReadProperty("GradientColor2", &HC000&)
  m_BorderWidth = .ReadProperty("BorderWidth", 1)
  m_CornerCurve = .ReadProperty("CornerCurve", 10)
  m_ForeColor1 = .ReadProperty("ValuesForeColor", &H8D4214)
  m_ForeColor2 = .ReadProperty("MarksForeColor", &HFFFFFF)
  Set m_Font1 = .ReadProperty("ValuesFont", UserControl.Font)
  Set m_Font2 = .ReadProperty("MarksFont", UserControl.Font)
  m_ColorRightMark = .ReadProperty("ColorRightMark", vbRed)
  m_ColorLeftMark = .ReadProperty("ColorLeftMark", &HFFC0C0)
  m_ValuesLineColor = .ReadProperty("ValuesLineColor", &H8D4214)
  m_Min = .ReadProperty("Min", "0")
  m_Max = .ReadProperty("Max", "100")
  m_Interval = .ReadProperty("Interval", 10)
  m_ValueType = .ReadProperty("ValueType", eNumValue)
  m_DateValueIntervalBy = .ReadProperty("DateValueIntervalBy", 0)
  m_ColorSelector = .ReadProperty("ColorSelector", &H0&)
End With

End Sub

Private Sub UserControl_Resize()
  lPos = 10
  rPos = 60

Refresh
End Sub

Private Sub UserControl_Show()
Refresh
End Sub

Private Sub UserControl_Terminate()
TerminateGDI
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag
  Call .WriteProperty("Enabled", m_Enabled)
  Call .WriteProperty("Style", m_Style)
  Call .WriteProperty("BorderColor", m_BorderColor)
  Call .WriteProperty("BackColor", m_BackColor)
  Call .WriteProperty("GradientColor1", m_GradientColor1)
  Call .WriteProperty("GradientColor2", m_GradientColor2)
  Call .WriteProperty("BorderWidth", m_BorderWidth)
  Call .WriteProperty("CornerCurve", m_CornerCurve)
  Call .WriteProperty("ValuesForeColor", m_ForeColor1)
  Call .WriteProperty("MarksForeColor", m_ForeColor2)
  Call .WriteProperty("ValuesFont", m_Font1)
  Call .WriteProperty("MarksFont", m_Font2)
  Call .WriteProperty("ColorRightMark", m_ColorRightMark)
  Call .WriteProperty("ColorLeftMark", m_ColorLeftMark)
  Call .WriteProperty("ValuesLineColor", m_ValuesLineColor)
  Call .WriteProperty("Min", m_Min)
  Call .WriteProperty("Max", m_Max)
  Call .WriteProperty("Interval", m_Interval, 10)
  Call .WriteProperty("ValueType", m_ValueType)
  Call .WriteProperty("DateValueIntervalBy", m_DateValueIntervalBy)
  Call .WriteProperty("ColorSelector", m_ColorSelector)
  
End With
  
End Sub

Public Property Get BackColor() As OLE_COLOR
  BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_Color As OLE_COLOR)
  m_BackColor = New_Color
  PropertyChanged "BackColor"
  Refresh
End Property

Public Property Get GradientColor1() As OLE_COLOR
  GradientColor1 = m_GradientColor1
End Property

Public Property Let GradientColor1(ByVal New_Color As OLE_COLOR)
  m_GradientColor1 = New_Color
  PropertyChanged "GradientColor1"
  Refresh
End Property

Public Property Get GradientColor2() As OLE_COLOR
  GradientColor2 = m_GradientColor2
End Property

Public Property Let GradientColor2(ByVal New_Color As OLE_COLOR)
  m_GradientColor2 = New_Color
  PropertyChanged "GradientColor2"
  Refresh
End Property

Public Property Get ColorRightMark() As OLE_COLOR
  ColorRightMark = m_ColorRightMark
End Property

Public Property Let ColorRightMark(ByVal NewColorRightMark As OLE_COLOR)
  m_ColorRightMark = NewColorRightMark
  PropertyChanged "ColorRightMark"
  Refresh
End Property

Public Property Get BorderColor() As OLE_COLOR
  BorderColor = m_BorderColor
End Property

Public Property Let BorderColor(ByVal NewBorderColor As OLE_COLOR)
  m_BorderColor = NewBorderColor
  PropertyChanged "BorderColor"
  Refresh
End Property

Public Property Get BorderWidth() As Long
  BorderWidth = m_BorderWidth
End Property

Public Property Let BorderWidth(ByVal NewBorderWidth As Long)
  m_BorderWidth = NewBorderWidth
  PropertyChanged "BorderWidth"
  Refresh
End Property

Public Property Get ValuesForeColor() As OLE_COLOR
  ValuesForeColor = m_ForeColor1
End Property

Public Property Let ValuesForeColor(ByVal NewForeColor As OLE_COLOR)
  m_ForeColor1 = NewForeColor
  PropertyChanged "ValuesForeColor"
  Refresh
End Property

Public Property Get ValuesFont() As StdFont
  Set ValuesFont = m_Font1
End Property

Public Property Set ValuesFont(ByVal New_Font As StdFont)
  Set m_Font1 = New_Font
  PropertyChanged "ValuesFont"
  Refresh
End Property

Public Property Get MarksForeColor() As OLE_COLOR
  MarksForeColor = m_ForeColor2
End Property

Public Property Let MarksForeColor(ByVal NewForeColor As OLE_COLOR)
  m_ForeColor2 = NewForeColor
  PropertyChanged "MarksForeColor"
  Refresh
End Property

Public Property Get MarksFont() As StdFont
  Set MarksFont = m_Font2
End Property

Public Property Set MarksFont(ByVal New_Font As StdFont)
  Set m_Font2 = New_Font
  PropertyChanged "MarksFont"
  Refresh
End Property

Public Property Get CornerCurve() As Long
  CornerCurve = m_CornerCurve
End Property

Public Property Let CornerCurve(ByVal NewCornerCurve As Long)
  m_CornerCurve = NewCornerCurve
  PropertyChanged "CornerCurve"
  Refresh
End Property

Public Property Get Enabled() As Boolean
  Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
  m_Enabled = New_Enabled
  PropertyChanged "Enabled"
End Property

Public Property Get hDC() As Long
    hDC = UserControl.hDC
End Property

Public Property Get hwnd() As Long
    hwnd = UserControl.hwnd
End Property

Public Property Get ValuesLineColor() As OLE_COLOR
  ValuesLineColor = m_ValuesLineColor
End Property

Public Property Let ValuesLineColor(ByVal NewValuesLineColor As OLE_COLOR)
  m_ValuesLineColor = NewValuesLineColor
  PropertyChanged "ValuesLineColor"
  Refresh
End Property

Public Property Get ColorLeftMark() As OLE_COLOR
  ColorLeftMark = m_ColorLeftMark
End Property

Public Property Let ColorLeftMark(ByVal New_Color As OLE_COLOR)
  m_ColorLeftMark = New_Color
  PropertyChanged "ColorLeftMark"
  Refresh
End Property

Public Property Get Style() As pStyle
  Style = m_Style
End Property

Public Property Let Style(ByVal NewStyle As pStyle)
  m_Style = NewStyle
  PropertyChanged "Style"
  UserControl_Resize
End Property

Public Property Get Version() As String
Version = App.Major & "." & App.Minor & "." & App.Revision
End Property

Public Property Get Visible() As Boolean
  Visible = Extender.Visible
End Property

Public Property Let Visible(ByVal NewVisible As Boolean)
  Extender.Visible = NewVisible
End Property

Public Property Get Min() As String
  Min = m_Min
End Property

Public Property Let Min(ByVal NewMin As String)
  m_Min = NewMin
  PropertyChanged "Min"
  Refresh
End Property

Public Property Get Max() As String
  Max = m_Max
End Property

Public Property Let Max(ByVal NewMax As String)
  m_Max = NewMax
  PropertyChanged "Max"
  Refresh
End Property

Public Property Get Interval() As Long
  Interval = m_Interval
End Property

Public Property Let Interval(ByVal NewInterval As Long)
  m_Interval = NewInterval
  If m_Interval = 0 Then m_Interval = 1
  PropertyChanged "Interval"
  Refresh
End Property

Public Property Get ValueType() As eTypeValue
  ValueType = m_ValueType
End Property

Public Property Let ValueType(ByVal NewValueType As eTypeValue)
  m_ValueType = NewValueType
  PropertyChanged "ValueType"
  Select Case m_ValueType
    Case Is = eNumValue:    Min = 0:    Max = 100
    Case Is = eLetterValue: Min = "A":  Max = "Z"
    Case Is = eDateValue:   Min = "01-01-2021": Max = "31-01-2021"
  End Select
  Refresh
End Property

Public Property Get DateValueIntervalBy() As eDateValueI
  DateValueIntervalBy = m_DateValueIntervalBy
End Property

Public Property Let DateValueIntervalBy(ByVal NewDateValueIntervalBy As eDateValueI)
  m_DateValueIntervalBy = NewDateValueIntervalBy
  PropertyChanged "DateValueIntervalBy"
  Refresh
End Property

Public Property Get MarkLValue() As String
  MarkLValue = sLMark
End Property

Public Property Get MarkRValue() As String
  MarkRValue = sRMark
End Property

'Public Property Let MarkLValue(ByVal NewMarkLValue As String)
'  m_MarkLValue = NewMarkLValue
'  PropertyChanged "MarkLValue"
'  SetMarkValues
'End Property
'
'Public Property Let MarkRValue(ByVal NewMarkRValue As String)
'  m_MarkRValue = NewMarkRValue
'  PropertyChanged "MarkRValue"
'  SetMarkValues
'End Property

Public Function SetMarkLValue(ByVal NewMarkLValue As String)
  m_MarkLValue = NewMarkLValue
  SetMarkValues
End Function

Public Function SetMarkRValue(ByVal NewMarkRValue As String)
  m_MarkRValue = NewMarkRValue
  SetMarkValues
End Function

Public Property Get ColorSelector() As OLE_COLOR
  ColorSelector = m_ColorSelector
End Property

Public Property Let ColorSelector(ByVal NewColorSelector As OLE_COLOR)
  m_ColorSelector = NewColorSelector
  PropertyChanged "ColorSelector"
  Refresh
End Property
