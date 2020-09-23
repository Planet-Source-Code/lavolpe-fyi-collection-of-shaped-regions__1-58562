VERSION 5.00
Begin VB.Form frmRgnShapes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7440
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   368
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   496
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkRtTri 
      Caption         =   "B/R"
      Height          =   255
      Index           =   3
      Left            =   6480
      TabIndex        =   18
      Top             =   4080
      Width           =   615
   End
   Begin VB.CheckBox chkRtTri 
      Caption         =   "B/L"
      Height          =   255
      Index           =   2
      Left            =   5880
      TabIndex        =   17
      Top             =   4080
      Width           =   615
   End
   Begin VB.CheckBox chkRtTri 
      Caption         =   "T/R"
      Height          =   255
      Index           =   1
      Left            =   5160
      TabIndex        =   16
      Top             =   4080
      Width           =   615
   End
   Begin VB.CheckBox chkRtTri 
      Caption         =   "T/L"
      Height          =   255
      Index           =   0
      Left            =   4560
      TabIndex        =   15
      Top             =   4080
      Value           =   1  'Checked
      Width           =   615
   End
   Begin VB.OptionButton optShape 
      Caption         =   "Triangle (Right)"
      Height          =   255
      Index           =   5
      Left            =   4320
      TabIndex        =   6
      Top             =   3840
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Swap Height && Width"
      Height          =   495
      Left            =   5760
      TabIndex        =   0
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Timer digiTimer 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   6360
      Top             =   4800
   End
   Begin VB.PictureBox picDigi 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   975
      Left            =   4320
      ScaleHeight     =   61
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   181
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   4440
      Width           =   2775
   End
   Begin VB.CheckBox chkTri 
      Caption         =   "Direction \"
      Height          =   255
      Index           =   1
      Left            =   5880
      TabIndex        =   14
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CheckBox chkTri 
      Caption         =   "Direction /"
      Height          =   255
      Index           =   0
      Left            =   4560
      TabIndex        =   13
      Top             =   3480
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.OptionButton optShape 
      Caption         =   "Triangle (Isoceles)"
      Height          =   255
      Index           =   4
      Left            =   4320
      TabIndex        =   5
      Top             =   3240
      Width           =   2295
   End
   Begin VB.CheckBox chkSideA 
      Caption         =   "Right Side /"
      Height          =   255
      Index           =   5
      Left            =   5880
      TabIndex        =   12
      Top             =   2880
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.CheckBox chkSideA 
      Caption         =   "Right Side |"
      Height          =   255
      Index           =   4
      Left            =   5880
      TabIndex        =   11
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CheckBox chkSideA 
      Caption         =   "Right Side \"
      Height          =   255
      Index           =   3
      Left            =   5880
      TabIndex        =   10
      Tag             =   "2"
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CheckBox chkSideA 
      Caption         =   "Left Side /"
      Height          =   255
      Index           =   0
      Left            =   4560
      TabIndex        =   9
      Tag             =   "1"
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CheckBox chkSideA 
      Caption         =   "Left Side \"
      Height          =   255
      Index           =   2
      Left            =   4560
      TabIndex        =   7
      Tag             =   "1"
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CheckBox chkSideA 
      Caption         =   "Left Side |"
      Height          =   255
      Index           =   1
      Left            =   4560
      TabIndex        =   8
      Top             =   2640
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.OptionButton optShape 
      Caption         =   "Diagonal Rectangles (Trapezoids)"
      Height          =   255
      Index           =   3
      Left            =   4320
      TabIndex        =   4
      Top             =   2160
      Width           =   2775
   End
   Begin VB.OptionButton optShape 
      Caption         =   "Diamond"
      Height          =   255
      Index           =   2
      Left            =   4800
      TabIndex        =   3
      Top             =   1920
      Width           =   1215
   End
   Begin VB.OptionButton optShape 
      Caption         =   "Octagon"
      Height          =   255
      Index           =   1
      Left            =   4560
      TabIndex        =   2
      Top             =   1680
      Width           =   975
   End
   Begin VB.OptionButton optShape 
      Caption         =   "Hexagon"
      Height          =   255
      Index           =   0
      Left            =   4320
      TabIndex        =   1
      Top             =   1440
      Width           =   1095
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Index           =   1
      LargeChange     =   21
      Left            =   4200
      Max             =   250
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   960
      Value           =   122
      Width           =   2895
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Index           =   0
      LargeChange     =   21
      Left            =   4200
      Max             =   250
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   360
      Value           =   200
      Width           =   2895
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   3750
      Left            =   120
      ScaleHeight     =   246
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   246
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   120
      Width           =   3750
      Begin VB.PictureBox picClickMe 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00C0C000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   1320
         ScaleHeight     =   615
         ScaleWidth      =   1095
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   1440
         Width           =   1095
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Practical example using  regions for drawing >>>"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   28
      Top             =   5160
      Width           =   4095
   End
   Begin VB.Label Label2 
      Caption         =   "Use sliders to size shape horizontally or vertically.  All shapes have a horizontal and vertical counterpart."
      Height          =   495
      Index           =   1
      Left            =   120
      TabIndex        =   26
      Top             =   4560
      Width           =   4215
   End
   Begin VB.Label Label2 
      Caption         =   "Inside shape above is a picture box && shaped region is applied to show some neat effects of using these routines"
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   25
      Top             =   4080
      Width           =   4095
   End
   Begin VB.Label Label1 
      Caption         =   "Shape's Height (0 to 250)"
      Height          =   255
      Index           =   1
      Left            =   4200
      TabIndex        =   23
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Shape's Width (0 to 250)"
      Height          =   255
      Index           =   0
      Left            =   4200
      TabIndex        =   22
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "frmRgnShapes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' If you've been following some of my posts, you know that I
' have submitted some unique stuff using window regions. Well,
' here is another one. Some short routines that you can use to
' shape custom controls. These shapes will always return a
' sharp diagonal edge, if appropriate.

' 27 shapes can be created. That number includes the both the
' horizontal & vertical version of all shapes. Except the
' diamond which only has one orientation.

' Each of these routines are designed so that similar shapes
' can be stacked or "joined" if they have the same
' height (horizontal shapes) or width (vertical shapes).

' For example, you can easily join trapezoids to other trapezoids,
' stack hexagons or octagons, and stack/join triangles similar to how
' the screenshot provided with this post showed joined trapezoids.

' Also included is a somewhat simple routine, combining regions,
' which allows you to apply a two color border to any of the
' shapes. The drawing routine is at very end of this form

' In the region functions you may notice that I may extend the region points
' +1 or -1 outside the physical height and width, if so good catch. When creating
' regions, APIs by their nature exclude the bottom & right edges. Therefore, to
' ensure we create the region that "fills" up the passed height and width, we
' need to do some tweaking -- just one of the headaches with regions.

' used to provide borders on the shaped windows
Private Declare Function CreateRectRgn Lib "gdi32.dll" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32.dll" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function CombineRgn Lib "gdi32.dll" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Declare Function FrameRgn Lib "gdi32.dll" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function OffsetRgn Lib "gdi32.dll" (ByVal hRgn As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function FillRgn Lib "gdi32.dll" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long

' used to create the shaped windows
Private Declare Function CreatePolygonRgn Lib "gdi32.dll" (ByRef lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Type POINTAPI
    x As Long
    y As Long
End Type

' cached regions/brushes used to display the digital numbers in the example
Private digiNum() As Long
Private digiBrush(0 To 3) As Long

Private Sub chkRtTri_Click(Index As Integer)
' check boxes that set which right triangle to view
If chkRtTri(Index) = 0 Then
    If chkRtTri(0) + chkRtTri(1) + chkRtTri(2) + chkRtTri(3) = 0 Then
        'all unselected, so select one
        chkRtTri(2) = 1
    End If
Else
    ' treat the checkboxes like option boxes; only one selected at a time
    Dim I As Integer
    For I = 0 To 3
        If I <> Index Then chkRtTri(I) = 0
    Next
    chkRtTri(0).Tag = Index
    ' check box changed; if the right triangles opt button chosen, redraw now
    If optShape(5) = True Then Call HScroll1_Change(0)
End If
End Sub

Private Sub chkSideA_Click(Index As Integer)
' check boxes that change the left/right edge of a trapezoid

If chkSideA(Index).Value = 0 Then
    ' auto-select a check box if user unchecks all boxes
    If Index < 3 Then
        If chkSideA(0) + chkSideA(1) + chkSideA(2) = 0 Then chkSideA(2) = 1
    Else
        If chkSideA(3) + chkSideA(4) + chkSideA(5) = 0 Then chkSideA(5) = 1
    End If
Else
    ' use checkboxes like option buttons;
    ' prevent two from being simultaneously selected
    If Index < 3 Then
        If Index = 0 Then
            chkSideA(1) = 0
            chkSideA(2) = 0
        ElseIf Index = 1 Then
            chkSideA(0) = 0
            chkSideA(2) = 0
        Else
            chkSideA(1) = 0
            chkSideA(0) = 0
        End If
        chkSideA(0).Tag = Index
    Else
        If Index = 3 Then
            chkSideA(4) = 0
            chkSideA(5) = 0
        ElseIf Index = 4 Then
            chkSideA(3) = 0
            chkSideA(5) = 0
        Else
            chkSideA(3) = 0
            chkSideA(4) = 0
        End If
        chkSideA(3).Tag = Index - 3
    End If
    ' check box changed; if the trapezoid opt button chosen, redraw now
    If optShape(3) = True Then Call HScroll1_Change(0)
End If
End Sub

Private Sub chkTri_Click(Index As Integer)
' option to allow the direction of the triangle

' used like option buttons, check or uncheck the other checkbox
chkTri(Abs(Index - 1)) = Abs(chkTri(Index) - 1)
If chkTri(Index) Then
    ' if the triangle button chosen, redraw now
    chkTri(0).Tag = Index
    If optShape(0).Tag = "4" Then Call HScroll1_Change(0)
End If
End Sub

Private Function CreateDiagRectRegion(cx As Long, cy As Long, SideAStyle As Integer, SideBStyle As Integer) As Long

' the cx & cy parameters are the respective width & height of the region
'   the passed values may be modified which coder can use for other purposes
'   like drawing borders or calculating the client/clipping region
' SideAStyle is -1, 0 or 1
'   depending on horizontal/vertical shape, reflects the left or top side of the region
'   -1 draws left/top edge like  /
'    0 draws left/top edge like  |
'    1 draws left/top edge like  \
' SideBStyle is -1, 0 or 1
'   depending on horizontal/vertical shape, reflects the right or bottom side of the region
'   -1 draws right/bottom edge like  \
'    0 draws right/bottom edge like  |
'    1 draws right/bottom edge like  /


' NOTE. When doing diagonal rectangles, we need to calculate
' minimum height or width to prevent a "bow-tie" affect.  To see
' what I mean, rem out the check that resizes cx or cy based on
' Abs(SideAStyle + SideBStyle)=2 and pass sides of /\ or \/ and
' resize horizontal trapezoid by reducing width until affect is achieved

Dim tpts(0 To 4) As POINTAPI

If cx > cy Then ' horizontal

    ' absolute minimum width & height a trapezoid
    If Abs(SideAStyle + SideBStyle) = 2 Then ' has 2 opposing slanted sides
        If cx < cy * 2 Then cy = cx \ 2
    End If
    
    If SideAStyle < 0 Then
        tpts(0).x = cy - 1
        tpts(1).x = -1
    ElseIf SideAStyle > 0 Then
        tpts(1).x = cy
    End If
    tpts(1).y = cy
    
    tpts(2).x = cx + Abs(SideBStyle < 0)
    If SideBStyle > 0 Then tpts(2).x = tpts(2).x - cy
    tpts(2).y = cy
    
    tpts(3).x = cx + Abs(SideBStyle < 0)
    If SideBStyle < 0 Then tpts(3).x = tpts(3).x - cy

Else

    ' absolute minimum width & height a trapezoid
    If Abs(SideAStyle + SideBStyle) = 2 Then ' has 2 opposing slanted sides
        If cy < cx * 2 Then cx = cy \ 2
    End If
    
    If SideAStyle < 0 Then
        tpts(0).y = cx - 1
        tpts(3).y = -1
    ElseIf SideAStyle > 0 Then
        tpts(3).y = cx - 1
        tpts(0).y = -1
    End If
    
    tpts(1).y = cy
    If SideBStyle < 0 Then tpts(1).y = tpts(1).y - cx
    tpts(2).x = cx
    
    tpts(2).y = cy
    If SideBStyle > 0 Then tpts(2).y = tpts(2).y - cx
    tpts(3).x = cx

End If

tpts(4) = tpts(0)
   
CreateDiagRectRegion = CreatePolygonRgn(tpts(0), UBound(tpts) + 1, 2)

End Function

Private Function CreateHexRegion(cx As Long, cy As Long) As Long
' Function creates a horizontal/vertical hexagon region with perfectly smooth edges
' the cx & cy parameters are the respective width & height of the region
'   the passed values may be modified which coder can use for other purposes
'   like drawing borders or calculating the client/clipping region

' NOTE: A diamond can be created by passing cx & cy as equal values

Dim tpts(0 To 7) As POINTAPI

If cy > cx Then             ' vertical hex vs horizontal
    
    ' absolute minimum width & height of a hex region
    If cx < 4 Then cx = 4
    ' ensure width is even
    If cx Mod 2 Then cx = cx - 1

    ' calculate the vertical hex.

    tpts(0).x = cx \ 2              ' bot apex
    tpts(0).y = cy
    tpts(1).x = cx                  ' bot right
    tpts(1).y = cy - tpts(0).x
    tpts(2).x = cx                  ' top right
    tpts(2).y = tpts(0).x - 1
    tpts(3).x = tpts(0).x           ' top apex
    tpts(3).y = -1
    ' add an extra point & modify; trial & error
    ' shows without this added point, getting a
    ' nice smooth diagonal edge is impossible
    tpts(4).x = tpts(0).x - 1       ' added
    tpts(4).y = 0
    tpts(5).x = 0                   ' top left
    tpts(5).y = tpts(2).y
    tpts(6).x = 0                   ' bot left
    tpts(6).y = tpts(1).y
    tpts(7) = tpts(0)               ' bot apex, close polygon
    
Else
    
    ' absolute minimum width & height of a hex region
    If cy < 4 Then cy = 4
    ' ensure height is even
    If cy Mod 2 Then cy = cy - 1
    
    ' calculate the horizontal hex.
    
    tpts(0).x = 0                   ' left apex
    tpts(0).y = cy \ 2
    tpts(1).x = tpts(0).y           ' bot left
    tpts(1).y = cy
    tpts(2).x = cx - tpts(0).y      ' bot right
    tpts(2).y = tpts(1).y
    tpts(3).x = cx                  ' right apex
    tpts(3).y = tpts(0).y
    ' add an extra point & modify; trial & error
    ' shows without this added point, getting a
    ' nice smooth diagonal edge is impossible
    tpts(4).x = cx
    tpts(4).y = tpts(3).y - 1
    tpts(5).x = tpts(2).x + 1       ' top right
    tpts(5).y = 0
    tpts(6).x = tpts(1).x - 1       ' top left
    tpts(6).y = 0
    tpts(7).x = tpts(0).x           ' left apex, close polygon
    tpts(7).y = tpts(0).y - 1
    
End If

    CreateHexRegion = CreatePolygonRgn(tpts(0), UBound(tpts) + 1, 2)

End Function

Private Function CreateOctRegion(cx As Long, cy As Long) As Long
' Function returns a handle to an octagonal region

' the cx & cy parameters are the respective width & height of the region
'   the passed values may be modified which coder can use for other purposes
'   like drawing borders or calculating the client/clipping region

Dim tpts(0 To 8) As POINTAPI


If cx < cy Then ' vertical

    ' absolute minimum width & height of a octagon region
    If cx < 4 Then cx = 4
    ' ensure height is even
    If cx Mod 2 Then cx = cx - 1
    
    tpts(0).x = cx \ 4 + 1          ' bot left
    tpts(0).y = cy
    tpts(1).x = cx - cx \ 4 - 1 ' bot right
    tpts(1).y = cy
    tpts(2).x = cx                  ' mid bot right
    tpts(2).y = cy - cx \ 4 - 1
    tpts(3).x = tpts(2).x           ' mid top right
    tpts(3).y = cx \ 4
    tpts(4).x = tpts(1).x + 1         ' top right
    tpts(4).y = 0
    
    tpts(5).x = cx \ 4                 ' top left
    tpts(5).y = 0
    tpts(6).x = 0                       ' mid top left
    tpts(6).y = tpts(3).y
    tpts(7).x = 0                       ' mid bot left
    tpts(7).y = tpts(2).y
    
    tpts(8).x = tpts(0).x              ' bot left
    tpts(8).y = tpts(0).y
    

Else

    ' absolute minimum width & height of a octagon region
    If cy < 4 Then cy = 4
    ' ensure height is even
    If cy Mod 2 Then cy = cy - 1
        
        tpts(0).x = cy \ 4 + 1          ' bot left
        tpts(0).y = cy
        tpts(1).x = cx - cy \ 4 - 1 ' bot right
        tpts(1).y = cy
        tpts(2).x = cx                  ' mid bot right
        tpts(2).y = cy - cy \ 4 - 1
        tpts(3).x = tpts(2).x           ' mid top right
        tpts(3).y = cy \ 4
        tpts(4).x = tpts(1).x + 1         ' top right
        tpts(4).y = 0
        
        tpts(5).x = cy \ 4                 ' top left
        tpts(5).y = 0
        tpts(6).x = 0                       ' mid top left
        tpts(6).y = tpts(3).y
        tpts(7).x = 0                       ' mid bot left
        tpts(7).y = tpts(2).y
        
        tpts(8).x = tpts(0).x              ' bot left
        tpts(8).y = tpts(0).y
    
End If

    CreateOctRegion = CreatePolygonRgn(tpts(0), UBound(tpts) + 1, 2)

End Function

Private Sub Command1_Click()
' swap height & width
HScroll1(0).Tag = HScroll1(1).Value
HScroll1(1).Value = HScroll1(0).Value
HScroll1(0).Value = Val(HScroll1(0).Tag)
HScroll1(0).Tag = ""
Call HScroll1_Change(0)
End Sub

Private Sub digiTimer_Timer()
UpdateDigitalNumbers 0, 0, 0, 0
End Sub

Private Sub Form_Load()
' select a shape to start out with
optShape(1) = True
' create cached regions for sample timer
CreateDigiSegments
End Sub

Private Sub Form_Unload(Cancel As Integer)
' delete regions/brushes associated with sample timer
Dim I As Integer
For I = 0 To UBound(digiNum)
    DeleteObject digiNum(I)
Next
For I = 0 To UBound(digiBrush)
    DeleteObject digiBrush(I)
Next
End Sub


Private Sub CreateDigiSegments()

ReDim digiNum(0 To 6) As Long
' create the burn-in & burn-out colors for the 2 styles
' style 1: hex digital segments
digiBrush(0) = CreateSolidBrush(RGB(192, 64, 64))
digiBrush(1) = CreateSolidBrush(RGB(70, 0, 0))
' style 2: trapezoid digital segments
digiBrush(2) = CreateSolidBrush(vbCyan)
digiBrush(3) = CreateSolidBrush(RGB(0, 40, 40))

Dim segSizeX As Long, segSizeY As Long
Dim xOffset As Long, yOffset As Long
' guesstimate size needed for the sample picture box

segSizeX = picDigi.Height \ 3
If segSizeX Mod 2 Then segSizeX = segSizeX + 1
segSizeY = segSizeX \ 3
If segSizeY Mod 2 Then segSizeY = segSizeY + 1

' now create the hex segments
digiNum(0) = CreateHexRegion(segSizeX, segSizeY)
digiNum(1) = CreateHexRegion(segSizeY, segSizeX)

' calculate the x&y offsets to sorta center the the sample timer
yOffset = (picDigi.Height - segSizeX * 2 - segSizeY - 6) \ 2
xOffset = (picDigi.Width - ((segSizeX + segSizeY * 2 + 3) * 4)) \ 2

' the trapezoid segments will be 1/2 the height of the hex segments
segSizeY = segSizeY \ 2
digiNum(2) = CreateHexRegion(segSizeX, segSizeY)
digiNum(3) = CreateDiagRectRegion(segSizeX, segSizeY, 1, 1)
digiNum(4) = CreateDiagRectRegion(segSizeX, segSizeY, -1, -1)
digiNum(5) = CreateDiagRectRegion(segSizeY, segSizeX, 1, 1)
digiNum(6) = CreateDiagRectRegion(segSizeY, segSizeX, -1, -1)

UpdateDigitalNumbers segSizeX, segSizeY * 2, xOffset, yOffset

digiTimer.Enabled = True
End Sub

Private Sub UpdateDigitalNumbers(wd As Long, ht As Long, _
    userOffsetX As Long, userOffsetY As Long)

' quickly whipped together (about 15 mins worth of work) to show how
' regions can be used for drawing too. So this routine is definitely not optimized

Static cx As Long
Static cy As Long
Static cxOffset As Long
Static cyOffset As Long

If wd Then
    ' only sent once when initialized
    cx = wd
    cy = ht
    cxOffset = userOffsetX
    cyOffset = userOffsetY
End If

Dim curTime As String, xOffset As Long, yOffset As Long
Dim segOffset As Long, I As Integer, J As Integer, styleNr As Integer
Dim rgn2Use As Long, brush2Use As Long
Dim numFormat As String
Dim numWidth As Long, nSpacing As Integer

' updatable values that do not affect the static ones above
Dim styleCx As Long, styleCy As Long
Dim styleOffsetX As Long, styleOffsetY As Long

nSpacing = 2                ' pixel spacing btwn numbers
styleCy = cy                ' updatable style ids
styleCx = cx
styleOffsetX = cxOffset     ' updatable top/left offsets
styleOffsetY = cyOffset

' use seconds for the examples
curTime = Format(Now(), "ss")

' do each style where the main difference is height of region
For styleNr = 0 To 1
    ' set initial offset for the new style
    segOffset = segOffset * styleNr + 6
    ' calculate width the number will consume
    numWidth = styleCx + styleCy * 2
    
    ' loop thru each digit in the seconds
    For I = 1 To 2
        ' -set the digital segment on/off
        ' -it is read middle,bottom,then clockwise to bottom right
        ' -the decimal indicates next segments are not turned on
        Select Case Val(Mid$(curTime, I, 1))
        Case 0: numFormat = "123456.0"
        Case 1: numFormat = "56.01234"
        Case 2: numFormat = "01245.36"
        Case 3: numFormat = "01456.23"
        Case 4: numFormat = "0356.124"
        Case 5: numFormat = "01346.25"
        Case 6: numFormat = "01236.45"
        Case 7: numFormat = "456.0123"
        Case 8: numFormat = "0123456."
        Case 9: numFormat = "013456.2"
        End Select
        
        'set the burn-in brush reference
        brush2Use = styleNr * 2
        
        ' loop thru each segment & identify region offsets & references
        For J = 1 To 8
        
            If styleNr Then ' this is the trapezoid segments
                Select Case Mid$(numFormat, J, 1)
                Case 0: ' middle segment
                    xOffset = 2
                    yOffset = styleCx - 1
                    rgn2Use = 2
                Case 1: ' bottom segment
                    xOffset = 2
                    yOffset = styleCx * 2 - 1
                    rgn2Use = 4
                Case 2: ' bottom left segment
                    xOffset = 0
                    yOffset = styleCx + 2
                    rgn2Use = 5
                Case 3: ' top left segment
                    xOffset = 0
                    yOffset = 1
                    rgn2Use = 5
                Case 4: ' top segment
                    xOffset = 2
                    yOffset = 0
                    rgn2Use = 3
                Case 5: ' top right segment
                    xOffset = styleCx
                    yOffset = 1
                    rgn2Use = 6
                Case 6: ' bottom right segment
                    xOffset = styleCx
                    yOffset = styleCx + 2
                    rgn2Use = 6
                Case Else: rgn2Use = -1
                End Select
            
            Else            ' this is the hexagon segments
                Select Case Mid$(numFormat, J, 1)
                Case 0: ' middle segment
                    xOffset = styleCy \ 2 + 1
                    yOffset = styleCx + 2
                    rgn2Use = 0
                Case 1: ' bottom segment
                    xOffset = styleCy \ 2 + 1
                    yOffset = styleCx * 2 + 4
                    rgn2Use = 0
                Case 2: ' bottom left segment
                    xOffset = 0
                    yOffset = styleCx + styleCy \ 2 + 3
                    rgn2Use = 1
                Case 3: ' top left segment
                    xOffset = 0
                    yOffset = styleCy \ 2 + 1
                    rgn2Use = 1
                Case 4: ' top segment
                    xOffset = styleCy \ 2 + 1
                    yOffset = 0
                    rgn2Use = 0
                Case 5: ' top right segment
                    xOffset = styleCx + 2
                    yOffset = styleCy \ 2 + 1
                    rgn2Use = 1
                Case 6: ' bottom right segment
                    xOffset = styleCx + 2
                    yOffset = styleCx + styleCy \ 2 + 3
                    rgn2Use = 1
                Case Else:  rgn2Use = -1
                End Select
            End If
            
            ' now draw the segment
            If rgn2Use < 0 Then
                ' decimal encountered, simply set ref to burn-out color
                brush2Use = brush2Use + 1
            Else
                ' move the region to the appropriate painting location
                OffsetRgn digiNum(rgn2Use), styleOffsetX + segOffset + nSpacing + xOffset, yOffset + styleOffsetY
                ' paint the region
                FillRgn picDigi.hdc, digiNum(rgn2Use), digiBrush(brush2Use)
                ' now move the region back to 0,0 in prep for next segment
                OffsetRgn digiNum(rgn2Use), -(styleOffsetX + segOffset + nSpacing + xOffset), -(yOffset + styleOffsetY)
            End If
        Next
        ' next numeral, shift the numeral offset
        segOffset = segOffset + numWidth
    Next
    ' change settings for the trapezoid segments
    styleCy = styleCy \ 2
    styleOffsetY = styleOffsetY * 2
Next
picDigi.Refresh
End Sub






Private Sub HScroll1_Change(Index As Integer)
' used to trigger redrawing shapes

' flag to prevent actions when the "Swap Ht & Wd" button is clicked
If Len(HScroll1(0).Tag) Then Exit Sub

If Val(optShape(0).Tag) = 2 Or Val(optShape(0).Tag) = 5 Then ' diamond or right triangles
    ' diamond produced by passing equal cx & cy values to the hexagon creation routine
    ' right triangles done likewise
    ' Ensure both cx & cy are equal
    If HScroll1(Abs(Index - 1)) <> HScroll1(Index) Then
    
        HScroll1(Abs(Index - 1)).Value = HScroll1(Index).Value
        Exit Sub
    End If
ElseIf Val(optShape(0).Tag) = 4 Then ' triangle
    ' triangles produced by passing cx & 2*cx for cy or vice versa to the
    ' trapezoid creation routines. Ensure that ratio
    If HScroll1(0) > HScroll1(1) Then
        If HScroll1(0) \ 2 <> HScroll1(1) Then
            HScroll1(1).Value = HScroll1(0).Value \ 2
            Exit Sub
        End If
    Else
        If HScroll1(1) \ 2 <> HScroll1(0) Then
            HScroll1(0).Value = HScroll1(1).Value \ 2
            Exit Sub
        End If
    End If
End If

Dim pRgn As Long    ' this will be the region applied to the picturebox
' these regions are used to create & draw the borders
Dim tRgnTL As Long, tRgnBR As Long
Dim tRgn As Long
' brushes used to draw the borders
Dim hBrushL As Long, hBrushR As Long
' the ultimate size, modified if needed, of the new shape
Dim newCx As Long, newCy As Long

' create 2 brushes for the borders
hBrushL = CreateSolidBrush(vbWhite)
hBrushR = CreateSolidBrush(RGB(64, 64, 64))

' calculate the requested shape size
newCx = HScroll1(0).Value
newCy = HScroll1(1).Value

Select Case Val(optShape(0).Tag)
Case 0, 2: ' hexagon & diamond
    ' Note: by passing equal cx & cy, a perfect diamond is drawn
    pRgn = CreateHexRegion(newCx, newCy)
Case 1: ' octagon
    pRgn = CreateOctRegion(newCx, newCy)
Case 3: ' diagonal rectangle
    pRgn = CreateDiagRectRegion(newCx, newCy, Val(chkSideA(0).Tag) - 1, Val(chkSideA(3).Tag) - 1)
Case 4: ' triangle
    ' Note: by passing a cx = 2*cy or vice versa an isoceles triangle is drawn
    pRgn = CreateDiagRectRegion(newCx, newCy, Choose(Val(chkTri(0).Tag) + 1, -1, 1), Choose(Val(chkTri(0).Tag) + 1, -1, 1))
Case 5
    Select Case Val(chkRtTri(0).Tag)
    Case 0: ' top left
        pRgn = CreateDiagRectRegion(newCx, newCy, 0, 1)
    Case 1: ' top right
        pRgn = CreateDiagRectRegion(newCx, newCy, 0, -1)
    Case 2: ' bottom left
        pRgn = CreateDiagRectRegion(newCx, newCy, 1, 0)
    Case 3: ' bottom right
        pRgn = CreateDiagRectRegion(newCx, newCy, -1, 0)
    End Select
End Select


' drawing borders on shaped regions isn't exactly easy; this
' little algorithm could probably be used on very complicated
' shapes also....

' Do the left & top border first.
' Create 2 rectangular regions of the shaped size
tRgnTL = CreateRectRgn(0, 0, newCx, newCy)
tRgn = CreateRectRgn(0, 0, newCx, newCy)
' shift the new region left one to catch the left side
OffsetRgn tRgnTL, -1, 0
CombineRgn tRgnTL, tRgnTL, pRgn, 3
OffsetRgn tRgnTL, 1, 0
' now using the temp region, shift it up one to catch the top side
OffsetRgn tRgn, -1, -1
CombineRgn tRgn, tRgn, pRgn, 3
OffsetRgn tRgn, 1, 1
' add this to the new region & complete the left/top borders
CombineRgn tRgnTL, tRgn, tRgnTL, 2
DeleteObject tRgn

' do the same for the bottom & right borders
tRgnBR = CreateRectRgn(0, 0, newCx + 0, newCy + 0)
tRgn = CreateRectRgn(0, 0, newCx + 0, newCy + 0)
OffsetRgn tRgnBR, 1, 0
CombineRgn tRgnBR, tRgnBR, pRgn, 3
OffsetRgn tRgnBR, -1, 0
OffsetRgn tRgn, 1, 1
CombineRgn tRgn, tRgn, pRgn, 3
OffsetRgn tRgn, -1, -1
CombineRgn tRgnBR, tRgn, tRgnBR, 2
DeleteObject tRgn

' Should you want to create a clipping region for the shape so that
' borders are not drawn over by other drawing functions, that sample
' is shown below & rem'd out since it isn't used in these samples.
' Un'rem to show that a red fill color will be applied after the
' borders are drawn & the red fill will not overpaint the borders

' /// CLIPPING REGION EXAMPLE
'Dim clipRgn As Long, clipBrush As Long
'clipBrush = CreateSolidBrush(vbRed)
'clipRgn = CreateRectRgn(0, 0, 0, 0)
'CombineRgn clipRgn, pRgn, clipRgn, 5
'CombineRgn clipRgn, clipRgn, tRgnBR, 4
'CombineRgn clipRgn, clipRgn, tRgnTL, 4


' cut down on flicker when changing shapes...
picClickMe.Visible = False
' apply the new window region & don't delete region; windows now owns it
SetWindowRgn picClickMe.hWnd, pRgn, True
' center our new shape in the sample window
picClickMe.Move (HScroll1(0).Max - newCx) / 2, (HScroll1(1).Max - newCy) / 2, newCx, newCy

' draw the borders & delete the extra regions & brushes
picClickMe.Cls
FrameRgn picClickMe.hdc, tRgnTL, hBrushL, 1, 1
FrameRgn picClickMe.hdc, tRgnBR, hBrushR, 1, 1

' /// CLIPPING REGION EXAMPLE continued, unrem if testing
'FillRgn picClickMe.hdc, clipRgn, clipBrush
'DeleteObject clipRgn
'DeleteObject clipBrush


picClickMe.Visible = True


DeleteObject tRgnBR
DeleteObject tRgnTL
DeleteObject hBrushL
DeleteObject hBrushR

' here we will change some check box captions if needed
' To make the samples a bit easier, the captions will
' reflect the horizontal/vertical orientation of the shape
If HScroll1(0) >= HScroll1(1) Then
    If Left$(chkSideA(0).Caption, 4) = "Left" Then Exit Sub
ElseIf HScroll1(1) > HScroll1(0) Then
    If Left$(chkSideA(0).Caption, 4) = " Top" Then Exit Sub
End If
Dim I As Integer
For I = 0 To 2
    chkSideA(I).Caption = Choose(Abs(HScroll1(0) >= HScroll1(1)) + 1, " Top", "Left") & Mid$(chkSideA(I).Caption, 5)
    chkSideA(I + 3).Caption = Choose(Abs(HScroll1(0) >= HScroll1(1)) + 1, "Lower", "Right") & Mid$(chkSideA(I + 3).Caption, 6)
Next
End Sub

Private Sub optShape_Click(Index As Integer)
' simply call routine to redraw shape based on new selection
If optShape(Index) = True Then
    optShape(0).Tag = Index
    Call HScroll1_Change(0)
End If
End Sub
