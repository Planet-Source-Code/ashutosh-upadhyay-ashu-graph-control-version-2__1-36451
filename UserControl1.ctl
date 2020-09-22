VERSION 5.00
Begin VB.UserControl AshuGraphControl 
   ClientHeight    =   3516
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4596
   PropertyPages   =   "UserControl1.ctx":0000
   ScaleHeight     =   293
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   383
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      Height          =   372
      Left            =   3840
      ScaleHeight     =   27
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   37
      TabIndex        =   5
      Top             =   3120
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   132
      Left            =   120
      TabIndex        =   2
      Top             =   3120
      Width           =   2652
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   2172
      Left            =   4200
      TabIndex        =   1
      Top             =   0
      Width           =   132
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      FillColor       =   &H00404040&
      ForeColor       =   &H00FF0000&
      Height          =   2772
      Left            =   0
      ScaleHeight     =   227
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   327
      TabIndex        =   0
      Top             =   0
      Width           =   3972
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Busy...."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1452
         Left            =   240
         TabIndex        =   4
         Top             =   600
         Width           =   3612
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         ForeColor       =   &H80000001&
         Height          =   216
         Left            =   240
         TabIndex        =   3
         Top             =   120
         Width           =   516
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuTrackMouse 
         Caption         =   "&Track Mouse"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuShowGrids 
         Caption         =   "Show &Grids"
         Checked         =   -1  'True
      End
      Begin VB.Menu separator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSendToExcel 
         Caption         =   "Send To &Excel"
      End
      Begin VB.Menu separator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrintMe 
         Caption         =   "&Print Me"
      End
      Begin VB.Menu mnuseparator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "&Copy"
      End
      Begin VB.Menu mnuCopyAll 
         Caption         =   "Copy All"
      End
   End
End
Attribute VB_Name = "AshuGraphControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
'Event Declarations:
'Event Declarations:
Event Click() 'MappingInfo=Picture1,Picture1,-1,Click
Event DblClick() 'MappingInfo=Picture1,Picture1,-1,DblClick
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=Picture1,Picture1,-1,MouseUp
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=Picture1,Picture1,-1,MouseDown


'Default Property Values:
Const m_def_XUnitToPixels = 1
Const m_def_YUnitToPixels = 1
Const m_def_XGap = 50
Const m_def_YGap = 50
Const m_def_XMinPosition = 0
Const m_def_YMinPosition = 0
'Property Variables:
Dim m_FontColor As OLE_COLOR
Dim m_XUnitToPixels As Single
Dim m_YUnitToPixels As Single
Dim m_XGap As Single
Dim m_YGap As Single
Dim m_XMinPosition As Single
Dim m_YMinPosition As Single
Dim m_ShowGrids As Boolean
Dim m_GridColor As OLE_COLOR
'''''
'Default Property Values:
Const m_def_PictureFileName = " "
'Property Variables:
Dim m_PictureFileName As String

''''''Variable of class CGraphData to store graph points
Const MAX_NO_OF_GRAPHS = 5
Dim m_Graph(MAX_NO_OF_GRAPHS) As New CGraphData
'''''following values are used for scrolling graph as well as deciding begining point of graph
Dim m_XInitialMargin As Single
Dim m_YInitialMargin As Single

'''''following variables holds maximum and minimum values of x & y among all graphs
'''''these are used in scrolling graph  as well as printing
Dim p_MaxX As Single
Dim p_MinX As Single
Dim p_MaxY As Single
Dim p_MinY As Single
'''''following variables are used in printing graphs
Dim p_PageHeight As Long '''size in pixels
Dim p_PageWidth As Long
Dim p_NumCols As Long    '''' no of pages in row and column
Dim p_NumRows As Long    ''' total no of pages will be printed = p_NumCols * p_NumRows
Dim p_XScaleHeight As Long
Dim p_YScaleWidth As Long
Dim p_XPageMargin As Long    '''left & top margins on paper
Dim p_YPagemargin As Long
Dim p_XGraphSize As Long  ''''max. size in pixels on paper
Dim p_YGraphSize As Long
Dim p_XInitialMargin As Single  ''' to decide begining point of graph and scale on current paper
Dim p_YInitialMargin As Single
Dim p_zoomfactor As Long
Dim p_Header As String
Dim p_XLabel As String
Dim p_YLabel As String
Dim p_PrintPageNo As Integer ' 0 means don't print  1 means print at top   2 means print as bottom
''''''''''variable to tell whether track mouse pointer or not
Dim m_TrackPointer As Boolean
''''header in excel
Dim excel_Header As String
'''''End of list of variables


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Appearance
Public Property Get Appearance() As Integer
Attribute Appearance.VB_Description = "Returns/sets whether or not an object is painted at run time with 3-D effects."
Attribute Appearance.VB_ProcData.VB_Invoke_Property = "Options"
    Appearance = UserControl.Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As Integer)
If (New_Appearance = 0) Or (New_Appearance = 1) Then
    UserControl.Appearance() = New_Appearance
    PropertyChanged "Appearance"
End If
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Picture1,Picture1,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = Picture1.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    Picture1.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BorderStyle
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
Attribute BorderStyle.VB_ProcData.VB_Invoke_Property = "Options"
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    If (New_BorderStyle = 0) Or (New_BorderStyle = 1) Then
    UserControl.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
    End If
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Picture1,Picture1,-1,Cls
Public Sub Cls()
Attribute Cls.VB_Description = "Clears graphics and text generated at run time from a Form, Image, or PictureBox."
    Picture1.Cls
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
Attribute Enabled.VB_ProcData.VB_Invoke_Property = "Options"
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Picture1,Picture1,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = Picture1.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Dim tmp As Long
    Set Picture1.Font = New_Font
     PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Picture1,Picture1,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = Picture1.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    Picture1.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
    Me.Invalidate
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hWnd
Public Property Get hwnd() As Long
Attribute hwnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
    hwnd = UserControl.hwnd
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Picture1,Picture1,-1,Image
Public Property Get Image() As Picture
Attribute Image.VB_Description = "Returns a handle, provided by Microsoft Windows, to a persistent bitmap."
    Set Image = Picture1.Image
End Property
Public Property Get Picture() As Picture
    Set Picture = Picture1.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set Picture1.Picture = New_Picture
    PropertyChanged "Picture"
    
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ScaleHeight
Public Property Get ScaleHeight() As Single
Attribute ScaleHeight.VB_Description = "Returns/sets the number of units for the vertical measurement of an object's interior."
    ScaleHeight = UserControl.ScaleHeight
End Property

Public Property Let ScaleHeight(ByVal New_ScaleHeight As Single)
    UserControl.ScaleHeight() = New_ScaleHeight
    PropertyChanged "ScaleHeight"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ScaleWidth
Public Property Get ScaleWidth() As Single
Attribute ScaleWidth.VB_Description = "Returns/sets the number of units for the horizontal measurement of an object's interior."
    ScaleWidth = UserControl.ScaleWidth
End Property

Public Property Let ScaleWidth(ByVal New_ScaleWidth As Single)
    UserControl.ScaleWidth() = New_ScaleWidth
    PropertyChanged "ScaleWidth"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=6,0,0,0
'Public Property Get FontColor() As OLE_COLOR
'     FontColor = m_FontColor
'End Property

'Public Property Let FontColor(ByVal New_FontColor As OLE_COLOR)
'     m_FontColor = New_FontColor
'    PropertyChanged "FontColor"
'End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1,Label1,-1,BackColor
Public Property Get TooltipBkColor() As OLE_COLOR
Attribute TooltipBkColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    TooltipBkColor = Label1.BackColor
End Property

Public Property Let TooltipBkColor(ByVal New_TooltipBkColor As OLE_COLOR)
    Label1.BackColor() = New_TooltipBkColor
    PropertyChanged "TooltipBkColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1,Label1,-1,ForeColor
Public Property Get TooltipForeColor() As OLE_COLOR
Attribute TooltipForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    TooltipForeColor = Label1.ForeColor
End Property

Public Property Let TooltipForeColor(ByVal New_TooltipForeColor As OLE_COLOR)
    Label1.ForeColor() = New_TooltipForeColor
    PropertyChanged "TooltipForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=12,0,0,1
Public Property Get XUnitToPixels() As Single
Attribute XUnitToPixels.VB_ProcData.VB_Invoke_Property = "XScale"
    XUnitToPixels = m_XUnitToPixels
End Property

Public Property Let XUnitToPixels(ByVal New_XUnitToPixels As Single)
    m_XUnitToPixels = New_XUnitToPixels
    
    PropertyChanged "XUnitToPixels"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=12,0,0,1
Public Property Get YUnitToPixels() As Single
Attribute YUnitToPixels.VB_ProcData.VB_Invoke_Property = "YScale"
    YUnitToPixels = m_YUnitToPixels
End Property

Public Property Let YUnitToPixels(ByVal New_YUnitToPixels As Single)
    m_YUnitToPixels = New_YUnitToPixels
    
    PropertyChanged "YUnitToPixels"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=12,0,0,0
Public Property Get XGap() As Single
Attribute XGap.VB_ProcData.VB_Invoke_Property = "XScale"
    XGap = m_XGap
End Property

Public Property Let XGap(ByVal New_XGap As Single)
  If (New_XGap > 0#) Then
    m_XGap = New_XGap
    
    PropertyChanged "XGap"
    
    End If
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=12,0,0,0
Public Property Get YGap() As Single
Attribute YGap.VB_ProcData.VB_Invoke_Property = "YScale"
    YGap = m_YGap
End Property

Public Property Let YGap(ByVal New_YGap As Single)
    If (New_YGap > 0#) Then
    m_YGap = New_YGap
    
    PropertyChanged "YGap"
    
    End If
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=12,0,0,0
Public Property Get XMinPosition() As Single
Attribute XMinPosition.VB_ProcData.VB_Invoke_Property = "XScale"
    XMinPosition = m_XMinPosition
End Property

Public Property Let XMinPosition(ByVal New_XMinPosition As Single)
'    If (New_XMinPosition = 0#) Or (New_XMinPosition > 0#) Then
    m_XMinPosition = New_XMinPosition
    p_MinX = m_XMinPosition
    If p_MaxX < p_MinX Then p_MaxX = p_MinX
    PropertyChanged "XMinPosition"
    
   ' End If
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=12,0,0,0
Public Property Get YMinPosition() As Single
Attribute YMinPosition.VB_ProcData.VB_Invoke_Property = "YScale"
    YMinPosition = m_YMinPosition
End Property

Public Property Let YMinPosition(ByVal New_YMinPosition As Single)
   ' If (New_YMinPosition = 0#) Or (New_YMinPosition > 0#) Then
    m_YMinPosition = New_YMinPosition
    
    p_MinY = m_YMinPosition
    If p_MaxY < p_MinY Then p_MaxY = p_MinY
    PropertyChanged "YMinPosition"
    
    'End If
End Property
Public Property Get XScaleHeight() As Long
Attribute XScaleHeight.VB_ProcData.VB_Invoke_Property = "XScale"
           XScaleHeight = p_XScaleHeight
End Property
Public Property Let XScaleHeight(ByVal New_XScaleHeight As Long)
        If (New_XScaleHeight = 0#) Or (New_XScaleHeight > 0#) Then
        p_XScaleHeight = New_XScaleHeight
        PropertyChanged "XScaleHeight"
       
        End If
End Property
Public Property Get YScaleWidth() As Long
Attribute YScaleWidth.VB_ProcData.VB_Invoke_Property = "YScale"
         YScaleWidth = p_YScaleWidth
End Property
Public Property Let YScaleWidth(ByVal new_yscalewidth As Long)
        If (new_yscalewidth = 0#) Or (new_yscalewidth > 0#) Then
        p_YScaleWidth = new_yscalewidth
        PropertyChanged "YScaleWidth"
        
        End If
End Property
Public Property Get TrackMousePointer() As Boolean
Attribute TrackMousePointer.VB_ProcData.VB_Invoke_Property = "Options"
     TrackMousePointer = m_TrackPointer
End Property
Public Property Let TrackMousePointer(ByVal new_val As Boolean)
        m_TrackPointer = new_val
        Label1.Visible = False
        PropertyChanged "TrackMousePointer"
End Property
Public Property Get ShowGrids() As Boolean
Attribute ShowGrids.VB_ProcData.VB_Invoke_Property = "Options"
        ShowGrids = m_ShowGrids
End Property
Public Property Let ShowGrids(ByVal new_val As Boolean)
           m_ShowGrids = new_val
           PropertyChanged "ShowGrids"
End Property
Public Property Get GridColor() As OLE_COLOR
         GridColor = m_GridColor
End Property
Public Property Let GridColor(ByVal new_val As OLE_COLOR)
             m_GridColor = new_val
             PropertyChanged "GridColor"
End Property
Public Property Get PictureFileName() As String
    PictureFileName = m_PictureFileName
End Property

Public Property Let PictureFileName(ByVal New_PictureFileName As String)
    m_PictureFileName = New_PictureFileName
    On Error GoTo lab1
    Set Picture1.Picture = LoadPicture(New_PictureFileName)
    GoTo lab2
lab1:
    m_PictureFileName = "Error:Unable to load picture"
lab2:
    PropertyChanged "PictureFileName"
End Property

Public Property Get MousePointer() As Integer
Attribute MousePointer.VB_ProcData.VB_Invoke_Property = "Options"
    MousePointer = Picture1.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As Integer)
    Picture1.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

Private Sub HScroll1_Change()
m_XInitialMargin = HScroll1.Value / Me.XUnitToPixels + p_MinX - Me.XMinPosition
Label1.Visible = False
UserControl_Paint
End Sub

Private Sub HScroll1_Scroll()
m_XInitialMargin = HScroll1.Value / Me.XUnitToPixels + p_MinX - Me.XMinPosition
Label1.Visible = False
UserControl_Paint
End Sub

Private Sub mnuCopy_Click()
Dim tempP As Picture
Set tempP = Picture1.Picture
Picture1.Picture = Picture1.Image
Clipboard.Clear
Clipboard.SetData Picture1.Picture, vbCFBitmap
Set Picture1.Picture = tempP
UserControl_Paint
End Sub

Private Sub mnuCopyAll_Click()
Dim i As Long
Dim hhh As Long
Dim tmpX As Single
Dim tmpY As Single

Dim tmpP As Picture
Label2.Visible = True
Picture2.Picture = LoadPicture
Picture2.AutoRedraw = True
Picture2.Width = (p_MaxX - p_MinY + 50) * Me.XUnitToPixels             '''HScroll1.Max

Picture2.Height = (p_MaxY - p_MinY + 50) * Me.YUnitToPixels     ''  VScroll1.Max

Set tmpP = Picture2.Picture
tmpX = HScroll1.Value
tmpY = VScroll1.Value
Picture2.AutoRedraw = True
Picture2.Font = Me.Font
Picture2.ForeColor = Me.ForeColor
Picture2.BackColor = Me.BackColor
Picture2.Cls
HScroll1.Value = HScroll1.Min
VScroll1.Value = VScroll1.Max

Call drawHScale(Picture2.hdc, p_YScaleWidth, 0, p_XScaleHeight, Picture2.Width - p_YScaleWidth - 5)
Call drawVScale(Picture2.hdc, 0, p_XScaleHeight, Picture2.Height - p_XScaleHeight - 5, p_YScaleWidth)
If Me.ShowGrids Then
Call ShowHGrid(Picture2.hdc, p_YScaleWidth, p_XScaleHeight, Picture2.Height - p_XScaleHeight - 6, Picture2.Width - p_YScaleWidth - 6)
Call ShowVGrid(Picture2.hdc, p_YScaleWidth, p_XScaleHeight, Picture2.Height - p_XScaleHeight - 6, Picture2.Width - p_YScaleWidth - 6)
End If
For i = 0 To MAX_NO_OF_GRAPHS
Call drawGraph(Picture2.hdc, p_YScaleWidth, p_XScaleHeight, Picture2.Height - p_XScaleHeight - 6, Picture2.Width - p_YScaleWidth - 6, i)
Next i
Picture2.Picture = Picture2.Image
Clipboard.Clear
Clipboard.SetData Picture2.Picture, vbCFBitmap

Picture2.Visible = False

Set Picture2.Picture = tmpP
HScroll1.Value = tmpX
VScroll1.Value = tmpY

UserControl_Paint
Label2.Visible = False
End Sub

Private Sub mnuPrintasBitmap_Click()
Me.PrintOnlyShownPortionAsBitmap
End Sub

Private Sub mnuPrintMe_Click()
Me.PrintMe
End Sub

Private Sub mnuSendToExcel_Click()
Me.SendToExcel
End Sub

Private Sub mnuShowGrids_Click()
If mnuShowGrids.Checked = True Then
    mnuShowGrids.Checked = False
    Me.ShowGrids = False
    Me.Invalidate
Else
    mnuShowGrids.Checked = True
    Me.ShowGrids = True
    Me.Invalidate
End If
End Sub

Private Sub mnuTrackMouse_Click()
If mnuTrackMouse.Checked = True Then
         mnuTrackMouse.Checked = False
         Me.TrackMousePointer = False
Else
        mnuTrackMouse.Checked = True
        Me.TrackMousePointer = True
End If
End Sub

Private Sub Picture1_LostFocus()
Label1.Visible = False
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'Call drawGraph(Picture1.hDC, p_YScaleWidth + 1, p_XScaleHeight + 1, Picture1.Height - p_XScaleHeight - 6, Picture1.Width - p_YScaleWidth - 6, i)
Dim xx As Single
Dim yy As Single
Dim XFirstPoint As Single
Dim YFirstpoint As Single
Dim str1 As String
Dim str2 As String

If TrackMousePointer Then
     If (x > (p_YScaleWidth - 1)) And (y > (p_XScaleHeight - 1)) And (y < (Picture1.Height)) And (x < (Picture1.Width)) Then
          Label1.Visible = False
          XFirstPoint = Me.XMinPosition + m_XInitialMargin
          YFirstpoint = Me.YMinPosition + m_YInitialMargin
          If (Me.XUnitToPixels <> 0) And (Me.YUnitToPixels <> 0) Then
              xx = CSng(CSng(x - p_YScaleWidth + 0) / Me.XUnitToPixels + XFirstPoint)
              yy = CSng(CSng(Picture1.Height - 5 - y) / Me.YUnitToPixels + YFirstpoint)
          End If
          str1 = Format(xx, "##0.00")
          str2 = Format(yy, "##0.00")
          Label1.Caption = "( " & str1 & ", " & str2 & ")"
          If (Picture1.Width - x - 16 - Label1.Width) > 0 Then
                   Label1.Left = x + 16
          Else
                   Label1.Left = x - Label1.Width - 1
          End If
          If (Picture1.Height - y - 16 - Label1.Height - 5) > 0 Then
                   Label1.Top = y + 16
          Else
                   Label1.Top = y - Label1.Height - 1
          End If
          Label1.Visible = True
                   
     Else
          Label1.Visible = False
     End If
End If
End Sub



Private Sub UserControl_ExitFocus()
Label1.Visible = False
End Sub

Private Sub UserControl_Initialize()
      If (m_XUnitToPixels = 0) Or (m_XUnitToPixels < 0) Then m_XUnitToPixels = 1
      If (m_YUnitToPixels = 0) Or (m_YUnitToPixels < 0) Then m_YUnitToPixels = 1
      If (m_XGap = 0) Or (m_XGap < 0) Then m_XGap = m_def_XGap
      If (m_YGap = 0) Or (m_YGap < 0) Then m_YGap = m_def_YGap
      If m_XMinPosition < 0 Then m_XMinPosition = m_def_XMinPosition
      If m_YMinPosition < 0 Then m_YMinPosition = m_def_YMinPosition
      Picture1.Left = 0
      Picture1.Top = 0
      HScroll1.Left = 0
      VScroll1.Top = 0
      HScroll1.Top = Me.ScaleHeight - HScroll1.Height
      VScroll1.Left = Me.ScaleWidth - VScroll1.Width
      HScroll1.Width = Me.ScaleWidth - VScroll1.Width
      VScroll1.Height = Me.ScaleHeight - HScroll1.Height
      Picture1.Width = Me.ScaleWidth - VScroll1.Width
      Picture1.Height = Me.ScaleHeight - HScroll1.Height
      
      initializeGraph
      initializeScrolls
      Label1.Visible = False
      Label2.Visible = False
      excel_Header = "Data comes from AshuGraph Control"
     
      
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
     m_FontColor = RGB(0, 0, 255)
    m_XUnitToPixels = m_def_XUnitToPixels
    m_YUnitToPixels = m_def_YUnitToPixels
    m_XGap = m_def_XGap
    m_YGap = m_def_YGap
    m_XMinPosition = m_def_XMinPosition
    m_YMinPosition = m_def_YMinPosition
    p_XScaleHeight = 20
    p_YScaleWidth = 40
    m_TrackPointer = True
    m_ShowGrids = True
    m_PictureFileName = m_def_PictureFileName
    m_GridColor = QBColor(14)
End Sub

Private Sub UserControl_LostFocus()
Label1.Visible = False
End Sub

Private Sub UserControl_Paint()
Dim i As Long
Picture1.Cls
'Picture1.Refresh

Call drawHScale(Picture1.hdc, p_YScaleWidth, 0, p_XScaleHeight, Picture1.Width - p_YScaleWidth - 5)
Call drawVScale(Picture1.hdc, 0, p_XScaleHeight, Picture1.Height - p_XScaleHeight - 5, p_YScaleWidth)
If Me.ShowGrids Then
Call ShowHGrid(Picture1.hdc, p_YScaleWidth, p_XScaleHeight, Picture1.Height - p_XScaleHeight - 6, Picture1.Width - p_YScaleWidth - 6)
Call ShowVGrid(Picture1.hdc, p_YScaleWidth, p_XScaleHeight, Picture1.Height - p_XScaleHeight - 6, Picture1.Width - p_YScaleWidth - 6)
End If
For i = 0 To MAX_NO_OF_GRAPHS
Call drawGraph(Picture1.hdc, p_YScaleWidth, p_XScaleHeight, Picture1.Height - p_XScaleHeight - 6, Picture1.Width - p_YScaleWidth - 6, i)
Next i

End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.Appearance = PropBag.ReadProperty("Appearance", 1)
    Picture1.BackColor = PropBag.ReadProperty("BackColor", &HE0E0E0)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set Picture1.Font = PropBag.ReadProperty("Font", Ambient.Font)
    Picture1.ForeColor = PropBag.ReadProperty("ForeColor", &HFF0000)
    UserControl.ScaleHeight = PropBag.ReadProperty("ScaleHeight", 2880)
    UserControl.ScaleWidth = PropBag.ReadProperty("ScaleWidth", 3840)
     m_FontColor = PropBag.ReadProperty("FontColor", RGB(0, 0, 255))
    Label1.BackColor = PropBag.ReadProperty("TooltipBkColor", &H80000005)
    Label1.ForeColor = PropBag.ReadProperty("TooltipForeColor", &H80000001)
    m_XUnitToPixels = PropBag.ReadProperty("XUnitToPixels", m_def_XUnitToPixels)
    m_YUnitToPixels = PropBag.ReadProperty("YUnitToPixels", m_def_YUnitToPixels)
    m_XGap = PropBag.ReadProperty("XGap", m_def_XGap)
    m_YGap = PropBag.ReadProperty("YGap", m_def_YGap)
    m_XMinPosition = PropBag.ReadProperty("XMinPosition", m_def_XMinPosition)
    m_YMinPosition = PropBag.ReadProperty("YMinPosition", m_def_YMinPosition)
    p_MinX = Me.XMinPosition
    p_MinY = Me.YMinPosition
    p_MaxX = p_MinX
    p_MaxY = p_MinY
    p_XScaleHeight = PropBag.ReadProperty("XScaleHeight", 20)
    p_YScaleWidth = PropBag.ReadProperty("YScaleWidth", 40)
    m_TrackPointer = PropBag.ReadProperty("TrackMousePointer", True)
    m_ShowGrids = PropBag.ReadProperty("ShowGrids", True)
    m_GridColor = PropBag.ReadProperty("GridColor", QBColor(14))
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
    Picture1.MousePointer = PropBag.ReadProperty("MousePointer", 0)
End Sub

Private Sub UserControl_Resize()
      HScroll1.Top = Me.ScaleHeight - HScroll1.Height
      VScroll1.Left = Me.ScaleWidth - VScroll1.Width
      HScroll1.Width = Me.ScaleWidth - VScroll1.Width
      VScroll1.Height = Me.ScaleHeight - HScroll1.Height
      Picture1.Width = Me.ScaleWidth - VScroll1.Width
      Picture1.Height = Me.ScaleHeight - HScroll1.Height
End Sub

Private Sub UserControl_Terminate()
Dim i As Long
For i = 0 To ((i < MAX_NO_OF_GRAPHS) Or (i = MAX_NO_OF_GRAPHS))
Set m_Graph(i) = Nothing
Next i
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Appearance", UserControl.Appearance, 1)
    Call PropBag.WriteProperty("BackColor", Picture1.BackColor, &HE0E0E0)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 0)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", Picture1.Font, Ambient.Font)
    Call PropBag.WriteProperty("ForeColor", Picture1.ForeColor, &HFF0000)
    Call PropBag.WriteProperty("ScaleHeight", UserControl.ScaleHeight, 2880)
    Call PropBag.WriteProperty("ScaleWidth", UserControl.ScaleWidth, 3840)
    Call PropBag.WriteProperty("FontColor", m_FontColor, RGB(0, 0, 255))
    Call PropBag.WriteProperty("TooltipBkColor", Label1.BackColor, &H80000005)
    Call PropBag.WriteProperty("TooltipForeColor", Label1.ForeColor, &H80000001)
    Call PropBag.WriteProperty("XUnitToPixels", m_XUnitToPixels, m_def_XUnitToPixels)
    Call PropBag.WriteProperty("YUnitToPixels", m_YUnitToPixels, m_def_YUnitToPixels)
    Call PropBag.WriteProperty("XGap", m_XGap, m_def_XGap)
    Call PropBag.WriteProperty("YGap", m_YGap, m_def_YGap)
    Call PropBag.WriteProperty("XMinPosition", m_XMinPosition, m_def_XMinPosition)
    Call PropBag.WriteProperty("YMinPosition", m_YMinPosition, m_def_YMinPosition)
    Call PropBag.WriteProperty("XScaleHeight", p_XScaleHeight, 20)
    Call PropBag.WriteProperty("YScaleWidth", p_YScaleWidth, 40)
    Call PropBag.WriteProperty("TrackMousePointer", m_TrackPointer, True)
    Call PropBag.WriteProperty("ShowGrids", m_ShowGrids, True)
    Call PropBag.WriteProperty("GridColor", m_GridColor, QBColor(14))
    Call PropBag.WriteProperty("Picture", Picture, Nothing)
    Call PropBag.WriteProperty("MousePointer", Picture1.MousePointer, 0)
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''funtions for initializing
Public Sub SetColor(ByVal color As Long, ByVal GraphNo As Long)
If (GraphNo < 0) Or (GraphNo > MAX_NO_OF_GRAPHS) Then Exit Sub
m_Graph(GraphNo).color = color


End Sub
Private Sub initializeGraph()
Dim i As Long
Dim tmp As Long
m_XInitialMargin = 0#
m_YInitialMargin = 0#
p_MinX = Me.XMinPosition
p_MinY = Me.YMinPosition
p_MaxX = p_MinX
p_MaxY = p_MinY
p_NumCols = 1
p_NumRows = 1
p_PageHeight = 1
p_PageWidth = 1
p_XGraphSize = 1
p_XInitialMargin = 0#
p_XPageMargin = 0
p_XScaleHeight = 20
p_YGraphSize = 1
p_YInitialMargin = 0#
p_YPagemargin = 0
p_YScaleWidth = 40
p_zoomfactor = 100
 p_Header = ""
      p_XLabel = ""
      p_YLabel = ""
      p_PrintPageNo = 0
End Sub
Private Sub initializeScrolls()
Dim Size As Long
Dim size1 As Long
HScroll1.Min = 0
HScroll1.SmallChange = 1
HScroll1.LargeChange = CInt(Me.XGap * Me.XUnitToPixels)

Size = CInt((p_MaxX - p_MinX + 1) * Me.XUnitToPixels) ''- Picture1.Width
size1 = CInt((Me.XMinPosition - p_MinX) * Me.XUnitToPixels)
If (Size < 0) Or (size1 < 0) Then
     HScroll1.Max = Abs(Size) + Abs(size1)
     HScroll1.Value = 0
ElseIf Size > size1 Then
     HScroll1.Max = Size
     HScroll1.Value = size1
Else
     HScroll1.Max = size1
     HScroll1.Value = size1

End If

VScroll1.Min = 0
VScroll1.SmallChange = 1
VScroll1.LargeChange = CInt(Me.YGap * Me.YUnitToPixels)

Size = CInt((p_MaxY - p_MinY + 1) * Me.YUnitToPixels) ''- Picture1.Height
size1 = CInt((Me.YMinPosition - p_MinY) * Me.YUnitToPixels)
If (size1 < 0) Or (Size < 0) Then
VScroll1.Max = Abs(Size) + Abs(size1)
VScroll1.Value = VScroll1.Max
ElseIf Size > size1 Then
 VScroll1.Max = Size
 VScroll1.Value = VScroll1.Max - size1
 Else
 VScroll1.Max = size1
 VScroll1.Value = VScroll1.Max - size1
 End If
 

End Sub
Public Sub InitializeMe()
Dim i As Long
For i = 0 To MAX_NO_OF_GRAPHS
m_Graph(i).index = -1
m_Graph(i).Total_Data = 0
Next i
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''Function for data manipulation
Public Sub AddData(ByVal xx As Single, ByVal yy As Single, ByVal GraphNo As Long)
If (GraphNo < 0) Or (GraphNo > MAX_NO_OF_GRAPHS) Then Exit Sub
m_Graph(GraphNo).AddData xx, yy
If p_MaxX < xx Then p_MaxX = xx
If p_MaxY < yy Then p_MaxY = yy
If p_MinX > xx Then p_MinX = xx
If p_MinY > yy Then p_MinY = yy

End Sub
Public Function SetData(ByVal xx As Single, ByVal yy As Single, ByVal index As Long, ByVal GraphNo As Long) As Boolean
SetData = False
If (GraphNo < 0) Or (GraphNo > MAX_NO_OF_GRAPHS) Then Exit Function
If (m_Graph(GraphNo).SetPoint(xx, yy, index)) Then
If p_MaxX < xx Then p_MaxX = xx
If p_MaxY < yy Then p_MaxY = yy
If p_MinX > xx Then p_MinX = xx
If p_MinY > yy Then p_MinY = yy
'HScroll1.Max = CInt((p_MaxX - p_MinX + 1) * Me.XUnitToPixels) ''- Picture1.Width
'VScroll1.Max = CInt((p_MaxY - p_MinY + 1) * Me.YUnitToPixels)
'VScroll1.Value = VScroll1.Max - CInt((Me.YMinPosition - p_MinY) * Me.YUnitToPixels)
'HScroll1.Value = CInt((Me.XMinPosition - p_MinX) * Me.XUnitToPixels)
SetData = True
End If
End Function
Public Function GetData(ByRef xx As Single, ByRef yy As Single, ByRef index As Long, GraphNo As Long) As Boolean
GetData = False
If (GraphNo < 0) Or (GraphNo > MAX_NO_OF_GRAPHS) Then Exit Function
If (m_Graph(GraphNo).GetPoint(xx, yy, index)) Then GetData = True

End Function
Public Sub RemoveGraph(ByVal GraphNo As Long)
If (GraphNo < 0) Or (GraphNo > MAX_NO_OF_GRAPHS) Then Exit Sub
m_Graph(GraphNo).index = -1
m_Graph(GraphNo).Total_Data = 0
m_Graph(GraphNo).DrawStyle = 0
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''Private Functions for drawing graph and scales
Private Sub drawGraph(ByRef hdc As Long, Xdisplacement As Long, Ydisplacement As Long, hHeight As Long, wWidth As Long, GraphNo As Long)
Dim i As Long
Dim xx As Single
Dim yy As Single
Dim XFirstPoint As Single
Dim YFirstpoint As Single
Dim Xdistance As Single
Dim Ydistance As Single
Dim tmp As Long
Dim firstPoint As Boolean
Dim prevX As Long
Dim prevY As Long

If (GraphNo < 0) Or (GraphNo > MAX_NO_OF_GRAPHS) Then Exit Sub
If (m_Graph(GraphNo).Total_Data = 0) Or (m_Graph(GraphNo).Total_Data < 0) Then Exit Sub

XFirstPoint = Me.XMinPosition + m_XInitialMargin
YFirstpoint = Me.YMinPosition + m_YInitialMargin
Xdistance = 0#
Ydistance = 0#
firstPoint = True
For i = 0 To m_Graph(GraphNo).Total_Data - 1


If (m_Graph(GraphNo).GetPoint(xx, yy, i)) Then
    If ((xx > XFirstPoint) Or (xx = XFirstPoint)) And ((yy > YFirstpoint) Or (yy = YFirstpoint)) Then
          Xdistance = (xx - XFirstPoint) * Me.XUnitToPixels
          Ydistance = (yy - YFirstpoint) * Me.YUnitToPixels
          If (Xdistance < wWidth) And (Ydistance < hHeight) Then
                If m_Graph(GraphNo).DrawStyle = 0 Then
                 tmp = SetPixelV(hdc, CLng(Xdistance) + Xdisplacement, hHeight + Ydisplacement - CLng(Ydistance), m_Graph(GraphNo).color)
                 End If
                 If (m_Graph(GraphNo).DrawStyle = 1) Or (m_Graph(GraphNo).DrawStyle = 3) Then
                 '''connect through lines
                       If (firstPoint) Then
                               prevX = CLng(Xdistance) + Xdisplacement
                               prevY = hHeight + Ydisplacement - CLng(Ydistance)
                               firstPoint = False
                        Else
                            Call DrawLines(hdc, prevX, prevY, CLng(Xdistance) + Xdisplacement, hHeight + Ydisplacement - CLng(Ydistance), m_Graph(GraphNo).color)
                            prevX = CLng(Xdistance) + Xdisplacement
                            prevY = hHeight + Ydisplacement - CLng(Ydistance)
                              
                        End If
                  End If
                  
                 If (m_Graph(GraphNo).DrawStyle = 2) Or (m_Graph(GraphNo).DrawStyle = 3) Then
                             Call DrawSquarePoint(hdc, CLng(Xdistance) + Xdisplacement, hHeight + Ydisplacement - CLng(Ydistance), m_Graph(GraphNo).color)
                  End If
                 'Picture1.PSet (CLng(Xdistance) + Xdisplacement, hHeight + Ydisplacement - CLng(Ydistance)), RGB(255, 0, 0)
          Else
                firstPoint = True
          End If
    
    Else
     firstPoint = True
   End If
End If
Next i


End Sub
Private Sub drawHScale(hdc As Long, Xdisplacement As Long, Ydisplacement As Long, hHeight As Long, wWidth As Long)
Dim ptAPI As POINTAPI

Dim tmp As Long
Dim tmp1 As Long
Dim oldColor As Long
Dim str As String
Dim firstPoint As Single
Dim distance As Single
Dim i As Single

''''calculate height of font
tmp = -1 * MulDiv(Me.Font.Size, GetDeviceCaps(hdc, 90), 72)
tmp = Abs(tmp)

'''''''draw boundary lines
'tmp1 = MoveToEx(hdc, Xdisplacement, Ydisplacement + hHeight, ptAPI)
'tmp1 = LineTo(hdc, Xdisplacement + wWidth, Ydisplacement + hHeight)
'tmp1 = LineTo(hdc, Xdisplacement + wWidth, Ydisplacement + tmp)

firstPoint = Me.XMinPosition + m_XInitialMargin
distance = 0#
i = firstPoint

Do
   distance = (i - firstPoint) * Me.XUnitToPixels
   If distance > CSng(wWidth) Then Exit Do
   str = Format(i, "#0.00")
   tmp1 = TextOut(hdc, Xdisplacement + CLng(distance) + 1, Ydisplacement, str, Len(str))
   'tmp1 = MoveToEx(hdc, Xdisplacement + CLng(distance), Ydisplacement + tmp, ptAPI)
   'tmp1 = LineTo(hdc, Xdisplacement + CLng(distance), Ydisplacement + hHeight)
   DrawLines hdc, Xdisplacement + CLng(distance), Ydisplacement + tmp, Xdisplacement + CLng(distance), Ydisplacement + hHeight, Me.ForeColor
   i = i + Me.XGap
   
Loop While distance < wWidth
'''''''draw boundary lines
'''DrawLines hdc, Xdisplacement, Ydisplacement, Xdisplacement + wWidth, Ydisplacement, Me.ForeColor
DrawLines hdc, Xdisplacement, Ydisplacement + hHeight, Xdisplacement + wWidth, Ydisplacement + hHeight, Me.ForeColor
DrawLines hdc, Xdisplacement + wWidth, Ydisplacement + hHeight, Xdisplacement + wWidth, Ydisplacement + tmp, Me.ForeColor
End Sub



Private Sub drawVScale(hdc As Long, Xdisplacement As Long, Ydisplacement As Long, hHeight As Long, wWidth As Long)
Dim ptAPI As POINTAPI

Dim tmp As Long
Dim oldColor As Long
Dim str As String
Dim firstPoint As Single
Dim distance As Single
Dim i As Single
Dim lowerpoint As Boolean

''''calculate height of font
tmp = -1 * MulDiv(Me.Font.Size, GetDeviceCaps(hdc, 90), 72)
tmp = Abs(tmp)

'''''''draw boundary lines
'Call MoveToEx(hdc, Xdisplacement + wWidth, Ydisplacement, ptAPI)
'Call LineTo(hdc, Xdisplacement + wWidth, Ydisplacement + hHeight)
'Call MoveToEx(hdc, Xdisplacement + wWidth, Ydisplacement, ptAPI)
'Call LineTo(hdc, Xdisplacement + 2 * temp, Ydisplacement)
DrawLines hdc, Xdisplacement + wWidth, Ydisplacement, Xdisplacement + wWidth, Ydisplacement + hHeight, Me.ForeColor
DrawLines hdc, Xdisplacement + wWidth, Ydisplacement, Xdisplacement + 2 * temp, Ydisplacement, Me.ForeColor
firstPoint = Me.YMinPosition + m_YInitialMargin
distance = 0#
i = firstPoint
lowerpoint = True
Do
   distance = (i - firstPoint) * Me.YUnitToPixels
   If distance > CSng(hHeight) Then Exit Do
   str = Format(i, "#0.00")
   'Call MoveToEx(hdc, Xdisplacement + 2 * tmp, hHeight - CLng(distance) + Ydisplacement, ptAPI)
   'Call LineTo(hdc, Xdisplacement + wWidth, hHeight - CLng(distance) + Ydisplacement)
   DrawLines hdc, Xdisplacement + 2 * tmp, hHeight - CLng(distance) + Ydisplacement, Xdisplacement + wWidth, hHeight - CLng(distance) + Ydisplacement, Me.ForeColor
   If lowerpoint = True Then
   Call TextOut(hdc, Xdisplacement, hHeight - tmp - 2 + Ydisplacement, str, Len(str))
   lowerpoint = False
   Else
   Call TextOut(hdc, Xdisplacement, hHeight - CLng(distance) + Ydisplacement, str, Len(str))
   End If
   
   i = i + Me.YGap
   
Loop While distance < hHeight


End Sub

Public Sub SineCurveFill(ByVal GraphNo As Long)
Dim i As Single
Dim p As Single
Dim Q As Single
If (GraphNo < 0) Or (GraphNo > MAX_NO_OF_GRAPHS) Then Exit Sub

For i = 0 To 30 Step 0.01
p = 100 + 100 * CSng(Sin(CDbl(i)))
Q = 10 * i
Call AddData(Q, p, GraphNo)
Next i
End Sub

Private Sub VScroll1_Change()
m_YInitialMargin = (VScroll1.Max - VScroll1.Value) / Me.YUnitToPixels + p_MinY - Me.YMinPosition
Label1.Visible = False
UserControl_Paint
End Sub

Private Sub VScroll1_Scroll()
m_YInitialMargin = (VScroll1.Max - VScroll1.Value) / Me.YUnitToPixels + p_MinY - Me.YMinPosition
Label1.Visible = False
UserControl_Paint
End Sub
Private Sub ShowHGrid(hdc As Long, Xdisplacement As Long, Ydisplacement As Long, hHeight As Long, wWidth As Long)
Dim ptAPI As POINTAPI
Dim str As String
Dim distance As Single
Dim i As Single
Dim tmp As Integer
Dim tmp1 As Long


distance = 0#
i = Me.XGap
tmp = Picture1.DrawStyle
tmp1 = Picture1.ForeColor
Picture1.DrawStyle = vbDot
Picture1.ForeColor = Me.GridColor
Do
   
   distance = (i) * Me.XUnitToPixels
   If distance > CSng(wWidth) Then Exit Do
   
   'Call MoveToEx(hdc, Xdisplacement + CLng(distance), Ydisplacement, ptAPI)
   'Call LineTo(hdc, Xdisplacement + CLng(distance), Ydisplacement + hHeight)
   DrawLines hdc, Xdisplacement + CLng(distance), Ydisplacement, Xdisplacement + CLng(distance), Ydisplacement + hHeight, Me.GridColor
   i = i + Me.XGap
   
Loop While distance < CSng(wWidth)

Picture1.DrawStyle = tmp
Picture1.ForeColor = tmp1
End Sub
Private Sub ShowVGrid(hdc As Long, Xdisplacement As Long, Ydisplacement As Long, hHeight As Long, wWidth As Long)
Dim ptAPI As POINTAPI
Dim str As String
Dim distance As Single
Dim i As Single
Dim tmp As Integer
Dim tmp1 As Long


distance = 0#
i = 0
tmp = Picture1.DrawStyle
tmp1 = Picture1.ForeColor
Picture1.DrawStyle = vbDot
Picture1.ForeColor = Me.GridColor
Do
   
   distance = (i) * Me.YUnitToPixels
   If distance > CSng(hHeight) Then Exit Do
   
   'Call MoveToEx(hdc, Xdisplacement + 1, hHeight - CLng(distance) + Ydisplacement, ptAPI)
   'Call LineTo(hdc, Xdisplacement + wWidth, hHeight - CLng(distance) + Ydisplacement)
    DrawLines hdc, Xdisplacement + 1, hHeight - CLng(distance) + Ydisplacement, Xdisplacement + wWidth, hHeight - CLng(distance) + Ydisplacement, Me.GridColor
   i = i + Me.YGap
   
Loop While distance < hHeight
Picture1.DrawStyle = tmp
Picture1.ForeColor = tmp1
End Sub
Public Sub PrintMe()
Dim dDrawHeight As Long
Dim dDrawWidth As Long
Dim xextra As Long
Dim yextra As Long
Dim negyextra As Long
Dim Totalpages As Long
Dim i As Long
Dim j As Long
Dim str1 As String
Dim p_PScaleWidth As Long
Dim curRow As Long
Dim curCol As Long
Dim g_XGap As Single
Dim g_YGap As Single
''''' first set all parameters
'' set mode landscape
'PrintOrient (2)
''' set font and forecolor
'Printer.Font = Me.Font
Printer.Font.Size = 5
Printer.ForeColor = Me.ForeColor
Call SetTextColor(Printer.hdc, Me.ForeColor)
'''set zoom factor
Printer.Zoom = p_zoomfactor
Printer.ScaleMode = 3 '''pixel
''''show busy label
'''show busy label
    i = (Picture1.Width - Label2.Width) / 2
    If i > 0 Then
          Label2.Left = i
    Else
          Label2.Left = 0
    End If
    i = (Picture1.Height - Label2.Height) / 2
    If i > 0 Then
           Label2.Top = i
    Else
           Label2.Top = 0
    End If
    Label2.Visible = True


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''' OnBeginPrinting
Printer.Orientation = 2
'''calculate page width and height
p_PageWidth = GetDeviceCaps(Printer.hdc, 8)
p_PageHeight = GetDeviceCaps(Printer.hdc, 10)
If (p_PageHeight = 0) Or (p_PageWidth = 0) Then
   MsgBox "Error in Printing, Can't Print::No default Printer"
   Exit Sub
End If
  ''' calculate only dimensions of graph without calculating others factors
dDrawWidth = CLng((p_MaxX - p_MinX + 1) * m_XUnitToPixels)
dDrawHeight = CLng((p_MaxY - p_MinY + 1) * m_YUnitToPixels)

''''now calculate  no of pages
p_NumRows = dDrawHeight / p_PageHeight + 1
p_NumCols = dDrawWidth / p_PageWidth + 1
''' now take all in effect
p_PScaleWidth = 100 ''5 * Printer.Font.Size
If p_PScaleWidth < p_YScaleWidth Then p_PScaleWidth = p_YScaleWidth
If p_PScaleWidth < p_XScaleHeight Then p_PScaleWidth = p_XScaleHeight
xextra = p_PScaleWidth + p_XPageMargin
yextra = p_PScaleWidth + p_YPagemargin
If p_Header <> "" Then yextra = yextra + 5 * Printer.Font.Size
If p_XLabel <> "" Then yextra = yextra + 5 * Printer.Font.Size
If p_PrintPageNo <> 0 Then yextra = yextra + 5 * Printer.Font.Size
If p_YLabel <> "" Then xextra = xextra + 5 * Printer.Font.Size

dDrawHeight = dDrawHeight + p_NumRows * p_NumCols * (yextra)
dDrawWidth = dDrawWidth + p_NumCols * p_NumRows * xextra

''''now recalculate  no of pages
If (dDrawHeight Mod p_PageHeight) > 0 Then
p_NumRows = dDrawHeight / p_PageHeight + 1
Else
p_NumRows = dDrawHeight / p_PageHeight
End If
If (dDrawWidth Mod p_PageWidth) > 0 Then
p_NumCols = dDrawWidth / p_PageWidth + 1
Else
 p_NumCols = dDrawWidth / p_PageWidth
 End If
 
''''calculate the graph size on each page
p_XGraphSize = p_PageWidth - xextra  ''' size of x axis
p_YGraphSize = p_PageHeight - yextra    ''' size of y axis

p_XInitialMargin = m_XInitialMargin  ''' save the values
p_YInitialMargin = m_YInitialMargin
Totalpages = p_NumRows * p_NumCols

'''''''''end of OnBeginPrinting
g_XGap = m_XGap
g_YGap = m_YGap
m_XGap = 2 * m_XGap
m_YGap = 2 * m_YGap
''''now start printing
For i = 1 To (Totalpages)

Printer.Print ""
Printer.Font = Me.Font
Printer.ForeColor = Me.ForeColor

xextra = p_XPageMargin
yextra = p_YPagemargin
negyextra = 0

'''' draw header,labels and pageno
If p_PrintPageNo = 1 Then
       Call TextOut(Printer.hdc, CLng(p_PageWidth / 2), yextra, (str(i)), Len((str(i))))
       yextra = yextra + 5 * Printer.Font.Size
ElseIf p_PrintPageNo = 2 Then
       negyextra = negyextra + 5 * Printer.Font.Size
       Call TextOut(Printer.hdc, CLng(p_PageWidth / 2), p_PageHeight - negyextra, str(i), Len(str(i)))
End If

If p_Header <> "" Then
      Call TextOut(Printer.hdc, CLng(Abs((p_PageWidth - Len(p_Header)) / 2)), yextra, p_Header, Len(p_Header))
      yextra = yextra + 5 * Printer.Font.Size
End If
If p_XLabel <> "" Then
       negyextra = negyextra + 5 * Printer.Font.Size
      Call TextOut(Printer.hdc, CLng(Abs(p_PageWidth - Len(p_XLabel) - xextra) / 2 + xextra), p_PageHeight - negyextra, p_XLabel, Len(p_XLabel))
      
End If
If p_YLabel <> "" Then
     For j = 1 To Len(p_YLabel)
     str1 = CStr(Mid(p_YLabel, j, 1))
     Call TextOut(Printer.hdc, CLng(xextra), yextra + j * 5 * Printer.Font.Size, str1, Len(str1))
     Next j
     xextra = xextra + 5 * Printer.Font.Size
End If
''''''Now draw scales
Call drawLowerHScale(Printer.hdc, p_PScaleWidth + xextra, p_PageHeight - negyextra - p_PScaleWidth, p_PScaleWidth, p_PageWidth - xextra - p_PScaleWidth)
Call drawVScale(Printer.hdc, xextra, yextra, p_PageHeight - negyextra - p_PScaleWidth - yextra, p_PScaleWidth)
If Me.ShowGrids Then
Call ShowHGrid(Printer.hdc, p_PScaleWidth + xextra, yextra, p_PageHeight - p_PScaleWidth - negyextra - yextra, p_PageWidth - xextra - p_PScaleWidth)
Call ShowVGrid(Printer.hdc, xextra + p_PScaleWidth, yextra, p_PageHeight - p_PScaleWidth - negyextra - yextra, p_PageWidth - xextra - p_PScaleWidth)
End If
For j = 0 To MAX_NO_OF_GRAPHS
Call drawGraph(Printer.hdc, p_PScaleWidth + xextra, yextra, p_PageHeight - p_PScaleWidth - negyextra - yextra, p_PageWidth - xextra - p_PScaleWidth, j)
Next j

If (i Mod p_NumCols) > 0 Then
curRow = i / p_NumCols + 1
Else
curRow = i / p_NumCols
End If
curCol = ((i - 1) Mod p_NumCols) + 1
m_XInitialMargin = CSng((curCol - 1) * p_XGraphSize)
m_YInitialMargin = CSng((curRow - 1) * p_YGraphSize)
Printer.NewPage
Next i

Printer.EndDoc
Printer.EndDoc
'''rexstore the values
m_XInitialMargin = p_XInitialMargin
m_YInitialMargin = p_YInitialMargin
m_XGap = g_XGap
m_YGap = g_YGap
''''in the last set mode portrait
PrintOrient (1)
Label2.Visible = False
End Sub
Private Sub PrintOrient(mode As Integer)
    Dim Orient As OrientStructure
    Dim Ret As Integer
    Dim x As Integer
    Printer.Print ""
    Orient.Orientation = mode
    x = Escape(Printer.hdc, 30, Len(Orient), Orient, 0&)
    On Error Resume Next
    Ret = AbortDoc(Printer.hdc)
    On Error Resume Next
    Printer.EndDoc

End Sub
Public Sub PrintSettings(ByVal LeftMarginInPixels As Long, ByVal TopMarginInPixels As Long, ByVal Header As String, ByVal XAxisLabel As String, ByVal Yaxislabel As String, ByVal PrintPageNo As Integer, ByVal zoomFactor As Long, ByVal PaperSize As Long, ByVal printQuality As Long)
p_YPagemargin = LeftMarginInPixels  ''' since we are printing in landscape
p_XPageMargin = TopMarginInPixels
p_zoomfactor = zoomFactor
p_Header = Header
p_XLabel = XAxisLabel
p_YLabel = Yaxislabel
p_PrintPageNo = PrintPageNo  '' 0 means  no print  1 means print at top  2 means print at bottom
Printer.PaperSize = PaperSize
Printer.printQuality = printQuality

End Sub

Public Sub Invalidate()
Dim Size As Long
Dim size1 As Long
Size = CInt((p_MaxX - p_MinX + 1) * Me.XUnitToPixels) ''- Picture1.Width
size1 = CInt((Me.XMinPosition - p_MinX) * Me.XUnitToPixels)
If (Size < 0) Or (size1 < 0) Then
     HScroll1.Max = Abs(Size) + Abs(size1)
     HScroll1.Value = 0
ElseIf Size > size1 Then
     HScroll1.Max = Size
     HScroll1.Value = size1
Else
     HScroll1.Max = size1
     HScroll1.Value = size1

End If

Size = CInt((p_MaxY - p_MinY + 1) * Me.YUnitToPixels) ''- Picture1.Height
size1 = CInt((Me.YMinPosition - p_MinY) * Me.YUnitToPixels)
If (size1 < 0) Or (Size < 0) Then
VScroll1.Max = Abs(Size) + Abs(size1)
VScroll1.Value = VScroll1.Max
ElseIf Size > size1 Then
 VScroll1.Max = Size
 VScroll1.Value = VScroll1.Max - size1
 Else
 VScroll1.Max = size1
 VScroll1.Value = VScroll1.Max - size1
 End If
HScroll1.LargeChange = CInt(Me.XGap * Me.XUnitToPixels)
VScroll1.LargeChange = CInt(Me.YGap * Me.YUnitToPixels)

UserControl_Paint
End Sub
Public Sub PrintOnlyShownPortionAsBitmap()
''''''''this code is stolen from msdn
Const SRCCOPY = &HCC0020
      Const NEWFRAME = 1
      Const PIXEL = 3

'''show busy label
    i = (Picture1.Width - Label2.Width) / 2
    If i > 0 Then
          Label2.Left = i
    Else
          Label2.Left = 0
    End If
    i = (Picture1.Height - Label2.Height) / 2
    If i > 0 Then
           Label2.Top = i
    Else
           Label2.Top = 0
    End If
    Label2.Visible = True
    
Picture1.Picture = Picture1.Image
Printer.ScaleMode = PIXEL
 Printer.Print ""
 hMemoryDC% = CreateCompatibleDC(Picture1.hdc)
 hOldBitMap% = SelectObject(hMemoryDC%, Picture1.Picture)
 
 ApiError% = StretchBlt(Printer.hdc, 0, 0, Printer.ScaleWidth, Printer.ScaleHeight, hMemoryDC%, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, SRCCOPY)
 hOldBitMap% = SelectObject(hMemoryDC%, hOldBitMap%)
 ApiError% = DeleteDC(hMemoryDC%)
 Result% = Escape(Printer.hdc, NEWFRAME, 0, 0&, 0&)
 Printer.EndDoc
 Label2.Visible = False
End Sub


Private Sub Picture1_Click()
    RaiseEvent Click
End Sub

Private Sub Picture1_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
     
    RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
    If Me.TrackMousePointer Then
    mnuTrackMouse.Checked = True
    Else
    mnuTrackMouse.Checked = False
    End If
    If Me.ShowGrids Then
    mnuShowGrids.Checked = True
    Else
    mnuShowGrids.Checked = False
    End If
    PopupMenu mnuFile, 0, x, y, mnuTrackMouse
Else
    RaiseEvent MouseDown(Button, Shift, x, y)
End If
End Sub
Public Sub SendToExcel()
'''this code is stolen from planet-source-code.com ;
'' i think originally wriiten by Joe Miguel   joe_miguel@hotmail.com
''''i modified  and deleted many lines as my requirements
''''means i changed that code

Dim varNum As Long
Dim i As Long
Dim j As Long
Dim k As Long
Dim p As Long
Dim xx As Single
Dim yy As Single
Dim m_str As String
Dim objExcel As Excel.Application
Dim objWorkbook As Excel.Workbook
Dim objWorksheet As Excel.Worksheet
Dim objChart(MAX_NO_OF_GRAPHS - 1) As Excel.Chart

'Start the excel COM and make it visible.
On Error GoTo lab45
Set objExcel = GetObject("", "excel.application")
'Set objExcel = excel.Application ' Seems to cause a memory leak
    objExcel.Visible = False
    
'Start a workbook.
Set objWorkbook = objExcel.Workbooks.Add

'Turn off the alerts, otherwise user will have to confirm my actions.
    objExcel.DisplayAlerts = False
    '''show busy label
    i = (Picture1.Width - Label2.Width) / 2
    If i > 0 Then
          Label2.Left = i
    Else
          Label2.Left = 0
    End If
    i = (Picture1.Height - Label2.Height) / 2
    If i > 0 Then
           Label2.Top = i
    Else
           Label2.Top = 0
    End If
    Label2.Visible = True
'Depending on the users excel's settings, there could be many worksheet when starting a workbook.
'Ensure there is only one worksheet.

Do While objWorkbook.Worksheets.Count > 1
    Set objWorksheet = objWorkbook.Worksheets.Item(objWorkbook.Worksheets.Count)
    objWorksheet.Delete
Loop
'objWorkbook.Worksheets.Add
Set objWorksheet = objWorkbook.Worksheets.Item(objWorkbook.Worksheets.Count)
'Set objWorksheet to the remaining worksheet.
'Set objWorksheet = ActiveSheet

'Rename the sheet to Results.
    objWorksheet.Name = "GraphData"
      
'Headers
    objWorksheet.Cells(1, 1) = excel_Header '''''  "Data comes From AshuGraphControl"
    objWorksheet.Cells(1, 1).Font.Bold = True
    objWorksheet.Cells(2, 1) = " " & Now
    objWorksheet.Cells(2, 1).Font.Bold = True
       
    k = 0
'Sent data to excel

 '''loop through grphs
 For i = 0 To MAX_NO_OF_GRAPHS
   If m_Graph(i).Total_Data > 0 Then
        k = k + 1
        If m_Graph(i).Name = "" Then m_Graph(i).Name = "Graph " & str(i)
       objWorksheet.Cells(3, 2 * k - 1) = m_Graph(i).Name
       objWorksheet.Cells(3, 2 * k - 1).Font.Bold = True
       
       For j = 0 To (m_Graph(i).Total_Data - 1)
            p = j
            If (m_Graph(i).GetPoint(xx, yy, p)) Then
                 objWorksheet.Cells(j + 5, 2 * k - 1) = xx
                objWorksheet.Cells(j + 5, 2 * k) = yy
            End If
       Next j
       '''Draw chart for this graph
          
          
  End If
Next i
 
            
'
    
'Turn back on alerts so user will be notified to save on exit.
   objExcel.DisplayAlerts = True
objExcel.Visible = True
Label2.Visible = False


'Free up memory, otherwise there will be a memory leak.
Set objExcel = Nothing
Set objWorksheet = Nothing
Set objWorkbook = Nothing
For i = 0 To (MAX_NO_OF_GRAPHS - 1)
Set objChart(i) = Nothing
Next i
Exit Sub
lab45:
 MsgBox "Error: Excel Not installed On your system"
 Label2.Visible = False
End Sub
Public Sub SetName(ByVal s_name As String, ByVal GraphNo As Long)
If (GraphNo < 0) Or (GraphNo > MAX_NO_OF_GRAPHS) Then Exit Sub
m_Graph(GraphNo).Name = s_name

End Sub

Public Sub ShowBusyMessage(show As Boolean)
'''show busy label
    i = (Picture1.Width - Label2.Width) / 2
    If i > 0 Then
          Label2.Left = i
    Else
          Label2.Left = 0
    End If
    i = (Picture1.Height - Label2.Height) / 2
    If i > 0 Then
           Label2.Top = i
    Else
           Label2.Top = 0
    End If
    Label2.Visible = show
End Sub
Public Function GetFreeGraph() As Long
Dim i As Long
GetFreeGraph = -1
For i = 0 To MAX_NO_OF_GRAPHS - 1
If m_Graph(i).Total_Data = 0 Then
GetFreeGraph = i
Exit For
End If
Next i

End Function
Public Sub SaveGraphInTextFile(ByVal GraphNo As Long, ByVal fileName As String)
Dim fileNum
Dim m_str As String
Dim i As Long
Dim xx As Single
Dim yy As Single
Dim j As Long
If (GraphNo < 0) Or (GraphNo > MAX_NO_OF_GRAPHS) Then Exit Sub
fileNum = FreeFile
On Error GoTo slab1
Open fileName For Output As fileNum
'''''now write  total graphs = 1, & scale settings into comments
m_str = "Don't Delete These comments ,This comments are written by AshuGraphControl "
Print #fileNum, "# "; m_str

Print #fileNum, "# "; "  XUnitToPixels =  "; Me.XUnitToPixels
Print #fileNum, "# "; "  YUnitToPixels =  "; Me.YUnitToPixels
Print #fileNum, "# "; "  XGap =  "; Me.XGap
Print #fileNum, "# "; "  YGap =  "; Me.YGap
Print #fileNum, "# "; "  XMinPosition =  "; Me.XMinPosition
Print #fileNum, "# "; "  YMinposition =  "; Me.YMinPosition
''''Now write Graph setting
Print #fileNum, "# "; "  GraphName =  "; m_Graph(GraphNo).Name
Print #fileNum, "# "; "  color =  "; m_Graph(GraphNo).color
Print #fileNum, "# "; "  DrawStyle =  "; m_Graph(GraphNo).DrawStyle
Print #fileNum, "# "; "  TotalPoints =  "; m_Graph(GraphNo).Total_Data
Print #fileNum, (Chr(13) + Chr(10))
Print #fileNum, "# "; "  Points X  ,    Y"
Print #fileNum, (Chr(13) + Chr(10))
'''''Now write data
For i = 0 To (m_Graph(GraphNo).Total_Data - 1)
    j = i
    If (m_Graph(GraphNo).GetPoint(xx, yy, j)) Then Print #fileNum, xx; " , "; yy
Next i
Close #fileNum






Exit Sub
slab1:
MsgBox "Error in Opening file"
Close #fileNum
Exit Sub

End Sub


Public Function ReadGraphFromTextFile(ByVal fileName As String, ByVal ResetSettings As Boolean) As Long
Dim fileNum
Dim m_str As String
Dim i As Long
Dim xx As Single
Dim yy As Single
Dim j As Long
Dim k As Integer
Dim m As String
Dim m1 As String
Dim m2 As String
Dim m3 As String
Dim m4 As String
Dim file1
Dim GraphNo As Long
GraphNo = GetFreeGraph
ReadGraphFromTextFile = GraphNo
If (GraphNo < 0) Or (GraphNo > MAX_NO_OF_GRAPHS) Then Exit Function

fileNum = FreeFile
On Error GoTo slab1
Open fileName For Input As fileNum
file1 = FreeFile
Open "data1.data" For Output As file1
Do While Not EOF(fileNum)
        Line Input #fileNum, m_str
        m_str = Trim(m_str)
        m = Mid(m_str, 1, 1)
        m = Trim(m)
        If m = "#" Then
           m1 = Mid(m_str, 2)
           m1 = LTrim(m1)
           m2 = Mid(m1, 1, 4)
           m3 = UCase(m2)
           m4 = Trim(m3)
           If m4 = Trim("XUNI") Then
                     j = FindChar(m1, "=")
                     If (j > 0) And (ResetSettings) Then
                     m1 = Mid(m1, j + 1)
                     m2 = Trim(m1)
                     Me.XUnitToPixels = CSng(Val(m2))
                     End If
                    
           ElseIf m4 = "YUNI" Then
                    j = FindChar(m1, "=")
                     If (j > 0) And (ResetSettings) Then
                     m = Mid(m1, j + 1)
                     m = Trim(m)
                     Me.YUnitToPixels = CSng(Val(m))
                     End If
                    
           ElseIf m4 = "XGAP" Then
                    j = FindChar(m1, "=")
                     If (j > 0) And (ResetSettings) Then
                     m = Mid(m1, j + 1)
                     m = Trim(m)
                     Me.XGap = CSng(Val(m))
                     End If
                     
           ElseIf m4 = "YGAP" Then
                    j = FindChar(m1, "=")
                     If (j > 0) And (ResetSettings) Then
                     m = Mid(m1, j + 1)
                     m = Trim(m)
                     Me.YGap = CSng(Val(m))
                     End If
                     
           ElseIf m4 = "XMIN" Then
                   j = FindChar(m1, "=")
                     If (j > 0) And (ResetSettings) Then
                     m = Mid(m1, j + 1)
                     m = Trim(m)
                     Me.XMinPosition = CSng(Val(m))
                     End If
                     
           ElseIf m4 = "YMIN" Then
                    j = FindChar(m1, "=")
                     If (j > 0) And (ResetSettings) Then
                     m = Mid(m1, j + 1)
                     m = Trim(m)
                     Me.YMinPosition = CSng(Val(m))
                     End If
                    
           ElseIf m4 = "GRAP" Then
                   j = FindChar(m1, "=")
                     If (j > 0) Then
                     m = Mid(m1, j + 1)
                     m_Graph(GraphNo).Name = Trim(m)
                     End If
           ElseIf m4 = "COLO" Then
                    j = FindChar(m1, "=")
                     If (j > 0) Then
                     m = Mid(m1, j + 1)
                     m = Trim(m)
                     m_Graph(GraphNo).color = CLng(Val(m))
                     Debug.Print "color ", CLng(Val(m))
                     End If
           ElseIf m4 = "DRAW" Then
                     j = FindChar(m1, "=")
                     If (j > 0) Then
                     m = Mid(m1, j + 1)
                     m = Trim(m)
                     m_Graph(GraphNo).DrawStyle = CInt(Val(m))
                     Debug.Print "Drawstyle", CInt(Val(m))
                     End If
           ElseIf m4 = "TOTA" Then
                   j = FindChar(m1, "=")
                     If (j > 0) Then
                     m = Mid(m1, j + 1)
                     m = Trim(m)
                     'm_Graph(GraphNo).DrawStyle = CInt(Val(m))
                     End If
               
           End If
                  
                  
          
           
      ElseIf m_str <> "" Then
             j = FindChar(m_str, ",")
             
             
             If j > 0 Then
             'Print #file1, m_str
               m = Mid(m_str, 1, j - 1)
               m = Trim(m)
               xx = CSng(Val(m))
               m = Mid(m_str, j + 1)
               m = Trim(m)
               yy = CSng(Val(m))
               Call AddData(xx, yy, GraphNo)
             End If
                       
    End If

Loop

Close #fileNum
Close #file1
Exit Function
slab1:
MsgBox "Error in Opening file"
Close #fileNum
Close #file1

End Function
Private Function FindChar(source1 As String, char1 As String) As Long
Dim i As Long
Dim j As Long
Dim m As String
FindChar = -1
i = Len(source1)
For j = 1 To (i - 1)
    m = Mid(source1, j, 1)
    If Trim((m)) = Trim((char1)) Then
    FindChar = j
    Exit For
    End If
Next j


End Function
Private Function FindChar2(source As String, char As String) As Long
Dim i As Long
Dim j As Long
Dim m As String
FindChar2 = -1
i = Len(source)
For j = 1 To (i - 1)
    m = Mid(source, j, 1)
    If ((m)) = ((char)) Then
    FindChar2 = j
    Exit For
    End If
Next j


End Function
Public Function DrawDerivativeForGraph(ByVal GraphNo As Long) As Long
Dim d_graph As Long
Dim i As Long
Dim j As Long
Dim k As Long
Dim xx1 As Single
Dim yy1 As Single
Dim xx2 As Single
Dim xx3 As Single
Dim yy2 As Single
Dim yy3 As Single
Dim xx As Single
Dim yy As Single
DrawDerivativeForGraph = -1
d_graph = GetFreeGraph
If d_graph < -1 Then
        MsgBox "Maximum Limits Exceeds,To Draw Derivative Graph,First Make Free Space by removing at least one graph"
        Exit Function
End If
k = 0
For i = 0 To (m_Graph(GraphNo).Total_Data - 3)
           j = i
          Call m_Graph(GraphNo).GetPoint(xx1, yy1, j)
          Call m_Graph(GraphNo).GetPoint(xx2, yy2, j + 1)
          'Call m_Graph(GraphNo).GetPoint(xx3, yy3, j + 2)
          j = j + 1
          Do While ((xx2 - xx1) > 1#) And (j < m_Graph(GraphNo).Total_Data - 3)
          j = j + 1
          Call m_Graph(GraphNo).GetPoint(xx2, yy2, j)
          Loop
          xx = xx2
          yy = 0
          If (xx2 - xx1) > 0 Then yy = (yy2 - yy1) / (xx2 - xx1)
          Call Me.AddData(xx, yy, d_graph)
Next i
m_Graph(d_graph).color = RGB(0, 0, 255)
DrawDerivativeForGraph = d_graph
  Me.Invalidate
End Function


Private Sub DrawLines(hdc As Long, x1 As Long, y1 As Long, x2 As Long, y2 As Long, color As Long)
Dim ptAPI As POINTAPI

If hdc = Printer.hdc Then
        Printer.Line (x1, y1)-(x2, y2), color
ElseIf hdc = Picture1.hdc Then
        Picture1.Line (x1, y1)-(x2, y2), color
ElseIf hdc = Picture2.hdc Then
        Picture2.Line (x1, y1)-(x2, y2), color
Else
        MoveToEx hdc, x1, y1, ptAPI
        LineTo hdc, x2, y2
End If
End Sub
Private Sub DrawSquarePoint(hdc As Long, x As Long, y As Long, color As Long)
If hdc = Printer.hdc Then
        Printer.Line (x - 1, y - 1)-(x + 1, y + 1), color, BF
ElseIf hdc = Picture1.hdc Then
       Picture1.Line (x - 1, y - 1)-(x + 1, y + 1), color, BF
ElseIf hdc = Picture2.hdc Then
      Picture2.Line (x - 1, y - 1)-(x + 1, y + 1), color, BF
End If
       
End Sub
Public Sub SetDrawStyle(ByVal GraphNo As Long, ByVal DrawStyle As Integer)
If (GraphNo < 0) Or (GraphNo > MAX_NO_OF_GRAPHS) Then Exit Sub
m_Graph(GraphNo).DrawStyle = DrawStyle

End Sub
Private Sub drawLowerHScale(hdc As Long, Xdisplacement As Long, Ydisplacement As Long, hHeight As Long, wWidth As Long)
Dim ptAPI As POINTAPI

Dim tmp As Long
Dim tmp1 As Long
Dim oldColor As Long
Dim str As String
Dim firstPoint As Single
Dim distance As Single
Dim i As Single

''''calculate height of font
tmp = -1 * MulDiv(Me.Font.Size, GetDeviceCaps(hdc, 90), 72)
tmp = Abs(tmp)


firstPoint = Me.XMinPosition + m_XInitialMargin
distance = 0#
i = firstPoint

Do
   distance = (i - firstPoint) * Me.XUnitToPixels
   If distance > CSng(wWidth) Then Exit Do
   str = Format(i, "#0.00")
   tmp1 = TextOut(hdc, Xdisplacement + CLng(distance), Ydisplacement + hHeight - tmp, str, Len(str))
   
   DrawLines hdc, Xdisplacement + CLng(distance), Ydisplacement, Xdisplacement + CLng(distance), Ydisplacement + hHeight - tmp, Me.ForeColor
   i = i + Me.XGap
   
Loop While distance < wWidth
'''''''draw boundary lines
DrawLines hdc, Xdisplacement, Ydisplacement, Xdisplacement + wWidth, Ydisplacement, Me.ForeColor
'DrawLines hdc, Xdisplacement, Ydisplacement + hHeight, Xdisplacement + wWidth, Ydisplacement + hHeight, Me.ForeColor
DrawLines hdc, Xdisplacement + wWidth, Ydisplacement + hHeight, Xdisplacement + wWidth, Ydisplacement + tmp, Me.ForeColor
End Sub
Public Sub SetExcelHeading(ByVal header_excel As String)
excel_Header = header_excel

End Sub
