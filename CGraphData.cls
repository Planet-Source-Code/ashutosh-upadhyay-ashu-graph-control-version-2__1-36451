VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CGraphData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Const MAX_POINTS_PER_GRAPH = 70000
Dim x(MAX_POINTS_PER_GRAPH) As Single
Dim y(MAX_POINTS_PER_GRAPH) As Single
Public index As Long
Public Total_Data As Long
Public color As Long
Public Name As String
Public DrawStyle As Integer




Public Function CGraphData()
Total_Data = 0
index = -1
color = RGB(255, 0, 0)
Name = ""
DrawStyle = 1
End Function

Public Sub AddData(ByVal xx As Single, ByVal yy As Single)
If Total_Data > MAX_POINTS_PER_GRAPH Then Exit Sub
x(Total_Data) = xx
y(Total_Data) = yy
Total_Data = Total_Data + 1

End Sub

Public Sub DeleteAll()
Total_Data = 0
index = -1
Name = ""
End Sub
Public Function GetPoint(ByRef xx As Single, ByRef yy As Single, ByRef iindex As Long) As Boolean
If (iindex < 0) Or (iindex > Total_Data) Then
iindex = -1
GetPoint = False
Exit Function
End If
xx = x(iindex)
yy = y(iindex)
index = iindex
GetPoint = True
End Function
Public Function SetPoint(ByVal xx As Single, ByVal yy As Single, ByVal iindex) As Boolean
If (iindex < 0) Or (iindex > Total_Data) Then
SetPoint = False
Exit Function
End If
x(iindex) = xx
y(iindex) = yy
index = iindex

End Function
