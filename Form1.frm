VERSION 5.00
Object = "{844D93BA-B24B-420C-A010-F8A96621C63B}#1.0#0"; "ExtBasic.ocx"
Begin VB.Form Form1 
   Caption         =   "double click direct call line()"
   ClientHeight    =   5580
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7620
   LinkTopic       =   "Form1"
   ScaleHeight     =   5580
   ScaleWidth      =   7620
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton Command2 
      Caption         =   "run js"
      Height          =   615
      Left            =   3000
      TabIndex        =   2
      Top             =   3840
      Width           =   1815
   End
   Begin ExtBasicLib.ExtBasic ExtBasic1 
      Height          =   3135
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   6615
      _Version        =   65536
      _ExtentX        =   11668
      _ExtentY        =   5530
      _StockProps     =   0
   End
   Begin VB.CommandButton Command1 
      Caption         =   "start"
      Height          =   615
      Left            =   600
      TabIndex        =   0
      Top             =   3840
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    hr = ExtBasic1.InitWebKit("", 0)
    ExtBasic1.RegisterObject "form1", Me
    ExtBasic1.LoadUrl App.Path + "\test_WebBrowser_ActiveX.html"   '"http://www.baidu.com"
End Sub

Private Sub Command2_Click()
    ExtBasic1.ExecJScript "alert('***** run javascript ******')"
End Sub

Private Sub ExtBasic1_OnLoadEnd(ByVal bstrUrl As String, ByVal statusCode As Long)
    Dim strText As String
    If 0 = statusCode Then
        strText = ExtBasic1.GetSource
    Else
        strText = "load fail"
    End If
    MsgBox Left(strText, 32), vbOKOnly, bstrUrl
End Sub

Private Sub Form_DblClick()
    Call Me.Line(0, 0, 0, 5000, 4000, 255)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ExtBasic1.UnInitWebKit
End Sub

Public Function ListSymEx(ByRef dispId As Long) As String
    MsgBox "called ListSymEx: " + CStr(dispId)
    ListSymEx = ""
End Function

Public Function Form_ListSymEx(ByRef dispId As Long) As String
    MsgBox "called Form_ListSymEx: " + CStr(dispId)
    Form_ListSymEx = ""
End Function

