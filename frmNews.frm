VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "msinet.ocx"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form w 
   Caption         =   "Breaking News"
   ClientHeight    =   10455
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   19875
   FillColor       =   &H00C0FFC0&
   ForeColor       =   &H00C0FFC0&
   Icon            =   "frmNews.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10455
   ScaleWidth      =   19875
   StartUpPosition =   3  'Windows Default
   Begin SHDocVwCtl.WebBrowser wb 
      Height          =   4575
      Left            =   240
      TabIndex        =   8
      Top             =   5160
      Width           =   18015
      ExtentX         =   31776
      ExtentY         =   8070
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.TextBox txturl 
      Height          =   375
      Left            =   10200
      TabIndex        =   7
      Top             =   240
      Width           =   9015
   End
   Begin VB.TextBox txtInterval 
      Height          =   375
      Left            =   4200
      TabIndex        =   6
      Text            =   "60000"
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox txtCounter 
      Height          =   285
      Left            =   5760
      TabIndex        =   5
      Top             =   240
      Width           =   975
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   8640
      Top             =   240
   End
   Begin VB.CommandButton Command3 
      Cancel          =   -1  'True
      Caption         =   "End"
      Height          =   495
      Left            =   2550
      TabIndex        =   4
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Stop"
      Height          =   495
      Left            =   1335
      TabIndex        =   3
      Top             =   240
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   6000
      Left            =   9360
      Top             =   240
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   3135
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   1680
      Width           =   19695
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   1080
      Width           =   19695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   7320
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      Protocol        =   4
      RemoteHost      =   "www.google.ca"
      URL             =   "http://www.google.ca"
   End
End
Attribute VB_Name = "w"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lastNews As String
Dim i As Integer
Dim oldText As String

Private Sub Command1_Click()
Init
getData
'parseNews ("")
End Sub

Function Init()
    Timer1.Enabled = True
    Timer2.Enabled = True
    lastNews = ""
    Timer1.Interval = Val(txtInterval)

End Function

Private Sub Command2_Click()
Timer1.Enabled = False
Timer2.Enabled = False

End Sub

Private Sub Command3_Click()
End
End Sub

Private Sub Form_Load()
    wb.AddressBar = True
    wb.Silent = True
    wb.Navigate "http://www.google.ca", 4
End Sub

Private Sub Form_Resize()
wb.Width = Me.Width - 900
wb.Top = Text2.Top + Text2.Height
If Not Me.WindowState = 1 Then
    wb.Height = Me.Height - wb.Top - 800
    txturl = Me.Height
    Me.txtCounter = wb.Height
End If
End Sub

Private Sub Inet1_StateChanged(ByVal State As Integer)
'MsgBox State
End Sub

Private Sub Timer1_Timer()
If Timer1.Interval < 6000 Then

MsgBox ("Timer too small. Defaulting to 1 min.")
txtInterval = 60000
Else
    i = 0
    getData
End If
End Sub


Private Sub getData()
On Error GoTo weberror
Dim msg As String


Text1 = "Getting Data..."
Text1.Text = Inet1.OpenURL("http://www.cnbc.com/franchise/20991458?callback=breakingNews&mode=breaking_news")

'http://quote.cnbc.com/quote-html-webservice/quote.htm?callback=webQuoteRequest&symbols=TD.TO&symbolType=symbol&requestMethod=quick&exthrs=1&extMode=&fund=1&entitlement=0&skipcache=&extendedMask=1&partnerId=2&output=jsonp&noform=1
'http://quote.cnbc.com/quote-html-webservice/quote.htm?callback=webQuoteRequest&symbols=MSFT &symbolType=symbol&requestMethod=quick&exthrs=1&extMode=&fund=1&entitlement=0&skipcache=&extendedMask=1&partnerId=2&output=jsonp&noform=1
'http://apps.cnbc.com/view.asp?uid=stocks/financials&view=cashFlowStatement&symbol=ry.to
'http://apps.cnbc.com/view.asp?uid=stocks/financials&view=balanceSheet&symbol=TD.TO
'http://apps.cnbc.com/view.asp?uid=stocks/financials&view=incomeStatement&symbol=TD.TO
'http://data.cnbc.com/quotes/TD.TO
If Len(Text1) > 20 And Text1 <> oldText Then
    oldText = Text1
    lastNews = parseNews(Text1)
    Text2 = lastNews & vbCrLf & Text2
Beep
Beep
Beep
End If
Exit Sub
weberror:
Text2 = Text2 & "____Error:" & Err.Description

End Sub

Private Sub Timer2_Timer()
i = i + 1
txtCounter = Timer1.Interval / 1000 - i
End Sub 'txtur,r



Function parseNews(m As String) As String
Dim url As String
Dim id As String
Dim headline As String
Dim what As String
Dim what1 As String
Dim what2 As String
Dim pos As Integer


'm = "breakingNews({\""url"":""http:\/\/www.cnbc.com\/2015\/10\/15\/early-movers-gs-bud-unh-mo-sbux-wmt-nflx-unh-more.html"",""id"":103080254,""headline"":""Early movers: GS, BUD, UNH, MO, SBUX, WMT, NFLX, UNH & more""});"


what = "url"":"
what1 = """,""id"":"
what2 = ",""headline"":"
pos = InStr(1, m, what, vbTextCompare)
url = Mid(m, pos + Len(what) + 1)
url = Mid(url, 1, InStr(1, url, what1, vbTextCompare) - 1)
url = Replace(url, "\", "")
 
id = Mid(m, InStr(1, m, what1, vbTextCompare) + Len(what) + 1)
id = Mid(id, 1, InStr(1, id, what2, vbTextCompare) - 1)
pos = InStr(1, m, what2, vbTextCompare)

headline = Mid(m, pos + Len(what2) + 1)
headline = Mid(headline, 1, Len(headline) - 4)

'MsgBox m
'MsgBox InStr(1, m, what2, vbTextCompare)
'MsgBox Mid(m, InStr(1, m, what2, vbTextCompare) + Len(what2))
'MsgBox "headline:" & headline

txturl = wb.LocationURL
wb.Navigate url, 4
parseNews = Format(Now(), "mmm-dd hh:mm") & " - " & headline & "(" & id & ")"
End Function

Private Sub txtInterval_Change()
Timer1.Interval = Val(txtInterval)
i = 0
If IsNumeric(Timer1.Interval) Then txtCounter = Timer1.Interval / 1000
End Sub

Private Sub txturl_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then wb.Navigate txturl, 4
End Sub

Private Sub wb_StatusTextChange(ByVal Text As String)
txturl = wb.LocationURL
End Sub
