VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "Msinet.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Web Link Extractor - EJ Smith"
   ClientHeight    =   4740
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5460
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4740
   ScaleWidth      =   5460
   StartUpPosition =   3  'Windows Default
   Begin InetCtlsObjects.Inet Inet 
      Left            =   840
      Top             =   2640
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "Extracted Links"
      Height          =   4215
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   5295
      Begin VB.TextBox txtLinks 
         Height          =   3855
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   4
         Top             =   240
         Width           =   5055
      End
   End
   Begin VB.CommandButton cmdExtract 
      Caption         =   "Extract"
      Height          =   285
      Left            =   4440
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox txtURL 
      Height          =   285
      Left            =   720
      TabIndex        =   1
      Text            =   "http://www.google.com"
      Top             =   120
      Width           =   3615
   End
   Begin VB.Label Label1 
      Caption         =   "URL:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Hope you like my code :-)
'Created by EJ Smith

Option Explicit

Function DownloadFile(strArg As String) As String
    'Function downloads url to a string, to be extracted
    
    'Sets timeout to 20 seconds, if ping is over 20 seconds, connection cannot be made...
    Inet.RequestTimeout = 20
    
    'Download URL to string
    DownloadFile = Inet.OpenURL(strArg)
End Function

Function ExtractLinks(strArg As String) As String
    'This function is self explanitory, it extracts all links from a string
    
    'Declare variables
    Dim strSplit() As String
    Dim intLinkStart As Integer
    Dim intLinkEnd As Integer
    Dim intProgLoop As Integer
    Dim strTagBegin As String
    Dim strTagEnd As String
    
    'Set tags to look for, can be changed to grab email addresses, etc.
    'Use this to extract emails -- strTagBegin = "<a href=" & Chr(34) & "mailto:"
    strTagBegin = "<a href="
    strTagEnd = ">"
    strSplit() = Split(strArg, strTagBegin)
    
    'Main loop to grab links, and strips out quotation marks
    For intProgLoop = (LBound(strSplit) + 1) To UBound(strSplit)
        intLinkEnd = InStr(1, strSplit(intProgLoop), strTagEnd)
        If ExtractLinks = "" Then
            ExtractLinks = Replace(Mid(strSplit(intProgLoop), 1, intLinkEnd - 1), Chr(34), "")
        Else
            ExtractLinks = ExtractLinks & vbNewLine & Replace(Mid(strSplit(intProgLoop), 1, intLinkEnd - 1), Chr(34), "")
        End If
    Next intProgLoop
End Function

Private Sub cmdExtract_Click()
    Dim strBuffer As String
    txtLinks.Text = ""
    cmdExtract.Enabled = False
    strBuffer = DownloadFile(txtURL.Text)
    txtLinks.Text = ExtractLinks(strBuffer)
    cmdExtract.Enabled = True
End Sub
