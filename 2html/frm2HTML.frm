VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm2HTML 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6840
   Icon            =   "frm2HTML.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   6840
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdConvert 
      Caption         =   "&Make HTML"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5460
      TabIndex        =   8
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "&Load Project"
      Height          =   375
      Left            =   2340
      TabIndex        =   7
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmdNewOutDir 
      Caption         =   "&Output To"
      Height          =   375
      Left            =   2340
      TabIndex        =   4
      Top             =   180
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog dlgProject 
      Left            =   5460
      Top             =   4380
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   ".vbp"
      DialogTitle     =   "Open Project"
   End
   Begin MSComctlLib.ListView lvProject 
      Height          =   2955
      Left            =   240
      TabIndex        =   0
      Top             =   2280
      Width           =   6435
      _ExtentX        =   11351
      _ExtentY        =   5212
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label lblHelp 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   3720
      TabIndex        =   6
      Top             =   180
      Width           =   2895
   End
   Begin VB.Label lblOutDirPath 
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   660
      Width           =   3255
   End
   Begin VB.Label lblProjectFile 
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   1740
      Width           =   3315
   End
   Begin VB.Label lblProject 
      Caption         =   "Project:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1380
      Width           =   1215
   End
   Begin VB.Label lblOutDir 
      Caption         =   "HTML will be written to:"
      Height          =   315
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   1935
   End
   Begin VB.Menu mnuItem 
      Caption         =   "Item"
      Visible         =   0   'False
      Begin VB.Menu mnuItemView 
         Caption         =   "&View in Browser"
      End
      Begin VB.Menu mnuItemEdit 
         Caption         =   "&Edit in WordPad"
      End
   End
End
Attribute VB_Name = "frm2HTML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const VBKEYWORDS = "If In Is As Or To On Do #If And Dim End Eqv For Get Imp Let Lib " & _
                   "New Not Put Set Spc Sub Tab Xor " & _
                   "#Const #Else #ElseIf #End Alias Base Binary Boolean Byte ByVal " & _
                   "Call CBool CCur CDbl CDec CInt Compare Const CStr Currency CVar " & _
                   "CVErr Decimal Declare DefBool DefCur DefDbl DefDec DefInt DefObj " & _
                   "DefStr DefVar Double Each ElseIf Enum Error False Function Global " & _
                   "GoSub GoTo Input Integer LBound Like Line Lock Long Loop LSet " & _
                   "Next Nothing Object Open Option Output Print Private Property " & _
                   "Public Random Read ReDim Resume Return RSet Seek Select Single " & _
                   "Static String Stop Then True Type UBound Unlock Variant Wend " & _
                   "While With " & _
                   "Case Close CDate CByte CLng CSng DefByte DefDate DefLng DefSng " & _
                   "Else Erase Exit Explicit "

Dim ln(64 To 90, 65 To 90, 0 To 13) As Byte
Dim w(64 To 90, 65 To 90, 0 To 13, 0 To 13) As Byte
Dim b(64 To 90, 65 To 90, 0 To 13, 0 To 53) As Byte
Dim p(0 To 127, 0 To 127, 0 To 127) As Byte
Dim bytPreBlue() As Byte
Dim bytPreGreen() As Byte
Dim bytPost() As Byte
Dim bytBreak() As Byte
Dim bytHR() As Byte
Dim bytNBSP() As Byte
Dim bytTemp() As Byte
Dim wl() As String
Dim strProjectFileName As String
Dim strOutputDir As String
Dim strProjectPath As String
Dim gCurItem As MSComctlLib.ListItem
Dim strItemOuts(100) As String

Private Sub cmdConvert_Click()
    DoConversions
End Sub

Private Sub cmdLoad_Click()
    Dim strtemp As String
    
    On Error GoTo errorhandle
    
    dlgProject.CancelError = True
    dlgProject.Filter = "*.vbp"
    dlgProject.FileName = "*.vbp"
    dlgProject.ShowOpen
    
    strProjectFileName = dlgProject.FileName
    lblProjectFile = strProjectFileName
    OpenProject
    
    Exit Sub
errorhandle:
    
End Sub

Sub OpenProject()
    Dim strLine As String
    Dim strFileName As String
    Dim intEqualPos As Integer
    Dim intSemiColonPos As Integer
    Dim lstX As ListItem
    Dim i As Integer, j As Integer
    Dim strPath As String
    
    lvProject.ListItems.Clear
    
    If InStr(1, strProjectFileName, ":") > 0 Then ChDrive Left(strProjectFileName, 1)
    strProjectPath = Mid(strProjectFileName, 1, InStrRev(strProjectFileName, "\"))
    ChDir strProjectPath
    
    Open strProjectFileName For Input As #1
    While Not EOF(1)
        Line Input #1, strLine
        If Left(strLine, 5) = "Form=" Or _
          Left(strLine, 6) = "Class=" Or _
          Left(strLine, 9) = "UserForm=" Or _
          Left(strLine, 7) = "Module=" Then
            intEqualPos = InStr(1, strLine, "=")
            intSemiColonPos = InStr(intEqualPos, strLine, ";")
            If intSemiColonPos > 0 Then
                strFileName = Trim(Mid(strLine, intSemiColonPos + 1))
            Else
                strFileName = Trim(Mid(strLine, intEqualPos + 1))
            End If
            If InStr(1, strFileName, "\") > 0 Then
                strPath = Mid(strFileName, 1, InStrRev(strFileName, "\") - 1)
                ChDir strPath
                strFileName = CurDir & IIf(Right(CurDir, 1) = "\", "", "\") & Mid(strFileName, InStrRev(strFileName, "\") + 1)
            Else
                strFileName = CurDir & IIf(Right(CurDir, 1) = "\", "", "\") & strFileName
            End If
            ChDir strProjectPath
            Set lstX = lvProject.ListItems.Add(, , strFileName)
            lstX.SubItems(1) = "Not yet converted.."
        End If
    Wend
    Close #1
    lblProjectFile = strProjectFileName
    cmdConvert.Enabled = True
End Sub

Private Sub DoConversions()
    
    Dim lstX As ListItem
    Dim strIn As String
    Dim strOut As String
    Dim s() As Byte
    Dim n() As Byte
    Dim strtemp As String
    Dim c As Integer
    Dim blnSuccess As Boolean
    
   ' On Error GoTo errorhandle
    
    ReDim Preserve s(1)
    ReDim Preserve n(1)
    strProjectPath = Mid(strProjectFileName, 1, InStrRev(strProjectFileName, "\"))
    ChDir strProjectPath
    For Each lstX In lvProject.ListItems
        strIn = lstX
        lstX.SubItems(1) = "Reading file..."
        Me.Refresh
        blnSuccess = ReadFile(strIn, s)
        If blnSuccess Then
            lstX.SubItems(1) = "Converting to HTML..."
            Me.Refresh
            blnSuccess = ConvertFile(s, n)
            If blnSuccess Then
                strtemp = Mid(lstX, InStrRev(lstX, "\") + 1)
                strOut = strOutputDir & IIf(Right(strOutputDir, 1) = "\", "", "\") & strtemp & ".html"
                'strOut = Replace(strOut, ".", "_")
                lstX.SubItems(1) = "Writing HTML..."
                Me.Refresh
                WriteFile n, strOut
                strItemOuts(c) = strOut
                c = c + 1
                lstX.SubItems(1) = "Done."
                Me.Refresh
            Else
                lstX.SubItems(1) = "Error during conversion!"
            End If
        Else
            lstX.SubItems(1) = "Error reading file!"
        End If
        Me.Refresh
        ReDim s(1)
        ReDim n(1)
    Next lstX
    mnuItem.Enabled = True
    mnuItemView.Enabled = True
    mnuItemEdit.Enabled = True
    lblHelp.Font.Size = 8
    lblHelp = vbCrLf & vbCrLf & "Click on an item to view the page in a browser window or edit the HTML with WordPad."
    Exit Sub
errorhandle:
    Debug.Print Err.Description
    Resume Next
 End Sub

Public Function ReadFile(strIn As String, f() As Byte) As Boolean
    Dim strLine As String
    Dim strFName As String
    Dim bytresponse As Byte
    
    On Error GoTo errorhandler
    
    strFName = strIn
    Open strFName For Input As #1
    ReDim bytTemp(1)
    
    If InStr(1, strFName, "frm") > 0 Then
        While InStr(1, strLine, "Attribute VB_Exposed = False") <= 0 And Not EOF(1)
            Line Input #1, strLine
        Wend
    End If
    While Not EOF(1)
        Line Input #1, strLine
        StoreBytes bytTemp, strLine & vbCrLf
        AppendBytes f, bytTemp
    Wend
    Close #1
    ReadFile = True
    Exit Function
    
errorhandler:
    bytresponse = MsgBox("Error reading file " & strIn & ": " & Err.Description, vbAbortRetryIgnore)
    If bytresponse = vbRetry Then
        Resume
    ElseIf bytresponse = vbIgnore Then
        Resume Next
    Else
        ReadFile = False
        Exit Function
    End If
    
End Function

Private Function WriteFile(f() As Byte, strOut As String)
    Dim bytresponse As Byte
    
    On Error GoTo errorhandler
    
    Open strOut For Binary As #1
    Put #1, , f
    Close #1
    Exit Function
    WriteFile = True
errorhandler:
    bytresponse = MsgBox("Error writing file " & strOut & ": " & Err.Description, vbAbortRetryIgnore)
    If bytresponse = vbRetry Then
        Resume
    ElseIf bytresponse = vbIgnore Then
        Resume Next
    Else
        WriteFile = False
        Exit Function
    End If
End Function

Private Function ConvertFile(ByRef s() As Byte, ByRef bytsOut() As Byte) As Boolean
    Dim i As Currency, x As Byte
    Dim j As Currency
    Dim k As Currency
    Dim t As Currency
    Dim a As Currency
    Dim e As Currency, z As Currency
    Dim l As Currency, u As Currency
    Dim foundword As Boolean, foundend As Boolean
    Dim word() As Byte
    Dim c As String * 1
    Dim x2 As Byte
    Dim c2 As Byte
    Dim endl As Currency, endq As Currency
    Dim endlin As Currency, q As Currency
    Dim prevchar As String * 1
    Dim strf As String
    Dim strtemp As String
    Dim stmp As String * 1
    Dim pr As Byte, xt As Currency, y As Currency
    Dim continue As Boolean
    Dim m As Currency
    Dim n(0 To 370000) As Byte, zi As Currency, zj As Currency
    Dim zk As Currency
    Dim bytresponse As Byte
    Dim endfound As Boolean, endwasfound As Boolean
    
    ReDim Preserve s(UBound(s) + 20)
    
    ReDim word(1)
 
    'On Error GoTo errorhandler
    
    i = 0
    For u = 1 To Len("<HTML><meta http-equiv=""Expire"" content=""now""> " & _
      "<meta http-equiv=""Pragma"" content=""no-cache""><BODY><font face=courier new size=-2><pre>")
      n(i) = Asc(Mid("<HTML><meta http-equiv=""Expire"" content=""now""> <meta http-equiv=""Pragma"" content=""no-cache""><BODY><font face=courier new size=-2><pre>", u, 1))
      i = i + 1
    Next u
    j = 0
    m = UBound(s) - 10
    ReDim Preserve s(m + 10)
    While j < m
        If p(s(j), s(j + 1), s(j + 2)) = 0 Then
            n(i) = s(j)
            i = i + 1
            j = j + 1
        Else
            Select Case s(j)
            Case 13: 'carriage return
                For u = 0 To UBound(bytBreak)
                    n(i) = bytBreak(u)
                    i = i + 1
                Next u
                n(i) = 13
                n(i + 1) = 10
                i = i + 2
                j = j + 2
'                While s(j) = 32 And j < m 'spaces at beginning of line
'                    n(i) = 32
''                    For u = 0 To UBound(bytNBSP)
''                        n(i) = bytNBSP(u)
''                        i = i + 1
''                    Next u
'                    j = j + 1
'                Wend
            Case 34: 'quotation mark
                n(i) = s(j)
                i = i + 1
                j = j + 1
                While (s(j) <> 34) And (j < m)
                    If s(j) = Asc("<") Then
                        n(i) = Asc("&")
                        n(i + 1) = Asc("l")
                        n(i + 2) = Asc("t")
                        i = i + 3
                        j = j + 1
                    ElseIf s(j) = Asc(">") Then
                        n(i) = Asc("&")
                        n(i + 1) = Asc("g")
                        n(i + 2) = Asc("t")
                        i = i + 3
                        j = j + 1
                    Else
                        n(i) = s(j)
                        i = i + 1
                        j = j + 1
                    End If
                Wend
                n(i) = s(j)
                i = i + 1
                j = j + 1
            Case 39: 'apostrophe
                For u = 0 To UBound(bytPreGreen)
                    n(i) = bytPreGreen(u)
                    i = i + 1
                Next u
                n(i) = s(j)
                i = i + 1
                j = j + 1
                While s(j) <> 13 And j < m
                    n(i) = s(j)
                    i = i + 1
                    j = j + 1
                Wend
                For u = 0 To UBound(bytPost)
                    n(i) = bytPost(u)
                    i = i + 1
                Next u
            Case Asc(">"):
                n(i) = Asc("&")
                n(i + 1) = Asc("g")
                n(i + 2) = Asc("t")
                i = i + 3
                j = j + 1
            Case Asc("<"):
                n(i) = Asc("&")
                n(i + 1) = Asc("l")
                n(i + 2) = Asc("t")
                i = i + 3
                j = j + 1
            Case Else: 'keyword check
                If j > 1 Then t = s(j - 1) Else t = 0
                If IsAlpha(t) Then
                    n(i) = s(j)
                    i = i + 1
                    j = j + 1
                Else
                    foundword = False
                    c2 = Asc(UCase(Chr((s(j + 1)))))
                    If Not (IsAlpha(s(j + 2))) Then
                        foundword = True
                        z = 0
                        While b(s(j), c2, 0, z) <> 0
                            n(i) = b(s(j), c2, 0, z)
                            z = z + 1
                            i = i + 1
                        Wend
                        j = j + 2
                    ElseIf Not (IsAlpha(s(j + 3))) And _
                      b(IIf(s(j) = 35, 64, s(j)), c2, 0, 0) <> 0 Then
                        endwasfound = endfound
                        endfound = False
                        foundword = True
                        If p(s(j), s(j + 1), s(j + 2)) = 11 Then
                            endfound = True
                        End If
                        z = 0
                        a = IIf(s(j) = 35, 64, s(j))
                        While b(a, c2, 0, z) <> 0
                            n(i) = b(a, c2, 0, z)
                            z = z + 1
                            i = i + 1
                        Wend
                        If endwasfound Then
                            If p(s(j), s(j + 1), s(j + 2)) = 23 Then
                                j = j + 2
                                For u = 0 To UBound(bytHR)
                                    n(i) = bytHR(u)
                                    i = i + 1
                                Next u
                                n(i) = 13
                                n(i + 1) = 10
                                i = i + 2
                            End If
                        End If
                        j = j + 3
                     Else
                        z = 0
                        a = IIf(s(j) = 35, 64, Asc(UCase(Chr(s(j)))))
                        For x = 0 To 13
                            y = 0
                            While w(a, c2, x, y) = s(j + y)
                                y = y + 1
                            Wend
                            If y = ln(a, c2, x) And y > 3 And Not (IsAlpha(s(j + y))) Then
                                endwasfound = endfound
                                endfound = False
                                foundword = True
                                z = 0
                                While b(a, c2, x, z) <> 0
                                    n(i) = b(a, c2, x, z)
                                    z = z + 1
                                    i = i + 1
                                Wend
                                If endwasfound Then
                                    If (s(j) = 80 And s(j + y - 1) = 121) Or _
                                      (s(j) = 70 And s(j + y - 1) = 110) Then
                                        j = j + 2
                                        For u = 0 To UBound(bytHR)
                                            n(i) = bytHR(u)
                                            i = i + 1
                                        Next u
                                        n(i) = 13
                                        n(i + 1) = 10
                                        i = i + 2
                                    End If
                                End If
                                
                                j = j + ln(a, c2, x)
                                                        
                                Exit For
                                
                            End If
                            
                        Next x
                        If Not foundword Then
                            n(i) = s(j)
                            i = i + 1
                            j = j + 1
                        End If
                    End If
                End If
            End Select
        End If
    Wend

    ReDim bytTemp(1)
    StoreBytes bytTemp, "</pre></BODY></HTML>"
    For u = 0 To UBound(bytTemp)
        n(i) = bytTemp(u)
        i = i + 1
    Next u
    
    u = 0
    ReDim Preserve bytsOut(i)
    bytsOut = n
    ConvertFile = True
    Exit Function
    
errorhandler:
    If Err.Number = 9 Then Resume Next
    bytresponse = MsgBox("Error during conversion: " & Err.Description, vbAbortRetryIgnore)
    If bytresponse = vbRetry Then
        Resume
    ElseIf bytresponse = vbIgnore Then
        Resume Next
    Else
        ConvertFile = False
        Exit Function
    End If
    
End Function


Public Sub Setup()
    Dim i As Currency, j As Currency, e As Currency
    Dim a As Currency, t As Currency, k As Currency
    Dim l As Currency
    Dim wl() As String
    
    StoreBytes bytPreBlue, "<font color=#000080>"
    StoreBytes bytPreGreen, "<font color=green>"
    StoreBytes bytPost, "</font>"
    StoreBytes bytBreak, "<br>"
    StoreBytes bytHR, "<hr>"
    StoreBytes bytNBSP, "&nbsp;"
               
    wl = Split(VBKEYWORDS)
    
    For i = 0 To UBound(wl) - 1
        If Asc(Mid(wl(i), 1, 1)) = 35 Then a = 64 Else a = Asc(Mid(wl(i), 1, 1))
        t = 0
        While w(a, Asc(UCase(Mid(wl(i), 2, 1))), t, 1) <> 0
            t = t + 1
        Wend
        j = 0
        ln(a, Asc(UCase(Mid(wl(i), 2, 1))), t) = Len(wl(i))
        While j < Len(wl(i))
            w(a, Asc(UCase(Mid(wl(i), 2, 1))), t, j) = Asc(Mid(wl(i), j + 1, 1))
            j = j + 1
        Wend
        For k = 0 To UBound(bytPreBlue)
            b(a, Asc(UCase(Mid(wl(i), 2, 1))), t, k) = bytPreBlue(k)
        Next k
        For j = 0 To Len(wl(i)) - 1
            b(a, Asc(UCase(Mid(wl(i), 2, 1))), t, j + k) = Asc(Mid(wl(i), j + 1, 1))
        Next j
        e = k
        For k = 0 To UBound(bytPost)
            b(a, Asc(UCase(Mid(wl(i), 2, 1))), t, j + e + k) = bytPost(k)
        Next k
    Next i
    
'   [    ] [ '  ] [    ]
'   [    ] [ "  ] [    ]
'   [ F  ] [ o  ] [ r  ]
'   [    ] [ CR ] [ LF ]
'   [ A  ] [ s  ] [ CR ]
'   [ A  ] [ s  ] [ Sp ]
        
    If Dir(App.Path & "\" & "pattern.dat") <> "" Then
        Open App.Path & "\" & "pattern.dat" For Binary As #1
        Get #1, , p
        Close #1
    Else
        For i = 0 To UBound(wl) - 1
            If Len(wl(i)) > 2 Then
                p(Asc(Mid(wl(i), 1, 1)), Asc(Mid(wl(i), 2, 1)), Asc(Mid(wl(i), 3, 1))) = i
            Else
                p(Asc(Mid(wl(i), 1, 1)), Asc(Mid(wl(i), 2, 1)), 13) = i
                p(Asc(Mid(wl(i), 1, 1)), Asc(Mid(wl(i), 2, 1)), 32) = i
            End If
        Next i
        
'        For i = 0 To 127
'            p(13, 10, i) = 1
'        Next i
        
        For i = 0 To 127
            For j = 0 To 127
                p(39, i, j) = 1
                p(34, i, j) = 1
                p(Asc("<"), i, j) = 126
                p(Asc(">"), i, j) = 127
            Next j
        Next i
        On Error Resume Next
        Kill App.Path & "\" & "pattern.dat"
        Open App.Path & "\" & "pattern.dat" For Binary As #1
        Put #1, , p
        Close #1
    End If
End Sub

Sub StoreBytes(ByRef bytX() As Byte, ByVal strX As String)
    Dim i As Currency
    
    For i = 1 To Len(strX)
        ReDim Preserve bytX(i - 1)
        bytX(i - 1) = Asc(Mid(strX, i, 1))
    Next i
End Sub

Sub AppendBytes(ByRef bytX() As Byte, ByRef bytAdd() As Byte)
    Dim i As Currency
    Dim j As Currency
    
    'On Error Resume Next
    j = UBound(bytX)
    ReDim Preserve bytX(UBound(bytX) + UBound(bytAdd) + 1)
    For i = 0 To UBound(bytAdd)
        bytX(j + i + 1) = bytAdd(i)
    Next
End Sub


Sub AppendByte(ByRef bytX() As Byte, ByVal bytAdd As Byte)
    Dim i As Currency
    Dim j As Currency
    
    j = UBound(bytX)
    ReDim Preserve bytX(UBound(bytX) + 1)
    bytX(j + 1) = bytAdd
End Sub

Function IsLC(ByVal bytchar As Byte) As Boolean
    IsLC = (bytchar > 96 And bytchar < 123)
End Function

Function IsUC(ByVal bytchar As Byte) As Boolean
    IsUC = (bytchar > 64 And bytchar < 91)
End Function

Function IsAlpha(ByVal bytchar As Byte) As Boolean
    IsAlpha = (bytchar > 96 And bytchar < 123) Or _
     (bytchar > 64 And bytchar < 91) Or _
     (bytchar = 46) Or _
     (bytchar = 95) Or _
     (bytchar > 47 And bytchar < 58)
End Function


Sub PrintBytes(ByRef bytX() As Byte)
    Dim i As Currency
    
    For i = 0 To UBound(bytX)
        Debug.Print Chr(bytX(i));
    Next i
    Debug.Print
End Sub

Private Sub cmdNewOutDir_Click()
    strOutputDir = BrowseForFolder(Me.hWnd, "Select a Folder to Write to", "c:\")
    lblOutDirPath = strOutputDir
End Sub

Private Sub Form_Load()
    lvProject.ColumnHeaders.Add , , "Name", lvProject.Width / 10 * 7, lvwColumnLeft
    lvProject.ColumnHeaders.Add , , "Status", lvProject.Width / 10 * 3, lvwColumnLeft
    lvProject.View = lvwReport
    mnuItem.Enabled = False
    mnuItemView.Enabled = False
    mnuItemEdit.Enabled = False
    ReadSettings
    If strProjectFileName = "" Then
        lblHelp = "Click on Load Picture to select a Visual Basic project file."
        cmdConvert.Enabled = False
        cmdLoad_Click
        cmdConvert.Enabled = True
    Else
        lblHelp.Font.Size = 7
        lblHelp = "Click the Output To button to change the directory the HTML files will be written to."
        lblHelp = lblHelp & vbCrLf & vbCrLf & "Click on Load Project to select a Visual Basic project file."
        lblHelp = lblHelp & vbCrLf & vbCrLf & "Click on Make HTML to convert."
    End If
    Setup
    OpenProject
End Sub

Private Sub SaveSettings()
    ' Save the name of the last project worked on
    SaveSetting App.EXEName, "General", "ProjectFileName", strProjectFileName
    ' Save the default output directory
    SaveSetting App.EXEName, "General", "OutputDirectory", strOutputDir
End Sub

Private Sub ReadSettings()
    ' Read the name of the last project worked on
    strProjectFileName = GetSetting(App.EXEName, "General", "ProjectFileName", "")
    ' Read the default output directory
    strOutputDir = GetSetting(App.EXEName, "General", "OutputDirectory", "C:\")
    lblOutDirPath = strOutputDir
End Sub

Private Sub Form_Unload(Cancel As Integer)
    strOutputDir = lblOutDirPath
    SaveSettings
End Sub

Private Sub lvProject_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Set gCurItem = Item
    PopupMenu mnuItem
End Sub

Private Sub mnuItemEdit_Click()
    'Dim strFName As String
    'Dim strtemp As String
    'strtemp = strItemOuts(gCurItem.Index - 1)
    Shell "c:\program files\accessories\wordpad.exe " & Chr(34) & strItemOuts(gCurItem.Index - 1) & Chr(34), vbNormalFocus
    'strFName = Mid(strtemp, InStrRev(strtemp, "\") + 1)
    'AppActivate strFName
End Sub

Private Sub mnuitemview_click()
    Dim strShortcutFName As String
    
    strShortcutFName = strItemOuts(gCurItem.Index - 1) & ".url"
    Open strShortcutFName For Output As #1
    Print #1, "[InternetShortcut]"
    Print #1, "URL=" & strItemOuts(gCurItem.Index - 1)
    Close #1
    Shell "rundll32.exe shdocvw.dll,OpenURL " & strShortcutFName, vbNormalFocus
End Sub

