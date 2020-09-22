VERSION 5.00
Begin VB.Form frmJunkMail 
   Caption         =   "Edit Junk Mail Senders"
   ClientHeight    =   6645
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6450
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmJunkMail.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6645
   ScaleWidth      =   6450
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAbout 
      Caption         =   "&About"
      Height          =   315
      Left            =   5640
      TabIndex        =   5
      Top             =   5840
      Width           =   630
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   315
      Left            =   5640
      TabIndex        =   4
      Top             =   5455
      Width           =   630
   End
   Begin VB.TextBox txtJunk 
      Height          =   5715
      Left            =   60
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   885
      Width           =   5310
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "&Load"
      Height          =   315
      Left            =   5640
      TabIndex        =   1
      Top             =   5070
      Width           =   630
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   315
      Left            =   5640
      TabIndex        =   0
      Top             =   6225
      Width           =   630
   End
   Begin VB.Label lblInstructions 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   $"frmJunkMail.frx":030A
      Height          =   690
      Left            =   90
      TabIndex        =   3
      Top             =   75
      Width           =   6285
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmJunkMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' API routine for grabbing the AppData path with the Windows shell
Private Const CSIDL_APPDATA = &H1A

Private Type SHITEMID
    CB   As Long
    abID As Byte
End Type

Private Type ITEMIDLIST
    mkID As SHITEMID
End Type

Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal PV As Long)

Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hWndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Public Function GetShellAppdataLocation(lnghWnd As Long) As String

    ' Comments  : Returns the path of the user's "Appdata" folder
    ' Parameters: lnghWnd - handle to window to serve as the
    '               parent for the dialog. Uses the frmJunkMail
    '               form's hWnd property
    ' Returns   : Path of the user's Appdata folder
    
    Dim lngResult As Long
    Dim strPath   As String
    Dim idlist    As ITEMIDLIST
    
    On Error GoTo Err_GetShellAppdataLocation
    
    ' populate an ITEMIDLIST struct with the specified folder information
    lngResult = SHGetSpecialFolderLocation(lnghWnd, CSIDL_APPDATA, idlist)
        If lngResult = 0 Then
            ' if the information is present, get the path information
            strPath = Space$(260)
            lngResult = SHGetPathFromIDList(ByVal idlist.mkID.CB, ByVal strPath)
            ' free memory allocated by shell
            CoTaskMemFree idlist.mkID.CB
            'if a path was found, trim off trailing null char
            strPath = Left$(strPath, InStr(1, strPath, vbNullChar) - 1)
            GetShellAppdataLocation = strPath
        End If

Exit_GetShellAppdataLocation:
    
    On Error GoTo 0
    Exit Function

Err_GetShellAppdataLocation:

    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In mod1, during GetShellAppdataLocation" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_GetShellAppdataLocation
    End Select

End Function
Public Sub File2TextBox(sFile As String, oText As TextBox)
   
    ' Purpose   : fills textbox with contents of file
    ' Parameters: sFile, oText
    ' Modified  : 2/20/2002 By BB
    
    Dim FNum  As Integer
    Dim sTemp As String
    
    On Error GoTo Err_File2TextBox
    
    FNum = FreeFile()
    oText.Text = ""
    Open sFile For Input As FNum
        While Not EOF(FNum)
            Line Input #FNum, sTemp
                If oText.Text = "" Then
                    oText.Text = sTemp
                Else
                    oText.Text = oText.Text & vbCrLf & sTemp
                End If
        Wend
    Close FNum

Exit_File2TextBox:
    
    On Error GoTo 0
    Exit Sub

Err_File2TextBox:

    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In frmJunkMail, during File2TextBox" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_File2TextBox
    End Select

End Sub
Public Function TextBox2File(sFile As String, oText As TextBox) As Boolean
   
    ' Purpose   : grabs textbox text into an array, and
    '               writes back to file one line at a time
    ' Parameters: sFile, oText
    ' Modified  : 2/20/2002 By BB
    
    Dim FNum      As Integer
    Dim X         As Integer
    Dim sTemp     As String
    Dim arrText() As String
    
    On Error GoTo Err_TextBox2File
    
    FNum = FreeFile()
    X = 0
    Open sFile For Output As FNum
    arrText = Split(oText, vbCrLf)
        For X = LBound(arrText) To UBound(arrText)
            If arrText(X) = vbCrLf Then
                ' if all the line contains is a Cr and Lf, skip it
            Else
                Print #FNum, arrText(X)
            End If
        Next
    Close FNum
    TextBox2File = True
    
Exit_TextBox2File:
    
    On Error Resume Next
    Erase arrText
    On Error GoTo 0
    Exit Function

Err_TextBox2File:

    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In frmJunkMail, during TextBox2File" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            TextBox2File = False
            Resume Exit_TextBox2File
    End Select
    
End Function
Private Function GetOutlookJunkMailSendersPath() As String
   
    ' Purpose   : calls API function that gets current user's
    '               Outlook Junk Senders list text file
    ' Returns   : String
    ' Modified  : 2/20/2002 By BB
    
    Dim strPath As String
    Dim strUser As String
    
    On Error GoTo Err_GetOutlookJunkMailSendersPath
    
    strPath = GetShellAppdataLocation(frmJunkMail.hwnd)
    strUser = Replace(strPath, ":\Documents and Settings\", "")
    strUser = Replace(strUser, "\Application Data", "")
    strUser = Right$(strUser, Len(strUser) - 1)
    strPath = strPath & "\Microsoft\Outlook\Junk Senders.txt"
        If Format$(strPath) <> "" Then
            GetOutlookJunkMailSendersPath = strPath
        Else
            ' whoops
            MsgBox "For some reason, the system couldn't find your Microsoft Outlook Junk Senders file" & vbCrLf & vbCrLf & "           What could be wrong..." & vbCrLf & vbCrLf & "Be sure you have Microsoft Outlook installed on this PC" & vbCrLf & vbCrLf & "Windows reports your user name as " & strUser & ", is this your name?" & vbCrLf & vbCrLf & vbCrLf & "If this information is correct, something else must be wrong...so sorry!", vbOKOnly + vbExclamation + vbDefaultButton1, "                  Junk Senders Not Found"
            GetOutlookJunkMailSendersPath = ""
        End If

Exit_GetOutlookJunkMailSendersPath:
    
    On Error GoTo 0
    Exit Function

Err_GetOutlookJunkMailSendersPath:

    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In frmJunkMail, during GetOutlookJunkMailSendersPath" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_GetOutlookJunkMailSendersPath
    End Select
    
End Function
Private Sub cmdAbout_Click()
   
    ' ego time!
    
    On Error GoTo Err_cmdAbout_Click
    
    MsgBox "Edit Microsoft Outlook Junk Senders List" & vbCrLf & vbCrLf & "     Written By Brian Battles WS1O" & vbCrLf & "          brianb@cmtelephone.com" & vbCrLf & "       http://www.battleszone.com/ ", vbOKOnly + vbInformation, "Edit Junk Senders List by Brian Battles WS1O"

Exit_cmdAbout_Click:
    
    On Error GoTo 0
    Exit Sub

Err_cmdAbout_Click:

    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In frmJunkMail, during cmdAbout_Click" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_cmdAbout_Click
    End Select
    
End Sub
Private Sub cmdClose_Click()
   
    On Error GoTo Err_cmdClose_Click
    
    Unload Me

Exit_cmdClose_Click:
    
    On Error GoTo 0
    Exit Sub

Err_cmdClose_Click:

    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In frmJunkMail, during cmdClose_Click" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_cmdClose_Click
    End Select
    
End Sub
Private Sub cmdLoad_Click()
   
    ' Purpose   : runs common dialog to get the file, then calls
    '               the routine that loads its contents into the listbox
    ' Modified  : 2/20/2002 By BB

    On Error GoTo Err_cmdLoad_Click
    
    Dim strPath As String
    
    Screen.MousePointer = vbHourglass
    strPath = GetOutlookJunkMailSendersPath
    If Format$(strPath) = "" Then
        ' whoops!
        GoTo Exit_cmdLoad_Click
    End If
    File2TextBox strPath, txtJunk
    If Format$(txtJunk.Text) = "" Then
        ' no text
        lblInstructions.Caption = "Either your MS Outlook Junk Senders list couldn't be found, or you don't have one. Please check your Microsoft Outlook documentation if you need help with this feature"
    Else
        lblInstructions.Caption = "Here's the list of Junk Senders you defined with the Microsoft Outlook Rules Wizard. Now you can edit, delete, add to this list until it has exactly what you want, then press Save to keep your changes (or just press Close to quit and cancel any changes)"
    End If

Exit_cmdLoad_Click:
    
    Screen.MousePointer = vbDefault
    On Error GoTo 0
    Exit Sub

Err_cmdLoad_Click:

    Select Case Err
        Case 0
            Resume Next
        Case 53, 32755 ' user cancelled
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In frmJunkMail, during cmdLoad_Click" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_cmdLoad_Click
    End Select
    
End Sub
Private Sub cmdSave_Click()
   
    ' Purpose   : gets path to user's Outlook Junk Senders file, then writes
    '               contents of text box to it
    ' Modified  : 2/20/2002 By BB

    On Error GoTo Err_cmdSave_Click
    
    Screen.MousePointer = vbHourglass
    If Format$(txtJunk) = "" Then
        If MsgBox("              There aren't any addresses on your list" & vbCrLf & vbCrLf & "That means your Junk Mail Senders file will contain no one" & vbCrLf & vbCrLf & "                   Are you sure you want to save this? " & vbCrLf, vbYesNo + vbQuestion + vbDefaultButton1, "         Confirm Empty File") = vbNo Then
            GoTo Exit_cmdSave_Click
        Else
            ' let 'em save it
        End If
    End If
    If TextBox2File(GetOutlookJunkMailSendersPath, txtJunk) Then
        MsgBox "Junk Mail Senders File Saved", vbInformation, "   Success!"
    Else
        MsgBox "Junk Mail Senders File NOT Saved", vbInformation, "   Phooey!"
    End If
    
Exit_cmdSave_Click:
    
    Screen.MousePointer = vbDefault
    On Error GoTo 0
    Exit Sub

Err_cmdSave_Click:

    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In frmJunkMail, during cmdSave_Click" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_cmdSave_Click
    End Select
    
End Sub
