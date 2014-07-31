VERSION 5.00
Begin VB.Form Main_Screen 
   Appearance      =   0  'Flat
   BackColor       =   &H000080FF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Application Loader"
   ClientHeight    =   2490
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6735
   Icon            =   "Application_Loader.frx":0000
   LinkTopic       =   "Main_Screen"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2490
   ScaleWidth      =   6735
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1575
      Left            =   0
      ScaleHeight     =   1575
      ScaleWidth      =   6735
      TabIndex        =   0
      Top             =   480
      Width           =   6735
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   2160
      Width           =   6495
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6495
   End
End
Attribute VB_Name = "Main_Screen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private application_name As String
Private application_excutable As String
Private application_icon As String
Private application_splashimage As String
Private application_netmajorversion As String
Private application_netminorversion As String
Private application_buildnumber As String
Private application_colour As String
Private application_forecolour As String
Private application_download As String

Private Sub Form_Load()

Dim myFSO As New Scripting.FileSystemObject

Dim path As String
path = App.path & "\Inputs\"
path = Replace(path, "\\", "\")


Set ts = myFSO.OpenTextFile(path & "inputs.txt", ForReading, False, TristateUseDefault)

application_name = Trim(Replace(ts.ReadLine, "application_name = ", ""))
application_excutable = Trim(Replace(ts.ReadLine, "application_excutable = ", ""))
application_icon = Trim(Replace(ts.ReadLine, "application_icon = ", ""))
application_splashimage = Trim(Replace(ts.ReadLine, "application_splashimage = ", ""))
application_netmajorversion = Trim(Replace(ts.ReadLine, "application_netmajorversion = ", ""))
application_netminorversion = Trim(Replace(ts.ReadLine, "application_netminorversion = ", ""))
application_buildnumber = Trim(Replace(ts.ReadLine, "application_buildnumber = ", ""))
application_colour = CvtColorWeb2VB(Trim(Replace(ts.ReadLine, "application_colour = ", "")))
application_forecolour = CvtColorWeb2VB(Trim(Replace(ts.ReadLine, "application_forecolour = ", "")))
application_download = Trim(Replace(ts.ReadLine, "application_download = ", ""))

ts.Close

Me.BackColor = application_colour
'Picture2.BackColor = application_colour
Picture1.BackColor = application_colour
Label1.BackColor = application_colour
Label1.ForeColor = application_forecolour
Label1.Caption = application_name & " Loader"
Me.Caption = application_name & " Requirement Check"
Picture1.BackColor = application_colour
Picture1.Picture = LoadPicture(path & application_splashimage)
'Picture2.BackColor = application_colour
'Picture2.Picture = LoadPicture(path & application_icon)
Label2.BackColor = application_colour
Label2.ForeColor = application_forecolour
Label2.Caption = "Checking for .NET Framework Installation..."


  Dim lngRootKey As Long
  Dim strKeyQuery As Variant
  
  strKeyQuery = vbNullString
  lngRootKey = HKEY_LOCAL_MACHINE

Dim proceed, major, minor, build, overmajor As Boolean
proceed = False
major = False
minor = False
build = False
overmajor = False

Dim testmajor, testminor As Integer
testmajor = application_netmajorversion
testminor = application_netminorversion

    strKeyQuery = regDoes_Key_Exist(lngRootKey, "SOFTWARE\Microsoft\.NETFramework\policy\v" & testmajor & "." & testminor)
    If strKeyQuery = True Then
        proceed = True
        major = True
        minor = True
    End If

Dim outerloop, innerloop As Integer

If proceed = False Then

For outerloop = testmajor To testmajor + 10
For innerloop = 0 To testminor + 10
    strKeyQuery = regDoes_Key_Exist(lngRootKey, "SOFTWARE\Microsoft\.NETFramework\policy\v" & outerloop & "." & innerloop)
    If strKeyQuery = True Then
        proceed = True
        major = True
        minor = True
        If outerloop > testmajor Then
            overmajor = True
        End If
    End If
    strKeyQuery = regQuery_A_Key(lngRootKey, "SOFTWARE\Microsoft\.NETFramework\policy\v" & outerloop & "." & innerloop, application_buildnumber)
        If Not strKeyQuery = "" Then
        proceed = True
        major = True
        minor = True
        build = True
        If outerloop > testmajor Then
            overmajor = True
        End If
    End If
    strKeyQuery = regDoes_Key_Exist(lngRootKey, "SOFTWARE\Microsoft\NET Framework Setup\NDP\v" & outerloop & "." & innerloop & "." & application_buildnumber)
    If strKeyQuery = True Then
        proceed = True
        major = True
        minor = True
        build = True
        If outerloop > testmajor Then
            overmajor = True
        End If
    End If
Next innerloop
Next outerloop

End If

Dim pid As Double
application_excutable = Replace(App.path & "\" & application_excutable, "\\", "\")
    
If (major = True And minor = True) Or (overmajor = True) Then
    Label2.Caption = "Usable .NET Framework Installation Found..."
    If Dir$(application_excutable) <> "" Then
        pid = Shell(application_excutable, vbNormalFocus)
        Unload Me
    Else
        Label2.Caption = "Program Executable Cannot Be Located."
    End If
    Else
        Label2.Caption = "Required .NET Framework Installation (v" & application_netmajorversion & "." & application_netminorversion & ") Not Found.  (More Info.)"
    End If
   
End Sub

Private Function CvtColorVB2Web(colorcode As String) As _
    String
Dim vcolor
    vcolor = Hex(Val(colorcode))
    If Len(vcolor) < 6 Then
       vcolor = String(6 - Len(vcolor), "0") & vcolor
    End If
    CvtColorVB2Web = Mid(vcolor, 5, 2) & Mid(vcolor, 3, 2) _
        & Mid(vcolor, 1, 2)
End Function

Private Function CvtColorWeb2VB(colorcode As String) As _
    String
    If Len(colorcode) < 6 Then
       colorcode = colorcode & String(6 - Len(colorcode), _
           "0")
    End If
    CvtColorWeb2VB = "&H" & Mid(colorcode, 5, 2) & _
        Mid(colorcode, 3, 2) & Mid(colorcode, 1, 2)
End Function


Private Sub Label2_Click()
Dim result As VbMsgBoxResult
If Label2.Caption = "Program Executable Cannot Be Located." Then

result = MsgBox("Program Executable Cannot Be Located (" & application_excutable & "). Please ensure that this file is present in the application folder.", vbInformation, "Executable Not Found")
Clipboard.Clear
Clipboard.SetText application_excutable

If result = vbOK Or result = vbCancel Or result = vbYes Then
'Unload Me
End If

    Else

result = MsgBox("No usable .NET Framework (v" & application_netmajorversion & "." & application_netminorversion - 1 & ") was located on your system. Without this, " & application_name & " cannot launch. You can obtain the necessary .NET installation files from the following location: " & vbCrLf & "[" & application_download & "]" & vbCrLf & "(This location has automatically been copied to the clipboard)", vbInformation, ".NET Framework Requirement Not Met")
Clipboard.Clear
Clipboard.SetText application_download

If result = vbOK Or result = vbCancel Or result = vbYes Then
'Unload Me
End If
End If

End Sub
