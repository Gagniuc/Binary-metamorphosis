VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Binary metamorphosis (V1.0)"
   ClientHeight    =   7800
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15735
   LinkTopic       =   "Form1"
   ScaleHeight     =   7800
   ScaleWidth      =   15735
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer_Row_hex_size 
      Interval        =   100
      Left            =   16560
      Top             =   7320
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Parameters"
      Height          =   5655
      Left            =   12000
      TabIndex        =   2
      Top             =   120
      Width           =   3495
      Begin VB.TextBox Text_path 
         BackColor       =   &H00404040&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   240
         TabIndex        =   11
         Text            =   "..."
         Top             =   1680
         Width           =   3015
      End
      Begin VB.CommandButton Open_file 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Open file"
         Height          =   855
         Left            =   240
         TabIndex        =   10
         Top             =   480
         Width           =   3015
      End
      Begin VB.PictureBox bar 
         BackColor       =   &H00808080&
         Height          =   375
         Left            =   240
         ScaleHeight     =   21
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   197
         TabIndex        =   8
         Top             =   4920
         Width           =   3015
      End
      Begin VB.HScrollBar Row_hex_size 
         Height          =   255
         Left            =   240
         Max             =   300
         Min             =   10
         TabIndex        =   4
         Top             =   3840
         Value           =   70
         Width           =   2415
      End
      Begin VB.TextBox File_name 
         BackColor       =   &H00404040&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   240
         TabIndex        =   3
         Text            =   "tini.exe"
         Top             =   2760
         Width           =   3015
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Path:"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   1440
         Width           =   2055
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00808080&
         Height          =   1815
         Left            =   120
         Top             =   360
         Width           =   3255
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00808080&
         Height          =   855
         Left            =   120
         Top             =   2400
         Width           =   3255
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00808080&
         Height          =   855
         Left            =   120
         Top             =   4560
         Width           =   3255
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00808080&
         Height          =   855
         Left            =   120
         Top             =   3480
         Width           =   3255
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Conversion status:"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   4680
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Hex row size:"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   3600
         Width           =   2055
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Extracted file name:"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Label size_row_c 
         BackStyle       =   0  'Transparent
         Caption         =   "50 chr"
         Height          =   255
         Left            =   2760
         TabIndex        =   5
         Top             =   3840
         Width           =   615
      End
   End
   Begin VB.TextBox VBCODE 
      BackColor       =   &H00000040&
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   7335
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   120
      Width           =   11535
   End
   Begin VB.CommandButton Convert_to_VB 
      Caption         =   "&Convert EXE to VB6 source code"
      Height          =   1335
      Left            =   12000
      TabIndex        =   0
      Top             =   6000
      Width           =   3495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'########################################## -->
'#                                        # -->
'#  Name   : Binary metamorphosis V1.0    # -->
'#  Use    : Pack EXE to VB6 source code  # -->
'#  Author : Dr. Paul A. Gagniuc          # -->
'#  Area   : European Union               # -->
'#  Date   : 01/01/2013                   # -->
'#                                        # -->
'#  Mode   : Visual Basic 6.0             # -->
'#                                        # -->
'########################################## -->

Dim sFile As String
Dim BAS_CONTENT As String


Private Sub Convert_to_VB_Click()

    Dim c As String
    
    FF = FreeFile
    Trial = Text_path.Text
    
    Open Trial For Binary Access Read As #FF
        c = Space$(LOF(FF))
        Get #1, 1, c
    Close #FF
    
    VBHEX = VBHEX & HexString(c)
    
    step_line = Row_hex_size.Value
    u = 0
    
    For i = 1 To Len(VBHEX) Step step_line
        u = u + 1
        VB_line = VB_line & "   Fragment$(" & u & ") = " & Chr(34) & Mid(VBHEX, i, step_line) & Chr(34) & vbCrLf
    Next i
    
    
    t = t & "Sub Main()" & vbCrLf
    t = t & vbCrLf
    t = t & "   Dim Fragment$(" & u & ")" & vbCrLf
    t = t & vbCrLf & VB_line & vbCrLf
    t = t & "   DF = FreeFile" & vbCrLf
    t = t & "   Open App.Path & " & Chr(34) & "\" & File_name.Text & Chr(34) & " For Output As #DF" & vbCrLf
    t = t & "       For i = 1 To " & u & vbCrLf
    t = t & "           a$ = Fragment$(i)" & vbCrLf
    t = t & "           While Len(a$) > 0" & vbCrLf
    t = t & "           b$ = " & Chr(34) & "&H" & Chr(34) & " & Left$(a$, 2)" & vbCrLf
    t = t & "           a$ = Right$(a$, Len(a$) - 2)" & vbCrLf
    t = t & "           Print #DF, Chr$(Val(b$));" & vbCrLf
    t = t & "           Wend" & vbCrLf
    t = t & "       Next i" & vbCrLf
    t = t & "   Close #DF" & vbCrLf
    t = t & vbCrLf
    t = t & "   MsgBox " & Chr(34) & File_name.Text & " extracted to app path !" & Chr(34) & ", vbInformation" & vbCrLf
    t = t & "   End" & vbCrLf
    t = t & vbCrLf
    t = t & "End Sub" & vbCrLf
    
    BAS_CONTENT = t
    VBCODE.Text = t

End Sub


Public Function HexString(ByVal EvalString As String) As String

    Dim intStrLen As Variant
    Dim intLoop As Variant
    Dim strHex As String
    
    bar.Cls
    
    intStrLen = Len(EvalString)
    
    For intLoop = 1 To intStrLen
        strHex = strHex & Right$(" " & Hex(Asc(Mid(EvalString, intLoop, 1))), 2)
        bar.Line (0, 100)-((bar.ScaleWidth / intStrLen) * intLoop, 0), &H40&, BF            ' &H008080FF&
        DoEvents
    Next
    
    HexString = strHex
    
End Function


Private Sub Form_Load()
    Text_path.Text = App.Path & "\tini.executable"
End Sub


Private Sub Open_file_Click()

        Dim CC As cCommonDialog
        Set CC = New cCommonDialog
        
        If CC.VBGetOpenFileName(sFile, , True, , , True, "EXE files (*.exe, *.bin, *.jpg)|*.exe;*.bin;*.jpg|All files|*.*", , , "Open a EXE file", "", Form1.hWnd, 0) Then
            Text_path.Text = sFile
        End If
    
End Sub


Private Sub Row_hex_size_Change()
    Convert_to_VB_Click
End Sub


Private Sub Save_BAS_Click()

    Dim fname As String
    Dim CC As cCommonDialog
    Set CC = New cCommonDialog
    
    fname = Split(File_name.Text, ".")(0) & ".bas"
    
    Dim CurEnvFile As String
    
    If CC.VBGetSaveFileName(fname, CurEnvFile, , "BAS file save (*.bas)|*.bas;|All files|*.*", , "", "bla bla", ".bas", Form1.hWnd, 0) Then
        file1 = FreeFile
        Open CurEnvFile For Append As file1
        Print #file1, BAS_CONTENT
        Close #file1
    End If
        
End Sub


Private Sub Timer_Row_hex_size_Timer()
    size_row_c.Caption = Row_hex_size.Value & " chr"
End Sub

