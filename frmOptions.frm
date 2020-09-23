VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6150
   Icon            =   "frmOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "Options"
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   2490
      TabIndex        =   1
      Tag             =   "OK"
      Top             =   4455
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3720
      TabIndex        =   3
      Tag             =   "Cancel"
      Top             =   4455
      Width           =   1095
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Apply"
      Height          =   375
      Left            =   4920
      TabIndex        =   5
      Tag             =   "&Apply"
      Top             =   4455
      Width           =   1095
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   3
      Left            =   -20000
      ScaleHeight     =   3840.968
      ScaleMode       =   0  'User
      ScaleWidth      =   5745.64
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample4 
         Caption         =   "Sample 4"
         Height          =   2022
         Left            =   505
         TabIndex        =   11
         Tag             =   "Sample 4"
         Top             =   502
         Width           =   2033
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   2
      Left            =   -20000
      ScaleHeight     =   3840.968
      ScaleMode       =   0  'User
      ScaleWidth      =   5745.64
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample3 
         Caption         =   "Sample 3"
         Height          =   2022
         Left            =   406
         TabIndex        =   10
         Tag             =   "Sample 3"
         Top             =   403
         Width           =   2033
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   1
      Left            =   -20000
      ScaleHeight     =   3840.968
      ScaleMode       =   0  'User
      ScaleWidth      =   5745.64
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample2 
         Caption         =   "Sample 2"
         Height          =   2022
         Left            =   307
         TabIndex        =   8
         Tag             =   "Sample 2"
         Top             =   305
         Width           =   2033
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   0
      Left            =   210
      ScaleHeight     =   3840.968
      ScaleMode       =   0  'User
      ScaleWidth      =   5745.64
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample1 
         Caption         =   "Main system Options"
         Height          =   3705
         Left            =   0
         TabIndex        =   4
         Tag             =   "Sample 1"
         Top             =   30
         Width           =   5640
         Begin VB.TextBox txtEmailChanges 
            Height          =   285
            Left            =   1860
            TabIndex        =   20
            Top             =   1260
            Width           =   3165
         End
         Begin MSComCtl2.UpDown UpDown1 
            Height          =   285
            Left            =   2371
            TabIndex        =   18
            Top             =   930
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   393216
            Value           =   5
            BuddyControl    =   "txtSlideShow"
            BuddyDispid     =   196617
            OrigLeft        =   2310
            OrigTop         =   900
            OrigRight       =   2565
            OrigBottom      =   1215
            Max             =   100
            Min             =   1
            Enabled         =   -1  'True
         End
         Begin VB.TextBox txtSlideShow 
            Height          =   285
            Left            =   1860
            Locked          =   -1  'True
            TabIndex        =   17
            Text            =   "5"
            Top             =   930
            Width           =   510
         End
         Begin VB.CommandButton cmdFind 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   4680
            Picture         =   "frmOptions.frx":058A
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Find an Existing PLU"
            Top             =   630
            Width           =   330
         End
         Begin VB.TextBox txtHomeIndividual 
            Height          =   285
            Left            =   1860
            Locked          =   -1  'True
            TabIndex        =   14
            Top             =   630
            Width           =   2775
         End
         Begin VB.TextBox txtFamName 
            Height          =   285
            Left            =   1860
            TabIndex        =   12
            Top             =   330
            Width           =   3165
         End
         Begin VB.Label Label5 
            Caption         =   "Family Name:"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   1290
            Width           =   1515
         End
         Begin VB.Label Label4 
            Caption         =   "Seconds"
            Height          =   195
            Left            =   2670
            TabIndex        =   22
            Top             =   960
            Width           =   705
         End
         Begin VB.Label Label3 
            Caption         =   "Slide Show Interval:"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   960
            Width           =   1935
         End
         Begin VB.Label Label2 
            Caption         =   "Principle Individual:"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   660
            Width           =   1515
         End
         Begin VB.Label Label1 
            Caption         =   "Family Name:"
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   360
            Width           =   1515
         End
      End
   End
   Begin MSComctlLib.TabStrip tbsOptions 
      Height          =   4245
      Left            =   105
      TabIndex        =   0
      Top             =   120
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   7488
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Main"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Function InvokeOptions() As Boolean
    
    Call LoadOptions
    
    Me.Show vbModal

End Function

Private Sub cmdApply_Click()
    If Not SaveOptions Then
        MsgBox "Error saving these options.", vbOKOnly Or vbCritical, Me.Caption
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
Dim lNewId As Long

    lNewId = frmIndex.invoke(0, 2099, "B")
    
    If lNewId <> Val(txtHomeIndividual.Tag) And lNewId > 0 Then
        txtHomeIndividual.Tag = lNewId
        txtHomeIndividual = GetFullName(lNewId, False, True)
    End If
    
End Sub

Private Sub cmdOK_Click()
    If SaveOptions Then
        Unload Me
    Else
        MsgBox "Error saving these options.", vbOKOnly Or vbCritical, Me.Caption
    End If
End Sub

Private Sub LoadOptions()
Dim SQL As String
Dim RS As ADODB.Recordset
Dim sErr As String

    On Error GoTo ErrSub
    
    SQL = "Select * from Options"
    
    Set RS = New ADODB.Recordset
    
    RS.Open SQL, gApp.cn, adOpenForwardOnly, adLockReadOnly
    
    Do While Not RS.EOF
        Select Case RS(gccOPTID)
            Case 1  'Family History for Surname
                txtFamName = RS(gccOPTVALUE)
            Case 2  'The Id of the 'Home' individual
                txtHomeIndividual.Tag = CLng(RS(gccOPTVALUE))
                txtHomeIndividual = GetFullName(Val(txtHomeIndividual.Tag), False, True)
            Case 3
                UpDown1 = CLng(RS(gccOPTVALUE))
                txtSlideShow = RS(gccOPTVALUE)
            Case 4
                txtEmailChanges = RS(gccOPTVALUE)
        End Select
        RS.MoveNext
    Loop
    
    
Exit Sub
ErrSub:
    sErr = Err.Number & vbCrLf & Err.Description & vbCrLf & vbCrLf & _
            "In Module " & Me.Name & vbCrLf & _
            "In Function GetIndividual"
            
    Call Showerror(sErr, 0)

End Sub

Private Function SaveOptions() As Boolean
Dim SQL As String
Dim sErr As String

    On Error GoTo ErrSub

    SQL = "Update " & gtcOPTIONS & " Set " & _
            gccOPTVALUE & " = '" & txtFamName & "' WHERE " & _
            gccOPTID & " = 1"
    
    gApp.cn.Execute SQL
    
    SQL = "Update " & gtcOPTIONS & " Set " & _
            gccOPTVALUE & " = '" & txtHomeIndividual.Tag & "' WHERE " & _
            gccOPTID & " = 2"
    
    gApp.cn.Execute SQL
    
    SaveOptions = True
    
    SQL = "Update " & gtcOPTIONS & " Set " & _
            gccOPTVALUE & " = '" & txtSlideShow & "' WHERE " & _
            gccOPTID & " = 3"
    
    gApp.cn.Execute SQL
    
    SQL = "Update " & gtcOPTIONS & " Set " & _
            gccOPTVALUE & " = '" & txtEmailChanges & "' WHERE " & _
            gccOPTID & " = 4"
    
    gApp.cn.Execute SQL
    
    
    SaveOptions = True
    
    
Exit Function
ErrSub:
    sErr = Err.Number & vbCrLf & Err.Description & vbCrLf & vbCrLf & _
            "In Module " & Me.Name & vbCrLf & _
            "In Function GetIndividual"
            
    Call Showerror(sErr, 0)
            
End Function

Private Sub UpDown1_Change()
    txtSlideShow = UpDown1.Value
End Sub
