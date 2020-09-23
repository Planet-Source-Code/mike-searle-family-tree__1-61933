VERSION 5.00
Begin VB.Form frmNewPerson 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   2250
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4860
   Icon            =   "frmNewPerson.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   4860
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtDOB 
      Height          =   285
      Left            =   1200
      MaxLength       =   30
      TabIndex        =   2
      Top             =   840
      Width           =   1875
   End
   Begin VB.OptionButton optGender 
      Caption         =   "Male"
      Enabled         =   0   'False
      Height          =   195
      Index           =   0
      Left            =   1215
      TabIndex        =   3
      Top             =   1410
      Value           =   -1  'True
      Width           =   735
   End
   Begin VB.OptionButton optGender 
      Caption         =   "Female"
      Enabled         =   0   'False
      Height          =   195
      Index           =   1
      Left            =   2025
      TabIndex        =   4
      Top             =   1410
      Width           =   885
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   345
      Left            =   3750
      TabIndex        =   6
      Top             =   1740
      Width           =   885
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Ok"
      Height          =   345
      Left            =   2790
      TabIndex        =   5
      Top             =   1740
      Width           =   885
   End
   Begin VB.TextBox txtSurname 
      Height          =   285
      Left            =   1200
      MaxLength       =   30
      TabIndex        =   0
      Top             =   210
      Width           =   3435
   End
   Begin VB.TextBox txtFirstNames 
      Height          =   285
      Left            =   1200
      MaxLength       =   30
      TabIndex        =   1
      Top             =   510
      Width           =   3435
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "D.O.B.:"
      Height          =   195
      Left            =   570
      TabIndex        =   10
      Top             =   870
      Width           =   525
   End
   Begin VB.Label lblGender 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Gender:"
      Enabled         =   0   'False
      Height          =   195
      Left            =   510
      TabIndex        =   9
      Top             =   1410
      Width           =   570
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Surname:"
      Height          =   195
      Left            =   420
      TabIndex        =   8
      Top             =   240
      Width           =   675
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "First Name(s):"
      Height          =   195
      Left            =   135
      TabIndex        =   7
      Top             =   540
      Width           =   960
   End
End
Attribute VB_Name = "frmNewPerson"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mlNewId As Long
Dim mlRelId As Long
Dim meRelation As eRelationships

Public Function invoke(lRelId As Long, sRelName As String, eRelation As eRelationships, _
                        sSurname As String, sFirstNames As String) As Long
Dim sCap As String

    meRelation = eRelation

    Select Case eRelation
        Case eFather
            sCap = "Father of "
            optGender(0).Value = True
        Case eMother
            sCap = "Mother of "
            optGender(1).Value = True
        Case eHusband
            sCap = "Husband of "
            optGender(0).Value = True
        Case eWife
            sCap = "Wife of "
            optGender(1).Value = True
        Case eChild
            sCap = "Child of "
            lblGender.Enabled = True
            optGender(0).Enabled = True
            optGender(1).Enabled = True
        Case eNone
            sCap = "New unrelated person "
            lblGender.Enabled = True
            optGender(0).Enabled = True
            optGender(1).Enabled = True
    End Select
    
    mlRelId = lRelId
    
    txtSurname = sSurname
    
    Me.Caption = sCap & sRelName
        
    Me.Show vbModal
    
    invoke = mlNewId
    
End Function

Private Sub cmdCancel_Click()
    mlNewId = 0
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If SaveNewPerson Then
        Unload Me
    End If
End Sub

Private Function SaveNewPerson() As Boolean
'This function adds a new person onto the database
Dim RS As ADODB.Recordset
Dim SQL As String
Dim sGender As String

Dim sErr As String

    On Error GoTo ErrSub

    If ValidDetails Then
        If optGender(1).Value = True Then
            sGender = "F"
        Else
            sGender = "M"
        End If

        SQL = "Insert into " & gtcINDIVIDUALS & " (" & _
                    gccINDSURNAME & ", " & _
                    gccINDFIRSTNAMES & ", " & _
                    gccINDGENDER & ", " & _
                    gccINDDOBTEXT & ", " & _
                    gccINDDOBDATE & _
                ") Values ('" & _
                    Trim(txtSurname) & "', '" & _
                    Trim(txtFirstNames) & "', '" & _
                    sGender & "', '" & _
                    Trim(txtDOB) & "', " & _
                    Val(txtDOB.Tag) & ")"
        
        gApp.cn.Execute SQL
        
        SQL = "SELECT max(" & gccINDID & ") as NewId FROM " & gtcINDIVIDUALS
        
        Set RS = New ADODB.Recordset
        
        RS.Open SQL, gApp.cn, adOpenForwardOnly, adLockReadOnly
        
        If Not RS.BOF And Not RS.EOF Then
            mlNewId = RS("NewId")
        End If
        SaveNewPerson = True
    Else
        SaveNewPerson = False
    End If

Exit Function
ErrSub:
    sErr = Err.Number & vbCrLf & Err.Description & vbCrLf & vbCrLf & _
            "In Module " & Me.Name & vbCrLf & _
            "In Function SaveNewPerson"

    Call Showerror(sErr, 0)

End Function

Private Function ValidDetails() As Boolean
Dim sMess As String
Dim sDate As String
Dim lDateNum As Long
    
    If Trim(txtSurname) = "" Then
        sMess = sMess & "You must enter a surname." & vbCrLf
    End If
    
    sDate = Trim(txtDOB)
    If sDate = "" Then
        sMess = sMess & "You must indicate an approx date of birth." & vbCrLf
    Else
        lDateNum = ValidDate(sDate)
        If lDateNum = 0 Then
            sMess = sMess & "The date of birth is not recognised as a valid date format." & vbCrLf
        Else
            txtDOB = sDate
            txtDOB.Tag = lDateNum
        End If
    End If

    If sMess = "" Then
        ValidDetails = True
    Else
        MsgBox "You cannot save this data because of the following errors..." & vbCrLf & vbCrLf & sMess, vbOKOnly Or vbCritical, Me.Caption
    End If

End Function

