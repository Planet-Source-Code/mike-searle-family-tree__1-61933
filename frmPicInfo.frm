VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmPicInfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Picture Properties for:"
   ClientHeight    =   6225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6375
   Icon            =   "frmPicInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   6375
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picGallery 
      AutoRedraw      =   -1  'True
      Height          =   1305
      Left            =   4890
      ScaleHeight     =   1245
      ScaleMode       =   0  'User
      ScaleWidth      =   1425
      TabIndex        =   9
      Top             =   780
      Width           =   1425
   End
   Begin VB.TextBox txtFileName 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   1050
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   60
      Width           =   5265
   End
   Begin VB.TextBox txtDate 
      Height          =   285
      Left            =   1050
      TabIndex        =   1
      Top             =   720
      Width           =   1845
   End
   Begin VB.TextBox txtCaption 
      Height          =   285
      Left            =   1050
      TabIndex        =   0
      Top             =   390
      Width           =   5265
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   345
      Left            =   5460
      TabIndex        =   4
      Top             =   5850
      Width           =   885
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Ok"
      Height          =   345
      Left            =   4500
      TabIndex        =   3
      Top             =   5850
      Width           =   885
   End
   Begin RichTextLib.RichTextBox rtbNotes 
      Height          =   3615
      Left            =   30
      TabIndex        =   2
      Top             =   2190
      Width           =   6315
      _ExtentX        =   11139
      _ExtentY        =   6376
      _Version        =   393217
      TextRTF         =   $"frmPicInfo.frx":058A
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "FileName:"
      Height          =   225
      Left            =   150
      TabIndex        =   8
      Top             =   90
      Width           =   765
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Date:"
      Height          =   195
      Left            =   150
      TabIndex        =   6
      Top             =   780
      Width           =   765
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Caption:"
      Height          =   225
      Left            =   150
      TabIndex        =   5
      Top             =   420
      Width           =   765
   End
End
Attribute VB_Name = "frmPicInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlImgId As Long
Private mbChanged As Boolean
Private mbOk As Boolean

Public Function invoke(lImgId As Long) As Boolean
    mlImgId = lImgId
    If GetPicInfo Then
        Me.Show vbModal
        invoke = mbOk
    End If
End Function

Private Function GetPicInfo() As Boolean
Dim RS As ADODB.Recordset
Dim SQL As String
Dim sErr As String
Dim X, Y As Single
Dim lWidth, lHeight As Long
Dim objPic As IPictureDisp
Dim Factor As Single

    On Error GoTo ErrSub:
    
    SQL = "SELECT * FROM " & gtcIMAGES & " WHERE " & gccIMGID & " = " & mlImgId

    Set RS = New ADODB.Recordset
    
    RS.Open SQL, gApp.cn, adOpenForwardOnly, adLockReadOnly
    
    If Not RS.BOF And Not RS.EOF Then
        Me.Caption = "Picture Properties for: " & RS(gccIMGNAME)
        txtFileName = App.Path & "\" & RS(gccIMGNAME)
        txtCaption = RS(gccIMGCAPTION)
        txtDate = RS(gccIMGDATETEXT)
        rtbNotes = RS(gccIMGNOTES)
        
        Set objPic = LoadPicture(txtFileName)
        lWidth = Int(objPic.Width)
        lHeight = Int(objPic.Height)
        
        If lHeight > lWidth Then
            Factor = picGallery.ScaleHeight / lHeight
            lHeight = picGallery.ScaleHeight
            lWidth = lWidth * Factor
            X = Int((picGallery.ScaleWidth - lWidth) / 2)
        Else
            Factor = picGallery.ScaleWidth / lWidth
            lWidth = picGallery.ScaleWidth
            lHeight = lHeight * Factor
            Y = Int((picGallery.ScaleHeight - lHeight) / 2)
        End If
        
        picGallery.Picture = LoadPicture()
        picGallery.PaintPicture objPic, X, Y, lWidth, lHeight
        DoEvents
        
        GetPicInfo = True
    End If
    RS.Close

Exit Function
ErrSub:
    sErr = Err.Number & vbCrLf & Err.Description & vbCrLf & vbCrLf & _
            "In Module " & Me.Name & vbCrLf & _
            "In Function GetPicInfo"
            
    Call Showerror(sErr, 0)

End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If SaveDetails Then
        mbOk = True
        Unload Me
    End If
End Sub

Private Function SaveDetails() As Boolean
Dim RS As ADODB.Recordset
Dim SQL As String
Dim sErr As String

    On Error GoTo ErrSub

    If ValidDetails Then

        SQL = "SELECT * FROM " & gtcIMAGES & " WHERE " & gccIMGID & " = " & mlImgId
    
        Set RS = New ADODB.Recordset
        
        RS.Open SQL, gApp.cn, adOpenKeyset, adLockOptimistic
        
        If Not RS.BOF And Not RS.EOF Then
            RS(gccIMGCAPTION) = txtCaption
            RS(gccIMGDATETEXT) = txtDate
            RS(gccIMGDATEDATE) = Val(txtDate.Tag)
            RS(gccIMGNOTES) = rtbNotes.TextRTF
            RS.Update
            SaveDetails = True
        End If
        RS.Close
    End If

Exit Function
ErrSub:
    sErr = Err.Number & vbCrLf & Err.Description & vbCrLf & vbCrLf & _
            "In Module " & Me.Name & vbCrLf & _
            "In Function SaveDetails"
            
    Call Showerror(sErr, 0)
End Function

Private Function ValidDetails() As Boolean
Dim sMess As String
Dim idx As Integer
Dim sDate As String
Dim lDateNum As Long

    sDate = Trim(txtDate)
    If sDate = "" Then
        sMess = sMess & "You must indicate an approx date." & vbCrLf
    Else
        lDateNum = ValidDate(sDate)
        If lDateNum = 0 Then
            sMess = sMess & "The date is not recognised as a valid date format." & vbCrLf
        Else
            txtDate = sDate
            txtDate.Tag = lDateNum
        End If
    End If

    If sMess = "" Then
        ValidDetails = True
    Else
        MsgBox "You cannot save this data because of the following errors..." & vbCrLf & vbCrLf & sMess, vbOKOnly Or vbCritical, Me.Caption
    End If


End Function

