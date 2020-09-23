VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPicViewer 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   8625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   Icon            =   "frmPicViewer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8625
   ScaleWidth      =   11910
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog comDlg 
      Left            =   60
      Top             =   4470
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Select Printer"
   End
   Begin VB.HScrollBar HScroll 
      Height          =   255
      LargeChange     =   150
      Left            =   240
      SmallChange     =   30
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   8370
      Visible         =   0   'False
      Width           =   11385
   End
   Begin VB.VScrollBar VScroll 
      Height          =   7065
      LargeChange     =   150
      Left            =   11640
      SmallChange     =   30
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1290
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picViewer 
      AutoRedraw      =   -1  'True
      Height          =   7065
      Left            =   240
      ScaleHeight     =   7005
      ScaleWidth      =   11325
      TabIndex        =   2
      Top             =   1290
      Width           =   11385
      Begin VB.Image imgHotSpot 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   945
         Left            =   1530
         Top             =   390
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.Image imgFrame 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   645
         Index           =   0
         Left            =   0
         Top             =   0
         Visible         =   0   'False
         Width           =   615
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   30
      Top             =   6570
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPicViewer.frx":058A
            Key             =   "frames"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPicViewer.frx":0E64
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPicViewer.frx":13A6
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPicViewer.frx":18E8
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPicViewer.frx":21C2
            Key             =   "zoomin"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPicViewer.frx":24DC
            Key             =   "zoomout"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPicViewer.frx":27F6
            Key             =   "prev"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPicViewer.frx":2C48
            Key             =   "next"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPicViewer.frx":309A
            Key             =   "slideshow"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11910
      _ExtentX        =   21008
      _ExtentY        =   1058
      ButtonWidth     =   1879
      ButtonHeight    =   953
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exit"
            Key             =   "Exit"
            Object.ToolTipText     =   "Show Frames"
            ImageKey        =   "Exit"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Show Frames"
            Key             =   "Frames"
            ImageKey        =   "frames"
            Style           =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Save"
            Key             =   "Save"
            ImageKey        =   "Save"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Print"
            Key             =   "Print"
            ImageKey        =   "Print"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Zoom In"
            Key             =   "ZoomIn"
            ImageKey        =   "zoomin"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Zoom Out"
            Key             =   "ZoomOut"
            ImageKey        =   "zoomout"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Slide Show"
            Key             =   "SlideShow"
            ImageKey        =   "slideshow"
            Style           =   1
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Prev Slide"
            Key             =   "Prev"
            ImageKey        =   "prev"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Next Slide"
            Key             =   "Next"
            ImageKey        =   "next"
         EndProperty
      EndProperty
      Begin VB.TextBox txtZoom 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   9630
         TabIndex        =   5
         Text            =   "100%"
         Top             =   150
         Width           =   975
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   8490
      Top             =   0
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      Caption         =   "Caption"
      BeginProperty Font 
         Name            =   "Baskerville Old Face"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   0
      TabIndex        =   0
      Top             =   630
      Width           =   11865
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "Popup Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuPopInfo 
         Caption         =   "Show Information"
      End
      Begin VB.Menu mnuPopRemove 
         Caption         =   "Remove Frame"
      End
      Begin VB.Menu mnuPopSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopCancel 
         Caption         =   "Cancel"
      End
   End
End
Attribute VB_Name = "frmPicViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private X, Y As Single
Private msngWidth, msngHeight As Single
Private msngFactor As Single
Private objPic As IPicture

Private X1 As Long
Private X2 As Long
Private Y1 As Long
Private Y2 As Long

Private miIndex As Integer
Private mbEnd As Boolean
Private mbPrev As Boolean

Private mRS As ADODB.Recordset

Private CPhoto As colPhotoNames
Private mlngImgId As Long           'Need this to save the captions

Public Function invoke(lngImgId As Long, sPath As String, sCaption As String, sName As String)
'Entry point for the form.
Dim Factor As Single

    Call ShowPicture(lngImgId, sPath, sCaption, sName)
    
    Me.Show vbModal

End Function

Private Sub ShowHideFrames(bShow As Boolean)
'This function reads through the collection and displays an image with a fixed border
'for each element in the array
Dim i As Integer
Dim sngX As Single
Dim sngY As Single

    On Error GoTo ErrSub

    'Because the picture is always centred in the picturebox we need to get the
    'actual top and left X values of the actual picture.
    sngX = (picViewer.Width - (msngWidth * msngFactor)) / 2
    sngY = (picViewer.Height - (msngHeight * msngFactor)) / 2

    'unload all images and make index 0 invisible
    imgFrame(0).Visible = False
    On Error Resume Next
    For i = 1 To imgFrame.Count - 1
        Unload imgFrame(i)
    Next i
    On Error GoTo ErrSub
    
    'If the parameter is set to show then reload and redraw all the images for
    'each item in the collection
    If bShow = True Then
        For i = 0 To CPhoto.Count - 1
            If i > 0 Then
                Load imgFrame(i)
            End If
            'Left and top are the recorded X and Y positions * the scaling factor
            'plus the offset of the picture in the frame and then minus the value
            'to where the picture is scrolled (if it is scrolled at all).
            imgFrame(i).Left = (CPhoto(i + 1).X1 * msngFactor) + sngX - (HScroll.Value * 15)
            imgFrame(i).Top = (CPhoto(i + 1).Y1 * msngFactor) + sngY - (VScroll.Value * 15)
            imgFrame(i).Width = (CPhoto(i + 1).X2 - CPhoto(i + 1).X1) * msngFactor
            imgFrame(i).Height = (CPhoto(i + 1).Y2 - CPhoto(i + 1).Y1) * msngFactor
            imgFrame(i).Visible = True
            imgFrame(i).ToolTipText = CPhoto(i + 1).Note
        Next i
    End If
    
Exit Sub
ErrSub:


End Sub

Private Sub ZoomIn()
Dim sngZoom As Single
    
'Hide all the image frames
    Call ShowHideFrames(False)

    sngZoom = Val(txtZoom)
        
'Set the appropriate zoom level based on the current setting
    If sngZoom <= 5 Then
        sngZoom = 10
    ElseIf sngZoom <= 10 Then sngZoom = 20
    ElseIf sngZoom <= 20 Then sngZoom = 25
    ElseIf sngZoom <= 25 Then sngZoom = 50
    ElseIf sngZoom <= 50 Then sngZoom = 75
    ElseIf sngZoom <= 75 Then sngZoom = 100
    ElseIf sngZoom <= 100 Then sngZoom = 125
    ElseIf sngZoom <= 125 Then sngZoom = 150
    ElseIf sngZoom <= 150 Then sngZoom = 175
    ElseIf sngZoom <= 175 Then sngZoom = 200
    ElseIf sngZoom <= 200 Then sngZoom = 250
    ElseIf sngZoom <= 250 Then sngZoom = 300
    ElseIf sngZoom <= 300 Then sngZoom = 350
    ElseIf sngZoom <= 350 Then sngZoom = 400
    End If
    
'set the scaling factor to the new zoom level
    msngFactor = sngZoom / 100
    txtZoom = Format(Int(msngFactor * 100), "###") & " %"
'blank out the picture box
    picViewer.Picture = LoadPicture("")
'let the timer take care of drawing the picture
    DrawPicture
        
End Sub

Private Sub ZoomOut()
Dim sngZoom As Single

'Hide all the frames
    Call ShowHideFrames(False)

    sngZoom = Val(txtZoom)
    
'Set the appropriate zoom level based on the current setting
    If sngZoom >= 400 Then
        sngZoom = 350
    ElseIf sngZoom >= 350 Then sngZoom = 300
    ElseIf sngZoom >= 300 Then sngZoom = 250
    ElseIf sngZoom >= 250 Then sngZoom = 200
    ElseIf sngZoom >= 200 Then sngZoom = 150
    ElseIf sngZoom >= 150 Then sngZoom = 125
    ElseIf sngZoom >= 125 Then sngZoom = 100
    ElseIf sngZoom >= 100 Then sngZoom = 75
    ElseIf sngZoom >= 75 Then sngZoom = 50
    ElseIf sngZoom >= 50 Then sngZoom = 25
    ElseIf sngZoom >= 25 Then sngZoom = 10
    ElseIf sngZoom >= 10 Then sngZoom = 5
    End If
    
'set the scaling factor to the new zoom level
    msngFactor = sngZoom / 100
    txtZoom = Format(Int(msngFactor * 100), "###") & " %"
'blank out the picture box
    picViewer.Picture = LoadPicture("")
'let the timer take care of drawing the picture
    DrawPicture
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'Don't allow quitting if the data hasn't been saved - could be more user friendly!
    If Toolbar1.Buttons("Save").Enabled Then Cancel = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    mRS.Close
    Set mRS = Nothing
End Sub

Private Sub HScroll_Change()
'Hide the frames
    Call ShowHideFrames(False)
'Blank out the picture
    picViewer.Picture = LoadPicture("")
'Let the time take care of redrawing the picture
    Timer1.Enabled = True
End Sub

Private Sub imgFrame_Click(Index As Integer)
'This sub opens up the caption form for adding/changing the caption
'This can only be done if the frames are visible
Dim sName As String
Dim lngId As Long

    lngId = CPhoto.Item(Index + 1).IndId
    sName = frmGetCaption.invoke(lngId, CPhoto.Item(Index + 1).Note)
    If sName <> "" Then
        CPhoto.Item(Index + 1).Note = sName
        CPhoto.Item(Index + 1).IndId = lngId
        Toolbar1.Buttons("Save").Enabled = True
    End If
    
End Sub

Private Sub imgFrame_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'Open up the pop-up menu if right clicking on an image
    If Button = vbRightButton Then
        miIndex = Index
        PopupMenu mnuPopUp
    End If
End Sub

Private Sub mnuPopInfo_Click()
    imgFrame_Click (miIndex)
End Sub

Private Sub mnuPopRemove_Click()
'Remove a frame.
'Firstly Remove the reference from the collection
    CPhoto.Remove (miIndex + 1)
'Redisplay the frames
    ShowHideFrames (False)
    ShowHideFrames (True)
    Toolbar1.Buttons("Save").Enabled = True
End Sub

Private Sub picViewer_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'This is used when drawing new frames on the picturebox.  It needs to be done with the
'shift key pressed.
'A separate hotspot image is used.
    If Shift = 1 Then
        Y1 = Y
        X1 = X
        imgHotSpot.Left = X1
        imgHotSpot.Top = Y1
        imgHotSpot.Width = 0
        imgHotSpot.Height = 0
        imgHotSpot.Visible = True
    End If
End Sub

Private Sub picViewer_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'This is used to draw the hotspot image (if shift is pressed) or
'to show the tooltip text when pointing at people
Dim i As Integer
Dim bfound As Boolean
Dim sngX As Single
Dim sngY As Single

'Set two variables which represent the left and top of the actual image
    sngX = (picViewer.Width - (msngWidth * msngFactor)) / 2
    sngY = (picViewer.Height - (msngHeight * msngFactor)) / 2

    If Shift = 1 Then
        'Stretch out the hotspot image
        If Button = 1 Then
            If X > X1 Then imgHotSpot.Width = X - X1
            If Y > Y1 Then imgHotSpot.Height = Y - Y1
        End If
    Else
        'set the tooltip text for the appropriate hotspot.
        With CPhoto
            For i = 1 To .Count
                If X >= (.Item(i).X1 * msngFactor) + sngX - (HScroll.Value * 15) And X <= (.Item(i).X2 * msngFactor) + sngX - (HScroll.Value * 15) And Y >= (.Item(i).Y1 * msngFactor) + sngY - (VScroll.Value * 15) And Y <= (.Item(i).Y2 * msngFactor) + sngY - (VScroll.Value * 15) Then
                    picViewer.ToolTipText = .Item(i).Note
                    bfound = True
                    Exit For
                End If
            Next i
        End With
        If Not bfound Then
            picViewer.ToolTipText = ""
        End If
    End If
    
End Sub

Private Sub picViewer_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'If the shift key is pressed this adds the hotspot to the collection
Dim sName As String
Dim lngIndId As Long
Dim sngX As Single
Dim sngY As Single
Dim i As Integer

'Set two variables which represent the left and top of the actual image
    sngX = (picViewer.Width - (msngWidth * msngFactor)) / 2
    sngY = (picViewer.Height - (msngHeight * msngFactor)) / 2

    If Shift = 1 Then
        imgHotSpot.Visible = False
        X2 = X
        Y2 = Y
        'Invoke the form to get the name caption
        sName = frmGetCaption.invoke(lngIndId)
        If sName <> "" Then
            'Add the new hotspot to the collection
            Call CPhoto.Add(lngIndId, (X1 - sngX) / msngFactor, (Y1 - sngY) / msngFactor, (X2 - sngX) / msngFactor, (Y2 - sngY) / msngFactor, sName)
            If Toolbar1.Buttons("Frames").Value = vbChecked Then
                If imgFrame.Count > 0 Then
                    i = imgFrame.Count
                    Load imgFrame(i)
                Else
                    i = 0
                End If
                imgFrame(i).Left = (CPhoto(i + 1).X1 * msngFactor) + sngX
                imgFrame(i).Top = (CPhoto(i + 1).Y1 * msngFactor) + sngY
                imgFrame(i).Width = (CPhoto(i + 1).X2 - CPhoto(i + 1).X1) * msngFactor
                imgFrame(i).Height = (CPhoto(i + 1).Y2 - CPhoto(i + 1).Y1) * msngFactor
                imgFrame(i).Visible = True
                imgFrame(i).ToolTipText = CPhoto(i + 1).Note
            End If
            Toolbar1.Buttons("Save").Enabled = True
        End If
    End If
End Sub

Private Sub Timer1_Timer()
'This does all the work of drawing the image.
'It cant be done before the form is made visible - hence this technique
    Timer1.Interval = Val(GetOption(3)) * 1000

'Paint the picture on the picturebox
    If Toolbar1.Buttons("SlideShow").Value = tbrPressed Then
        Timer1.Enabled = True
        NextPicture
    Else
        Timer1.Enabled = False
        DrawPicture
    End If
    
End Sub

Private Function LoadHotSpots() As Boolean
'This loads the hotspots for the image from the database into the collection
Dim RS As ADODB.Recordset
Dim SQL As String
Dim sErr As String

    On Error GoTo ErrSub:
    
    Set CPhoto = Nothing
    Set CPhoto = New colPhotoNames

    SQL = "Select * FROM " & gtcHOTSPOTS & " WHERE " & _
            gccHSPIMGID & " = " & mlngImgId

    Set RS = New ADODB.Recordset
    
    RS.Open SQL, gApp.cn, adOpenForwardOnly, adLockReadOnly
    
    Do While Not RS.EOF
        Call CPhoto.Add(RS(gccHSPINDID), RS(gccHSPX1), RS(gccHSPY1), RS(gccHSPX2), RS(gccHSPY2), RS(gccHSPNOTE))
        RS.MoveNext
    Loop
    

Exit Function
ErrSub:
    sErr = Err.Number & vbCrLf & Err.Description & vbCrLf & vbCrLf & _
            "In Module " & Me.Name & vbCrLf & _
            "In Sub cmdSave_Click"
            
    Call Showerror(sErr, 0)

End Function

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    Select Case UCase(Button.Key)
        Case "EXIT"
            If Toolbar1.Buttons("Save").Enabled Then
                If MsgBox("You have changed/added some captions - save these changes?", vbYesNo Or vbQuestion, Me.Caption) = vbYes Then
                    SaveData
                Else
                    Toolbar1.Buttons("Save").Enabled = False
                End If
            End If
            Set CPhoto = Nothing
            Unload Me
        Case "FRAMES"
            Call ShowHideFrames(Toolbar1.Buttons("Frames").Value)
        Case "SAVE"
            SaveData
        Case "PRINT"
            PrintPicture
        Case "ZOOMIN"
            Call ZoomIn
        Case "ZOOMOUT"
            Call ZoomOut
        Case "SLIDESHOW"
            mbPrev = False
            If Toolbar1.Buttons("SlideShow").Value = tbrPressed Then
                Toolbar1.Buttons("SlideShow").Caption = "End Show"
                mbEnd = False
                Call SlideShow
            Else
                Toolbar1.Buttons("SlideShow").Caption = "Slide Show"
                mbEnd = True
            End If
        Case "PREV"
            mbPrev = True
            NextPicture
        Case "NEXT"
            mbPrev = False
            NextPicture
    End Select
End Sub

Private Function SaveData() As Boolean
'Save the data from the collection back to the database
Dim RS As ADODB.Recordset
Dim SQL As String
Dim sErr As String
Dim i As Integer

    On Error GoTo ErrSub

'Firstly delete the hotspot info
    SQL = "Delete FROM " & gtcHOTSPOTS & " WHERE " & _
            gccHSPIMGID & " = " & mlngImgId
    
    gApp.cn.Execute SQL

'Now re-insert it from the collection
    With CPhoto
        For i = 1 To CPhoto.Count
            SQL = "Insert into " & gtcHOTSPOTS & " ( " & _
                gccHSPIMGID & ", " & _
                gccHSPSEQ & ", " & _
                gccHSPINDID & ", " & _
                gccHSPX1 & ", " & _
                gccHSPX2 & ", " & _
                gccHSPY1 & ", " & _
                gccHSPY2 & ", " & _
                gccHSPNOTE & ") VALUES (" & _
                mlngImgId & ", " & _
                i & ", " & _
                .Item(i).IndId & ", " & _
                .Item(i).X1 & ", " & _
                .Item(i).X2 & ", " & _
                .Item(i).Y1 & ", " & _
                .Item(i).Y2 & ", '" & _
                .Item(i).Note & "')"
    
            gApp.cn.Execute SQL
            
            If .Item(i).IndId <> 0 Then
                SQL = "Select * from " & gtcIMAGELINK & " WHERE " & _
                    gccIMLIMGID & " = " & mlngImgId & " AND " & _
                    gccIMLINDID & " = " & .Item(i).IndId
                
                Set RS = New ADODB.Recordset
                
                RS.Open SQL, gApp.cn, adOpenKeyset, adLockOptimistic
                
                If RS.BOF And RS.EOF Then
                    RS.AddNew
                    RS(gccIMLIMGID) = mlngImgId
                    RS(gccIMLINDID) = .Item(i).IndId
                    RS.Update
                    RS.Close
                End If
            End If
        Next i
    End With
    Toolbar1.Buttons("Save").Enabled = False
    
Exit Function
ErrSub:
    sErr = Err.Number & vbCrLf & Err.Description & vbCrLf & vbCrLf & _
            "In Module " & Me.Name & vbCrLf & _
            "In Sub SaveData"
            
    Call Showerror(sErr, 0)

End Function

Private Sub VScroll_Change()
'Hide all the frames
    Call ShowHideFrames(False)
'Blank out the picture
    picViewer.Picture = LoadPicture("")
'Let the timer take care of redrawing the picture
    Timer1.Enabled = True
End Sub

Private Function PrintPicture() As Boolean
'Print the picture using standard windows print technique.
'FIXIT - Need to add the captions overlay on a separate page or
'further down the same page if both will fit.
Dim sngX As Single
Dim sngY As Single
Dim X As Printer
Dim sErr As String

    On Error GoTo ErrSub
    
    comDlg.ShowPrinter

    Printer.Orientation = comDlg.Orientation

    sngX = (Printer.Width - (msngWidth * msngFactor)) / 2
    sngY = (Printer.Height - (msngHeight * msngFactor)) / 2
    
    Printer.PaintPicture objPic, sngX, sngY, msngWidth * msngFactor, msngHeight * msngFactor
    Printer.EndDoc

Exit Function
ErrSub:
    sErr = Err.Number & vbCrLf & Err.Description & vbCrLf & vbCrLf & _
            "In Module " & Me.Name & vbCrLf & _
            "In Function PrintPicture"
            
    Call Showerror(sErr, 0)

End Function

Private Function ShowPicture(lngImgId As Long, sPath As String, sCaption As String, sName As String) As Boolean
Dim sErr As String

    On Error GoTo ErrSub

    'make the image id accessible by all other functions
    mlngImgId = lngImgId
    Me.Caption = sName

    'load the picture
    Set objPic = Nothing
    Set objPic = LoadPicture(sPath)
    
    'Don't know why I used 26.4577 but it scales it correctly
    msngWidth = (objPic.Width / 26.4577) * 15   'This is now width in twips
    msngHeight = (objPic.Height / 26.4577) * 15 'This is now height in twips
        
    msngFactor = 1
    
    'Now see if the picture will fit in the picturebox
    If msngWidth > picViewer.Width Then
        msngFactor = picViewer.Width / msngWidth
    End If

    If msngHeight * msngFactor > picViewer.Height Then
        msngFactor = picViewer.Height / msngHeight
    End If
    
    'Set the zoom percentage
    txtZoom = Format(Int(msngFactor * 100), "###") & " %"
        
    lblCaption.Caption = sCaption
    picViewer.Picture = LoadPicture()
    LoadHotSpots
    
Exit Function
ErrSub:
    sErr = Err.Number & vbCrLf & Err.Description & vbCrLf & vbCrLf & _
            "In Module " & Me.Name & vbCrLf & _
            "In Function PrintPicture"
            
    Call Showerror(sErr, 0)

End Function


Private Function SlideShow() As Boolean
Dim SQL As String
Dim sErr As String

    On Error GoTo ErrSub
    
    If mRS Is Nothing Then
        Set mRS = New ADODB.Recordset
    End If
    
    If mRS.State = 0 Then
        SQL = "Select * from " & gtcIMAGES & " ORDER BY " & _
                gccIMGDATEDATE
                
        
        mRS.Open SQL, gApp.cn, adOpenDynamic, adLockOptimistic
    End If
    
    Toolbar1.Buttons("Next").Enabled = True
    Toolbar1.Buttons("Prev").Enabled = True
    
    If Not mRS.EOF And Not mRS.BOF Then
        Call ShowPicture(mRS(gccIMGID), App.Path & "\" & mRS(gccIMGNAME), mRS(gccIMGCAPTION), mRS(gccIMGCAPTION))
        DoEvents
        DrawPicture
        Timer1.Interval = Val(GetOption(3)) * 1000
        Timer1.Enabled = True
    End If
    
Exit Function
ErrSub:
    sErr = Err.Number & vbCrLf & Err.Description & vbCrLf & vbCrLf & _
            "In Module " & Me.Name & vbCrLf & _
            "In Function PrintPicture"
            
    Call Showerror(sErr, 0)

End Function

Private Function NextPicture() As Boolean
    If mbPrev Then
        If Not mRS.BOF Then
            mRS.MovePrevious
            If Not mRS.BOF Then
                Call ShowPicture(mRS(gccIMGID), App.Path & "\" & mRS(gccIMGNAME), mRS(gccIMGCAPTION), mRS(gccIMGCAPTION))
                DoEvents
            Else
                mRS.Close
                Toolbar1.Buttons("SlideShow").Value = tbrUnpressed
                Toolbar1.Buttons("Prev").Enabled = False
                Toolbar1.Buttons("Next").Enabled = False
            End If
        Else
            mRS.Close
            Toolbar1.Buttons("SlideShow").Value = tbrUnpressed
            Toolbar1.Buttons("Prev").Enabled = False
            Toolbar1.Buttons("Next").Enabled = False
        End If
    Else
        If Not mRS.EOF Then
            mRS.MoveNext
            If Not mRS.EOF Then
                Call ShowPicture(mRS(gccIMGID), App.Path & "\" & mRS(gccIMGNAME), mRS(gccIMGCAPTION), mRS(gccIMGCAPTION))
                DoEvents
            Else
                mRS.Close
                Toolbar1.Buttons("SlideShow").Value = tbrUnpressed
                Toolbar1.Buttons("Prev").Enabled = False
                Toolbar1.Buttons("Next").Enabled = False
            End If
        Else
            mRS.Close
            Toolbar1.Buttons("SlideShow").Value = tbrUnpressed
            Toolbar1.Buttons("Prev").Enabled = False
            Toolbar1.Buttons("Next").Enabled = False
        End If
    End If

    DrawPicture

End Function

Private Function DrawPicture()
Dim sngX As Single
Dim sngY As Single
Static sngHFactor As Single
Static sngVFactor As Single

'Set two variables which represent the left and top of the actual image
    sngX = (picViewer.Width - (msngWidth * msngFactor)) / 2
    sngY = (picViewer.Height - (msngHeight * msngFactor)) / 2
    
'Make the Horizontal and Vertical scrollbars visible if appropriate
    If sngX < 1 Then
        HScroll.Visible = True
        HScroll.Max = ((msngWidth * msngFactor) - picViewer.Width) / 30 'Make it pixels
        HScroll.Min = HScroll.Max * -1
    Else
        HScroll.Visible = False
        HScroll.Value = 0
        HScroll.Max = 0
    End If
    If sngY < 1 Then
        VScroll.Visible = True
        VScroll.Max = ((msngHeight * msngFactor) - picViewer.Height) / 30  'Make it pixels
        VScroll.Min = VScroll.Max * -1
    Else
        VScroll.Visible = False
        VScroll.Value = 0
        VScroll.Max = 0
    End If
        
    picViewer.PaintPicture objPic, sngX - (HScroll.Value * 15), sngY - (VScroll.Value * 15), msngWidth * msngFactor, msngHeight * msngFactor
    DoEvents
'Switch the timer off
'Show or hide the frames depending on the state of the toolbar frames button
    Call ShowHideFrames(Toolbar1.Buttons("Frames").Value)
    
End Function
