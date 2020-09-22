VERSION 5.00
Begin VB.Form fTest 
   Caption         =   "ucToolbar 3.1 - Test"
   ClientHeight    =   4740
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   9045
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   316
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   603
   StartUpPosition =   2  'CenterScreen
   Begin Test.ucToolbar ucToolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      Top             =   0
      Width           =   9045
      _ExtentX        =   15954
      _ExtentY        =   741
   End
   Begin VB.TextBox txtEvents 
      Appearance      =   0  'Flat
      Height          =   3015
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   1080
      Width           =   4335
   End
   Begin VB.Menu mnuFileTop 
      Caption         =   "&File"
      Begin VB.Menu mnuFile 
         Caption         =   "E&xit"
         Index           =   0
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuTestTop 
      Caption         =   "&Test"
      Begin VB.Menu mnuTest 
         Caption         =   "&Apply some changes"
         Index           =   0
      End
      Begin VB.Menu mnuTest 
         Caption         =   "&Unapply those changes"
         Index           =   1
      End
   End
   Begin VB.Menu mnuPopupTop 
      Caption         =   "PopupMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuPopup 
         Caption         =   "Item 0"
         Index           =   0
      End
      Begin VB.Menu mnuPopup 
         Caption         =   "Item 1"
         Index           =   1
      End
   End
End
Attribute VB_Name = "fTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    With ucToolbar1
        
        '-- Initialize toolbar
        
        Call .Initialize(ImageSize:=24, _
                         FlatStyle:=True, _
                         ListStyle:=True, _
                         Divider:=True)
                                                  
        Call .AddBitmap(LoadResPicture("TB_BITMAP", vbResBitmap), vbMagenta)
                         
        '-- Add buttons (*)
        
        Call .AddButton("Back", 0, , [eDropDown])
        Call .AddButton("Forward", 1, , [eDropDown])
        Call .AddButton(, 2, "Up")
        Call .AddButton(, , , [eSeparator])
        Call .AddButton(, 3, "Stop")
        Call .AddButton(, 4, "Refresh")
        Call .AddButton(, 5, "Home")
        Call .AddButton(, , , [eSeparator])
        Call .AddButton("Search", 6)
        Call .AddButton("Favourites", 7)
        Call .AddButton(, , , [eSeparator])
        Call .AddButton(, 8, "View", [eWholeDropDown])
        Call .AddButton(, 9, "Full screen", [eCheck])
        
        '-- Adjust height
        
        Let .Height = .ToolbarHeight
    End With
    
' (*) For no ListStyle (caption below image), don't set caption as null string for those
'     buttons you don't want to display caption; set as space char, otherwise, all buttons
'     will suffer same appearance.
End Sub

Private Sub Form_Resize()
    
    On Error Resume Next
  
    ucToolbar1.ButtonChecked(13) = (WindowState = vbMaximized)
    Call txtEvents.Move(0, ucToolbar1.Height, Me.ScaleWidth, Me.ScaleHeight - ucToolbar1.Height)
    
    On Error GoTo 0
End Sub

Private Sub mnuFile_Click(Index As Integer)
    Call Unload(Me)
End Sub

Private Sub mnuTest_Click(Index As Integer)
   
    With ucToolbar1
    
        Select Case Index
            
            Case 0 '-- Apply some changes
                .ButtonCaption(1) = vbNullString
                .ButtonCaption(2) = vbNullString
                .ButtonTipText(1) = "Back"
                .ButtonTipText(2) = "Forward"
                .ButtonEnabled(3) = False
                .ButtonVisible(5) = False
                .ButtonVisible(6) = False
                .ButtonImage(7) = 10
                .ButtonChecked(13) = True:  Call ucToolbar1_ButtonClick(13)
            
            Case 1 '-- Unapply those changes
                .ButtonCaption(1) = "Back"
                .ButtonCaption(2) = "Forward"
                .ButtonTipText(1) = vbNullString
                .ButtonTipText(2) = vbNullString
                .ButtonEnabled(3) = True
                .ButtonVisible(5) = True
                .ButtonVisible(6) = True
                .ButtonImage(7) = 5
                .ButtonChecked(13) = False: Call ucToolbar1_ButtonClick(13)
        End Select
    End With
End Sub





Private Sub ucToolbar1_ButtonClick(ByVal Button As Long)
    pvAddItem "ucToolbar1_ButtonClick " & Button
    
    Select Case Button
        Case 13
            If (ucToolbar1.ButtonChecked(13)) Then
                Me.WindowState = vbMaximized
              Else
                Me.WindowState = vbNormal
            End If
    End Select
End Sub

Private Sub ucToolbar1_ButtonLeave(ByVal Button As Long)
    pvAddItem "ucToolbar1_ButtonLeave " & Button
End Sub

Private Sub ucToolbar1_ButtonEnter(ByVal Button As Long)
    pvAddItem "ucToolbar1_ButtonEnter " & Button
End Sub

Private Sub ucToolbar1_ButtonDropDown(ByVal Button As Long, ByVal x As Long, ByVal y As Long)
    pvAddItem "ucToolbar1_ButtonDropDown " & Button
    
    Call PopupMenu(mnuPopupTop, , x, y)
End Sub

Private Sub ucToolbar1_ToolbarEnter()
    pvAddItem "ucToolbar1_ToolbarEnter"
End Sub

Private Sub ucToolbar1_ToolbarLeave()
    pvAddItem "ucToolbar1_ToolbarLeave"
End Sub





Private Sub pvAddItem(ByVal sItem As String)
    
    txtEvents = txtEvents & sItem & vbCrLf
    txtEvents.SelStart = Len(txtEvents)
End Sub
