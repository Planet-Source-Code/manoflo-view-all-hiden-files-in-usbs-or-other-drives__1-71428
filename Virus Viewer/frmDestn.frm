VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmDestn 
   Caption         =   "Folder Browser"
   ClientHeight    =   5670
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6735
   Icon            =   "frmDestn.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5670
   ScaleWidth      =   6735
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picTemp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   840
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   2
      Top             =   7680
      Visible         =   0   'False
      Width           =   375
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   5295
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ListView lstViewer 
      Height          =   5295
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   9340
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      OLEDragMode     =   1
      _Version        =   327682
      Icons           =   "imgBig"
      SmallIcons      =   "imgBig"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      OLEDragMode     =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Name"
         Object.Width           =   6068
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Size"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Type"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Date Modified"
         Object.Width           =   3528
      EndProperty
   End
   Begin ComctlLib.ImageList imgBig 
      Left            =   3120
      Top             =   7560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
   End
End
Attribute VB_Name = "frmDestn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FSO As New Scripting.FileSystemObject
Dim mDrive As Drive
Dim mFolder As Folder
Dim mFile As File
Dim lITEM As ListItem

Private Sub ViewDrives()
Dim DrvLet As String
    For Each mDrive In FSO.Drives
        If mDrive.IsReady = True Then
            DrvLet = mDrive.VolumeName
            If Len(Trim$(DrvLet)) = 0 Then
                DrvLet = "Local Disk"
            End If
            With lstViewer.ListItems
                Set lITEM = .Add(, , DrvLet & "(" & mDrive.Path & ")", , ExtractIcon(mDrive.RootFolder, imgBig, picTemp, 16))
            End With
        End If
    Next
End Sub

Private Sub ViewFolders()
    For Each mDrive In FSO.Drives
        If mDrive.IsReady = True Then
            For Each mFolder In mDrive.RootFolder.SubFolders
                lstViewer.AddItem mFolder.Path
                lstViewer.ZOrder
            Next
        End If
    Next
End Sub

Private Sub Form_Load()
    Call ViewDrives
End Sub

Private Sub Form_Resize()
    With lstViewer
        .Left = 0
        .Top = 0
        .Width = frmDestn.Width
        .Height = frmDestn.Height
    End With
End Sub

