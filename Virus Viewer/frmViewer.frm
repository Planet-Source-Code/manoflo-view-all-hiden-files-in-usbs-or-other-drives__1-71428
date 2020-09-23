VERSION 5.00
Begin VB.Form frmViewer 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "USB Virus Browser"
   ClientHeight    =   4080
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5865
   Icon            =   "frmViewer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   5865
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdInvAll 
      Caption         =   "Invert All"
      Height          =   375
      Left            =   2966
      TabIndex        =   11
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CommandButton cmdClearAll 
      Caption         =   "Clear All"
      Height          =   375
      Left            =   1543
      TabIndex        =   10
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CommandButton cmdSelAll 
      Caption         =   "Select All"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   3600
      Width           =   1335
   End
   Begin VB.ListBox lstViewer 
      Height          =   3210
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   6
      Top             =   120
      Width           =   4180
   End
   Begin VB.TextBox txtProp 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   1410
      Left            =   4393
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "Copy File(s)"
      Height          =   375
      Left            =   4390
      TabIndex        =   3
      Top             =   560
      Width           =   1335
   End
   Begin VB.CommandButton cmdDeleteSel 
      Caption         =   "Delete File(s)"
      Height          =   375
      Left            =   4390
      TabIndex        =   2
      Top             =   1000
      Width           =   1335
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   4390
      TabIndex        =   0
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CommandButton cmdCheck 
      Caption         =   "Browse USB"
      Height          =   375
      Left            =   4390
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.PictureBox picCopy 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4695
      Left            =   -120
      ScaleHeight     =   4695
      ScaleWidth      =   6135
      TabIndex        =   7
      Top             =   -480
      Visible         =   0   'False
      Width           =   6135
      Begin VB.Label lblCopySt 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   8
         Top             =   2280
         Width           =   6015
      End
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   4395
      TabIndex        =   4
      Top             =   1485
      Width           =   1335
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000009&
      X1              =   -120
      X2              =   5880
      Y1              =   3495
      Y2              =   3495
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      X1              =   -120
      X2              =   5880
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Menu mv 
      Caption         =   "v"
      Visible         =   0   'False
      Begin VB.Menu mvSFLS 
         Caption         =   "View Sub Files"
      End
      Begin VB.Menu mvSFLD 
         Caption         =   "View Sub Folders"
      End
      Begin VB.Menu mvB0 
         Caption         =   "-"
      End
      Begin VB.Menu mvSALL 
         Caption         =   "View Files And Folders"
      End
   End
End
Attribute VB_Name = "frmViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'HI THERE
'ENJOY WATCHING ALL FILES AND FOLDERS AND ALL HIDDEN FILES OR THREATS IN USBS AND REMOVING THEM BEFORE INFECTING YOUR PC

'STEPS TO FOLLOW:
'1. RUN THE PROGRAM
'2. INSERT YOUR USB
'3. CLICK ON THE BROWSE BUTTON
'4. CLICK ON THE VIEW FILES AND FOLDERS
'5. THE PROGRAM WILL DISPLAY ALL FILES VISIBLE AND THE ONCE HIDDEN
'6. BECAREFULL OF THE FOLLOWING KIND OF FILES
    '6.1 AUTORUN.INF    IT CARRIES VIRUS INFORMATION(WHEN DOUBLE CLICKING YOUR USB THIS GETS ACTIVATED AND RUN THE .EXE ; .CMD OR .BAT FILES TO INFECT YOUR PCs
    '6.2 .CMD; .EXE ; .BAT FILES WHICH YOU NEVER ADDED TO YOUR USB AND ARE NOT VISIBLE WHEN YOU OPEN YOUR USB BUT ONLY VISIBLE WITHIN THIS PROGRAM
'7. DELETE ALL FILES STATED IN 6.1 AND 6.2 AS THEY ARE VIRUSES OR THREATS
'8. BECAREFULL IF YOU HAVE CHANGED THE mDrive.DriveType = Removable AS THIS WILL ENABLE YOU TO VIEW SYSTEM HIDDEN FILES IN OTHER DRIVES AND DELETING THEM CAN RESULT IN SYSTEM FAILURE
'9. VOTE FOR ME(MY E-MAIL IS: manoflo@webmail.co.za)
'10. ENJOY WATCHING VIRUSES DYING UNDER YOUR CLICKS OF THE MOUSE!

Option Explicit
Dim FSO As New Scripting.FileSystemObject 'Declare the referenced scrun.dll
Dim mDrive As Drive
Dim mFolder As Folder
Dim mFile As File
Dim seLi As Long
Dim DestPath As String
Dim Lv As Long
Dim lvC As Long
Dim nL As Long
Dim CopyType As Integer

'This function checks the path if it contains the leading "\", if not then add it
Private Function ConfigPath(ByVal ckP As String) As String
    On Error Resume Next
    If Right$(ckP, 1) <> "\" Then
        ckP = ckP & "\"
    End If
    ConfigPath = ckP
End Function
'The sub below browses files from the USB Mery Stick but you can change the mDrive.DriveType to other drives to view them
Private Sub ViewFiles()
    On Error Resume Next
    For Each mDrive In FSO.Drives
        If mDrive.IsReady = True And mDrive.DriveType = Removable Then
            For Each mFile In mDrive.RootFolder.Files
                DoEvents
                lstViewer.AddItem mFile.Path
            Next
        End If
    Next
End Sub
'The sub below browses folders from the USB Mery Stick but you can change the mDrive.DriveType to other drives to view them
Private Sub ViewFolders()
    On Error Resume Next
    For Each mDrive In FSO.Drives
        If mDrive.IsReady = True And mDrive.DriveType = Removable Then
            For Each mFolder In mDrive.RootFolder.SubFolders
                DoEvents
                lstViewer.AddItem mFolder.Path
            Next
        End If
    Next
End Sub

Private Sub cmdCheck_Click()
    On Error Resume Next
    CopyType = 0
    PopupMenu mv 'calls the popup menu "mv"
End Sub
'The sub will clear all items listed in the list
Private Sub cmdClearAll_Click()
    On Error Resume Next
    For seLi = 0 To lstViewer.ListCount - 1
        lstViewer.ListIndex = seLi
        lstViewer.Selected(seLi) = False
    Next
    lstViewer.ListIndex = 0
End Sub

Private Sub cmdClose_Click()
    On Error Resume Next
    Unload Me 'Un load from the memory
    Set frmViewer = Nothing 'it is important to destroy the form afrter you have uloaded it from the memory
End Sub
'The sub below will copy all selected files in the list to the your selected destination
Private Sub cmdCopy_Click()
    On Error Resume Next
    lvC = 0
    For Lv = 0 To lstViewer.ListCount - 1
       If lstViewer.Selected(Lv) = True Then
            lvC = lvC + 1
       End If
    Next
    
    If lvC > 0 Then
        DestPath = BrowseForFolder("", frmViewer.hwnd, "Please select destination path.") 'BrowseForFolder is a function in the module for browsing the destination folder
        If Len(Trim$(DestPath)) > 0 Then
            Me.Caption = DestPath
            If MsgBox("Are you sure you want copy the selected file(s)", vbYesNo + vbExclamation, App.CompanyName) = vbYes Then
                nL = 0
                For Lv = 0 To lstViewer.ListCount - 1
                    Me.Caption = "Please wait..."
                    If FSO.FileExists(lstViewer.List(Lv)) = True And lstViewer.Selected(Lv) = True Then
                        DoEvents 'The doevents helps to prevent a loop from hanging
                        nL = nL + 1
                        picCopy.ZOrder 'This will make sure that the picCopy is on top of the list
                        picCopy.Visible = True
                        lblCopySt.Caption = Round(nL / lvC * 100, 0) & "% Completed"
                        FSO.CopyFile lstViewer.List(Lv), ConfigPath(DestPath) & FSO.GetFile(lstViewer.List(Lv)).Name, True
                    ElseIf FSO.FolderExists(lstViewer.List(Lv)) = True And lstViewer.Selected(Lv) = True Then
                        DoEvents
                        nL = nL + 1
                        picCopy.ZOrder
                        picCopy.Visible = True
                        lblCopySt.Caption = Round(nL / lvC * 100, 0) & "% Completed"
                        FSO.CopyFolder lstViewer.List(Lv), ConfigPath(DestPath) & FSO.GetFolder(lstViewer.List(Lv)).Name, True
                    End If
                Next
                Me.Caption = "USB Virus Browser"
                picCopy.Visible = False
                nL = 0
            End If
        End If
    Else
        MsgBox "Please tick at least one item to be copied", vbExclamation, App.CompanyName
    End If
End Sub
'This will delete all the selected files from the USB(Exp Hidden files like viruses
Private Sub cmdDeleteSel_Click()
On Error Resume Next
    lvC = 0
    For Lv = 0 To lstViewer.ListCount - 1
       If lstViewer.Selected(Lv) = True Then
            lvC = lvC + 1
       End If
    Next
    If lvC = 0 Then
        MsgBox "Please select at least one item", vbExclamation, App.CompanyName
        Exit Sub
    End If
    If MsgBox("Are you sure you want to delete?", vbYesNo + vbExclamation, App.CompanyName) = vbYes Then
        For seLi = 0 To lstViewer.ListCount - 1
        DoEvents
            If lstViewer.Selected(seLi) = True Then
                lstViewer.ListIndex = seLi
                DoEvents
                If lstViewer.Selected(seLi) = True Then
                    If FSO.FileExists(lstViewer.List(seLi)) = True Then
                        FSO.DeleteFile lstViewer.List(seLi), True
                        lstViewer.RemoveItem seLi
                    ElseIf FSO.FolderExists(lstViewer.List(seLi)) = True Then
                        FSO.DeleteFolder lstViewer.List(seLi), True
                        lstViewer.RemoveItem seLi
                    End If
                End If
            End If
        Next
        
        lstViewer.ListIndex = 0
        If CopyType = 1 Then
            mvSFLS_Click
        ElseIf CopyType = 2 Then
            mvSFLD_Click
        ElseIf CopyType = 3 Then
            mvSALL_Click
        End If
        If lstViewer.ListCount = 0 Then
            txtProp.Text = ""
        End If
    End If
End Sub
'Inverts the selection
Private Sub cmdInvAll_Click()
    On Error Resume Next
    For seLi = 0 To lstViewer.ListCount - 1
        lstViewer.ListIndex = seLi
        lstViewer.Selected(seLi) = Not lstViewer.Selected(seLi)
    Next
    lstViewer.ListIndex = 0
End Sub
'Select all items in the list
Private Sub cmdSelAll_Click()
    On Error Resume Next
    For seLi = 0 To lstViewer.ListCount - 1
        lstViewer.ListIndex = seLi
        lstViewer.Selected(seLi) = True
    Next
    lstViewer.ListIndex = 0
End Sub

Private Sub Form_Load()
    CopyType = 0
End Sub

Private Sub lstViewer_Click()
    On Error Resume Next
    Dim cuItem As String
    cuItem = lstViewer.List(lstViewer.ListIndex)
        If Len(Trim$(cuItem)) > 0 Then
            If FSO.FileExists(cuItem) = True Then
                If Len(FSO.GetFile(cuItem).Size) > 4 Then
                    DoEvents
                    txtProp.Text = "Size: (" & Round(FSO.GetFile(cuItem).Size / 1053888.80952381, 2) & " MB) " & "Date Created: (" & FSO.GetFile(cuItem).DateCreated & ") File Type: (" & FSO.GetFile(cuItem).Type & ")"
                Else
                    DoEvents
                    txtProp.Text = "Size: (" & FSO.GetFile(cuItem).Size & " Bytes) " & "Date Created: (" & FSO.GetFile(cuItem).DateCreated & ") File Type: (" & FSO.GetFile(cuItem).Type & ")"
                End If
            ElseIf FSO.FolderExists(cuItem) = True Then
                If Len(FSO.GetFolder(cuItem).Size) > 4 Then
                    DoEvents
                    txtProp.Text = "Date Created: (" & FSO.GetFolder(cuItem).DateCreated & ")"
                Else
                    DoEvents
                    txtProp.Text = "Date Created: (" & FSO.GetFolder(cuItem).DateCreated & ")"
                End If
            End If
        End If
End Sub


Private Sub mvSALL_Click()
    lstViewer.Clear
    Call ViewFolders
    Call ViewFiles
    lblStatus.Caption = lstViewer.ListCount & " File(s) And Folder(s) Found!"
    CopyType = 3
End Sub

Private Sub mvSFLD_Click()
    lstViewer.Clear
    Call ViewFolders
    lblStatus.Caption = lstViewer.ListCount & " Folder(s) Found!"
    CopyType = 2
End Sub

Private Sub mvSFLS_Click()
    lstViewer.Clear
    Call ViewFiles
    lblStatus.Caption = lstViewer.ListCount & " File(s) Found!"
    CopyType = 1
End Sub
