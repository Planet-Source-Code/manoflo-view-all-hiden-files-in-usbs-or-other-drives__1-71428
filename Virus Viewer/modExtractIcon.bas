Attribute VB_Name = "modExtractIcon"
Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As typSHFILEINFO, ByVal cbSizeFileInfo As Long, ByVal uFlags As Long) As Long
Private Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal himl&, ByVal i&, ByVal hDCDest&, ByVal x&, ByVal y&, ByVal Flags&) As Long
Private Declare Function ShellExecuteEX Lib "shell32.dll" Alias "ShellExecuteEx" (SEI As SHELLEXECUTEINFO) As Long

Private Type typSHFILEINFO
  hIcon As Long
  iIcon As Long
  dwAttributes As Long
  szDisplayName As String * 260
  szTypeName As String * 80
End Type

Private Const SHGFI_DISPLAYNAME = &H200
Private Const SHGFI_EXETYPE = &H2000
Private Const SHGFI_SYSICONINDEX = &H4000
Private Const SHGFI_SHELLICONSIZE = &H4
Private Const SHGFI_TYPENAME = &H400
Private Const SHGFI_LARGEICON = &H0
Private Const SHGFI_SMALLICON = &H1
Private Const ILD_TRANSPARENT = &H1
Private Const Flags = SHGFI_TYPENAME Or SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX Or SHGFI_DISPLAYNAME Or SHGFI_EXETYPE

Private Type SHELLEXECUTEINFO
    cbSize As Long
    fMask As Long
    hwnd As Long
    lpVerb As String
    lpFile As String
    lpParameters As String
    lpDirectory As String
    nShow As Long
    hInstApp As Long
    lpIDList As Long
    lpClass As String
    hkeyClass As Long
    dwHotKey As Long
    hIcon As Long
    hProcess As Long
End Type

Const SEE_MASK_INVOKEIDLIST = &HC
Const SEE_MASK_NOCLOSEPROCESS = &H40
Const SEE_MASK_FLAG_NO_UI = &H400

Private FileInfo As typSHFILEINFO

Public Function ExtractIcon(filename As String, AddtoImageList As ImageList, PictureBox As PictureBox, PixelsXY As Integer) As Long
    Dim SmallIcon As Long
    Dim NewImage As ListImage
    Dim IconIndex As Integer
    DoEvents
    If PixelsXY = 16 Then
        SmallIcon = SHGetFileInfo(filename, 0&, FileInfo, Len(FileInfo), Flags Or SHGFI_SMALLICON)
    Else
        SmallIcon = SHGetFileInfo(filename, 0&, FileInfo, Len(FileInfo), Flags Or SHGFI_LARGEICON)
    End If
    
    If SmallIcon <> 0 Then
      With PictureBox
        .Height = 15 * PixelsXY
        .Width = 15 * PixelsXY
        .ScaleHeight = 15 * PixelsXY
        .ScaleWidth = 15 * PixelsXY
        .Picture = LoadPicture("")
        .AutoRedraw = True
        SmallIcon = ImageList_Draw(SmallIcon, FileInfo.iIcon, PictureBox.hDC, 0, 0, ILD_TRANSPARENT)
        .Refresh
      End With
      
      IconIndex = AddtoImageList.ListImages.Count + 1
      Set NewImage = AddtoImageList.ListImages.Add(IconIndex, , PictureBox.Image)
      ExtractIcon = IconIndex
    End If
End Function

Public Function GetSize(ByVal FileSize As String)
    DoEvents
    Select Case Val(FileSize)
        Case 0 To 999
            GetSize = Round(Val(FileSize), 2) & " Bytes"
            Exit Function
        Case 1000 To 999999
            GetSize = Left(Format(Val(FileSize) / 1024, "#,##.00"), 4)
            If Right(GetSize, 1) <> "." Then
                GetSize = GetSize & " KB"
            ElseIf Right(GetSize, 1) = "." Then
                GetSize = Replace(GetSize, ".", "", 1) & " KB"
            End If
            Exit Function
        Case 1000000 To 1029999999
            GetSize = Left(Format((Val(FileSize) / 1048576), "#,##.00"), 4)
            If Right(GetSize, 1) <> "." Then
                GetSize = GetSize & " MB"
            ElseIf Right(GetSize, 1) = "." Then
                GetSize = Replace(GetSize, ".", "", 1) & " MB"
            End If
            Exit Function
        Case Is >= 1030000000
            GetSize = Left(Format((Val(FileSize) / 1071576), "#,##.00"), 4)
            If Right(GetSize, 1) <> "." Then
                GetSize = GetSize & " GB"
            ElseIf Right(GetSize, 1) = "." Then
                GetSize = Replace(GetSize, ".", "", 1) & " GB"
            End If
            Exit Function
    End Select
End Function

Function ShowFileProperties(filename As String, OwnerhWnd As Long) As Long
    Dim SEI As SHELLEXECUTEINFO
    DoEvents
    With SEI
        .cbSize = Len(SEI)
        .fMask = SEE_MASK_NOCLOSEPROCESS Or SEE_MASK_INVOKEIDLIST Or SEE_MASK_FLAG_NO_UI
        .hwnd = OwnerhWnd
        .lpVerb = "properties"
        .lpFile = filename
        .lpParameters = vbNullChar
        .lpDirectory = vbNullChar
        .nShow = 0
        .hInstApp = 0
        .lpIDList = 0
    End With
    
    ShellExecuteEX SEI
    ShowFileProperties = SEI.hInstApp
End Function


