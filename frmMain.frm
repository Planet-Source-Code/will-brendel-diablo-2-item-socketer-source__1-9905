VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Socketer by Disk2 - [No File]"
   ClientHeight    =   1950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4170
   ControlBox      =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   4170
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSocketAll 
      Caption         =   "So&cket All"
      Enabled         =   0   'False
      Height          =   390
      Left            =   2475
      TabIndex        =   6
      Top             =   525
      Width           =   1365
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   390
      Left            =   2475
      TabIndex        =   4
      Top             =   1425
      Width           =   1365
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "&About"
      Height          =   390
      Left            =   2475
      TabIndex        =   3
      Top             =   975
      Width           =   1365
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "&Open Character"
      Height          =   390
      Left            =   2475
      TabIndex        =   0
      Top             =   75
      Width           =   1365
   End
   Begin VB.ComboBox cmbList 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "frmMain.frx":0442
      Left            =   300
      List            =   "frmMain.frx":044F
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   300
      Width           =   1365
   End
   Begin VB.CommandButton cmdSocket 
      Caption         =   "Socket Item"
      Enabled         =   0   'False
      Height          =   390
      Left            =   300
      TabIndex        =   2
      Top             =   750
      Width           =   1365
   End
   Begin MSComDlg.CommonDialog dlgMain 
      Left            =   1800
      Top             =   450
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Image Image1 
      Height          =   1815
      Left            =   2025
      Picture         =   "frmMain.frx":047A
      Stretch         =   -1  'True
      Top             =   75
      Width           =   30
   End
   Begin VB.Label lblInfo 
      Caption         =   "Item To Socket:"
      Height          =   240
      Left            =   300
      TabIndex        =   5
      Top             =   75
      Width           =   1365
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Socketer v1.10 Source Code                                    '
' =============================================================='
' This code is probably very inefficient and sloppy,            '
' but I released it to give any newbie game hacker              '
' an example of a simple saved game editor. This program        '
' edits items in Diablo 2 (www.blizzard.com). It sockets        '
' them.                                                         '
'                                                               '
' The program works by cycling through the items and replacing  '
' a byte with 0x08 (the marker of a socketed item :). Well,     '
' it's written in VB so it can't be TOO hard :), so here it is. '
'                                                               '
' BTW: If you use this code to make an editor, kindly give me   '
' some credit. Thanks ;)                                        '
' Also, the comments I provided might be confusing. Sorry, I'm  '
' not a good writer :)                                          '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' This is a flag used in the Open button. It makes sure the filename selected exists...
Private Const OFN_FILEMUSTEXIST = &H1000

' The inventory is marked by a "JM  JM". It's also ended by that...
Private Type ItemHeader
    szFirstJM As String * 2 ' The first JM
    iItemCount As Byte ' This is a FAKE item count. This caused errors in v1.00 of Socketer
    iEmpty As Byte ' NULL byte
    szLastJM As String * 2 ' The final JM
End Type

' This is obviously not a complete item type :)
' This only holds the inventory position and the Equipped position.
' It's all I need for this purpose...
Private Type Item
    iSocketed As Byte
    iInvPos As Byte
End Type

' Hold the filename
Dim strFileName As String
' Used to find the beginning of the file...
Dim ItemHead As ItemHeader

Private Sub cmdExit_Click()
    ' Exit the program :)
    End
End Sub

Private Sub cmdAbout_Click()
    ' Just show the about info... No need for a form when a message box will do :)
    MsgBox "Socketer v1.10" & vbCrLf & vbCrLf & "Written by Disk2" & vbCrLf & "Many thanks to Qster for helping me test Socketer :)" & vbCrLf & vbCrLf & "This program transforms any equipped weapon, shield, or helm into a socketed item.", vbOKOnly + vbInformation, "About Socketer"
End Sub

Private Sub cmdOpen_Click()
    ' Set the properties...
    dlgMain.DialogTitle = "Open Character"
    dlgMain.Filter = "Diablo 2 Saved Games (*.d2s)|*.d2s|"
    dlgMain.InitDir = getstring(HKEY_LOCAL_MACHINE, "SOFTWARE\Blizzard Entertainment\Diablo II", "Save Path")
    dlgMain.Flags = &H1000
    dlgMain.CancelError = False
    
    ' Show the window...
    dlgMain.ShowOpen
    
    ' If the user chose a file, open it and enable all the controls
    If Len(dlgMain.FileName) > 0 Then
        strFileName = dlgMain.FileName ' Set the filename
        
        ' Set the caption
        frmMain.Caption = "Socketer by Disk2 - [" & dlgMain.FileTitle & "]"
        
        ' Enable controls
        cmbList.Enabled = True
        cmdSocket.Enabled = True
        cmdSocketAll.Enabled = True
    End If
End Sub

Private Sub cmdSocket_Click()
    ' Ok, depending on which item was chosen socket the item and show the appropriate message
    If cmbList.Text = "Left-Hand Item" Then
        Socket &H4, "Your left-hand item is now socketed."
    ElseIf cmbList.Text = "Right-Hand Item" Then
        Socket &H5, "Your right-hand item is now socketed."
    ElseIf cmbList.Text = "Helm" Then
        Socket &H1, "Your helm is now socketed."
    End If
End Sub

Private Sub cmdSocketAll_Click()
    ' Socket all the items and show the message
    Socket &H0, "All of your equipped socketable items are now socketed!"
End Sub

Private Sub Form_Load()
    ' If this is the first time the program has been run, show the legal notice, etc...
    ' This isn't important really :)
    If GetSetting("Socketer by Disk2", "Data", "FirstUse", "1") = "1" Then
        frmNotice.Show vbModal, frmMain
        
        SaveSetting "Socketer by Disk2", "Data", "FirstUse", "0"
    End If
End Sub

Private Sub Socket(Position As Integer, Message As String)
    On Error Resume Next ' If we encounter an error, resume next :)
    
    Dim iPos As Integer ' IMPORTANT: This holds our position in the file...
    Dim xItem As Item ' The temp item. Used to compare item positions, etc...
    Dim TheEnd As ItemHeader ' Used to check if we're at the end of the inventory
    Dim TheString As String * 4 ' Should be "JMJM" if we're at the end of the inventory

    ' See declaration of ItemHeader type (at top of code)
    TheEnd.iEmpty = &H0
    TheEnd.iItemCount = &H0
    TheEnd.szFirstJM = ""
    TheEnd.szLastJM = ""

    ' Clear ItemHead (it's a global variable so we need to clear it each time...)
    ItemHead.szFirstJM = ""
    ItemHead.szLastJM = ""
    ItemHead.iItemCount = 0
    ItemHead.iEmpty = 0

    ' Start at the beginning of the file
    iPos = &H1

    ' Open the filename (strFileName)
    Open strFileName For Binary As #1
        ' Get the position of the start of the inventory data
        Do Until ItemHead.szFirstJM = "JM" And ItemHead.szLastJM = "JM"
            Get #1, iPos, ItemHead
            
            iPos = iPos + 1
        Loop
    
        ' OK. We found it. Now we have to increase our position by 3 to get to the first item
        iPos = iPos + 3

        ' If the item count is zero then there's no point in continuing :)
        If ItemHead.iItemCount = 0 Then
            MsgBox "This character doesn't appear to have any items! If this is an error please email me at cregistry@yahoo.com and attach the saved game file. Thanks!", vbOKOnly + vbInformation, "Notice"
            
            ' Close the file and exit the sub...
            Close #1
            Exit Sub
        End If

        ' The ItemHead.iItemCount is a fake value for the number of items (I guess).
        ' The number doesn't account for gems that are in socketed items.
        ' So now we have to read items until we find the closing "JM  JM" in the file...
        Do Until TheString = "JMJM"
            ' First of all, make sure we aren't at the end of the inventory
            Get #1, iPos, TheEnd
            ' If TheString equals "JMJM" then we are at the end
            TheString = TheEnd.szFirstJM & TheEnd.szLastJM
            
            ' Increase our position by 2 to get to the item data
            ' BTW: Each item is 25 bytes long...
            iPos = iPos + 2
            
            ' Read the position of the item.
            Get #1, iPos + 4, xItem.iInvPos
            
            ' Depending on the value of Position when the function was called,
            ' socket the appropriate item(s).
            ' BTW: &H0 means socket all the items that are equipped...
            Select Case Position
                Case &H0
                    If xItem.iInvPos = &H1 Or xItem.iInvPos = &H4 Or xItem.iInvPos = &H5 Then
                        Put #1, iPos + 1, &H8
                    End If
                Case &H1
                    If xItem.iInvPos = &H1 Then
                        Put #1, iPos + 1, &H8
                    End If
                Case &H4
                    If xItem.iInvPos = &H4 Then
                        Put #1, iPos + 1, &H8
                    End If
                Case &H5
                    If xItem.iInvPos = &H5 Then
                        Put #1, iPos + 1, &H8
                    End If
            End Select
            
            ' Increase the position by 25 so that we can read the next item...
            iPos = iPos + 25
            
            ' Then loop back to the beginning :)
        Loop
        
    ' Close the file
    Close #1
    
    ' Show the message (this is shown regardless of wether the item was socketed
    ' successfully :)
    MsgBox "Done!" & Message, vbOKOnly + vbInformation, "Done"
End Sub
