VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmBoard 
   Caption         =   "UserForm1"
   ClientHeight    =   6255
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7365
   OleObjectBlob   =   "frmBoard.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmBoard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public playerNum, spaceNum As Integer

Private Sub UserForm_Initialize()
    playerNum = 0
    spaceNum = 0
End Sub

Private Sub UserForm_Click()
    nextMove
End Sub

Private Sub nextMove()
    ' Declare variables
    Dim spaceObj, playerObj As MSForms.Image
    ' Set image objects
    Set spaceObj = frmBoard.Controls("imgSpace" & spaceNum)
    Set playerObj = frmBoard.Controls("imgPlayer" & playerNum)

    ' Move little squares
    If playerNum = 0 Then
        playerObj.Left = spaceObj.Left + 4
        playerObj.Top = spaceObj.Top + 4
    ElseIf playerNum = 1 Then
        playerObj.Left = spaceObj.Left + 38
        playerObj.Top = spaceObj.Top + 4
    ElseIf playerNum = 2 Then
        playerObj.Left = spaceObj.Left + 4
        playerObj.Top = spaceObj.Top + 38
    ElseIf playerNum = 3 Then
        playerObj.Left = spaceObj.Left + 38
        playerObj.Top = spaceObj.Top + 38
    End If
    
    ' Calculate next space to move to
    If playerNum = 3 Then
        spaceNum = (spaceNum + 1) Mod 8
    End If
    
    ' Calculate next player's move
    playerNum = (playerNum + 1) Mod 4
End Sub

