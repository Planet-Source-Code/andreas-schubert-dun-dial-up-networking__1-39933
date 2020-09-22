VERSION 5.00
Begin VB.Form frm_main 
   Caption         =   "Form2"
   ClientHeight    =   6120
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9840
   LinkTopic       =   "Form2"
   ScaleHeight     =   6120
   ScaleWidth      =   9840
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton cmd_createNew 
      Caption         =   "Create new Entry"
      Height          =   255
      Left            =   2640
      TabIndex        =   12
      Top             =   1440
      Width           =   2055
   End
   Begin VB.CommandButton cmd_SetDialingParams 
      Caption         =   "Set dialing parameters"
      Height          =   255
      Left            =   5040
      TabIndex        =   9
      Top             =   1800
      Width           =   1935
   End
   Begin VB.CommandButton cmd_GetDialingParams 
      Caption         =   "Get dialing parameters"
      Height          =   255
      Left            =   5040
      TabIndex        =   8
      Top             =   1440
      Width           =   1935
   End
   Begin VB.CommandButton cmd_SetProps 
      Caption         =   "Set Entry Properties"
      Height          =   255
      Left            =   2640
      TabIndex        =   7
      Top             =   3240
      Width           =   2055
   End
   Begin VB.CommandButton cmd_GetEntryProps 
      Caption         =   "Get Entry Properties"
      Height          =   255
      Left            =   2640
      TabIndex        =   6
      Top             =   2880
      Width           =   2055
   End
   Begin VB.CommandButton cmd_ValidateEntryName 
      Caption         =   "Validate Entry name"
      Height          =   255
      Left            =   2640
      TabIndex        =   5
      Top             =   2520
      Width           =   2055
   End
   Begin VB.CommandButton cmd_deleteEntry 
      Caption         =   "Delete Entry"
      Height          =   255
      Left            =   2640
      TabIndex        =   4
      Top             =   2160
      Width           =   2055
   End
   Begin VB.CommandButton cmd_renameEntry 
      Caption         =   "Rename Entry"
      Height          =   255
      Left            =   2640
      TabIndex        =   3
      Top             =   1800
      Width           =   2055
   End
   Begin VB.CommandButton cmd_refresh 
      Caption         =   "refresh list"
      Height          =   255
      Left            =   2640
      TabIndex        =   2
      Top             =   720
      Width           =   4335
   End
   Begin VB.ListBox List1 
      Height          =   2400
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   2175
   End
   Begin VB.Shape Shape2 
      Height          =   1095
      Left            =   4920
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Shape Shape1 
      Height          =   2655
      Left            =   2520
      Top             =   1080
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "Dialing information"
      Height          =   255
      Index           =   1
      Left            =   4920
      TabIndex        =   11
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Entry information"
      Height          =   255
      Index           =   0
      Left            =   2640
      TabIndex        =   10
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "existing dial-up phonebook entries:"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   360
      Width           =   3375
   End
End
Attribute VB_Name = "frm_main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim retVal As Long
Dim myRASEntryName() As tRasEntryName

Private Sub cmd_createNew_Click()
' creating a new entry works just like setting data for an existing entry,
' except you have to specify a new and valid name for the connection
' then you set up your parameters for the tRasEntry structure
' and call fRasSetEntryProperties
End Sub

Private Sub cmd_deleteEntry_Click()
If MsgBox("Really delete entry '" & myRASEntryName(Me.List1.ListIndex) & "'?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
retVal = fRASDeleteEntry(myRASEntryName(Me.List1.ListIndex))
If retVal <> 0 Then
    MsgBox fRASErrorHandler(retVal)
End If
cmd_refresh_Click
End Sub

Private Sub cmd_GetDialingParams_Click()
' the function fRasGetEntryDialParams gives you information
' about dialing parameters of an existing connection, such as
' Username and Password to use:
Dim b() As Byte
Dim retVal As Long
Dim clsRasDialParams As tRasDialParams
retVal = fRasGetEntryDialParams(b, vbNullString, "MyConnection")
If retVal <> 0 Then MsgBox fRASErrorHandler(retVal)
' remember to transform the byte array to a tRasDialParams structure,
' except you are able to read byte arrays more easily then plaintext ;-)
If fBytesToRasDialParams(b, clsRasDialParams) Then
       ' do whatever you want here, e.g.
       Debug.Print clsRasDialParams.UserName
Else
       MsgBox "Structure could not be transformed!", vbError
End If

End Sub

Private Sub cmd_GetEntryProps_Click()
' see documentation for deails!
Dim clsRasEntry As tRasEntry
Dim retVal As Long
retVal = fRasGetEntryProperties(myRASEntryName(Me.List1.ListIndex).EntryName, clsRasEntry)
If retVal <> 0 Then MsgBox fRASErrorHandler(retVal)
' now the clsRasentry structure is filled and you could display the values
' e.g.
Debug.Print clsRasEntry.CountryCode & "/" & clsRasEntry.AreaCode & "/" & clsRasEntry.LocalPhoneNumber
End Sub

Private Sub cmd_refresh_Click()
Me.List1.Clear
retVal = fRasGetAllEntries(myRASEntryName)
For i = 0 To retVal - 1
    Me.List1.AddItem myRASEntryName(i).EntryName
Next i
If Me.List1.ListIndex < 1 Then Call Form_Load
End Sub


Private Sub cmd_renameEntry_Click()
aw = InputBox("Enter the new name for entry '" & myRASEntryName(Me.List1.ListIndex).EntryName & "'", , myRASEntryName(Me.List1.ListIndex).EntryName)
If aw = "" Then Exit Sub
If MsgBox("Really change the name to '" & aw & "'?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
retVal = fRASRenameEntry(myRASEntryName(Me.List1.ListIndex).EntryName, aw)
If retVal <> 0 Then
    MsgBox fRASErrorHandler(retVal)
End If
cmd_refresh_Click
End Sub

Private Sub cmd_SetDialingParams_Click()
' as easy as reading dialing parameters:
Dim b() As Byte
Dim retVal As Long
Dim tDialP As tRasDialParams
tDialP.UserName = "XXX"  ' set your information here
retVal = fRasDialParamsToBytes(tDialP, b)
retVal = fRasSetEntryDialParams(vbNullString, b)
If retVal <> 0 Then MsgBox fRASErrorHandler(retVal)

End Sub

Private Sub cmd_SetProps_Click()
' see documentation for details.
' it works like this:
Dim clsRasEntry As tRasEntry
Dim retVal As Long
' you set up the appropriate information
'clsRasEntry.CountryCode = "+49"
'clsRasEntry.AreaCode = "089"
'clsRasEntry.LocalPhoneNumber = "12345"
' and then you call fRasSetEntryProperties("Your connection",clsrasentry)

retVal = fRasSetEntryProperties("MyConnection", clsRasEntry)
If retVal <> 0 Then MsgBox fRASErrorHandler(retVal)

End Sub

Private Sub cmd_ValidateEntryName_Click()
aw = InputBox("Enter the connection name to check!")
If aw = "" Then Exit Sub
retVal = fRASValidateEntryName(aw)
Select Case retVal
    Case 0   ' if the name is valid and does not exist already
        MsgBox "'" & aw & "' is a valid name and a connection with this name does not yet exist!", vbInformation
    Case 123 ' if the name syntax is invalid
        MsgBox "The name syntax '" & aw & "' is valid!", vbInformation
    Case 183 ' if the entry name already exists
        MsgBox "'" & aw & "' is the name of a existing connection!", vbInformation
    Case Else
        MsgBox "The following error occured while validating entry name: " & fRASErrorHandler(retVal), vbCritical + vbOKOnly
End Select
End Sub

Private Sub Form_Load()
    Me.cmd_deleteEntry.Enabled = False
    Me.cmd_GetEntryProps.Enabled = False
    Me.cmd_renameEntry.Enabled = False
    Me.cmd_SetProps.Enabled = False
End Sub

Private Sub List1_Click()
    Me.cmd_deleteEntry.Enabled = True
    Me.cmd_GetEntryProps.Enabled = True
    Me.cmd_renameEntry.Enabled = True
    Me.cmd_SetProps.Enabled = True
End Sub
