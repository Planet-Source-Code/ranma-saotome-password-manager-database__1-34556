VERSION 5.00
Begin VB.Form frmPW 
   BackColor       =   &H0099816A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Password-Manager"
   ClientHeight    =   4020
   ClientLeft      =   1095
   ClientTop       =   435
   ClientWidth     =   7110
   Icon            =   "frmPW.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   7110
   Begin VB.CommandButton Command3 
      Caption         =   "no mask"
      Height          =   300
      Left            =   2520
      TabIndex        =   24
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "show Password"
      Height          =   300
      Left            =   1080
      TabIndex        =   22
      Top             =   3000
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "E&xit"
      Height          =   300
      Left            =   5880
      TabIndex        =   21
      Top             =   3600
      Width           =   1095
   End
   Begin VB.ListBox List1 
      Height          =   2205
      Left            =   4680
      TabIndex        =   20
      Top             =   840
      Width           =   2295
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   7050
      TabIndex        =   18
      Top             =   0
      Width           =   7110
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "for a professional tool visit http://www.steganos.com"
         Height          =   255
         Left            =   840
         TabIndex        =   23
         Top             =   360
         Width           =   3855
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   240
         Picture         =   "frmPW.frx":030A
         Top             =   120
         Width           =   480
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Password Database Tool "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   19
         Top             =   120
         Width           =   3255
      End
   End
   Begin VB.CommandButton cmdLast 
      Caption         =   ">>"
      Height          =   300
      Left            =   4800
      TabIndex        =   16
      Top             =   3600
      UseMaskColor    =   -1  'True
      Width           =   345
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   ">"
      Height          =   300
      Left            =   4440
      TabIndex        =   15
      Top             =   3600
      UseMaskColor    =   -1  'True
      Width           =   345
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "<"
      Height          =   300
      Left            =   4080
      TabIndex        =   14
      Top             =   3600
      UseMaskColor    =   -1  'True
      Width           =   345
   End
   Begin VB.CommandButton cmdFirst 
      Caption         =   "<<"
      Height          =   300
      Left            =   3720
      TabIndex        =   13
      Top             =   3600
      UseMaskColor    =   -1  'True
      Width           =   345
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   300
      Left            =   2520
      TabIndex        =   12
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
      Height          =   300
      Left            =   1320
      TabIndex        =   11
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   300
      Left            =   1320
      TabIndex        =   10
      Top             =   3600
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   300
      Left            =   120
      TabIndex        =   9
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update"
      Height          =   300
      Left            =   120
      TabIndex        =   8
      Top             =   3600
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Password"
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   3
      Left            =   1080
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   2640
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Notes"
      Height          =   870
      Index           =   2
      Left            =   1080
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   1545
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Username"
      Height          =   285
      Index           =   1
      Left            =   1080
      TabIndex        =   2
      Top             =   1215
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Location"
      Height          =   285
      Index           =   0
      Left            =   1080
      TabIndex        =   1
      Top             =   900
      Width           =   3375
   End
   Begin VB.Label lblStatus 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   120
      TabIndex        =   17
      Top             =   4560
      Width           =   3360
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   7
      Top             =   2640
      Width           =   1005
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Notes:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   6
      Top             =   1545
      Width           =   1005
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Username:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   1215
      Width           =   1005
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Location:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   900
      Width           =   1005
   End
End
Attribute VB_Name = "frmPW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Keeps the Code clean
Dim WithEvents rs As Recordset
Attribute rs.VB_VarHelpID = -1
Dim mbChangedByCode As Boolean
Dim mvBookMark As Variant
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean
Dim mbDataChanged As Boolean

Private Sub Command1_Click()
Unload Me
End
End Sub

Private Sub Command2_Click()
Dim clsDS2 As New clsDS2
MsgBox "Your Password is:   " & vbCrLf & vbCrLf & _
clsDS2.DecryptString(txtFields(3), True), vbInformation, "Password Manager..."
End Sub

Private Sub Command3_Click()
'On Error Resume Next
If Command3.Caption = "no mask" Then
Command3.Caption = "mask"
txtFields(3).PasswordChar = ""
Else
Command3.Caption = "no mask"
txtFields(3).PasswordChar = "*"
End If
End Sub

Private Sub Form_Load()
  Dim db As Connection
  
    
   On Error GoTo ErrHandler
   
  Set db = New Connection
  db.CursorLocation = adUseClient
  db.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\PW.mdb"
  Set rs = New Recordset
  rs.Open "select Location,Username,Notes,Password from Manager", db, adOpenStatic, adLockOptimistic

  Dim oText As TextBox
  'Bind the text boxes to the data provider
  For Each oText In Me.txtFields
    Set oText.DataSource = rs
  Next

Me.Caption = App.Title & " Version " & App.Major & "." & App.Minor & "." & App.Revision
Call Styleme
Call Lockme
Call listme
'works great
App.TaskVisible = False

  mbDataChanged = False

Exit_:
 Screen.MousePointer = vbNormal
 On Error Resume Next
 Exit Sub

ErrHandler:
 Screen.MousePointer = vbNormal
 MsgBox "Error..." & Err.Number & " in " & Err.Description, vbCritical
 Resume Exit_
End Sub

Private Sub Form_Resize()
  On Error Resume Next

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If mbEditFlag Or mbAddNewFlag Then Exit Sub

  Select Case KeyCode
    Case vbKeyEnd
      cmdLast_Click
    Case vbKeyHome
      cmdFirst_Click
    Case vbKeyUp, vbKeyPageUp
      If Shift = vbCtrlMask Then
        cmdFirst_Click
      Else
        cmdPrevious_Click
      End If
    Case vbKeyDown, vbKeyPageDown
      If Shift = vbCtrlMask Then
        cmdLast_Click
      Else
        cmdNext_Click
      End If
  End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

Private Sub List1_Click()
On Error Resume Next
List1.ToolTipText = List1.Text
Call Search(List1.Text, rs, rs.Fields("Location"))
End Sub

Private Sub RS_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'This will display the current record position for this recordset
  lblStatus.Caption = "Record: " & CStr(rs.AbsolutePosition)
End Sub

Private Sub RS_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'This is where you put validation code
  'This event gets called when the following actions occur
  Dim bCancel As Boolean

  Select Case adReason
  Case adRsnAddNew
  Case adRsnClose
  Case adRsnDelete
  Case adRsnFirstChange
  Case adRsnMove
  Case adRsnRequery
  Case adRsnResynch
  Case adRsnUndoAddNew
  Case adRsnUndoDelete
  Case adRsnUndoUpdate
  Case adRsnUpdate
  End Select

  If bCancel Then adStatus = adStatusCancel
End Sub

Private Sub cmdAdd_Click()
  On Error GoTo AddErr
  With rs
    If Not (.BOF And .EOF) Then
      mvBookMark = .Bookmark
    End If
    .AddNew
    
    Call unLockme
    
    lblStatus.Caption = "Add record"
    mbAddNewFlag = True
    SetButtons False
  End With

  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cmdDelete_Click()
On Error Resume Next

If MsgBox("Delete Entry??", vbCritical + vbYesNo, "Password Manager?") = vbYes Then
  With rs
    .Delete
    .MoveNext
    If .EOF Then .MoveLast
    Call listme
  End With
  Exit Sub
End If

End Sub

Private Sub cmdEdit_Click()
On Error Resume Next

If MsgBox("This will delete your current password! Continue ?", vbCritical + vbYesNo, "Password Manager?") = vbYes Then

Call unLockme
txtFields(3).Text = ""

  lblStatus.Caption = "Edit record"
  mbEditFlag = True
  SetButtons False
  Exit Sub
End If


End Sub
Private Sub cmdCancel_Click()
  On Error Resume Next

  SetButtons True
  mbEditFlag = False
  mbAddNewFlag = False
  rs.CancelUpdate
  Call Lockme
  If mvBookMark > 0 Then
    rs.Bookmark = mvBookMark
  Else
    rs.MoveFirst
  End If
  mbDataChanged = False

End Sub

Private Sub cmdUpdate_Click()
Dim clsDS2 As New clsDS2
  On Error GoTo UpdateErr
 
'encrypting Password before Safing
txtFields(3).Text = clsDS2.EncryptString(txtFields(3), True)
  
  rs.UpdateBatch adAffectAll
  
 Call listme
 Call Lockme
 
  If mbAddNewFlag Then
    rs.MoveLast              'move to the new record
    
  End If

  mbEditFlag = False
  mbAddNewFlag = False
  SetButtons True
  mbDataChanged = False

  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub cmdFirst_Click()
  On Error GoTo GoFirstError

  rs.MoveFirst
  mbDataChanged = False

  Exit Sub

GoFirstError:
  MsgBox Err.Description
End Sub

Private Sub cmdLast_Click()
  On Error GoTo GoLastError

  rs.MoveLast
  mbDataChanged = False

  Exit Sub

GoLastError:
  MsgBox Err.Description
End Sub

Private Sub cmdNext_Click()
  On Error GoTo GoNextError

  If Not rs.EOF Then rs.MoveNext
  If rs.EOF And rs.RecordCount > 0 Then
    Beep
     'moved off the end so go back
    rs.MoveLast
  End If
  'show the current record
  mbDataChanged = False

  Exit Sub
GoNextError:
  MsgBox Err.Description
End Sub

Private Sub cmdPrevious_Click()
  On Error GoTo GoPrevError

  If Not rs.BOF Then rs.MovePrevious
  If rs.BOF And rs.RecordCount > 0 Then
    Beep
    'moved off the end so go back
    rs.MoveFirst
  End If
  'show the current record
  mbDataChanged = False

  Exit Sub

GoPrevError:
  MsgBox Err.Description
End Sub

Private Sub SetButtons(bVal As Boolean)
  cmdAdd.Visible = bVal
  cmdEdit.Visible = bVal
  cmdUpdate.Visible = Not bVal
  cmdCancel.Visible = Not bVal
  cmdDelete.Visible = bVal
  cmdNext.Enabled = bVal
  cmdFirst.Enabled = bVal
  cmdLast.Enabled = bVal
  cmdPrevious.Enabled = bVal
End Sub


Private Sub Styleme()
Dim intfields As Integer

        For intfields = 0 To 3
        MakeFlat txtFields(intfields).hwnd
        Next
        
MakeFlat List1.hwnd
CButton cmdAdd
CButton cmdEdit
CButton cmdUpdate
CButton cmdCancel
CButton cmdDelete
CButton cmdNext
CButton cmdFirst
CButton cmdLast
CButton cmdPrevious
CButton Command1
CButton Command2
CButton Command3
End Sub


Private Sub Lockme()
Dim intfields As Integer

         For intfields = 0 To 3
         frmPW.txtFields(intfields).Locked = True
         Next
End Sub

Private Sub unLockme()
Dim intfields As Integer

         For intfields = 0 To 3
         frmPW.txtFields(intfields).Locked = False
         Next
End Sub

Private Sub listme()
On Error Resume Next
 List1.Clear
 rs.MoveFirst
 While Not rs.EOF
   List1.AddItem rs.Fields("Location")
   rs.MoveNext
 Wend
rs.MoveFirst
End Sub
