VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cptStatusSheet_frm 
   Caption         =   "Create Status Sheet"
   ClientHeight    =   7230
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12375
   OleObjectBlob   =   "cptStatusSheet_frm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "cptStatusSheet_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<cpt_version>v1.2.11</cpt_version>
Option Explicit
Private Const BLN_TRAP_ERRORS As Boolean = True
'If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
Private Const adVarChar As Long = 200
Private Const adInteger As Long = 3

Private Sub cboCreate_Change()
'objects
'strings
'longs
Dim lngField As Long
'integers
'doubles
'booleans
'variants
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  Select Case Me.cboCreate
    Case 0 'A single workbook
      Me.lboItems.ForeColor = -2147483630
      Me.chkSendEmails.Caption = "Create Email"
      Me.lblForEach.Visible = False
      Me.cboEach.Enabled = False
      Me.lboItems.Enabled = False
      FilterClear

    Case 1 'A worksheet for each
      Me.lboItems.ForeColor = -2147483630
      Me.chkSendEmails.Caption = "Create Email"
      Me.lblForEach.Visible = True
      Me.cboEach.Enabled = True
      Me.lboItems.Enabled = True
      If Me.Visible Then Me.cboEach.DropDown

    Case 2 'A workbook for each
      Me.lboItems.ForeColor = -2147483630
      Me.chkSendEmails.Caption = "Create Email(s)"
      Me.lblForEach.Visible = True
      Me.cboEach.Enabled = True
      Me.lboItems.Enabled = True
      If Me.Visible Then Me.cboEach.DropDown

    End Select
        
exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptStatusSheet_frm", "cboCreate_Change", Err, Erl)
  Resume exit_here
End Sub

Private Sub cboEach_Change()
'objects
Dim Task As Task
'strings
Dim strFieldName As String
'longs
Dim lngItem As Long
Dim lngField As Long
'integers
'doubles
'booleans
'variants
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  Me.lboItems.Clear
  Me.lboItems.ForeColor = -2147483630
  FilterClear
  
  On Error Resume Next
  lngField = FieldNameToFieldConstant(Me.cboEach)
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  
  If lngField > 0 Then
    With CreateObject("System.Collections.SortedList")
      For Each Task In ActiveProject.Tasks
        If Task Is Nothing Then GoTo next_task
        If Task.Summary Then GoTo next_task
        If Len(Task.GetField(lngField)) > 0 Then
          If Not .Contains(Task.GetField(lngField)) Then .Add Task.GetField(lngField), Task.GetField(lngField)
        End If
next_task:
      Next Task
      'validate field has items
      If .Count = 0 Then
        If Len(CustomFieldGetName(lngField)) > 0 Then
          strFieldName = CustomFieldGetName(lngField)
        Else
          strFieldName = FieldConstantToFieldName(lngField)
        End If
        MsgBox "The field '" & strFieldName & "' contains no values.", vbExclamation + vbOKOnly, "Invalid Selection"
      Else
        For lngItem = 0 To .Count - 1
          Me.lboItems.AddItem .getByIndex(lngItem)
        Next lngItem
      End If
    End With
  End If 'lngField > 0
  
exit_here:
  On Error Resume Next
  Set Task = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cboEach_Change", "cboEach_Change", Err, Erl)
  Resume exit_here
End Sub

Private Sub cboEVP_AfterUpdate()
  
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  
  If Len(Me.cboEVP.Value) > 0 Then
    Me.lblEVP.ForeColor = -2147483630 '"Black"
  Else
    Me.lblEVP.ForeColor = 192
  End If

exit_here:
  On Error Resume Next
  Exit Sub
err_here:
  Call cptHandleErr("cptStatusSheet_frm", "cboEVP_AfterUpdate", Err, Erl)
  Resume exit_here
End Sub

Private Sub cboEVP_Change()

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  If cptStatusSheet_frm.Visible Then Call cptRefreshStatusTable

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptStatusSheet_frm", "cboEVP_Change", Err, Erl)
  Resume exit_here
End Sub

Private Sub cboEVT_AfterUpdate()

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  If Len(Me.cboEVT.Value) > 0 Then
    Me.lblEVT.ForeColor = -2147483630 '"Black"
  Else
    Me.lblEVT.ForeColor = 192
  End If

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptStatusSheet_frm", "cboEVT_AfterUpdate", Err, Erl)
  Resume exit_here
End Sub

Private Sub cboEVT_Change()

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  If cptStatusSheet_frm.Visible Then Call cptRefreshStatusTable

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptStatusSheet_frm", "cboEVT_Change", Err, Erl)
  Resume exit_here
End Sub

Private Sub chkHide_Click()

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  Me.txtHideCompleteBefore.Enabled = Me.chkHide
  If Me.Visible Then Call cptRefreshStatusTable
  
exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("chkHide_Click", "chkHide_Click", Err, Erl)
  Resume exit_here
  
End Sub

Sub cmdAdd_Click()
Dim lgField As Long, lgExport As Long, lgExists As Long
Dim blnExists As Boolean

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  For lgField = 0 To Me.lboFields.ListCount - 1
    If Me.lboFields.Selected(lgField) Then
      'ensure doesn't already exist
      blnExists = False
      For lgExists = 0 To Me.lboExport.ListCount - 1
        If Me.lboExport.List(lgExists, 0) = Me.lboFields.List(lgField) Then
          GoTo next_item
        End If
      Next lgExists
      Me.lboExport.AddItem
      lgExport = Me.lboExport.ListCount - 1
      Me.lboExport.List(lgExport, 0) = Me.lboFields.List(lgField, 0)
      Me.lboExport.List(lgExport, 1) = Me.lboFields.List(lgField, 1)
      Me.lboExport.List(lgExport, 2) = Me.lboFields.List(lgField, 2)
    End If
next_item:
  Next lgField

  Call cptRefreshStatusTable

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptStatusSheet_frm", "cmdAdd_Click", Err, Erl)
  Resume exit_here

End Sub

Private Sub cmdAddAll_Click()
Dim lgField As Long, lgExport As Long, lgExists As Long
Dim blnExists As Boolean

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  For lgField = 0 To Me.lboFields.ListCount - 1
    'ensure doesn't already exist
    blnExists = False
    For lgExists = 0 To Me.lboExport.ListCount - 1
      If Me.lboExport.List(lgExists, 0) = Me.lboFields.List(lgField) Then
        GoTo next_item
      End If
    Next lgExists
    Me.lboExport.AddItem
    lgExport = Me.lboExport.ListCount - 1
    Me.lboExport.List(lgExport, 0) = Me.lboFields.List(lgField, 0)
    Me.lboExport.List(lgExport, 1) = Me.lboFields.List(lgField, 1)
    Me.lboExport.List(lgExport, 2) = Me.lboFields.List(lgField, 2)
next_item:
  Next lgField

  Call cptRefreshStatusTable

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptStatusSheet_frm", "cmdAddAll_Click", Err, Erl)
  Resume exit_here

End Sub

Private Sub cmdCancel_Click()
Dim strFileName As String

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  strFileName = cptDir & "\settings\cpt-status-sheet-search.adtg"
  If Dir(strFileName) <> vbNullString Then Kill strFileName
  Unload Me

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptStatusSheet_frm", "cmdCancel_Click", Err, Erl)
  Resume exit_here

End Sub

Private Sub cmdDown_Click()
Dim lgExport As Long
Dim lgField As Long, strField As String, strField2 As String
  
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  
  For lgExport = Me.lboExport.ListCount - 1 To 0 Step -1
    If lgExport < Me.lboExport.ListCount - 1 Then
      If Me.lboExport.Selected(lgExport) Then
        'capture values
        lgField = Me.lboExport.List(lgExport + 1, 0)
        strField = Me.lboExport.List(lgExport + 1, 1)
        strField2 = Me.lboExport.List(lgExport + 1, 2)
        'move selected values
        Me.lboExport.List(lgExport + 1, 0) = Me.lboExport.List(lgExport, 0)
        Me.lboExport.List(lgExport + 1, 1) = Me.lboExport.List(lgExport, 1)
        Me.lboExport.List(lgExport + 1, 2) = Me.lboExport.List(lgExport, 2)
        Me.lboExport.Selected(lgExport + 1) = True
        Me.lboExport.List(lgExport, 0) = lgField
        Me.lboExport.List(lgExport, 1) = strField
        Me.lboExport.List(lgExport, 2) = strField2
        Me.lboExport.Selected(lgExport) = False
      End If
    End If
  Next lgExport

  Call cptRefreshStatusTable

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("frmStatusSeet", "cmdDown_Click", Err, Erl)
  Resume exit_here

End Sub

Private Sub cmdRemove_Click()
Dim lgExport As Long

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  For lgExport = Me.lboExport.ListCount - 1 To 0 Step -1
    If Me.lboExport.Selected(lgExport) Then
      Me.lboExport.RemoveItem lgExport
    End If
  Next lgExport

  Call cptRefreshStatusTable

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptStatusSheet_frm", "cmdRemove_Click", Err, Erl)
  Resume exit_here
  
End Sub

Private Sub cmdRemoveAll_Click()
Dim lgExport As Long

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  For lgExport = Me.lboExport.ListCount - 1 To 0 Step -1
    Me.lboExport.RemoveItem lgExport
  Next lgExport

  Call cptRefreshStatusTable

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptStatusSheet_frm", "cmdRemoveAll_Click", Err, Erl)
  Resume exit_here

End Sub

Private Sub cmdRun_Click()
'objects
'strings
'longs
Dim lngSelectedItems As Long
Dim lngItem As Long
'integers
'doubles
'booleans
Dim blnIncluded As Boolean
'variants
'dates

Dim blnError As Boolean, intOutput As Integer, intHide As Integer
Dim strFileName As String

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  blnError = False

  '-2147483630 = Black
  Me.lblStatusDate.ForeColor = -2147483630
  Me.lblEVT.ForeColor = -2147483630
  Me.lblEVP.ForeColor = -2147483630
  Me.chkHide.ForeColor = -2147483630
  Me.lblStatus.ForeColor = -2147483630
  Me.cboCostTool.ForeColor = -2147483630
  Me.cboCreate.ForeColor = -2147483630
  Me.cboEach.BorderColor = -2147483642
  
  'validation
  If Not IsDate(Me.txtStatusDate.Value) Then
    Me.lblStatusDate.ForeColor = 192  'Red
    blnError = True
  ElseIf IsDate(Me.txtStatusDate.Value) Then
    If CDate(Me.txtStatusDate.Value) < #1/1/1984# Then
      Me.lblStatusDate.ForeColor = 192  'Red
      blnError = True
    End If
  End If
  If Me.chkHide.Value = True Then
    If Not IsDate(Me.txtHideCompleteBefore.Value) Then
      Me.chkHide.ForeColor = 192  'Red
      blnError = True
    ElseIf IsDate(Me.txtHideCompleteBefore.Value) Then
      If CDate(Me.txtHideCompleteBefore.Value) < #1/1/1984# Then
        Me.chkHide.ForeColor = 192
        blnError = True
      End If
    End If
  End If
  If Len(Me.cboCostTool.Value) = 0 Then
    Me.lblCostTool.ForeColor = 192 'Red
    blnError = True
  End If
  'hide complete before must be prior to or equal to status date
  If IsDate(Me.txtStatusDate.Value) And IsDate(Me.txtHideCompleteBefore.Value) Then
    If CDate(Me.txtHideCompleteBefore.Value) > CDate(Me.txtStatusDate.Value) Then
      MsgBox "'Hide Complete Before' date must be prior to, or equal to, status date.", vbExclamation + vbOKOnly, "Invalid Hide Complete Before Date"
      Me.chkHide.ForeColor = 192
      blnError = True
    End If
  End If
  If Len(Me.cboEVT.Value) = 0 Then
    Me.lblEVT.ForeColor = 192 'Red
    blnError = True
  End If
  If Len(Me.cboEVP.Value) = 0 Then
    Me.lblEVP.ForeColor = 192 'Red
    blnError = True
  End If
  If Me.cboCreate.Value <> "0" Then
    'a limiting field must be selected
    If Me.cboEach.Value = 0 Then
      Me.cboEach.BorderColor = 192
      blnError = True
    End If
    'at least one item selected
    For lngItem = 0 To Me.lboItems.ListCount - 1
      If Me.lboItems.Selected(lngItem) Then lngSelectedItems = lngSelectedItems + 1
    Next lngItem
    If lngSelectedItems = 0 Then
      Me.lboItems.Selected(0) = True
      'Me.lboItems.ForeColor = 92
      blnError = True
    End If
    'the limiting field should be included in the export list
    blnIncluded = False
    For lngItem = 0 To Me.lboExport.ListCount - 1
      If Me.lboExport.List(lngItem, 1) = Me.cboEach Then blnIncluded = True
    Next lngItem
    If Not blnIncluded Then
      If MsgBox("The For Each field '" & Me.cboEach & "' is not included in the export list." & vbCrLf & vbCrLf & "Include it?", vbYesNo + vbQuestion, "Include For Each Field?") = vbYes Then
        For lngItem = 0 To Me.lboFields.ListCount - 1
          Me.lboFields.Selected(lngItem) = Me.lboFields.List(lngItem, 1) = Me.cboEach
        Next lngItem
        Me.cmdAdd_Click
      End If
    End If
  End If
  If blnError Then
    Me.lblStatus.ForeColor = 192 'red
    Me.lblStatus.Caption = " Please complete all required fields."
  Else
    'save settings
    strFileName = cptDir & "\settings\cpt-status-sheet.adtg"
    With CreateObject("ADODB.Recordset")
      .Fields.Append "cboEVT", adVarChar, 100
      .Fields.Append "cboEVP", adVarChar, 100
      .Fields.Append "chkOutput", adInteger
      .Fields.Append "chkHide", adInteger
      .Fields.Append "cboCostTool", adVarChar, 100
      .Fields.Append "cboEach", adVarChar, 100
      .Open
      If Me.chkHide Then intHide = 1 Else intHide = 0
      intOutput = Me.cboCreate.Value + 1
      .AddNew Array(0, 1, 2, 3, 4), Array(Me.cboEVT.Value, Me.cboEVP.Value, intOutput, intHide, Me.cboCostTool.Value)
      .Update
      .MoveFirst
      If Not IsNull(Me.cboEach.Value) Then
        .Fields("cboEach") = Me.cboEach.Value
      End If
      .Update
      If Dir(strFileName) <> vbNullString Then Kill strFileName
      .Save strFileName
      .Close
    End With
    'create the sheet
    Call cptCreateStatusSheet
  End If

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptStatusSheet_frm", "cmdRun_Click", Err, Erl)
  Resume exit_here
  
End Sub

Private Sub cmdUp_Click()
Dim lgExport As Long
Dim lgField As Long, strField As String, strField2 As String
  
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  
  For lgExport = 0 To Me.lboExport.ListCount - 1
    If lgExport > 0 Then
      If Me.lboExport.Selected(lgExport) Then
        'capture values
        lgField = Me.lboExport.List(lgExport - 1, 0)
        strField = Me.lboExport.List(lgExport - 1, 1)
        strField2 = Me.lboExport.List(lgExport - 1, 2)
        'move selected values
        Me.lboExport.List(lgExport - 1, 0) = Me.lboExport.List(lgExport, 0)
        Me.lboExport.List(lgExport - 1, 1) = Me.lboExport.List(lgExport, 1)
        Me.lboExport.List(lgExport - 1, 2) = Me.lboExport.List(lgExport, 2)
        Me.lboExport.Selected(lgExport - 1) = True
        Me.lboExport.List(lgExport, 0) = lgField
        Me.lboExport.List(lgExport, 1) = strField
        Me.lboExport.List(lgExport, 2) = strField2
        Me.lboExport.Selected(lgExport) = False
      End If
    End If
  Next lgExport
  
  Call cptRefreshStatusTable

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptStatusSheet_frm", "cmdUp_Click", Err, Erl)
  Resume exit_here
  
End Sub

Private Sub lblURL_Click()

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  If cptInternetIsConnected Then Application.FollowHyperlink "http://www.ClearPlanConsulting.com"

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptStatusSheet_frm", "lblURL", Err, Erl)
  Resume exit_here
End Sub

Private Sub lboItems_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'objects
'strings
Dim strCriteria As String
Dim strFieldName As String
'longs
Dim lngSelectedItems As Long
Dim lngItem As Long
'integers
'doubles
'booleans
'variants
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  
  cptSpeed True
  
  strFieldName = Me.cboEach.Value

  For lngItem = 0 To Me.lboItems.ListCount - 1
    If Me.lboItems.Selected(lngItem) Then
      strCriteria = strCriteria & Me.lboItems.List(lngItem) & Chr$(9)
      lngSelectedItems = lngSelectedItems + 1
    End If
  Next lngItem
  
  strCriteria = Left(strCriteria, Len(strCriteria) - 1)
  
  If lngSelectedItems < Me.lboItems.ListCount Then
    SetAutoFilter FieldName:=strFieldName, FilterType:=pjAutoFilterIn, Criteria1:=strCriteria
  Else
    FilterClear
  End If
  
  ActiveWindow.TopPane.Activate
  SelectBeginning
  
exit_here:
  On Error Resume Next
  cptSpeed False
  Exit Sub
err_here:
  Call cptHandleErr("cptStatusSheet_frm", "lboItems_AfterUpdate", Err, Erl)
  Resume exit_here
End Sub

Private Sub stxtSearch_Change()
Dim lgItem As Long

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  Me.lboFields.Clear
  
  With CreateObject("ADODB.REcordset")
    .Open cptDir & "\settings\cpt-status-sheet-search.adtg"
    If Len(Me.stxtSearch.Text) > 0 Then
      .Filter = "[Custom Field Name] LIKE '*" & cptRemoveIllegalCharacters(Me.stxtSearch.Text) & "*'"
    Else
      .Filter = 0
    End If
    If .RecordCount > 0 Then .MoveFirst
    lgItem = 0
    Do While Not .EOF
      Me.lboFields.AddItem
      Me.lboFields.List(lgItem, 0) = .Fields(0)
      Me.lboFields.List(lgItem, 1) = .Fields(1)
      Me.lboFields.List(lgItem, 2) = .Fields(2)
      .MoveNext
      lgItem = lgItem + 1
    Loop
    .Close
  End With
  
exit_here:
  On Error Resume Next
  Exit Sub
err_here:
  Call cptHandleErr("cptStatusSheet_frm", "stxtSearch_Change", Err, Erl)
  Resume exit_here
  
End Sub

Private Sub stxtSearch_Enter()
Dim lgField As Long, strFileName As String
  
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  strFileName = cptDir & "\settings\cpt-status-sheet-search.adtg"
  If Dir(strFileName) <> vbNullString Then Exit Sub
  With CreateObject("ADODB.Recordset")
    .Fields.Append "Field Constant", adVarChar, 100
    .Fields.Append "Custom Field Name", adVarChar, 100
    .Fields.Append "Local Field Name", adVarChar, 100
    .Open
    For lgField = 0 To cptStatusSheet_frm.lboFields.ListCount - 1
      .AddNew Array(0, 1, 2), Array(Me.lboFields.List(lgField, 0), cptStatusSheet_frm.lboFields.List(lgField, 1), cptStatusSheet_frm.lboFields.List(lgField, 2))
    Next lgField
    .Update
    .Save strFileName
    .Close
  End With
  
exit_here:
  On Error Resume Next
  Exit Sub
err_here:
  Call cptHandleErr("cptStatusSheet_frm", "stxtSearch_Enter", Err, Erl)
  Resume exit_here
  
End Sub

Private Sub txtHideCompleteBefore_Change()
Dim stxt As String
  
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  stxt = cptRegEx(Me.txtHideCompleteBefore.Text, "[0-9\/]*")
  Me.txtHideCompleteBefore.Text = stxt
  If Len(Me.txtHideCompleteBefore.Text) > 0 Then
    If IsDate(Me.txtHideCompleteBefore.Text) Then
      If CDate(Me.txtHideCompleteBefore.Text) > #1/1/1984# Then
        Me.chkHide.ForeColor = -2147483630 '"Black"
      Else
        Me.chkHide.ForeColor = 192 'red
      End If
    Else
      Me.chkHide.ForeColor = 192 'red
    End If
  End If

exit_here:
  On Error Resume Next
  Exit Sub
err_here:
  Call cptHandleErr("cptStatusSheet_frm", "txtHideCompleteBefore", Err, Erl)
  Resume exit_here

End Sub

Private Sub txtStatusDate_Change()
Dim stxt As String

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  stxt = cptRegEx(Me.txtStatusDate.Text, "[0-9\/]*")
  Me.txtStatusDate.Text = stxt
  If Len(Me.txtStatusDate.Text) > 0 Then
    If IsDate(Me.txtStatusDate.Text) Then
      If CDate(Me.txtStatusDate.Text) > #1/1/1984# Then
        Me.lblStatusDate.ForeColor = -2147483630 '"Black"
      Else
        Me.lblStatusDate.ForeColor = 192 'red
      End If
    Else
      Me.lblStatusDate.ForeColor = 192 'red
    End If
  End If
  
exit_here:
  On Error Resume Next
  Exit Sub
err_here:
  Call cptHandleErr("cptStatusSheet_frm", "txtStatusDate_Change", Err, Erl)
  Resume exit_here
  
End Sub
