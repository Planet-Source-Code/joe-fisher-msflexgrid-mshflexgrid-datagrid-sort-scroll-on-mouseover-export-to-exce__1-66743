VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 1  'vbDataSource
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPurchases"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim WithEvents adoPrimaryRS As Recordset
Attribute adoPrimaryRS.VB_VarHelpID = -1
Private DoingRequery As Boolean
Public Event MoveComplete()

Private Sub Class_Initialize()
  Dim db As Connection
  Set db = New Connection
  db.CursorLocation = adUseClient

  db.Provider = "Microsoft.Jet.OLEDB.4.0"
  
  db.Open (App.Path & "\AdataBasetest.mdb")

  Set adoPrimaryRS = New Recordset
  adoPrimaryRS.Open "select PO_Number,Date,[Part Number],Qty,Price,[Shipping Method],ID1 from Purchases", db, adOpenStatic, adLockOptimistic

  DataMembers.Add "Primary"
End Sub

Private Sub Class_GetDataMember(DataMember As String, Data As Object)
  Select Case DataMember
  Case "Primary"
    Set Data = adoPrimaryRS
  End Select
End Sub

Private Sub adoPrimaryRS_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  RaiseEvent MoveComplete
End Sub

Private Sub adoPrimaryRS_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
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

Public Sub FilterPO(FieldNM As String, FieldVal As Integer)
    Dim STField As String
    Dim Vstring As Integer
    STField = FieldNM
    Vstring = FieldVal
    Dim FilterS As String
    FilterS = STField & " =" & (Val(Vstring))
    adoPrimaryRS.Filter = FilterS
    
End Sub

Public Property Get EditingRecord() As Boolean
  EditingRecord = (adoPrimaryRS.EditMode <> adEditNone)
End Property

Public Property Get AbsolutePosition() As Long
  AbsolutePosition = adoPrimaryRS.AbsolutePosition
End Property

Public Sub AddNew()
  adoPrimaryRS.AddNew
End Sub

Public Sub Delete()
  adoPrimaryRS.Delete
    MoveNext
End Sub

Public Sub Requery()
  adoPrimaryRS.Requery
  DataMemberChanged "Primary"
End Sub

Public Sub Update()
  With adoPrimaryRS
    .UpdateBatch adAffectAll
    If .EditMode = adEditAdd Then
      MoveLast
    End If
  End With
End Sub

Public Sub Cancel()
  With adoPrimaryRS
    .CancelUpdate
    If .EditMode = adEditAdd Then
      MoveFirst
    End If
  End With
End Sub

Public Sub MoveFirst()
  adoPrimaryRS.MoveFirst
End Sub

Public Sub MoveLast()
  adoPrimaryRS.MoveLast
End Sub

Public Sub MoveNext()
  If Not adoPrimaryRS.EOF Then adoPrimaryRS.MoveNext
  If adoPrimaryRS.EOF And adoPrimaryRS.RecordCount > 0 Then
    Beep
    'moved off the end so go back
    adoPrimaryRS.MoveLast
  End If
End Sub

Public Sub MovePrevious()
  If Not adoPrimaryRS.BOF Then adoPrimaryRS.MovePrevious
  If adoPrimaryRS.BOF And adoPrimaryRS.RecordCount > 0 Then
    Beep
    'moved off the end so go back
    adoPrimaryRS.MoveFirst
  End If
End Sub

