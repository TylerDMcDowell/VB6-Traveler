VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSubSec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim db As ADODB.Connection
Dim WithEvents adoPrimaryRS As ADODB.Recordset
Attribute adoPrimaryRS.VB_VarHelpID = -1

Public Function GetRecordset(msSql As String) As ADODB.Recordset
  Dim adoDisconnect As ADODB.Recordset
  Set db = New ADODB.Connection
  Set adoPrimaryRS = New Recordset
  Dim icount As Integer
  db.CursorLocation = adUseClient
  db.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "/abc2.mdb"
  adoPrimaryRS.Open msSql, db, adOpenStatic, adLockOptimistic
  
 ' On Error Resume Next
  Set adoDisconnect = Nothing
  Set adoDisconnect = New Recordset
  adoDisconnect.CursorLocation = adUseClient
  adoDisconnect.ActiveConnection = Nothing
  adoDisconnect.LockType = adLockBatchOptimistic
  
  For icount = 0 To adoPrimaryRS.Fields.Count - 1
        adoDisconnect.Fields.Append adoPrimaryRS.Fields(icount).Name, _
                                    adoPrimaryRS.Fields(icount).Type, _
                                    adoPrimaryRS.Fields(icount).DefinedSize, _
                                    adoPrimaryRS.Fields(icount).Attributes
  Next icount
  
  adoDisconnect.Open
  adoDisconnect.MarshalOptions = adMarshalModifiedOnly
  
  Do Until adoPrimaryRS.EOF
    adoDisconnect.AddNew
    For icount = 0 To adoPrimaryRS.Fields.Count - 1
        adoDisconnect.Fields(icount) = adoPrimaryRS.Fields(icount)
    Next icount
    adoPrimaryRS.MoveNext
  Loop
  
    If adoDisconnect.BOF = False Then adoDisconnect.MoveFirst
  
  Set GetRecordset = adoDisconnect
  Set adoPrimaryRS = Nothing
  Set db = Nothing
End Function

