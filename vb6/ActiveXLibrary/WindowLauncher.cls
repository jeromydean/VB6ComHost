VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WindowLauncher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Sub ShowModal()
    Dim frm1 As Form1
    Set frm1 = New Form1
    frm1.Show vbModal
End Sub

Public Sub ShowNonModal()
    Dim frm1 As Form1
    Set frm1 = New Form1
    frm1.Show vbModeless
End Sub
