VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IRMPlugin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum RMResult
    rmres_canceled = 0
    rmres_success = 1
End Enum

Public Enum RMShowMode
    rmshow_modal = vbModal
    rmshow_modeless = vbModeless
End Enum

Public Property Get Name() As String
End Property

Public Property Get Description() As String
End Property

Public Sub InitPlugin(host As IRMPluginHost)
End Sub

Public Sub DoAction(ByVal showmode As RMShowMode)
End Sub

Public Property Get Result() As RMResult
End Property

Public Property Get Value() As Variant
End Property
