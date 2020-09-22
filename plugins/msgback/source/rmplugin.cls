VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "rmplugin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements rmplugin_defs.IRMPlugin

Private clsHost     As rmplugin_defs.IRMPluginHost
Private udeResult   As RMResult
Private varValue    As Variant

Private Property Get IRMPlugin_Description() As String
    IRMPlugin_Description = "returns a text"
End Property

Private Sub IRMPlugin_DoAction(ByVal showmode As rmplugin_defs.RMShowMode)
    Dim strRet  As String

    strRet = InputBox("Write something:")
    If StrPtr(strRet) = 0 Then
        udeResult = rmres_canceled
    Else
        varValue = strRet
        udeResult = rmres_success
    End If

    clsHost.RaiseFinished Me
End Sub

Private Sub IRMPlugin_InitPlugin(host As rmplugin_defs.IRMPluginHost)
    Set clsHost = host
End Sub

Private Property Get IRMPlugin_Name() As String
    IRMPlugin_Name = "Return a text"
End Property

Private Property Get IRMPlugin_Result() As rmplugin_defs.RMResult
    IRMPlugin_Result = udeResult
End Property

Private Property Get IRMPlugin_Value() As Variant
    IRMPlugin_Value = varValue
End Property
