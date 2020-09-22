VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "rmplugin_calculator"
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
    IRMPlugin_Description = "Sums up values"
End Property

Private Sub IRMPlugin_DoAction(ByVal showmode As rmplugin_defs.RMShowMode)
    frmCalc.Show showmode

    If frmCalc.blnCancel Then
        udeResult = rmres_canceled
    Else
        varValue = Val(frmCalc.txtVal1.Text) + Val(frmCalc.txtVal2.Text)
        udeResult = rmres_success
    End If

    ' tell the host we're done
    clsHost.RaiseFinished Me
End Sub

Private Sub IRMPlugin_InitPlugin(host As rmplugin_defs.IRMPluginHost)
    Set clsHost = host
End Sub

Private Property Get IRMPlugin_Name() As String
    IRMPlugin_Name = "Calculator"
End Property

Private Property Get IRMPlugin_Result() As rmplugin_defs.RMResult
    IRMPlugin_Result = udeResult
End Property

Private Property Get IRMPlugin_Value() As Variant
    IRMPlugin_Value = varValue
End Property
