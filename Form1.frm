VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Plugin Host"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   315
      Left            =   383
      TabIndex        =   2
      Top             =   2550
      Width           =   1440
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Execute"
      Height          =   315
      Left            =   2858
      TabIndex        =   1
      Top             =   2550
      Width           =   1440
   End
   Begin VB.ListBox lstPlugins 
      Height          =   2010
      Left            =   233
      TabIndex        =   0
      Top             =   300
      Width           =   4215
   End
   Begin VB.Label Label1 
      Caption         =   "Plugins:"
      Height          =   240
      Left            =   225
      TabIndex        =   3
      Top             =   75
      Width           =   840
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements rmplugin_defs.IRMPluginHost

Private clsPluginLoader As PlugInLoader

Private Sub cmdLoad_Click()
    Dim clsPlugin   As rmplugin_defs.IRMPlugin

    Set clsPlugin = clsPluginLoader.CreatePlugin(lstPlugins.ListIndex)
    If clsPlugin Is Nothing Then
        MsgBox "Could not load the Plugin!", vbExclamation
        Exit Sub
    End If

    ' reference to us, so the plugin can inform us about its state
    clsPlugin.InitPlugin Me

    ' execute the plugin
    clsPlugin.DoAction rmshow_modal
End Sub

Private Sub cmdRefresh_Click()
    Dim i           As Long
    Dim clsPlugin   As rmplugin_defs.IRMPlugin

    clsPluginLoader.FindPlugins

    For i = 0 To clsPluginLoader.PluginCount - 1
        Set clsPlugin = clsPluginLoader.CreatePlugin(i)
        lstPlugins.AddItem clsPlugin.Name & " - " & clsPlugin.Description
    Next
End Sub

Private Sub Form_Load()
    Set clsPluginLoader = New PlugInLoader

    ' the interface the plugins have to implement
    Set clsPluginLoader.Interface = New rmplugin_defs.IRMPlugin
End Sub

''''''''''''''''''''
''' Plugin Host
''''''''''''''''''''

Private Property Get IRMPluginHost_OwnerFormHandle() As Long
    IRMPluginHost_OwnerFormHandle = Me.hWnd
End Property

Private Sub IRMPluginHost_RaiseFinished(plugin As rmplugin_defs.IRMPlugin)
    If plugin.Result = rmres_success Then
        MsgBox "Plugin " & plugin.Name & " is finished." & vbCrLf & _
               "Result: " & plugin.Value, _
               vbInformation
    Else
        MsgBox "Plugin " & plugin.Name & " was aborted.", _
               vbExclamation
    End If
End Sub
