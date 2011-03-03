Imports System.ComponentModel
Imports System.Configuration.Install
Imports System.Runtime.InteropServices

Public Class Installer

    Public Sub New()
        MyBase.New()

        'This call is required by the Component Designer.
        InitializeComponent()
    End Sub

    Public Overrides Sub Install(ByVal stateSaver As System.Collections.IDictionary)
        Try
            MyBase.Install(stateSaver)
            Dim regsrv As New RegistrationServices()
            If Not (regsrv.RegisterAssembly(MyBase.GetType().Assembly, _
                AssemblyRegistrationFlags.SetCodeBase)) Then
                Throw New InstallException("Failed To Register for COM")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error during installation")
        End Try
    End Sub

    Public Overrides Sub Uninstall(ByVal savedState As System.Collections.IDictionary)
        Try
            MyBase.Uninstall(savedState)
            Dim regsrv As New RegistrationServices()
            If Not (regsrv.UnregisterAssembly(MyBase.GetType().Assembly)) Then
                Throw New InstallException("Failed To Unregister for COM")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error during unistallation")
        End Try
    End Sub

End Class
