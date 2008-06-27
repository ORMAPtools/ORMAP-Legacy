Imports System.ComponentModel
Imports System.Configuration.Install
Imports System.Runtime.InteropServices

Public Class InstallationConfiguration

    Public Sub New()
        MyBase.New()

        'This call is required by the Component Designer.
        InitializeComponent()


        'Add initialization code after the call to InitializeComponent

    End Sub

    Public Overrides Sub Install(ByVal stateSaver As System.Collections.IDictionary)
        Dim pRegSvr As New RegistrationServices

        Try
            MyBase.Install(stateSaver)

            If Not pRegSvr.RegisterAssembly(MyBase.GetType().Assembly, AssemblyRegistrationFlags.SetCodeBase) Then
                Throw New InstallException("COM registration failed.  Some or all of the application classes are not properly registered in the ESRI component categories.")
            End If
        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, "Install Error", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Error)
        End Try

    End Sub

    Public Overrides Sub Uninstall(ByVal savedState As System.Collections.IDictionary)
        Dim pRegSvr As New RegistrationServices

        Try
            MyBase.Uninstall(savedState)

            If Not pRegSvr.UnregisterAssembly(MyBase.GetType().Assembly) Then
                Throw New InstallException("COM unregistration failed.  Some or all of the application classes were not properly removed from the ESRI component categories.")
            End If
        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message, "Uninstall Error", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Error)
        End Try

    End Sub

End Class
