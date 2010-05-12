Imports System.Windows.Forms

Public Class FilterAnnoScaleForm

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click

        If Not Me.uxMapScale.AutoCompleteCustomSource.Contains(uxMapScale.Text) Then
            MessageBox.Show("Map Scale does not exist.  Please try again.", "Invalid Map Scale", MessageBoxButtons.OK, MessageBoxIcon.Error)
            uxMapScale.Text = String.Empty
            uxMapScale.Focus()
        Else
            Me.DialogResult = System.Windows.Forms.DialogResult.OK
            Me.Close()
        End If

    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub FilterAnnoScaleForm_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim mapScaleStringList() As String = New String() {"10", "20", "30", "40", "50", "60", "100", "200", "400", "2000"}
        uxMapScale.AutoCompleteCustomSource.Clear()
        uxMapScale.AutoCompleteCustomSource.AddRange(mapScaleStringList)
    End Sub

    Private Sub uxMapScale_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles uxMapScale.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            OK_Button.PerformClick()
        End If
    End Sub
End Class