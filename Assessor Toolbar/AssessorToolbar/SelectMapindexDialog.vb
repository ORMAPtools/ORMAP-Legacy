Imports System.Windows.Forms

Public Class SelectMapindexDialog

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click

        If Not Me.uxMapNumber.AutoCompleteCustomSource.Contains(uxMapNumber.Text) Then
            MessageBox.Show("Mapindex does not exist.  Please try again.", "Invalid MapIndex", MessageBoxButtons.OK, MessageBoxIcon.Error)
            uxMapNumber.Text = String.Empty
            uxMapNumber.Focus()
        Else
            Me.DialogResult = System.Windows.Forms.DialogResult.OK
            Me.Close()
        End If

    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Public ReadOnly Property MapNumber() As String
        Get
            Return uxMapNumber.Text()
        End Get
    End Property

    Private Sub uxMapNumber_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles uxMapNumber.TextChanged

        If Not uxMapNumber.AutoCompleteCustomSource.Contains(uxMapNumber.Text.Trim) Then
            If uxMapNumber.AutoCompleteCustomSource.Contains(uxMapNumber.Text.ToLower.Trim) Then
                uxMapNumber.Text = uxMapNumber.Text.ToLower.Trim
                uxMapNumber.SelectionStart = uxMapNumber.Text.Length
            End If
            If uxMapNumber.AutoCompleteCustomSource.Contains(uxMapNumber.Text.ToUpper.Trim) Then
                uxMapNumber.Text = uxMapNumber.Text.ToUpper.Trim
                uxMapNumber.SelectionStart = uxMapNumber.Text.Length
            End If
        End If

    End Sub

End Class
