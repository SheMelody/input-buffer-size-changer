Public Class Form1
    Dim ServicesKey As Microsoft.Win32.RegistryKey

    Dim KeyboardLatencies As New Dictionary(Of Integer, Integer) From {
            {0, 100}, {1, 100}, {2, 58}, {3, 16}, {4, 16}
        }
    Dim MouseLatencies As New Dictionary(Of Integer, Integer) From {
            {1, 100}, {2, 58}, {3, 20}, {4, 16}
        }
    Function GetCurrentBuffers(ByVal ServicesKey As Microsoft.Win32.RegistryKey) As Integer
        Dim mouclass As Microsoft.Win32.RegistryKey
        Dim kbdclass_parameters As Microsoft.Win32.RegistryKey
        Dim mouclass_parameters As Microsoft.Win32.RegistryKey
        Try
            mouclass = ServicesKey.OpenSubKey("mouclass")
            kbdclass_parameters = ServicesKey.OpenSubKey("kbdclass\Parameters")
            Dim KeyboardBuffers As Integer = kbdclass_parameters.GetValue("KeyboardDataQueueSize")
            If (Not mouclass.GetSubKeyNames().Contains("Parameters")) And KeyboardBuffers = 100 Then
                mouclass.Dispose()
                kbdclass_parameters.Dispose()
                Return 0
            End If
            mouclass_parameters = mouclass.OpenSubKey("Parameters")
            Dim MouseBuffers As Integer = mouclass_parameters.GetValue("MouseDataQueueSize")
            mouclass_parameters.Dispose()
            mouclass.Dispose()
            kbdclass_parameters.Dispose()
            If MouseBuffers = 58 And KeyboardBuffers = 58 Then
                Return 2
            End If
            If MouseBuffers = 100 And KeyboardBuffers = 100 Then
                Return 1
            End If
            If MouseBuffers = 20 And KeyboardBuffers = 16 Then
                Return 3
            End If
            If MouseBuffers = 16 And KeyboardBuffers = 16 Then
                Return 4
            End If
            Return -1
        Catch ex As Exception
            Try
                mouclass_parameters.Dispose()
            Catch ex2 As Exception
            End Try
            Try
                mouclass.Dispose()
            Catch ex2 As Exception
            End Try
            Try
                kbdclass_parameters.Dispose()
            Catch ex2 As Exception
            End Try
            Return -1
        End Try

    End Function
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ServicesKey = My.Computer.Registry.LocalMachine.OpenSubKey("SYSTEM\CurrentControlSet\Services")
        Dim CurrentConfiguration = GetCurrentBuffers(ServicesKey)
        If CurrentConfiguration = -1 Then
            MsgBox("An unknown configuration for the buffers was detected, the configuration shown in the GUI doesn't equal to your current buffers configuration.", MsgBoxStyle.Information, "Info")
            ComboBox1.SelectedIndex = 0
        Else
            ComboBox1.SelectedIndex = CurrentConfiguration
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim kbdclass As Microsoft.Win32.RegistryKey
        Dim kbdclass_parameters As Microsoft.Win32.RegistryKey
        Dim mouclass As Microsoft.Win32.RegistryKey
        Dim mouclass_parameters As Microsoft.Win32.RegistryKey
        Try
            kbdclass = ServicesKey.OpenSubKey("kbdclass")
            If Not kbdclass.GetSubKeyNames().Contains("Parameters") Then
                kbdclass_parameters = kbdclass.CreateSubKey("Parameters", True)
            Else
                kbdclass_parameters = kbdclass.OpenSubKey("Parameters", True)
            End If
            kbdclass_parameters.SetValue("KeyboardDataQueueSize", KeyboardLatencies(ComboBox1.SelectedIndex), Microsoft.Win32.RegistryValueKind.DWord)
            kbdclass_parameters.Dispose()
            kbdclass.Dispose()
            mouclass = ServicesKey.OpenSubKey("mouclass", True)
            If ComboBox1.SelectedIndex = 0 Then
                If mouclass.GetSubKeyNames().Contains("Parameters") Then
                    mouclass.DeleteSubKey("Parameters")
                End If
            Else
                If Not mouclass.GetSubKeyNames().Contains("Parameters") Then
                    mouclass_parameters = mouclass.CreateSubKey("Parameters", True)
                Else
                    mouclass_parameters = mouclass.OpenSubKey("Parameters", True)
                End If
                mouclass_parameters.SetValue("MouseDataQueueSize", MouseLatencies(ComboBox1.SelectedIndex))
                mouclass_parameters.Dispose()
            End If
            mouclass.Dispose()
            MsgBox("Your input buffers have been set to " & ComboBox1.SelectedItem.ToString() & ". In order to apply this configuration, you need to reboot the system.", MsgBoxStyle.Information, "Done")
        Catch ex As Exception
            Try
                mouclass_parameters.Dispose()
            Catch ex2 As Exception
            End Try
            Try
                mouclass.Dispose()
            Catch ex2 As Exception
            End Try
            Try
                kbdclass_parameters.Dispose()
            Catch ex2 As Exception
            End Try
            Try
                kbdclass.Dispose()
            Catch ex2 As Exception
            End Try
            MsgBox("The following error has occurred: " & ex.Message, MsgBoxStyle.Critical, "Error")
        End Try

    End Sub
End Class
