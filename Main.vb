Imports System
Imports System.Windows.Forms


Public Class Main

    Public Shared Sub Main()

        Dim f As New Form
        Dim list As New ListView

        AddHandler list.LostFocus, Sub(sender, e) f.Close()
        AddHandler list.KeyDown, Sub(sender, e) If e.KeyData = Keys.Escape Then f.Close()
        AddHandler f.Load,
            Sub(sender, e)

                Dim get_config_int = Function(name As String) CInt(System.Configuration.ConfigurationManager.AppSettings(name))

                f.FormBorderStyle = FormBorderStyle.None
                f.ImeMode = ImeMode.Disable

                list.Dock = DockStyle.Fill
                list.View = View.Details
                list.Columns.Add("")
                list.HeaderStyle = ColumnHeaderStyle.None
                list.FullRowSelect = True
                list.MultiSelect = False
                f.Controls.Add(list)

                Dim area = Screen.PrimaryScreen.WorkingArea

                f.Left = area.Left + get_config_int("Left")
                f.Top = area.Bottom - get_config_int("Bottom") - get_config_int("Height")
                f.Width = get_config_int("Width")
                f.Height = get_config_int("Height")

                For Each lnk In IO.Directory.GetFiles(".", "*.lnk")

                    list.Items.Add(IO.Path.GetFileNameWithoutExtension(lnk)).Tag = IO.Path.GetFullPath(lnk)
                Next
                list.AutoResizeColumns(ColumnHeaderAutoResizeStyle.ColumnContent)
            End Sub

        Dim select_item =
            Sub(path As String)

                If (Control.ModifierKeys And Keys.Shift) = Keys.Shift Then

                    Dim shell As New IWshRuntimeLibrary.WshShell
                    Dim x = shell.CreateShortcut(path)
                    If TypeOf x Is IWshRuntimeLibrary.WshShortcut Then

                        Dim shortcut = CType(x, IWshRuntimeLibrary.WshShortcut)
                        Diagnostics.Process.Start(shortcut.WorkingDirectory)
                    End If
                Else

                    Diagnostics.Process.Start(path)
                End If
            End Sub
        AddHandler list.DoubleClick, Sub(sender, e) select_item(list.SelectedItems(0).Tag.ToString)
        AddHandler list.KeyDown, Sub(sender, e) If (e.KeyData And Keys.KeyCode) = Keys.Enter Then select_item(list.SelectedItems(0).Tag.ToString)

        Application.Run(f)
    End Sub

End Class
