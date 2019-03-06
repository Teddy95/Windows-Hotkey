Imports System.ComponentModel

Public Class Form1
    Dim send = False
    Dim interval = 5000 ' Standardmäßig 30 Sekunden
    Dim extraKey, mainKey, key As String

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ' send auf true setzen
        send = True

        ' Steuerelemente deaktivieren
        Button1.Enabled = False
        RadioButton1.Enabled = False
        RadioButton2.Enabled = False
        RadioButton3.Enabled = False
        RadioButton4.Enabled = False
        ComboBox1.Enabled = False
        NumericUpDown1.Enabled = False

        ' Stop-Button aktivieren
        Button2.Enabled = True

        ' Interval Eingabe prüfen und in Millisekunden umrechnen
        If NumericUpDown1.Value > 0 Then
            interval = NumericUpDown1.Value * 1000
        End If

        ' Zusatztaste festlegen
        If RadioButton1.Checked Then
            extraKey = ""
        End If
        If RadioButton2.Checked Then
            extraKey = "+"
        End If
        If RadioButton3.Checked Then
            extraKey = "^"
        End If
        If RadioButton4.Checked Then
            extraKey = "%"
        End If

        ' Haupttaste ermitteln
        If ComboBox1.SelectedItem <> "" Then
            mainKey = ComboBox1.SelectedItem
        Else
            MsgBox("Bitte Taste zum simulieren angeben!")
            Button2.PerformClick()
            Exit Sub
        End If

        ' SendKey String zusammenstellen
        key = extraKey + "{" + mainKey + "}"

        BackgroundWorker1.RunWorkerAsync()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        ' send auf false setzen
        send = False

        ' Steuerelemente aktivieren
        Button1.Enabled = True
        RadioButton1.Enabled = True
        RadioButton2.Enabled = True
        RadioButton3.Enabled = True
        RadioButton4.Enabled = True
        ComboBox1.Enabled = True
        NumericUpDown1.Enabled = True

        ' Stop-Button deaktivieren
        Button2.Enabled = False
    End Sub

    Private Sub BackgroundWorker1_DoWork(sender As Object, e As DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        System.Threading.Thread.Sleep(10000) ' 10 Sekunden warten, bevor der erste Tastenbefehl simuliert wird

        While send
            SendKeys.SendWait(key)
            System.Threading.Thread.Sleep(interval) ' Taste simulieren im festgelegten Zeitinterval
        End While
    End Sub
End Class
