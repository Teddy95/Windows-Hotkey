Imports System.ComponentModel

Public Class Form1
    Dim send = False

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        send = True
        Button1.Enabled = False
        Button2.Enabled = True
        BackgroundWorker1.RunWorkerAsync()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        send = False
        Button1.Enabled = True
        Button2.Enabled = False
    End Sub

    Private Sub BackgroundWorker1_DoWork(sender As Object, e As DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        System.Threading.Thread.Sleep(10000) ' 10 Sekunden warten, bevor der erste Tastenbefehl simuliert wird

        While send
            SendKeys.SendWait("{TAB}")
            System.Threading.Thread.Sleep(30000) ' Alle 30 Sekunden Taste simulieren
        End While
    End Sub
End Class
