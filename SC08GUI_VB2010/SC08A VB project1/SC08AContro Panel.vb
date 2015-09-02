﻿'**************************************************************************
'********* SC08A 8 or 16 Servco Motor Controller GUI- *********************
'***************** SC08A SERVO COMTROL PANEL ******************************
'************************ Date:8/6/2011 ***********************************
'************ Visual Basic Expression 2008 Edition ************************
'*********************** www.cytron.com.my ********************************
'**************************************************************************
'command to activate all servo Hexa:C0,1-activate 0-deactivate
'command to activate individual 1-16servo Hexa:C1-D0,1-activate 0-deactivate
'command to request position 1-16 servo Hexa:A1-B0
'command to starting 1-16servo position Hexa:81-90
'
Public Class SC08A_VB
    'define com port 
    Dim WithEvents ComPort As New IO.Ports.SerialPort
    'define variable
    Dim servo1 As Integer
    Dim servo2 As Integer
    Dim servo3 As Integer
    Dim servo4 As Integer
    Dim servo5 As Integer
    Dim servo6 As Integer
    Dim servo7 As Integer
    Dim servo8 As Integer
    Dim servo9 As Integer
    Dim servo10 As Integer
    Dim servo11 As Integer
    Dim servo12 As Integer
    Dim servo13 As Integer
    Dim servo14 As Integer
    Dim servo15 As Integer
    Dim servo16 As Integer

    Dim data1 As Byte       'higher byte
    Dim data2 As Byte       'lower byte
    Dim data3 As Byte       'ramp/speed value
    Dim transmitt As Integer

    Dim higher_data As Byte 'receive byte1
    Dim lower_data As Byte 'receive byte2
    Dim position As Integer 'receive position

    'for closing the panel
    Private Sub SC08A_VB_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Try
            ComPort.Close()
            transmitt = 0
        Catch ex As Exception

        End Try
    End Sub
    'comport auto searching when panel loading
    Private Sub SC08A_VB_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        For i As Integer = 0 To _
          My.Computer.Ports.SerialPortNames.Count - 1
            ComPortBox.Items.Add( _
              My.Computer.Ports.SerialPortNames(i))
        Next
        Button_Cancel.Enabled = False
        Active8servo.Enabled = False
        Active1.Enabled = False
        Button_Reset.Enabled = False
    End Sub

    'connect Com Port 
    Private Sub Button_Connect_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Connect.Click
        Button_Search.Enabled = False

        If ComPort.IsOpen Then
            ComPort.Close()
        End If
        Try
            With ComPort
                .PortName = ComPortBox.Text
                .BaudRate = 9600 '115200
                .Parity = IO.Ports.Parity.None
                .DataBits = 8
                .StopBits = IO.Ports.StopBits.One
                .ReadTimeout = 1000
                .WriteTimeout = 1000
                ' .Encoding = System.Text.Encoding.Unicode
            End With
            BaudRate_msg.Text = ComPort.BaudRate

            ComPort.Open()
            ComPortDisp.Text = ComPortBox.Text & " connected."
            Button_Connect.Enabled = False
            Button_Cancel.Enabled = True
            transmitt = 1
        Catch ex As Exception
            'MsgBox(ex.ToString)
            MsgBox("Please select COM PORTS")

        End Try
        Active8servo.Enabled = True
        Active1.Enabled = True
        Button_Reset.Enabled = True
    End Sub

    'Cancel Com Port connection
    Private Sub Button_Cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Cancel.Click
        Button_Reset.Enabled = False
        Button_Search.Enabled = True
        Active8servo.Enabled = False
        Active1.Enabled = False
        Try
            ComPort.Close()
            ComPortDisp.Text = ComPort.PortName & " disconnected."
            Button_Connect.Enabled = True
            Button_Cancel.Enabled = False
            transmitt = 0
        Catch ex As Exception
            'MsgBox(ex.ToString)
            ComPortDisp.Text = ComPort.PortName & " disconnected."
            Button_Connect.Enabled = True
            Button_Cancel.Enabled = False

        End Try

    End Sub


    Private Sub Button_Search_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Search.Click
        ComPortBox.Items.Clear()
        For i As Integer = 0 To _
          My.Computer.Ports.SerialPortNames.Count - 1
            ComPortBox.Items.Add( _
               My.Computer.Ports.SerialPortNames(i))
        Next
        Button_Cancel.Enabled = False
    End Sub
    'activate all servo motor
    Private Sub Active1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Active1.Click
        Button_MarkAll.Enabled = True
        If ProgressBar3.Value <> ProgressBar3.Maximum Then
            Active8servo.Enabled = False
            If transmitt = 1 Then
                Dim send() As Byte = {&HC0, &H1} 'activate all servo HC0=1byte,H1=2nd byte

                Try
                    ComPort.Write(send, 0, send.Length) 'transmitt data
                    Active1.Enabled = False

                    Active16_1.Enabled = False
                    Active16_2.Enabled = False
                    Active16_3.Enabled = False
                    Active16_4.Enabled = False
                    Active16_5.Enabled = False
                    Active16_6.Enabled = False
                    Active16_7.Enabled = False
                    Active16_8.Enabled = False
                    Active16_9.Enabled = False
                    Active16_10.Enabled = False
                    Active16_11.Enabled = False
                    Active16_12.Enabled = False
                    Active16_13.Enabled = False
                    Active16_14.Enabled = False
                    Active16_15.Enabled = False
                    Active16_16.Enabled = False

                    Deactive16_1.Enabled = True
                    Deactive16_2.Enabled = True
                    Deactive16_3.Enabled = True
                    Deactive16_4.Enabled = True
                    Deactive16_5.Enabled = True
                    Deactive16_6.Enabled = True
                    Deactive16_7.Enabled = True
                    Deactive16_8.Enabled = True
                    Deactive16_9.Enabled = True
                    Deactive16_10.Enabled = True
                    Deactive16_11.Enabled = True
                    Deactive16_12.Enabled = True
                    Deactive16_13.Enabled = True
                    Deactive16_14.Enabled = True
                    Deactive16_15.Enabled = True
                    Deactive16_16.Enabled = True

                Catch ex As Exception
                    'MsgBox(ex.ToString)
                    MsgBox("send fail")


                End Try

            End If
            ProgressBar3.Increment(5)
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

            Dim send21() As Byte = {&HA1} 'reporting position command
            If transmitt = 1 Then
                Try
                    ComPort.Write(send21, 0, send21.Length) 'transmitt reporting position command
                Catch ex As Exception
                    MsgBox("sendreport.fail1")

                End Try

            End If
            Try
                System.Threading.Thread.Sleep(300)
                higher_data = ComPort.ReadByte
                position = (higher_data And &H7F) << 6 'receive 1st byte data from SC08A


                System.Threading.Thread.Sleep(300)
                lower_data = ComPort.ReadByte 'receive 2nd byte data from SC08A
                position = position Or (lower_data And &H3F)

                servo16bar1.Value = position
                value16_1.Text = position
                Pulse16_1.Text = (position * 0.25 * 10 ^ -3) + 0.5


            Catch ex As Exception
                MsgBox("Rc_report.fail1")
            End Try

            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

            Dim send22() As Byte = {&HA2}
            If transmitt = 1 Then
                Try
                    ComPort.Write(send22, 0, send22.Length)
                Catch ex As Exception
                    MsgBox("sendreport.fail2")

                End Try

            End If
            Try
                System.Threading.Thread.Sleep(300)
                higher_data = ComPort.ReadByte
                position = (higher_data And &H7F) << 6


                System.Threading.Thread.Sleep(300)
                lower_data = ComPort.ReadByte
                position = position Or (lower_data And &H3F)

                servo16bar2.Value = position
                value16_2.Text = position
                Pulse16_2.Text = (position * 0.25 * 10 ^ -3) + 0.5

            Catch ex As Exception
                MsgBox("Rc_report.fail2")
            End Try
            ProgressBar3.Increment(5)
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim send23() As Byte = {&HA3}
            If transmitt = 1 Then
                Try
                    ComPort.Write(send23, 0, send23.Length)
                Catch ex As Exception
                    MsgBox("sendreport.fail3")

                End Try

            End If
            Try
                System.Threading.Thread.Sleep(300)
                higher_data = ComPort.ReadByte
                position = (higher_data And &H7F) << 6


                System.Threading.Thread.Sleep(300)
                lower_data = ComPort.ReadByte
                position = position Or (lower_data And &H3F)

                servo16bar3.Value = position
                value16_3.Text = position
                Pulse16_3.Text = (position * 0.25 * 10 ^ -3) + 0.5


            Catch ex As Exception
                MsgBox("Rc_report.fail3")
            End Try
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim send24() As Byte = {&HA4}
            If transmitt = 1 Then
                Try
                    ComPort.Write(send24, 0, send24.Length)
                Catch ex As Exception
                    MsgBox("sendreport.fail4")

                End Try

            End If
            Try
                System.Threading.Thread.Sleep(300)
                higher_data = ComPort.ReadByte
                position = (higher_data And &H7F) << 6


                System.Threading.Thread.Sleep(300)
                lower_data = ComPort.ReadByte
                position = position Or (lower_data And &H3F)

                servo16bar4.Value = position
                value16_4.Text = position
                Pulse16_4.Text = (position * 0.25 * 10 ^ -3) + 0.5


            Catch ex As Exception
                MsgBox("Rc_report.fail4")
            End Try
            ProgressBar3.Increment(5)
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim send25() As Byte = {&HA5}
            If transmitt = 1 Then
                Try
                    ComPort.Write(send25, 0, send25.Length)
                Catch ex As Exception
                    MsgBox("sendreport.fail5")

                End Try

            End If
            Try
                System.Threading.Thread.Sleep(300)
                higher_data = ComPort.ReadByte
                position = (higher_data And &H7F) << 6


                System.Threading.Thread.Sleep(300)
                lower_data = ComPort.ReadByte
                position = position Or (lower_data And &H3F)

                servo16bar5.Value = position
                value16_5.Text = position
                Pulse16_5.Text = (position * 0.25 * 10 ^ -3) + 0.5


            Catch ex As Exception
                MsgBox("Rc_report.fail5")
            End Try
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim send26() As Byte = {&HA6}
            If transmitt = 1 Then
                Try
                    ComPort.Write(send26, 0, send26.Length)
                Catch ex As Exception
                    MsgBox("sendreport.fail6")

                End Try

            End If
            Try
                System.Threading.Thread.Sleep(300)
                higher_data = ComPort.ReadByte
                position = (higher_data And &H7F) << 6


                System.Threading.Thread.Sleep(300)
                lower_data = ComPort.ReadByte
                position = position Or (lower_data And &H3F)

                servo16bar6.Value = position
                value16_6.Text = position
                Pulse16_6.Text = (position * 0.25 * 10 ^ -3) + 0.5


            Catch ex As Exception
                MsgBox("Rc_report.fail6")
            End Try
            ProgressBar3.Increment(5)
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

            Dim send27() As Byte = {&HA7}
            If transmitt = 1 Then
                Try
                    ComPort.Write(send27, 0, send27.Length)
                Catch ex As Exception
                    MsgBox("sendreport.fail7")

                End Try

            End If
            Try
                System.Threading.Thread.Sleep(300)
                higher_data = ComPort.ReadByte
                position = (higher_data And &H7F) << 6


                System.Threading.Thread.Sleep(300)
                lower_data = ComPort.ReadByte
                position = position Or (lower_data And &H3F)

                servo16bar7.Value = position
                value16_7.Text = position
                Pulse16_7.Text = (position * 0.25 * 10 ^ -3) + 0.5


            Catch ex As Exception
                MsgBox("Rc_report.fail7")
            End Try
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

            Dim send28() As Byte = {&HA8}
            If transmitt = 1 Then
                Try
                    ComPort.Write(send28, 0, send28.Length)
                Catch ex As Exception
                    MsgBox("sendreport.fail8")

                End Try

            End If
            Try
                System.Threading.Thread.Sleep(300)
                higher_data = ComPort.ReadByte
                position = (higher_data And &H7F) << 6


                System.Threading.Thread.Sleep(300)
                lower_data = ComPort.ReadByte
                position = position Or (lower_data And &H3F)

                servo16bar8.Value = position
                value16_8.Text = position
                Pulse16_8.Text = (position * 0.25 * 10 ^ -3) + 0.5


            Catch ex As Exception
                MsgBox("Rc_report.fail8")
            End Try
            ProgressBar3.Increment(5)
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim send29() As Byte = {&HA9}
            If transmitt = 1 Then
                Try
                    ComPort.Write(send29, 0, send29.Length)
                Catch ex As Exception
                    MsgBox("sendreport.fail9")

                End Try

            End If
            Try
                System.Threading.Thread.Sleep(300)
                higher_data = ComPort.ReadByte
                position = (higher_data And &H7F) << 6


                System.Threading.Thread.Sleep(300)
                lower_data = ComPort.ReadByte
                position = position Or (lower_data And &H3F)

                servo16bar9.Value = position
                value16_9.Text = position
                Pulse16_9.Text = (position * 0.25 * 10 ^ -3) + 0.5


            Catch ex As Exception
                MsgBox("Rc_report.fail9")
            End Try
            ProgressBar3.Increment(5)
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

            Dim send30() As Byte = {&HAA}
            If transmitt = 1 Then
                Try
                    ComPort.Write(send30, 0, send30.Length)
                Catch ex As Exception
                    MsgBox("sendreport.fail10")

                End Try

            End If
            Try
                System.Threading.Thread.Sleep(300)
                higher_data = ComPort.ReadByte
                position = (higher_data And &H7F) << 6


                System.Threading.Thread.Sleep(300)
                lower_data = ComPort.ReadByte
                position = position Or (lower_data And &H3F)

                servo16bar10.Value = position
                value16_10.Text = position
                Pulse16_10.Text = (position * 0.25 * 10 ^ -3) + 0.5


            Catch ex As Exception
                MsgBox("Rc_report.fail10")
            End Try

            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

            Dim send31() As Byte = {&HAB}
            If transmitt = 1 Then
                Try
                    ComPort.Write(send31, 0, send31.Length)
                Catch ex As Exception
                    MsgBox("sendreport.fail11")

                End Try

            End If
            Try
                System.Threading.Thread.Sleep(300)
                higher_data = ComPort.ReadByte
                position = (higher_data And &H7F) << 6


                System.Threading.Thread.Sleep(300)
                lower_data = ComPort.ReadByte
                position = position Or (lower_data And &H3F)

                servo16bar11.Value = position
                value16_11.Text = position
                Pulse16_11.Text = (position * 0.25 * 10 ^ -3) + 0.5


            Catch ex As Exception
                MsgBox("Rc_report.fail11")
            End Try
            ProgressBar3.Increment(5)
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim send32() As Byte = {&HAC}
            If transmitt = 1 Then
                Try
                    ComPort.Write(send32, 0, send32.Length)
                Catch ex As Exception
                    MsgBox("sendreport.fail12")

                End Try

            End If
            Try
                System.Threading.Thread.Sleep(300)
                higher_data = ComPort.ReadByte
                position = (higher_data And &H7F) << 6


                System.Threading.Thread.Sleep(300)
                lower_data = ComPort.ReadByte
                position = position Or (lower_data And &H3F)

                servo16bar12.Value = position
                value16_12.Text = position
                Pulse16_12.Text = (position * 0.25 * 10 ^ -3) + 0.5


            Catch ex As Exception
                MsgBox("Rc_report.fail12")
            End Try
            ProgressBar3.Increment(5)
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

            Dim send33() As Byte = {&HAD}
            If transmitt = 1 Then
                Try
                    ComPort.Write(send33, 0, send33.Length)
                Catch ex As Exception
                    MsgBox("sendreport.fail13")

                End Try

            End If
            Try
                System.Threading.Thread.Sleep(300)
                higher_data = ComPort.ReadByte
                position = (higher_data And &H7F) << 6


                System.Threading.Thread.Sleep(300)
                lower_data = ComPort.ReadByte
                position = position Or (lower_data And &H3F)

                servo16bar13.Value = position
                value16_13.Text = position
                Pulse16_13.Text = (position * 0.25 * 10 ^ -3) + 0.5


            Catch ex As Exception
                MsgBox("Rc_report.fail13")
            End Try
            ProgressBar3.Increment(5)

            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

            Dim send34() As Byte = {&HAE}
            If transmitt = 1 Then
                Try
                    ComPort.Write(send34, 0, send34.Length)
                Catch ex As Exception
                    MsgBox("sendreport.fail14")

                End Try

            End If
            Try
                System.Threading.Thread.Sleep(300)
                higher_data = ComPort.ReadByte
                position = (higher_data And &H7F) << 6


                System.Threading.Thread.Sleep(300)
                lower_data = ComPort.ReadByte
                position = position Or (lower_data And &H3F)

                servo16bar14.Value = position
                value16_14.Text = position
                Pulse16_14.Text = (position * 0.25 * 10 ^ -3) + 0.5


            Catch ex As Exception
                MsgBox("Rc_report.fail14")
            End Try
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

            Dim send35() As Byte = {&HAF}
            If transmitt = 1 Then
                Try
                    ComPort.Write(send35, 0, send35.Length)
                Catch ex As Exception
                    MsgBox("sendreport.fail15")

                End Try

            End If
            Try
                System.Threading.Thread.Sleep(300)
                higher_data = ComPort.ReadByte
                position = (higher_data And &H7F) << 6


                System.Threading.Thread.Sleep(300)
                lower_data = ComPort.ReadByte
                position = position Or (lower_data And &H3F)

                servo16bar15.Value = position
                value16_15.Text = position
                Pulse16_15.Text = (position * 0.25 * 10 ^ -3) + 0.5


            Catch ex As Exception
                MsgBox("Rc_report.fail15")
            End Try
            ProgressBar3.Increment(5)
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

            Dim send36() As Byte = {&HB0}
            If transmitt = 1 Then
                Try
                    ComPort.Write(send36, 0, send36.Length)
                Catch ex As Exception
                    MsgBox("sendreport.fail16")

                End Try

            End If
            Try
                System.Threading.Thread.Sleep(300)
                higher_data = ComPort.ReadByte
                position = (higher_data And &H7F) << 6


                System.Threading.Thread.Sleep(300)
                lower_data = ComPort.ReadByte
                position = position Or (lower_data And &H3F)

                servo16bar16.Value = position
                value16_16.Text = position
                Pulse16_16.Text = (position * 0.25 * 10 ^ -3) + 0.5


            Catch ex As Exception
                MsgBox("Rc_report.fail16")
            End Try
            ProgressBar3.Increment(5)
        End If

        MsgBox("Sucessful!!")
        ProgressBar3.Value = 0 ' progress bar counting finish
        Timer2.Stop() 'timer in progress bar stop

    End Sub


    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        System.Diagnostics.Process.Start("www.cytron.com.my") 'website link
    End Sub

    '*********************************************
    '******** 16 servo motor operation code ******
    '*********************************************
    'servo motor 1
    Private Sub servo16bar1_Scroll(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ScrollEventArgs) Handles servo16bar1.Scroll

        value16_1.Text = servo16bar1.Value
        servo1 = servo16bar1.Value

        data1 = (servo1 >> 6) And &HFF
        data2 = servo1 And &H3F
        data3 = speed16_1.Value
        Pulse16_1.Text = (servo16bar1.Value * 0.25 * 10 ^ -3) + 0.5

        Dim send() As Byte = {&HE1, data1, data2, data3}
        If transmitt = 1 Then
            Try
                ComPort.Write(send, 0, send.Length)

            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

        End If
    End Sub
    'servo motor 2
    Private Sub servo16bar2_Scroll(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ScrollEventArgs) Handles servo16bar2.Scroll
        value16_2.Text = servo16bar2.Value
        servo2 = servo16bar2.Value

        data1 = (servo2 >> 6) And &HFF
        data2 = servo2 And &H3F
        data3 = speed16_2.Value
        Pulse16_2.Text = (servo16bar2.Value * 0.25 * 10 ^ -3) + 0.5

        Dim send() As Byte = {&HE2, data1, data2, data3}
        If transmitt = 1 Then
            Try
                ComPort.Write(send, 0, send.Length)

            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

        End If
    End Sub
    'servo motor 3
    Private Sub servo16bar3_Scroll(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ScrollEventArgs) Handles servo16bar3.Scroll
        value16_3.Text = servo16bar3.Value
        servo3 = servo16bar3.Value

        data1 = (servo3 >> 6) And &HFF
        data2 = servo1 And &H3F
        data3 = speed16_3.Value
        Pulse16_3.Text = (servo16bar3.Value * 0.25 * 10 ^ -3) + 0.5

        Dim send() As Byte = {&HE3, data1, data2, data3}
        If transmitt = 1 Then
            Try
                ComPort.Write(send, 0, send.Length)

            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

        End If
    End Sub
    'servo motor 4
    Private Sub servo16bar4_Scroll(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ScrollEventArgs) Handles servo16bar4.Scroll
        value16_4.Text = servo16bar4.Value
        servo4 = servo16bar4.Value

        data1 = (servo4 >> 6) And &HFF
        data2 = servo4 And &H3F
        data3 = speed16_4.Value
        Pulse16_4.Text = (servo16bar4.Value * 0.25 * 10 ^ -3) + 0.5

        Dim send() As Byte = {&HE4, data1, data2, data3}
        If transmitt = 1 Then
            Try
                ComPort.Write(send, 0, send.Length)

            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

        End If
    End Sub
    'servo motor 5
    Private Sub servo16bar5_Scroll(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ScrollEventArgs) Handles servo16bar5.Scroll
        value16_5.Text = servo16bar5.Value
        servo5 = servo16bar5.Value

        data1 = (servo5 >> 6) And &HFF
        data2 = servo5 And &H3F
        data3 = speed16_5.Value
        Pulse16_5.Text = (servo16bar5.Value * 0.25 * 10 ^ -3) + 0.5

        Dim send() As Byte = {&HE5, data1, data2, data3}
        If transmitt = 1 Then
            Try
                ComPort.Write(send, 0, send.Length)

            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

        End If
    End Sub
    'servo motor 6
    Private Sub servo16bar6_Scroll(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ScrollEventArgs) Handles servo16bar6.Scroll
        value16_6.Text = servo16bar6.Value
        servo6 = servo16bar6.Value

        data1 = (servo6 >> 6) And &HFF
        data2 = servo6 And &H3F
        data3 = speed16_6.Value
        Pulse16_6.Text = (servo16bar6.Value * 0.25 * 10 ^ -3) + 0.5

        Dim send() As Byte = {&HE6, data1, data2, data3}
        If transmitt = 1 Then
            Try
                ComPort.Write(send, 0, send.Length)

            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

        End If
    End Sub
    'servo motor 7
    Private Sub servo16bar7_Scroll(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ScrollEventArgs) Handles servo16bar7.Scroll
        value16_7.Text = servo16bar7.Value
        servo7 = servo16bar7.Value

        data1 = (servo7 >> 6) And &HFF
        data2 = servo7 And &H3F
        data3 = speed16_7.Value
        Pulse16_7.Text = (servo16bar7.Value * 0.25 * 10 ^ -3) + 0.5

        Dim send() As Byte = {&HE7, data1, data2, data3}
        If transmitt = 1 Then
            Try
                ComPort.Write(send, 0, send.Length)

            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

        End If
    End Sub
    'servo motor 8
    Private Sub servo16bar8_Scroll(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ScrollEventArgs) Handles servo16bar8.Scroll
        value16_8.Text = servo16bar8.Value
        servo8 = servo16bar8.Value

        data1 = (servo8 >> 6) And &HFF
        data2 = servo8 And &H3F
        data3 = speed16_8.Value
        Pulse16_8.Text = (servo16bar8.Value * 0.25 * 10 ^ -3) + 0.5

        Dim send() As Byte = {&HE8, data1, data2, data3}
        If transmitt = 1 Then
            Try
                ComPort.Write(send, 0, send.Length)

            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

        End If
    End Sub
    ' servo motor 9
    Private Sub servo16bar9_Scroll(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ScrollEventArgs) Handles servo16bar9.Scroll
        value16_9.Text = servo16bar9.Value
        servo9 = servo16bar9.Value

        data1 = (servo9 >> 6) And &HFF
        data2 = servo9 And &H3F
        data3 = speed16_9.Value
        Pulse16_9.Text = (servo16bar9.Value * 0.25 * 10 ^ -3) + 0.5

        Dim send() As Byte = {&HE9, data1, data2, data3}
        If transmitt = 1 Then
            Try
                ComPort.Write(send, 0, send.Length)

            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

        End If
    End Sub
    'servo motor 10
    Private Sub servo16bar10_Scroll(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ScrollEventArgs) Handles servo16bar10.Scroll
        value16_10.Text = servo16bar10.Value
        servo10 = servo16bar10.Value

        data1 = (servo10 >> 6) And &HFF
        data2 = servo10 And &H3F
        data3 = speed16_10.Value
        Pulse16_10.Text = (servo16bar10.Value * 0.25 * 10 ^ -3) + 0.5

        Dim send() As Byte = {&HEA, data1, data2, data3}
        If transmitt = 1 Then
            Try
                ComPort.Write(send, 0, send.Length)

            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

        End If
    End Sub
    'servo motor 16
    Private Sub servo16bar11_Scroll(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ScrollEventArgs) Handles servo16bar11.Scroll
        value16_11.Text = servo16bar11.Value
        servo11 = servo16bar11.Value

        data1 = (servo11 >> 6) And &HFF
        data2 = servo11 And &H3F
        data3 = speed16_11.Value
        Pulse16_11.Text = (servo16bar11.Value * 0.25 * 10 ^ -3) + 0.5

        Dim send() As Byte = {&HEB, data1, data2, data3}
        If transmitt = 1 Then
            Try
                ComPort.Write(send, 0, send.Length)

            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

        End If
    End Sub
    'servo motor 12
    Private Sub servo16bar12_Scroll(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ScrollEventArgs) Handles servo16bar12.Scroll
        value16_12.Text = servo16bar12.Value
        servo12 = servo16bar12.Value

        data1 = (servo12 >> 6) And &HFF
        data2 = servo12 And &H3F
        data3 = speed16_12.Value
        Pulse16_12.Text = (servo16bar12.Value * 0.25 * 10 ^ -3) + 0.5

        Dim send() As Byte = {&HEC, data1, data2, data3}
        If transmitt = 1 Then
            Try
                ComPort.Write(send, 0, send.Length)

            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

        End If
    End Sub
    'servo motor 13
    Private Sub servo16bar13_Scroll(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ScrollEventArgs) Handles servo16bar13.Scroll
        value16_13.Text = servo16bar13.Value
        servo13 = servo16bar13.Value

        data1 = (servo13 >> 6) And &HFF
        data2 = servo13 And &H3F
        data3 = speed16_13.Value
        Pulse16_13.Text = (servo16bar13.Value * 0.25 * 10 ^ -3) + 0.5

        Dim send() As Byte = {&HED, data1, data2, data3}
        If transmitt = 1 Then
            Try
                ComPort.Write(send, 0, send.Length)

            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

        End If
    End Sub
    'servo motor 14
    Private Sub servo16bar14_Scroll(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ScrollEventArgs) Handles servo16bar14.Scroll
        value16_14.Text = servo16bar14.Value
        servo14 = servo16bar14.Value

        data1 = (servo14 >> 6) And &HFF
        data2 = servo14 And &H3F
        data3 = speed16_14.Value
        Pulse16_14.Text = (servo16bar14.Value * 0.25 * 10 ^ -3) + 0.5

        Dim send() As Byte = {&HEE, data1, data2, data3}
        If transmitt = 1 Then
            Try
                ComPort.Write(send, 0, send.Length)

            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

        End If
    End Sub
    'servo motor 15
    Private Sub servo16bar15_Scroll(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ScrollEventArgs) Handles servo16bar15.Scroll
        value16_15.Text = servo16bar15.Value
        servo15 = servo16bar15.Value

        data1 = (servo15 >> 6) And &HFF
        data2 = servo8 And &H3F
        data3 = speed16_15.Value
        Pulse16_15.Text = (servo16bar15.Value * 0.25 * 10 ^ -3) + 0.5

        Dim send() As Byte = {&HEF, data1, data2, data3}
        If transmitt = 1 Then
            Try
                ComPort.Write(send, 0, send.Length)

            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

        End If
    End Sub
    'servo motor 16
    Private Sub servo16bar16_Scroll(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ScrollEventArgs) Handles servo16bar16.Scroll
        value16_16.Text = servo16bar16.Value
        servo16 = servo16bar16.Value

        data1 = (servo16 >> 6) And &HFF
        data2 = servo8 And &H3F
        data3 = speed16_16.Value
        Pulse16_16.Text = (servo16bar16.Value * 0.25 * 10 ^ -3) + 0.5

        Dim send() As Byte = {&HF0, data1, data2, data3}
        If transmitt = 1 Then
            Try
                ComPort.Write(send, 0, send.Length)

            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

        End If
    End Sub
    'reset all value
    Private Sub Button_Reset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Reset.Click
        Active1.Enabled = True
        Active8servo.Enabled = True

        servo16bar1.Value = 4000
        servo16bar2.Value = 4000
        servo16bar3.Value = 4000
        servo16bar4.Value = 4000
        servo16bar5.Value = 4000
        servo16bar6.Value = 4000
        servo16bar7.Value = 4000
        servo16bar8.Value = 4000
        servo16bar9.Value = 4000
        servo16bar10.Value = 4000
        servo16bar11.Value = 4000
        servo16bar12.Value = 4000
        servo16bar13.Value = 4000
        servo16bar14.Value = 4000
        servo16bar15.Value = 4000
        servo16bar16.Value = 4000


        value16_1.Text = 4000
        value16_2.Text = 4000
        value16_3.Text = 4000
        value16_4.Text = 4000
        value16_5.Text = 4000
        value16_6.Text = 4000
        value16_7.Text = 4000
        value16_8.Text = 4000
        value16_9.Text = 4000
        value16_10.Text = 4000
        value16_11.Text = 4000
        value16_12.Text = 4000
        value16_13.Text = 4000
        value16_14.Text = 4000
        value16_15.Text = 4000
        value16_16.Text = 4000


        Pulse16_1.Text = 1.5
        Pulse16_2.Text = 1.5
        Pulse16_3.Text = 1.5
        Pulse16_4.Text = 1.5
        Pulse16_5.Text = 1.5
        Pulse16_6.Text = 1.5
        Pulse16_7.Text = 1.5
        Pulse16_8.Text = 1.5
        Pulse16_9.Text = 1.5
        Pulse16_10.Text = 1.5
        Pulse16_11.Text = 1.5
        Pulse16_12.Text = 1.5
        Pulse16_13.Text = 1.5
        Pulse16_14.Text = 1.5
        Pulse16_15.Text = 1.5
        Pulse16_16.Text = 1.5


        speed16_1.Value = 0
        speed16_2.Value = 0
        speed16_3.Value = 0
        speed16_4.Value = 0
        speed16_5.Value = 0
        speed16_6.Value = 0
        speed16_7.Value = 0
        speed16_8.Value = 0
        speed16_9.Value = 0
        speed16_10.Value = 0
        speed16_11.Value = 0
        speed16_12.Value = 0
        speed16_13.Value = 0
        speed16_14.Value = 0
        speed16_15.Value = 0
        speed16_16.Value = 0


        Active1.Enabled = True

        Active16_1.Enabled = True
        Active16_2.Enabled = True
        Active16_3.Enabled = True
        Active16_4.Enabled = True
        Active16_5.Enabled = True
        Active16_6.Enabled = True
        Active16_7.Enabled = True
        Active16_8.Enabled = True
        Active16_9.Enabled = True
        Active16_10.Enabled = True
        Active16_11.Enabled = True
        Active16_12.Enabled = True
        Active16_13.Enabled = True
        Active16_14.Enabled = True
        Active16_15.Enabled = True
        Active16_16.Enabled = True


        Deactive16_1.Enabled = True
        Deactive16_2.Enabled = True
        Deactive16_3.Enabled = True
        Deactive16_4.Enabled = True
        Deactive16_5.Enabled = True
        Deactive16_6.Enabled = True
        Deactive16_7.Enabled = True
        Deactive16_8.Enabled = True
        Deactive16_9.Enabled = True
        Deactive16_10.Enabled = True
        Deactive16_11.Enabled = True
        Deactive16_12.Enabled = True
        Deactive16_13.Enabled = True
        Deactive16_14.Enabled = True
        Deactive16_15.Enabled = True
        Deactive16_16.Enabled = True

        servo16bar9.Enabled = True
        servo16bar10.Enabled = True
        servo16bar11.Enabled = True
        servo16bar12.Enabled = True
        servo16bar13.Enabled = True
        servo16bar14.Enabled = True
        servo16bar15.Enabled = True
        servo16bar16.Enabled = True

        value16_9.Enabled = True
        value16_10.Enabled = True
        value16_11.Enabled = True
        value16_12.Enabled = True
        value16_13.Enabled = True
        value16_14.Enabled = True
        value16_15.Enabled = True
        value16_16.Enabled = True

        Pulse16_9.Enabled = True
        Pulse16_10.Enabled = True
        Pulse16_11.Enabled = True
        Pulse16_12.Enabled = True
        Pulse16_13.Enabled = True
        Pulse16_14.Enabled = True
        Pulse16_15.Enabled = True
        Pulse16_16.Enabled = True

        speed16_9.Enabled = True
        speed16_10.Enabled = True
        speed16_11.Enabled = True
        speed16_12.Enabled = True
        speed16_13.Enabled = True
        speed16_14.Enabled = True
        speed16_15.Enabled = True
        speed16_16.Enabled = True

        Starting9.Enabled = True
        Starting10.Enabled = True
        Starting11.Enabled = True
        Starting12.Enabled = True
        Starting13.Enabled = True
        Starting14.Enabled = True
        Starting15.Enabled = True
        Starting16.Enabled = True

        CheckBox9.Enabled = True
        CheckBox10.Enabled = True
        CheckBox11.Enabled = True
        CheckBox12.Enabled = True
        CheckBox13.Enabled = True
        CheckBox14.Enabled = True
        CheckBox15.Enabled = True
        CheckBox16.Enabled = True
        If transmitt = 1 Then
            Dim send() As Byte = {&HC0, &H0} 'HC0=1byte,H0=2nd byte

            Try
                ComPort.Write(send, 0, send.Length) 'transmitt data
            Catch ex As Exception
                'MsgBox(ex.ToString)
                MsgBox("send fail")
            End Try

        End If
        '*************************************************************


    End Sub



    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If transmitt = 1 Then
            Dim send() As Byte = {&HC0, &H1} 'HC0=1byte,H1=2nd byte

            Try
                ComPort.Write(send, 0, send.Length) 'transmitt data

            Catch ex As Exception
                'MsgBox(ex.ToString)
                MsgBox("send fail")
            End Try

        End If
    End Sub
    '**************************************************************
    'Activate 16 Servo Motor Individual

    Private Sub Active16_1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Active16_1.Click

        If transmitt = 1 Then
            Dim send() As Byte = {&HC1, &H1} 'HC0=1byte,H1=2nd byte

            Try
                ComPort.Write(send, 0, send.Length) 'transmitt data
                Active16_1.Enabled = False
                Deactive16_1.Enabled = True
            Catch ex As Exception
                'MsgBox(ex.ToString)
                MsgBox("send fail")
            End Try

        End If
        Dim send21() As Byte = {&HA1}
        If transmitt = 1 Then
            Try
                ComPort.Write(send21, 0, send21.Length)
            Catch ex As Exception
                MsgBox("sendreport.fail1")

            End Try

        End If
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Try
            System.Threading.Thread.Sleep(300)
            higher_data = ComPort.ReadByte
            position = (higher_data And &H7F) << 6


            System.Threading.Thread.Sleep(300)
            lower_data = ComPort.ReadByte
            position = position Or (lower_data And &H3F)

            servo16bar1.Value = position
            value16_1.Text = position
            Pulse16_1.Text = (position * 0.25 * 10 ^ -3) + 0.5


        Catch ex As Exception
            MsgBox("Rc_report.fail1")
        End Try


    End Sub



    Private Sub Active16_2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Active16_2.Click
        If transmitt = 1 Then
            Dim send() As Byte = {&HC2, &H1} 'HC0=1byte,H1=2nd byte

            Try
                ComPort.Write(send, 0, send.Length) 'transmitt data
                Active16_2.Enabled = False
                Deactive16_2.Enabled = True
            Catch ex As Exception
                'MsgBox(ex.ToString)
                MsgBox("send fail")
            End Try

        End If

        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        Dim send22() As Byte = {&HA2}
        If transmitt = 1 Then
            Try
                ComPort.Write(send22, 0, send22.Length)
            Catch ex As Exception
                MsgBox("sendreport.fail1")

            End Try

        End If
        Try
            System.Threading.Thread.Sleep(300)
            higher_data = ComPort.ReadByte
            position = (higher_data And &H7F) << 6


            System.Threading.Thread.Sleep(300)
            lower_data = ComPort.ReadByte
            position = position Or (lower_data And &H3F)

            servo16bar2.Value = position
            value16_2.Text = position
            Pulse16_2.Text = (position * 0.25 * 10 ^ -3) + 0.5

        Catch ex As Exception
            MsgBox("Rc_report.fail1")
        End Try
    End Sub

    Private Sub Active16_3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Active16_3.Click
        If transmitt = 1 Then
            Dim send() As Byte = {&HC3, &H1} 'HC0=1byte,H1=2nd byte

            Try
                ComPort.Write(send, 0, send.Length) 'transmitt data
                Active16_3.Enabled = False
                Deactive16_3.Enabled = True
            Catch ex As Exception
                'MsgBox(ex.ToString)
                MsgBox("send fail")
            End Try

        End If
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim send23() As Byte = {&HA3}
        If transmitt = 1 Then
            Try
                ComPort.Write(send23, 0, send23.Length)
            Catch ex As Exception
                MsgBox("sendreport.fail1")

            End Try

        End If
        Try
            System.Threading.Thread.Sleep(300)
            higher_data = ComPort.ReadByte
            position = (higher_data And &H7F) << 6


            System.Threading.Thread.Sleep(300)
            lower_data = ComPort.ReadByte
            position = position Or (lower_data And &H3F)

            servo16bar3.Value = position
            value16_3.Text = position
            Pulse16_3.Text = (position * 0.25 * 10 ^ -3) + 0.5


        Catch ex As Exception
            MsgBox("Rc_report.fail1")
        End Try
    End Sub

    Private Sub Active16_4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Active16_4.Click
        If transmitt = 1 Then
            Dim send() As Byte = {&HC4, &H1} 'HC0=1byte,H1=2nd byte

            Try
                ComPort.Write(send, 0, send.Length) 'transmitt data
                Active16_4.Enabled = False
                Deactive16_4.Enabled = True
            Catch ex As Exception
                'MsgBox(ex.ToString)
                MsgBox("send fail")
            End Try

        End If
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim send24() As Byte = {&HA4}
        If transmitt = 1 Then
            Try
                ComPort.Write(send24, 0, send24.Length)
            Catch ex As Exception
                MsgBox("sendreport.fail1")

            End Try

        End If
        Try
            System.Threading.Thread.Sleep(300)
            higher_data = ComPort.ReadByte
            position = (higher_data And &H7F) << 6


            System.Threading.Thread.Sleep(300)
            lower_data = ComPort.ReadByte
            position = position Or (lower_data And &H3F)

            servo16bar4.Value = position
            value16_4.Text = position
            Pulse16_4.Text = (position * 0.25 * 10 ^ -3) + 0.5


        Catch ex As Exception
            MsgBox("Rc_report.fail1")
        End Try
    End Sub

    Private Sub Active16_5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Active16_5.Click
        If transmitt = 1 Then
            Dim send() As Byte = {&HC5, &H1} 'HC0=1byte,H1=2nd byte

            Try
                ComPort.Write(send, 0, send.Length) 'transmitt data
                Active16_5.Enabled = False
                Deactive16_5.Enabled = True
            Catch ex As Exception
                'MsgBox(ex.ToString)
                MsgBox("send fail")
            End Try

        End If
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim send25() As Byte = {&HA5}
        If transmitt = 1 Then
            Try
                ComPort.Write(send25, 0, send25.Length)
            Catch ex As Exception
                MsgBox("sendreport.fail1")

            End Try

        End If
        Try
            System.Threading.Thread.Sleep(300)
            higher_data = ComPort.ReadByte
            position = (higher_data And &H7F) << 6


            System.Threading.Thread.Sleep(300)
            lower_data = ComPort.ReadByte
            position = position Or (lower_data And &H3F)

            servo16bar5.Value = position
            value16_5.Text = position
            Pulse16_5.Text = (position * 0.25 * 10 ^ -3) + 0.5


        Catch ex As Exception
            MsgBox("Rc_report.fail1")
        End Try
    End Sub

    Private Sub Active16_6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Active16_6.Click
        If transmitt = 1 Then
            Dim send() As Byte = {&HC6, &H1} 'HC0=1byte,H1=2nd byte

            Try
                ComPort.Write(send, 0, send.Length) 'transmitt data
                Active16_6.Enabled = False
                Deactive16_6.Enabled = True
            Catch ex As Exception
                'MsgBox(ex.ToString)
                MsgBox("send fail")
            End Try

        End If
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim send26() As Byte = {&HA6}
        If transmitt = 1 Then
            Try
                ComPort.Write(send26, 0, send26.Length)
            Catch ex As Exception
                MsgBox("sendreport.fail1")

            End Try

        End If
        Try
            System.Threading.Thread.Sleep(300)
            higher_data = ComPort.ReadByte
            position = (higher_data And &H7F) << 6


            System.Threading.Thread.Sleep(300)
            lower_data = ComPort.ReadByte
            position = position Or (lower_data And &H3F)

            servo16bar6.Value = position
            value16_6.Text = position
            Pulse16_6.Text = (position * 0.25 * 10 ^ -3) + 0.5


        Catch ex As Exception
            MsgBox("Rc_report.fail1")
        End Try
    End Sub

    Private Sub Active16_7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Active16_7.Click
        If transmitt = 1 Then
            Dim send() As Byte = {&HC7, &H1} 'HC0=1byte,H1=2nd byte

            Try
                ComPort.Write(send, 0, send.Length) 'transmitt data
                Active16_7.Enabled = False
                Deactive16_7.Enabled = True
            Catch ex As Exception
                'MsgBox(ex.ToString)
                MsgBox("send fail")
            End Try

        End If
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        Dim send27() As Byte = {&HA7}
        If transmitt = 1 Then
            Try
                ComPort.Write(send27, 0, send27.Length)
            Catch ex As Exception
                MsgBox("sendreport.fail1")

            End Try

        End If
        Try
            System.Threading.Thread.Sleep(300)
            higher_data = ComPort.ReadByte
            position = (higher_data And &H7F) << 6


            System.Threading.Thread.Sleep(300)
            lower_data = ComPort.ReadByte
            position = position Or (lower_data And &H3F)

            servo16bar7.Value = position
            value16_7.Text = position
            Pulse16_7.Text = (position * 0.25 * 10 ^ -3) + 0.5


        Catch ex As Exception
            MsgBox("Rc_report.fail1")
        End Try
    End Sub

    Private Sub Active16_8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Active16_8.Click
        If transmitt = 1 Then
            Dim send() As Byte = {&HC8, &H1} 'HC0=1byte,H1=2nd byte

            Try
                ComPort.Write(send, 0, send.Length) 'transmitt data
                Active16_8.Enabled = False
                Deactive16_8.Enabled = True
            Catch ex As Exception
                'MsgBox(ex.ToString)
                MsgBox("send fail")
            End Try

        End If
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        Dim send28() As Byte = {&HA8}
        If transmitt = 1 Then
            Try
                ComPort.Write(send28, 0, send28.Length)
            Catch ex As Exception
                MsgBox("sendreport.fail1")

            End Try

        End If
        Try
            System.Threading.Thread.Sleep(300)
            higher_data = ComPort.ReadByte
            position = (higher_data And &H7F) << 6


            System.Threading.Thread.Sleep(300)
            lower_data = ComPort.ReadByte
            position = position Or (lower_data And &H3F)

            servo16bar8.Value = position
            value16_8.Text = position
            Pulse16_8.Text = (position * 0.25 * 10 ^ -3) + 0.5


        Catch ex As Exception
            MsgBox("Rc_report.fail1")
        End Try
    End Sub

    Private Sub Active16_9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Active16_9.Click
        If transmitt = 1 Then
            Dim send() As Byte = {&HC9, &H1} 'HC0=1byte,H1=2nd byte

            Try
                ComPort.Write(send, 0, send.Length) 'transmitt data
                Active16_9.Enabled = False
                Deactive16_9.Enabled = True
            Catch ex As Exception
                'MsgBox(ex.ToString)
                MsgBox("send fail")
            End Try

        End If
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        Dim send29() As Byte = {&HA9}
        If transmitt = 1 Then
            Try
                ComPort.Write(send29, 0, send29.Length)
            Catch ex As Exception
                MsgBox("sendreport.fail9")

            End Try

        End If
        Try
            System.Threading.Thread.Sleep(300)
            higher_data = ComPort.ReadByte
            position = (higher_data And &H7F) << 6


            System.Threading.Thread.Sleep(300)
            lower_data = ComPort.ReadByte
            position = position Or (lower_data And &H3F)

            servo16bar9.Value = position
            value16_9.Text = position
            Pulse16_9.Text = (position * 0.25 * 10 ^ -3) + 0.5


        Catch ex As Exception
            MsgBox("Rc_report.fail9")
        End Try

        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      
    End Sub

    Private Sub Active16_10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Active16_10.Click
        If transmitt = 1 Then
            Dim send() As Byte = {&HCA, &H1} 'HC0=1byte,H1=2nd byte

            Try
                ComPort.Write(send, 0, send.Length) 'transmitt data
                Active16_10.Enabled = False
                Deactive16_10.Enabled = True
            Catch ex As Exception
                'MsgBox(ex.ToString)
                MsgBox("send fail")
            End Try

        End If
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        Dim send30() As Byte = {&HAA}
        If transmitt = 1 Then
            Try
                ComPort.Write(send30, 0, send30.Length)
            Catch ex As Exception
                MsgBox("sendreport.fail10")

            End Try

        End If
        Try
            System.Threading.Thread.Sleep(300)
            higher_data = ComPort.ReadByte
            position = (higher_data And &H7F) << 6


            System.Threading.Thread.Sleep(300)
            lower_data = ComPort.ReadByte
            position = position Or (lower_data And &H3F)

            servo16bar10.Value = position
            value16_10.Text = position
            Pulse16_10.Text = (position * 0.25 * 10 ^ -3) + 0.5


        Catch ex As Exception
            MsgBox("Rc_report.fail10")
        End Try
       
    End Sub

    Private Sub Active16_11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Active16_11.Click
        If transmitt = 1 Then
            Dim send() As Byte = {&HCB, &H1} 'HC0=1byte,H1=2nd byte

            Try
                ComPort.Write(send, 0, send.Length) 'transmitt data
                Active16_11.Enabled = False
                Deactive16_11.Enabled = True
            Catch ex As Exception
                'MsgBox(ex.ToString)
                MsgBox("send fail")
            End Try

        End If
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        Dim send31() As Byte = {&HAB}
        If transmitt = 1 Then
            Try
                ComPort.Write(send31, 0, send31.Length)
            Catch ex As Exception
                MsgBox("sendreport.fail11")

            End Try

        End If
        Try
            System.Threading.Thread.Sleep(300)
            higher_data = ComPort.ReadByte
            position = (higher_data And &H7F) << 6


            System.Threading.Thread.Sleep(300)
            lower_data = ComPort.ReadByte
            position = position Or (lower_data And &H3F)

            servo16bar11.Value = position
            value16_11.Text = position
            Pulse16_11.Text = (position * 0.25 * 10 ^ -3) + 0.5


        Catch ex As Exception
            MsgBox("Rc_report.fail11")
        End Try
    End Sub

    Private Sub Active16_12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Active16_12.Click
        If transmitt = 1 Then
            Dim send() As Byte = {&HCC, &H1} 'HC0=1byte,H1=2nd byte

            Try
                ComPort.Write(send, 0, send.Length) 'transmitt data
                Active16_12.Enabled = False
                Deactive16_12.Enabled = True
            Catch ex As Exception
                'MsgBox(ex.ToString)
                MsgBox("send fail")
            End Try

        End If
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim send32() As Byte = {&HAC}
        If transmitt = 1 Then
            Try
                ComPort.Write(send32, 0, send32.Length)
            Catch ex As Exception
                MsgBox("sendreport.fail12")

            End Try

        End If
        Try
            System.Threading.Thread.Sleep(300)
            higher_data = ComPort.ReadByte
            position = (higher_data And &H7F) << 6


            System.Threading.Thread.Sleep(300)
            lower_data = ComPort.ReadByte
            position = position Or (lower_data And &H3F)

            servo16bar12.Value = position
            value16_12.Text = position
            Pulse16_12.Text = (position * 0.25 * 10 ^ -3) + 0.5


        Catch ex As Exception
            MsgBox("Rc_report.fail12")
        End Try

    End Sub

    Private Sub Active16_13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Active16_13.Click
        If transmitt = 1 Then
            Dim send() As Byte = {&HCD, &H1} 'HC0=1byte,H1=2nd byte

            Try
                ComPort.Write(send, 0, send.Length) 'transmitt data
                Active16_13.Enabled = False
                Deactive16_13.Enabled = True
            Catch ex As Exception
                'MsgBox(ex.ToString)
                MsgBox("send fail")
            End Try

        End If
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim send33() As Byte = {&HAD}
        If transmitt = 1 Then
            Try
                ComPort.Write(send33, 0, send33.Length)
            Catch ex As Exception
                MsgBox("sendreport.fail13")

            End Try

        End If
        Try
            System.Threading.Thread.Sleep(300)
            higher_data = ComPort.ReadByte
            position = (higher_data And &H7F) << 6


            System.Threading.Thread.Sleep(300)
            lower_data = ComPort.ReadByte
            position = position Or (lower_data And &H3F)

            servo16bar13.Value = position
            value16_13.Text = position
            Pulse16_13.Text = (position * 0.25 * 10 ^ -3) + 0.5


        Catch ex As Exception
            MsgBox("Rc_report.fail13")
        End Try
    End Sub

    Private Sub Active16_14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Active16_14.Click
        If transmitt = 1 Then
            Dim send() As Byte = {&HCE, &H1} 'HC0=1byte,H1=2nd byte

            Try
                ComPort.Write(send, 0, send.Length) 'transmitt data
                Active16_14.Enabled = False
                Deactive16_14.Enabled = True
            Catch ex As Exception
                'MsgBox(ex.ToString)
                MsgBox("send fail")
            End Try

        End If
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        Dim send34() As Byte = {&HAE}
        If transmitt = 1 Then
            Try
                ComPort.Write(send34, 0, send34.Length)
            Catch ex As Exception
                MsgBox("sendreport.fail14")

            End Try

        End If
        Try
            System.Threading.Thread.Sleep(300)
            higher_data = ComPort.ReadByte
            position = (higher_data And &H7F) << 6


            System.Threading.Thread.Sleep(300)
            lower_data = ComPort.ReadByte
            position = position Or (lower_data And &H3F)

            servo16bar14.Value = position
            value16_14.Text = position
            Pulse16_14.Text = (position * 0.25 * 10 ^ -3) + 0.5


        Catch ex As Exception
            MsgBox("Rc_report.fail14")
        End Try
  
    End Sub

    Private Sub Active16_15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Active16_15.Click
        If transmitt = 1 Then
            Dim send() As Byte = {&HCF, &H1} 'HC0=1byte,H1=2nd byte

            Try
                ComPort.Write(send, 0, send.Length) 'transmitt data
                Active16_15.Enabled = False
                Deactive16_15.Enabled = True
            Catch ex As Exception
                'MsgBox(ex.ToString)
                MsgBox("send fail")
            End Try

        End If
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


        Dim send35() As Byte = {&HAF}
        If transmitt = 1 Then
            Try
                ComPort.Write(send35, 0, send35.Length)
            Catch ex As Exception
                MsgBox("sendreport.fail15")

            End Try

        End If
        Try
            System.Threading.Thread.Sleep(300)
            higher_data = ComPort.ReadByte
            position = (higher_data And &H7F) << 6


            System.Threading.Thread.Sleep(300)
            lower_data = ComPort.ReadByte
            position = position Or (lower_data And &H3F)

            servo16bar15.Value = position
            value16_15.Text = position
            Pulse16_15.Text = (position * 0.25 * 10 ^ -3) + 0.5


        Catch ex As Exception
            MsgBox("Rc_report.fail15")
        End Try
    End Sub

    Private Sub Active16_16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Active16_16.Click
        If transmitt = 1 Then
            Dim send() As Byte = {&HD0, &H1} 'HC0=1byte,H1=2nd byte

            Try
                ComPort.Write(send, 0, send.Length) 'transmitt data
                Active16_16.Enabled = False
                Deactive16_16.Enabled = True
            Catch ex As Exception
                'MsgBox(ex.ToString)
                MsgBox("send fail")
            End Try

        End If
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        Dim send36() As Byte = {&HB0}
        If transmitt = 1 Then
            Try
                ComPort.Write(send36, 0, send36.Length)
            Catch ex As Exception
                MsgBox("sendreport.fail16")

            End Try

        End If
        Try
            System.Threading.Thread.Sleep(300)
            higher_data = ComPort.ReadByte
            position = (higher_data And &H7F) << 6


            System.Threading.Thread.Sleep(300)
            lower_data = ComPort.ReadByte
            position = position Or (lower_data And &H3F)

            servo16bar16.Value = position
            value16_16.Text = position
            Pulse16_16.Text = (position * 0.25 * 10 ^ -3) + 0.5


        Catch ex As Exception
            MsgBox("Rc_report.fail16")
        End Try
    
    End Sub
    '***********************************************************
    'Deactive 16Servo Motor Indivudual
    Private Sub Deactive16_1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Deactive16_1.Click
        If transmitt = 1 Then
            Dim send() As Byte = {&HC1, &H0} 'HC0=1byte,H1=2nd byte

            Try
                ComPort.Write(send, 0, send.Length) 'transmitt data
                Active16_1.Enabled = True
                Deactive16_1.Enabled = False
            Catch ex As Exception
                'MsgBox(ex.ToString)
                MsgBox("send fail")
            End Try

        End If
    End Sub

    Private Sub Deactive16_2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Deactive16_2.Click
        If transmitt = 1 Then
            Dim send() As Byte = {&HC2, &H0} 'HC0=1byte,H1=2nd byte

            Try
                ComPort.Write(send, 0, send.Length) 'transmitt data
                Active16_2.Enabled = True
                Deactive16_2.Enabled = False
            Catch ex As Exception
                'MsgBox(ex.ToString)
                MsgBox("send fail")
            End Try

        End If
    End Sub

    Private Sub Deactive16_3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Deactive16_3.Click
        If transmitt = 1 Then
            Dim send() As Byte = {&HC3, &H0} 'HC0=1byte,H1=2nd byte

            Try
                ComPort.Write(send, 0, send.Length) 'transmitt data
                Active16_3.Enabled = True
                Deactive16_3.Enabled = False
            Catch ex As Exception
                'MsgBox(ex.ToString)
                MsgBox("send fail")
            End Try

        End If
    End Sub

    Private Sub Deactive16_4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Deactive16_4.Click
        If transmitt = 1 Then
            Dim send() As Byte = {&HC4, &H0} 'HC0=1byte,H1=2nd byte

            Try
                ComPort.Write(send, 0, send.Length) 'transmitt data
                Active16_4.Enabled = True
                Deactive16_4.Enabled = False
            Catch ex As Exception
                'MsgBox(ex.ToString)
                MsgBox("send fail")
            End Try

        End If

    End Sub

    Private Sub Deactive16_5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Deactive16_5.Click
        If transmitt = 1 Then
            Dim send() As Byte = {&HC5, &H0} 'HC0=1byte,H1=2nd byte

            Try
                ComPort.Write(send, 0, send.Length) 'transmitt data
                Active16_5.Enabled = True
                Deactive16_5.Enabled = False
            Catch ex As Exception
                'MsgBox(ex.ToString)
                MsgBox("send fail")
            End Try

        End If
    End Sub

    Private Sub Deactive16_6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Deactive16_6.Click
        If transmitt = 1 Then
            Dim send() As Byte = {&HC6, &H0} 'HC0=1byte,H1=2nd byte

            Try
                ComPort.Write(send, 0, send.Length) 'transmitt data
                Active16_6.Enabled = True
                Deactive16_6.Enabled = False
            Catch ex As Exception
                'MsgBox(ex.ToString)
                MsgBox("send fail")
            End Try

        End If
    End Sub

    Private Sub Deactive16_7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Deactive16_7.Click
        If transmitt = 1 Then
            Dim send() As Byte = {&HC7, &H0} 'HC0=1byte,H1=2nd byte

            Try
                ComPort.Write(send, 0, send.Length) 'transmitt data
                Active16_7.Enabled = True
                Deactive16_7.Enabled = False
            Catch ex As Exception
                'MsgBox(ex.ToString)
                MsgBox("send fail")
            End Try

        End If
    End Sub

    Private Sub Deactive16_8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Deactive16_8.Click
        If transmitt = 1 Then
            Dim send() As Byte = {&HC8, &H0} 'HC0=1byte,H1=2nd byte

            Try
                ComPort.Write(send, 0, send.Length) 'transmitt data
                Active16_8.Enabled = True
                Deactive16_8.Enabled = False
            Catch ex As Exception
                'MsgBox(ex.ToString)
                MsgBox("send fail")
            End Try

        End If
    End Sub

    Private Sub Deactive16_9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Deactive16_9.Click
        If transmitt = 1 Then
            Dim send() As Byte = {&HC9, &H0} 'HC0=1byte,H1=2nd byte

            Try
                ComPort.Write(send, 0, send.Length) 'transmitt data
                Active16_9.Enabled = True
                Deactive16_9.Enabled = False
            Catch ex As Exception
                'MsgBox(ex.ToString)
                MsgBox("send fail")
            End Try

        End If
    End Sub

    Private Sub Deactive16_10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Deactive16_10.Click
        If transmitt = 1 Then
            Dim send() As Byte = {&HCA, &H0} 'HC0=1byte,H1=2nd byte

            Try
                ComPort.Write(send, 0, send.Length) 'transmitt data
                Active16_10.Enabled = True
                Deactive16_10.Enabled = False
            Catch ex As Exception
                'MsgBox(ex.ToString)
                MsgBox("send fail")
            End Try

        End If
    End Sub

    Private Sub Deactive16_11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Deactive16_11.Click
        If transmitt = 1 Then
            Dim send() As Byte = {&HCB, &H0} 'HC0=1byte,H1=2nd byte

            Try
                ComPort.Write(send, 0, send.Length) 'transmitt data
                Active16_11.Enabled = True
                Deactive16_11.Enabled = False
            Catch ex As Exception
                'MsgBox(ex.ToString)
                MsgBox("send fail")
            End Try

        End If
    End Sub

    Private Sub Deactive16_12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Deactive16_12.Click
        If transmitt = 1 Then
            Dim send() As Byte = {&HCC, &H0} 'HC0=1byte,H1=2nd byte

            Try
                ComPort.Write(send, 0, send.Length) 'transmitt data
                Active16_12.Enabled = True
                Deactive16_12.Enabled = False
            Catch ex As Exception
                'MsgBox(ex.ToString)
                MsgBox("send fail")
            End Try

        End If
    End Sub

    Private Sub Deactive16_13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Deactive16_13.Click
        If transmitt = 1 Then
            Dim send() As Byte = {&HCD, &H0} 'HC0=1byte,H1=2nd byte

            Try
                ComPort.Write(send, 0, send.Length) 'transmitt data
                Active16_13.Enabled = True
                Deactive16_13.Enabled = False
            Catch ex As Exception
                'MsgBox(ex.ToString)
                MsgBox("send fail")
            End Try

        End If
    End Sub
    Private Sub Deactive16_14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Deactive16_14.Click
        If transmitt = 1 Then
            Dim send() As Byte = {&HCE, &H0} 'HC0=1byte,H1=2nd byte

            Try
                ComPort.Write(send, 0, send.Length) 'transmitt data
                Active16_14.Enabled = True
                Deactive16_14.Enabled = False
            Catch ex As Exception
                'MsgBox(ex.ToString)
                MsgBox("send fail")
            End Try

        End If
    End Sub
    Private Sub Deactive16_15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Deactive16_15.Click
        If transmitt = 1 Then
            Dim send() As Byte = {&HCF, &H0} 'HC0=1byte,H1=2nd byte

            Try
                ComPort.Write(send, 0, send.Length) 'transmitt data
                Active16_15.Enabled = True
                Deactive16_15.Enabled = False
            Catch ex As Exception
                'MsgBox(ex.ToString)
                MsgBox("send fail")
            End Try

        End If
    End Sub



    Private Sub Deactive16_16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Deactive16_16.Click
        If transmitt = 1 Then
            Dim send() As Byte = {&HD0, &H0} 'HC0=1byte,H1=2nd byte

            Try
                ComPort.Write(send, 0, send.Length) 'transmitt data
                Active16_16.Enabled = True
                Deactive16_16.Enabled = False
            Catch ex As Exception
                'MsgBox(ex.ToString)
                MsgBox("send fail")
            End Try

        End If
    End Sub



    Private Sub Button2_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Clear.Click

        Starting1.Value = 4000
        Starting2.Value = 4000
        Starting3.Value = 4000
        Starting4.Value = 4000
        Starting5.Value = 4000
        Starting6.Value = 4000
        Starting7.Value = 4000
        Starting8.Value = 4000
        Starting9.Value = 4000
        Starting10.Value = 4000
        Starting11.Value = 4000
        Starting12.Value = 4000
        Starting13.Value = 4000
        Starting14.Value = 4000
        Starting15.Value = 4000
        Starting16.Value = 4000

    End Sub


    Dim rxbuff As String

    Private Sub Button2_Click_2(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

        CheckBox1.Enabled = False
        CheckBox2.Enabled = False
        CheckBox3.Enabled = False
        CheckBox4.Enabled = False
        CheckBox5.Enabled = False
        CheckBox6.Enabled = False
        CheckBox7.Enabled = False
        CheckBox8.Enabled = False
        CheckBox9.Enabled = False
        CheckBox10.Enabled = False
        CheckBox11.Enabled = False
        CheckBox12.Enabled = False
        CheckBox13.Enabled = False
        CheckBox14.Enabled = False
        CheckBox15.Enabled = False
        CheckBox16.Enabled = False

        If ProgressBar1.Value <> ProgressBar1.Maximum Then
            If CheckBox1.Checked = True Then


                servo1 = Starting1.Text

                data1 = (servo1 >> 6) And &HFF
                data2 = servo1 And &H3F

                Dim send() As Byte = {&H81, data1, data2}
                If transmitt = 1 Then
                    Try
                        ComPort.Write(send, 0, send.Length)

                    Catch ex As Exception
                        MsgBox(ex.ToString)
                    End Try
                End If

            End If
            System.Threading.Thread.Sleep(300)
            Try
                rxbuff = (ComPort.ReadExisting)

            Catch ex As Exception
                MsgBox("Rc fail1")

            End Try
            ProgressBar1.Increment(10)
            '''''''''''''''''''''''''''''''''''''''''''''''''''
            If CheckBox2.Checked = True Then

                servo2 = Starting2.Text

                data1 = (servo2 >> 6) And &HFF
                data2 = servo2 And &H3F

                Dim send2() As Byte = {&H82, data1, data2}
                If transmitt = 1 Then
                    Try
                        ComPort.Write(send2, 0, send2.Length)

                    Catch ex As Exception
                        MsgBox(ex.ToString)
                    End Try

                End If
            End If

            System.Threading.Thread.Sleep(300)
            Try
                rxbuff = (ComPort.ReadExisting)

            Catch ex As Exception
                MsgBox("Rc fail2")

            End Try
            ProgressBar1.Increment(10)
            '''''''''''''''''''''''''''''''''''''''''''''''''''
            If CheckBox3.Checked = True Then

                servo3 = Starting3.Text

                data1 = (servo3 >> 6) And &HFF
                data2 = servo3 And &H3F

                Dim send3() As Byte = {&H83, data1, data2}
                If transmitt = 1 Then
                    Try
                        ComPort.Write(send3, 0, send3.Length)

                    Catch ex As Exception
                        MsgBox(ex.ToString)
                    End Try

                End If
            End If

            System.Threading.Thread.Sleep(300)
            Try
                rxbuff = (ComPort.ReadExisting)

            Catch ex As Exception
                MsgBox("Rc fail3")

            End Try
            ProgressBar1.Increment(10)
            ''''''''''''''''''''''''''''''''''''''''''''''''''
            If CheckBox4.Checked = True Then

                servo4 = Starting4.Text

                data1 = (servo4 >> 6) And &HFF
                data2 = servo4 And &H3F

                Dim send4() As Byte = {&H84, data1, data2}
                If transmitt = 1 Then
                    Try
                        ComPort.Write(send4, 0, send4.Length)

                    Catch ex As Exception
                        MsgBox(ex.ToString)
                    End Try

                End If
            End If

            System.Threading.Thread.Sleep(300)
            Try
                rxbuff = (ComPort.ReadExisting)
            Catch ex As Exception
                MsgBox("Rc fail4")

            End Try
            ProgressBar1.Increment(10)
            ''''''''''''''''''''''''''''''''''''''''''''''
            If CheckBox5.Checked = True Then

                servo5 = Starting5.Text

                data1 = (servo5 >> 6) And &HFF
                data2 = servo5 And &H3F

                Dim send5() As Byte = {&H85, data1, data2}
                If transmitt = 1 Then
                    Try
                        ComPort.Write(send5, 0, send5.Length)

                    Catch ex As Exception
                        MsgBox(ex.ToString)
                    End Try

                End If
            End If

            System.Threading.Thread.Sleep(300)
            Try
                rxbuff = (ComPort.ReadExisting)
            Catch ex As Exception
                MsgBox("Rc fail5")

            End Try
            ProgressBar1.Increment(10)
            ''''''''''''''''''''''''''''''''''''''''''''''''''

            If CheckBox6.Checked = True Then

                servo6 = Starting6.Text

                data1 = (servo6 >> 6) And &HFF
                data2 = servo6 And &H3F
                data3 = speed16_6.Value

                Dim send6() As Byte = {&H86, data1, data2}
                If transmitt = 1 Then
                    Try
                        ComPort.Write(send6, 0, send6.Length)

                    Catch ex As Exception
                        MsgBox(ex.ToString)
                    End Try

                End If
            End If


            System.Threading.Thread.Sleep(300)
            Try
                rxbuff = (ComPort.ReadExisting)
            Catch ex As Exception
                MsgBox("Rc fail6")

            End Try
            ProgressBar1.Increment(10)
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''

            If CheckBox7.Checked = True Then

                servo7 = Starting7.Text

                data1 = (servo7 >> 6) And &HFF
                data2 = servo7 And &H3F
                Dim send7() As Byte = {&H87, data1, data2}
                If transmitt = 1 Then
                    Try
                        ComPort.Write(send7, 0, send7.Length)

                    Catch ex As Exception
                        MsgBox(ex.ToString)
                    End Try

                End If
            End If

            System.Threading.Thread.Sleep(300)
            Try
                rxbuff = (ComPort.ReadExisting)
            Catch ex As Exception
                MsgBox("Rc fail7")

            End Try
            ProgressBar1.Increment(10)
            '''''''''''''''''''''''''''''''''''''''''''''''''''''

            If CheckBox8.Checked = True Then

                servo8 = Starting8.Text

                data1 = (servo8 >> 6) And &HFF
                data2 = servo8 And &H3F

                Dim send8() As Byte = {&H88, data1, data2}
                If transmitt = 1 Then
                    Try
                        ComPort.Write(send8, 0, send8.Length)

                    Catch ex As Exception
                        MsgBox(ex.ToString)
                    End Try

                End If
            End If

            System.Threading.Thread.Sleep(300)
            Try
                rxbuff = (ComPort.ReadExisting)
            Catch ex As Exception
                MsgBox("Rc fail8")

            End Try
            ProgressBar1.Increment(10)
            ''''''''''''''''''''''''''''''''''''''''''''

            If CheckBox9.Checked = True Then
                servo9 = Starting9.Text

                data1 = (servo9 >> 6) And &HFF
                data2 = servo9 And &H3F

                Dim send9() As Byte = {&H89, data1, data2}
                If transmitt = 1 Then
                    Try
                        ComPort.Write(send9, 0, send9.Length)

                    Catch ex As Exception
                        MsgBox(ex.ToString)
                    End Try

                End If
            End If
            System.Threading.Thread.Sleep(300)
            Try
                rxbuff = (ComPort.ReadExisting)
            Catch ex As Exception
                MsgBox("Rc fail9")
            End Try
            ProgressBar1.Increment(10)

            '''''''''''''''''''''''''''''''''''''''''''''''''''''


            If CheckBox10.Checked = True Then
                servo10 = Starting10.Text

                data1 = (servo10 >> 6) And &HFF
                data2 = servo10 And &H3F

                Dim send10() As Byte = {&H8A, data1, data2}
                If transmitt = 1 Then
                    Try
                        ComPort.Write(send10, 0, send10.Length)

                    Catch ex As Exception
                        MsgBox(ex.ToString)
                    End Try

                End If
            End If
            System.Threading.Thread.Sleep(300)
            Try
                rxbuff = (ComPort.ReadExisting)
            Catch ex As Exception
                MsgBox("Rc fail10")
            End Try
            ProgressBar1.Increment(10)
            ''''''''''''''''''''''''''''''''''''''''''''''''''''
            If CheckBox11.Checked = True Then
                servo11 = Starting11.Text

                data1 = (servo11 >> 6) And &HFF
                data2 = servo11 And &H3F

                Dim send11() As Byte = {&H8B, data1, data2}
                If transmitt = 1 Then
                    Try
                        ComPort.Write(send11, 0, send11.Length)

                    Catch ex As Exception
                        MsgBox(ex.ToString)
                    End Try

                End If
            End If
            System.Threading.Thread.Sleep(300)
            Try
                rxbuff = (ComPort.ReadExisting)
            Catch ex As Exception
                MsgBox("Rc fail11")
            End Try
            ProgressBar1.Increment(10)
            '''''''''''''''''''''''''''''''''''''''''''''''''''''
            If CheckBox12.Checked = True Then
                servo12 = Starting12.Text

                data1 = (servo12 >> 6) And &HFF
                data2 = servo12 And &H3F

                Dim send12() As Byte = {&H8C, data1, data2}
                If transmitt = 1 Then
                    Try
                        ComPort.Write(send12, 0, send12.Length)

                    Catch ex As Exception
                        MsgBox(ex.ToString)
                    End Try

                End If
            End If
            System.Threading.Thread.Sleep(300)
            Try
                rxbuff = (ComPort.ReadExisting)
            Catch ex As Exception
                MsgBox("Rc fail12")
            End Try
            ProgressBar1.Increment(10)
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''
            If CheckBox13.Checked = True Then
                servo13 = Starting13.Text

                data1 = (servo13 >> 6) And &HFF
                data2 = servo13 And &H3F

                Dim send13() As Byte = {&H8D, data1, data2}
                If transmitt = 1 Then
                    Try
                        ComPort.Write(send13, 0, send13.Length)

                    Catch ex As Exception
                        MsgBox(ex.ToString)
                    End Try

                End If
            End If


            System.Threading.Thread.Sleep(300)
            Try
                rxbuff = (ComPort.ReadExisting)
            Catch ex As Exception
                MsgBox("Rc fail13")
            End Try
            ProgressBar1.Increment(10)
            '''''''''''''''''''''''''''''''''''''''''''''''''''
            If CheckBox14.Checked = True Then
                servo14 = Starting14.Text

                data1 = (servo14 >> 6) And &HFF
                data2 = servo14 And &H3F

                Dim send14() As Byte = {&H8E, data1, data2}
                If transmitt = 1 Then
                    Try
                        ComPort.Write(send14, 0, send14.Length)

                    Catch ex As Exception
                        MsgBox(ex.ToString)
                    End Try

                End If
            End If

            System.Threading.Thread.Sleep(300)
            Try
                rxbuff = (ComPort.ReadExisting)
            Catch ex As Exception
                MsgBox("Rc fail14")
            End Try
            ProgressBar1.Increment(10)
            '''''''''''''''''''''''''''''''''''''''''''''''''''
            If CheckBox15.Checked = True Then
                servo15 = Starting15.Text

                data1 = (servo15 >> 6) And &HFF
                data2 = servo15 And &H3F

                Dim send15() As Byte = {&H8F, data1, data2}
                If transmitt = 1 Then
                    Try
                        ComPort.Write(send15, 0, send15.Length)

                    Catch ex As Exception
                        MsgBox(ex.ToString)
                    End Try

                End If
            End If
            System.Threading.Thread.Sleep(300)
            Try
                rxbuff = (ComPort.ReadExisting)
            Catch ex As Exception
                MsgBox("Rc fail15")
            End Try
            ProgressBar1.Increment(10)
            '''''''''''''''''''''''''''''''''''''''''''''''''''''
            If CheckBox16.Checked = True Then
                servo16 = Starting16.Text

                data1 = (servo16 >> 6) And &HFF
                data2 = servo16 And &H3F

                Dim send16() As Byte = {&H90, data1, data2}
                If transmitt = 1 Then
                    Try
                        ComPort.Write(send16, 0, send16.Length)

                    Catch ex As Exception
                        MsgBox(ex.ToString)
                    End Try

                End If
            End If
            System.Threading.Thread.Sleep(300)
            Try
                rxbuff = (ComPort.ReadExisting)
            Catch ex As Exception
                MsgBox("Rc fail16")
            End Try
            ''''''''''''''''''''''''''''''''''''''''''''''''''
            ProgressBar1.Increment(10)

        End If


        ProgressBar1.Value = 0
        CheckBox1.Enabled = True
        CheckBox2.Enabled = True
        CheckBox3.Enabled = True
        CheckBox4.Enabled = True
        CheckBox5.Enabled = True
        CheckBox6.Enabled = True
        CheckBox7.Enabled = True
        CheckBox8.Enabled = True
        CheckBox9.Enabled = True
        CheckBox10.Enabled = True
        CheckBox11.Enabled = True
        CheckBox12.Enabled = True
        CheckBox13.Enabled = True
        CheckBox14.Enabled = True
        CheckBox15.Enabled = True
        CheckBox16.Enabled = True
        Timer1.Stop()
    End Sub
    'clear all chek box button
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        CheckBox1.CheckState = False
        CheckBox2.CheckState = False
        CheckBox3.CheckState = False
        CheckBox4.CheckState = False
        CheckBox5.CheckState = False
        CheckBox6.CheckState = False
        CheckBox7.CheckState = False
        CheckBox8.CheckState = False
        CheckBox9.CheckState = False
        CheckBox10.CheckState = False
        CheckBox11.CheckState = False
        CheckBox12.CheckState = False
        CheckBox13.CheckState = False
        CheckBox14.CheckState = False
        CheckBox15.CheckState = False
        CheckBox16.CheckState = False
    End Sub

    Private Sub ProgressBar1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ProgressBar1.Click
        Timer1.Start()
     
       
    End Sub

    '*********************************************
    '******** 8 servo motor operation code ******
    '*********************************************
    Private Sub Active8servo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Active8servo.Click
        Button_MarkAll.Enabled = False

        If ProgressBar3.Value <> ProgressBar3.Maximum Then


            Active1.Enabled = False
            Active8servo.Enabled = False
            If transmitt = 1 Then
                Dim send() As Byte = {&HC0, &H1} 'activate all servo HC0=1byte,H1=2nd byte

                Try
                    ComPort.Write(send, 0, send.Length) 'transmitt data
                    Active1.Enabled = False

                    Active16_1.Enabled = False
                    Active16_2.Enabled = False
                    Active16_3.Enabled = False
                    Active16_4.Enabled = False
                    Active16_5.Enabled = False
                    Active16_6.Enabled = False
                    Active16_7.Enabled = False
                    Active16_8.Enabled = False
                    Active16_9.Enabled = False
                    Active16_10.Enabled = False
                    Active16_11.Enabled = False
                    Active16_12.Enabled = False
                    Active16_13.Enabled = False
                    Active16_14.Enabled = False
                    Active16_15.Enabled = False
                    Active16_16.Enabled = False

                    Deactive16_1.Enabled = True
                    Deactive16_2.Enabled = True
                    Deactive16_3.Enabled = True
                    Deactive16_4.Enabled = True
                    Deactive16_5.Enabled = True
                    Deactive16_6.Enabled = True
                    Deactive16_7.Enabled = True
                    Deactive16_8.Enabled = True

                    Deactive16_9.Enabled = False
                    Deactive16_10.Enabled = False
                    Deactive16_11.Enabled = False
                    Deactive16_12.Enabled = False
                    Deactive16_13.Enabled = False
                    Deactive16_14.Enabled = False
                    Deactive16_15.Enabled = False
                    Deactive16_16.Enabled = False

                    servo16bar9.Enabled = False
                    servo16bar10.Enabled = False
                    servo16bar11.Enabled = False
                    servo16bar12.Enabled = False
                    servo16bar13.Enabled = False
                    servo16bar14.Enabled = False
                    servo16bar15.Enabled = False
                    servo16bar16.Enabled = False

                    value16_9.Enabled = False
                    value16_10.Enabled = False
                    value16_11.Enabled = False
                    value16_12.Enabled = False
                    value16_13.Enabled = False
                    value16_14.Enabled = False
                    value16_15.Enabled = False
                    value16_16.Enabled = False


                    Pulse16_9.Enabled = False
                    Pulse16_10.Enabled = False
                    Pulse16_11.Enabled = False
                    Pulse16_12.Enabled = False
                    Pulse16_13.Enabled = False
                    Pulse16_14.Enabled = False
                    Pulse16_15.Enabled = False
                    Pulse16_16.Enabled = False

                    speed16_9.Enabled = False
                    speed16_10.Enabled = False
                    speed16_11.Enabled = False
                    speed16_12.Enabled = False
                    speed16_13.Enabled = False
                    speed16_14.Enabled = False
                    speed16_15.Enabled = False
                    speed16_16.Enabled = False

                    Starting9.Enabled = False
                    Starting10.Enabled = False
                    Starting11.Enabled = False
                    Starting12.Enabled = False
                    Starting13.Enabled = False
                    Starting14.Enabled = False
                    Starting15.Enabled = False
                    Starting16.Enabled = False

                    CheckBox9.Enabled = False
                    CheckBox10.Enabled = False
                    CheckBox11.Enabled = False
                    CheckBox12.Enabled = False
                    CheckBox13.Enabled = False
                    CheckBox14.Enabled = False
                    CheckBox15.Enabled = False
                    CheckBox16.Enabled = False

                Catch ex As Exception
                    'MsgBox(ex.ToString)
                    MsgBox("send fail")


                End Try
                ProgressBar3.Increment(5)
            End If
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

            Dim send21() As Byte = {&HA1}
            If transmitt = 1 Then
                Try
                    ComPort.Write(send21, 0, send21.Length)
                Catch ex As Exception
                    MsgBox("sendreport.fail1")

                End Try

            End If
            Try
                System.Threading.Thread.Sleep(300)
                higher_data = ComPort.ReadByte
                position = (higher_data And &H7F) << 6


                System.Threading.Thread.Sleep(300)
                lower_data = ComPort.ReadByte
                position = position Or (lower_data And &H3F)

                servo16bar1.Value = position
                value16_1.Text = position
                Pulse16_1.Text = (position * 0.25 * 10 ^ -3) + 0.5


            Catch ex As Exception
                MsgBox("Rc_report.fail1")
            End Try
            ProgressBar3.Increment(5)
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

            Dim send22() As Byte = {&HA2}
            If transmitt = 1 Then
                Try
                    ComPort.Write(send22, 0, send22.Length)
                Catch ex As Exception
                    MsgBox("sendreport.fail2")

                End Try

            End If
            Try
                System.Threading.Thread.Sleep(300)
                higher_data = ComPort.ReadByte
                position = (higher_data And &H7F) << 6


                System.Threading.Thread.Sleep(300)
                lower_data = ComPort.ReadByte
                position = position Or (lower_data And &H3F)

                servo16bar2.Value = position
                value16_2.Text = position
                Pulse16_2.Text = (position * 0.25 * 10 ^ -3) + 0.5

            Catch ex As Exception
                MsgBox("Rc_report.fail2")
            End Try
            ProgressBar3.Increment(5)
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim send23() As Byte = {&HA3}
            If transmitt = 1 Then
                Try
                    ComPort.Write(send23, 0, send23.Length)
                Catch ex As Exception
                    MsgBox("sendreport.fail3")

                End Try

            End If
            Try
                System.Threading.Thread.Sleep(300)
                higher_data = ComPort.ReadByte
                position = (higher_data And &H7F) << 6


                System.Threading.Thread.Sleep(300)
                lower_data = ComPort.ReadByte
                position = position Or (lower_data And &H3F)

                servo16bar3.Value = position
                value16_3.Text = position
                Pulse16_3.Text = (position * 0.25 * 10 ^ -3) + 0.5


            Catch ex As Exception
                MsgBox("Rc_report.fail3")
            End Try
            ProgressBar3.Increment(5)
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim send24() As Byte = {&HA4}
            If transmitt = 1 Then
                Try
                    ComPort.Write(send24, 0, send24.Length)
                Catch ex As Exception
                    MsgBox("sendreport.fail4")

                End Try

            End If
            Try
                System.Threading.Thread.Sleep(300)
                higher_data = ComPort.ReadByte
                position = (higher_data And &H7F) << 6


                System.Threading.Thread.Sleep(300)
                lower_data = ComPort.ReadByte
                position = position Or (lower_data And &H3F)

                servo16bar4.Value = position
                value16_4.Text = position
                Pulse16_4.Text = (position * 0.25 * 10 ^ -3) + 0.5


            Catch ex As Exception
                MsgBox("Rc_report.fail4")
            End Try
            ProgressBar3.Increment(5)
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim send25() As Byte = {&HA5}
            If transmitt = 1 Then
                Try
                    ComPort.Write(send25, 0, send25.Length)
                Catch ex As Exception
                    MsgBox("sendreport.fail5")

                End Try

            End If
            Try
                System.Threading.Thread.Sleep(300)
                higher_data = ComPort.ReadByte
                position = (higher_data And &H7F) << 6


                System.Threading.Thread.Sleep(300)
                lower_data = ComPort.ReadByte
                position = position Or (lower_data And &H3F)

                servo16bar5.Value = position
                value16_5.Text = position
                Pulse16_5.Text = (position * 0.25 * 10 ^ -3) + 0.5


            Catch ex As Exception
                MsgBox("Rc_report.fail5")
            End Try
            ProgressBar3.Increment(5)
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim send26() As Byte = {&HA6}
            If transmitt = 1 Then
                Try
                    ComPort.Write(send26, 0, send26.Length)
                Catch ex As Exception
                    MsgBox("sendreport.fail6")

                End Try

            End If
            Try
                System.Threading.Thread.Sleep(300)
                higher_data = ComPort.ReadByte
                position = (higher_data And &H7F) << 6


                System.Threading.Thread.Sleep(300)
                lower_data = ComPort.ReadByte
                position = position Or (lower_data And &H3F)

                servo16bar6.Value = position
                value16_6.Text = position
                Pulse16_6.Text = (position * 0.25 * 10 ^ -3) + 0.5


            Catch ex As Exception
                MsgBox("Rc_report.fail6")
            End Try
            ProgressBar3.Increment(5)
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

            Dim send27() As Byte = {&HA7}
            If transmitt = 1 Then
                Try
                    ComPort.Write(send27, 0, send27.Length)
                Catch ex As Exception
                    MsgBox("sendreport.fail7")

                End Try

            End If
            Try
                System.Threading.Thread.Sleep(300)
                higher_data = ComPort.ReadByte
                position = (higher_data And &H7F) << 6


                System.Threading.Thread.Sleep(300)
                lower_data = ComPort.ReadByte
                position = position Or (lower_data And &H3F)

                servo16bar7.Value = position
                value16_7.Text = position
                Pulse16_7.Text = (position * 0.25 * 10 ^ -3) + 0.5


            Catch ex As Exception
                MsgBox("Rc_report.fail7")
            End Try
            ProgressBar3.Increment(5)
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

            Dim send28() As Byte = {&HA8}
            If transmitt = 1 Then
                Try
                    ComPort.Write(send28, 0, send28.Length)
                Catch ex As Exception
                    MsgBox("sendreport.fail8")

                End Try

            End If
            Try
                System.Threading.Thread.Sleep(300)
                higher_data = ComPort.ReadByte
                position = (higher_data And &H7F) << 6


                System.Threading.Thread.Sleep(300)
                lower_data = ComPort.ReadByte
                position = position Or (lower_data And &H3F)

                servo16bar8.Value = position
                value16_8.Text = position
                Pulse16_8.Text = (position * 0.25 * 10 ^ -3) + 0.5


            Catch ex As Exception
                MsgBox("Rc_report.fail8")
            End Try
            ProgressBar3.Increment(5)
        End If


        ProgressBar3.Value = 0
        Timer2.Stop()
        MsgBox("Sucessful!!")

    End Sub
    '8 or 16 servo activate loading
    Private Sub ProgressBar3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ProgressBar3.Click
        Timer2.Start()
    End Sub

  
    Private Sub Button_MarkAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_MarkAll.Click

        CheckBox1.Checked = True
        CheckBox2.Checked = True
        CheckBox3.Checked = True
        CheckBox4.Checked = True
        CheckBox5.Checked = True
        CheckBox6.Checked = True
        CheckBox7.Checked = True
        CheckBox8.Checked = True
        CheckBox9.Checked = True
        CheckBox10.Checked = True
        CheckBox11.Checked = True
        CheckBox12.Checked = True
        CheckBox13.Checked = True
        CheckBox14.Checked = True
        CheckBox15.Checked = True
        CheckBox16.Checked = True

    End Sub

  
End Class
