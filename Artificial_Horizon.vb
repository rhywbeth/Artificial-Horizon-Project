Imports NationalInstruments.DAQmx
Imports System.Windows.Forms.DataVisualization.Charting

Public Class Form1

    Inherits System.Windows.Forms.Form

    Public n As Long = 0
    Public x, t, y, z As Double
    Public file As System.IO.StreamWriter

    Private myTask As Task
    Private reader As AnalogMultiChannelReader
    Private dataColumn As DataColumn()
    Private dataTable As DataTable = New DataTable
    Const dv = "Dev2"

    Private xf As Double = 0
    Private yf As Double = 0
    Private zf As Double = 0

    Private xk, yk, zk, xk1, yk1, zk1 As Double

    Private A As Double = 0

    Private fi, theta, fis, thetas, fi1, theta1, q, qq As Double

    Private xffir, yffir, zffir As Double

    Const ile = 7
    Dim aa() As Double = {0.015252383666206, 0.091687857545886, 0.17120775094111, 0.221852007846798, 0.221852007846798,
                          0.17120775094111, 0.091687857545886, 0.015252383666206}
    Dim xt(ile), yt(ile), zt(ile) As Double
    Dim id As Integer = 0

    Private k As Integer = 0
    Private kk As Integer = 0

    Private gx As Double = 0
    Private gy As Double = 0
    Private gz As Double = 0

    Private ax As Double = 1
    Private bx As Double = 1
    Private ay As Double = 1
    Private by As Double = 1
    Private az As Double = 1
    Private bz As Double = 1

    Private srx As Double
    Private sry As Double
    Private srz As Double

    Private start_time As DateTime

    Private dane() As Double = New Double() {0, 0, 0}

    Dim file1 As System.IO.StreamWriter
    Dim file2 As System.IO.StreamWriter

    Dim kal1 As Double = 1.2419
    Dim kal2 As Double = 1.2213
    Dim kal3 As Double = 1.2657

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick

        n += 1
        t = n / 5

        Dim data() As Double = reader.ReadSingleSample()

        x = data(0)
        y = data(1)
        z = data(2)

        'iir filter
        If n = 1 Then
            xf = x
            yf = y
            zf = z
        End If

        xf += A * (x - xf)
        yf += A * (y - yf)
        zf += A * (z - zf)

        'fir filter
        xt(id) = x
        yt(id) = y
        zt(id) = z
        xffir = 0
        yffir = 0
        zffir = 0
        Dim idk As Integer = id
        For k As Integer = 0 To ile
            idk = k + 1
            If idk > ile Then
                idk -= ile
            End If

            xffir += aa(k) * xt(idk)
            yffir += aa(k) * yt(idk)
            zffir += aa(k) * zt(idk)

        Next
        id += 1
        If (id > ile) Then
            id = 0
        End If

        file.WriteLine("   " & t.ToString("0.00") & "     " & x.ToString("0.00000") & "      " & xf.ToString("0.00000") & "      " & y.ToString("0.00000") & "      " & yf.ToString("0.00000") & "     " & z.ToString("0.00000") & "      " & zf.ToString("0.00000"))
        Label1.Text = x.ToString("0.00000")
        Label7.Text = xf.ToString("0.00000")
        Label2.Text = y.ToString("0.00000")
        Label8.Text = yf.ToString("0.00000")
        Label3.Text = z.ToString("0.00000")
        Label9.Text = zf.ToString("0.00000")
        Label17.Text = xffir.ToString("0.00000")
        Label18.Text = yffir.ToString("0.00000")
        Label19.Text = zffir.ToString("0.00000")

        theta1 = Math.Atan2(x, Math.Sqrt((y * y) + (z * z)))
        fi1 = Math.Atan2(y, z)

        xk = xffir - kal1
        yk = yffir - kal2
        zk = zffir - kal3

        xk1 = xf - kal1
        yk1 = yf - kal2
        zk1 = zf - kal3

        'angles calculation
        If n < 8 Then
            theta = Math.Atan2(xk1, Math.Sqrt((yk1 * yk1) + (zk1 * zk1)))
            fi = Math.Atan2(yk1, zk1)
        Else
            theta = Math.Atan2(xk, Math.Sqrt((yk * yk) + (zk * zk)))
            fi = Math.Atan2(yk, zk)
        End If

        fis = -fi * 180 / Math.PI
        thetas = theta * 180 / Math.PI

        Label10.Text = fis.ToString("0.000") & Chr(176)
        Label11.Text = thetas.ToString("0.000") & Chr(176)

        PictureBox2.Invalidate()

    End Sub

    Private Sub Form1_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        myTask.Dispose()
        file.Close()
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Button2.Enabled = False
        Button1.Enabled = True
        Button3.Enabled = False

        Label1.Visible = False
        Label2.Visible = False
        Label3.Visible = False
        Label7.Visible = False
        Label8.Visible = False
        Label9.Visible = False
        Label17.Visible = False
        Label18.Visible = False
        Label19.Visible = False
        Label10.Visible = False
        Label11.Visible = False

        Label25.Visible = False
        Label26.Visible = False
        Label27.Visible = False
        Label28.Visible = False
        Label29.Visible = False
        Label30.Visible = False

        file = My.Computer.FileSystem.OpenTextFileWriter("~\text.txt", False)
        file.WriteLine("   t           x           xf           y          yf           z           zf")
        Label4.Text = "x"
        Label5.Text = "y"
        Label6.Text = "z"

        A = 1 - Math.Exp(-0.5 / 1)

        Try

            myTask = New Task()
            myTask.AIChannels.CreateVoltageChannel(dv & "/ai4", "", AITerminalConfiguration.Rse, 0, 5, AIVoltageUnits.Volts)
            myTask.AIChannels.CreateVoltageChannel(dv & "/ai5", "", AITerminalConfiguration.Rse, 0, 5, AIVoltageUnits.Volts)
            myTask.AIChannels.CreateVoltageChannel(dv & "/ai6", "", AITerminalConfiguration.Rse, 0, 5, AIVoltageUnits.Volts)
            myTask.Control(TaskAction.Verify)
            reader = New AnalogMultiChannelReader(myTask.Stream)

        Catch ex As DaqException

            MessageBox.Show(ex.Message)

        End Try

    End Sub

    Private Sub PictureBox2_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBox2.Paint

        Dim g As Graphics = e.Graphics
        g.SmoothingMode = Drawing2D.SmoothingMode.HighQuality
        g.ResetTransform()

        'corners rounding graphics
        Dim path As New Drawing2D.GraphicsPath()
        Dim r As Single = 50
        path.AddArc(0, 0, r, r, 180, 90)
        path.AddArc(PictureBox2.Width - r, 0, r, r, 270, 90)
        path.AddArc(PictureBox2.Width - r, PictureBox2.Height - r, r, r, 0, 90)
        path.AddArc(0, PictureBox2.Height - r, r, r, 90, 90)
        path.CloseFigure()
        g.SetClip(path)

        g.Clear(Color.DodgerBlue)

        g.ResetTransform()
        g.TranslateTransform(200, 180)
        g.RotateTransform(fis)
        g.TranslateTransform(0, thetas * 4.3)
        g.FillRectangle(Brushes.Sienna, -300, 0, 600, 500)

        g.ResetTransform()
        Dim w2 As Single = PictureBox2.Width / 2.5
        Dim w3 As Single = PictureBox2.Width / 1.8
        Dim s As Single = PictureBox2.Width / 50
        Dim www As Single = PictureBox2.Width / 38
        g.TranslateTransform(PictureBox2.Width / 2, PictureBox2.Height / 2)
        g.RotateTransform(0)
        g.TranslateTransform(-w2 + s, 0)
        g.DrawLine(New Pen(Color.White, 2), 0, 0, s * 2, 0)
        g.TranslateTransform(+w2 - s, 0)

        g.RotateTransform(15)
        g.DrawLine(New Pen(Color.White, 2), -w2 + s, 0, -w2 + s * 2, 0)
        g.RotateTransform(15)
        g.DrawLine(New Pen(Color.White, 2), -w2 + s, 0, -w2 + s * 2, 0)
        g.RotateTransform(15)
        g.DrawLine(New Pen(Color.White, 2), -w2 + s, 0, -w2 + s * 3, 0)
        g.RotateTransform(15)
        g.DrawLine(New Pen(Color.White, 2), -w2 + s, 0, -w2 + s * 2, 0)
        g.RotateTransform(15)
        g.DrawLine(New Pen(Color.White, 2), -w2 + s, 0, -w2 + s * 2, 0)
        g.RotateTransform(15)

        'g.DrawLine(New Pen(Color.Red, 2), -w2 + s, 0, -w2 + s * 3, 0)
        path = New Drawing2D.GraphicsPath()
        path.AddLine((-w3 / 2) - www * 3, 0, (-w3 / 2) - www * 4, www)
        path.AddLine((-w3 / 2) - www * 4, -www, (-w3 / 2) - www * 4, www)
        path.AddLine((-w3 / 2) - www * 4, -www, (-w3 / 2) - www * 3, 0)
        g.FillRegion(Brushes.White, New Region(path))
        g.DrawLine(New Pen(Color.White, 1), (-w3 / 2) - www * 3, 0, (-w3 / 2) - www * 4, www)
        g.DrawLine(New Pen(Color.White, 1), (-w3 / 2) - www * 4, -www, (-w3 / 2) - www * 4, www)
        g.DrawLine(New Pen(Color.White, 1), (-w3 / 2) - www * 4, -www, (-w3 / 2) - www * 3, 0)

        g.RotateTransform(15)
        g.DrawLine(New Pen(Color.White, 2), -w2 + s, 0, -w2 + s * 2, 0)
        g.RotateTransform(15)
        g.DrawLine(New Pen(Color.White, 2), -w2 + s, 0, -w2 + s * 2, 0)
        g.RotateTransform(15)
        g.DrawLine(New Pen(Color.White, 2), -w2 + s, 0, -w2 + s * 3, 0)
        g.RotateTransform(15)
        g.DrawLine(New Pen(Color.White, 2), -w2 + s, 0, -w2 + s * 2, 0)
        g.RotateTransform(15)
        g.DrawLine(New Pen(Color.White, 2), -w2 + s, 0, -w2 + s * 2, 0)
        g.RotateTransform(15)
        g.DrawLine(New Pen(Color.White, 2), -w2 + s, 0, -w2 + s * 3, 0)

        'silhouette
        g.ResetTransform()
        Dim length As Single = PictureBox2.Width / 4
        Dim notch As Single = PictureBox2.Width / 30
        g.TranslateTransform(PictureBox2.Width / 2, PictureBox2.Height / 2)
        g.DrawLine(New Pen(Color.Yellow, 3), -length + notch * 2, 0, -notch, 0)
        g.DrawLine(New Pen(Color.Yellow, 3), notch, 0, length - notch * 2, 0)
        g.DrawArc(New Pen(Color.Yellow, 3), -notch, -notch, notch * 2, notch * 2, 180, -180)

        'triangle
        g.ResetTransform()
        Dim ww As Single = PictureBox2.Width / 38
        g.TranslateTransform(PictureBox2.Width / 2, PictureBox2.Height / 2)
        g.RotateTransform(-90 + fis)
        path = New Drawing2D.GraphicsPath()
        path.AddLine(w2 - ww * 3, 0, w2 - ww * 4, ww)
        path.AddLine(w2 - ww * 4, -ww, w2 - ww * 4, ww)
        path.AddLine(w2 - ww * 4, -ww, w2 - ww * 3, 0)
        g.FillRegion(Brushes.Yellow, New Region(path))
        g.DrawLine(New Pen(Color.Yellow, 1), w2 - ww * 3, 0, w2 - ww * 4, ww)
        g.DrawLine(New Pen(Color.Yellow, 1), w2 - ww * 4, -ww, w2 - ww * 4, ww)
        g.DrawLine(New Pen(Color.Yellow, 1), w2 - ww * 4, -ww, w2 - ww * 3, 0)

        g.ResetTransform()
        g.ResetClip()
        path = New Drawing2D.GraphicsPath()
        path.AddPie(New Rectangle(ww * 3, ww * 3, PictureBox2.Width - ww * 6, PictureBox2.Height - ww * 6), 0, 360)
        g.SetClip(path)

        'vertical lines
        g.TranslateTransform(PictureBox2.Width / 2, PictureBox2.Height / 2)
        g.RotateTransform(fis)
        g.TranslateTransform(0, thetas * 4.3)
        For i As Integer = -80 To 80 Step 10
            drawpitchline(g, i)
        Next i

    End Sub

    Private Function pitch_to_pix(ByVal pitch As Double) As Integer
        Return pitch / 40 * PictureBox2.Height / 2
    End Function

    Private Sub drawpitchline(ByVal g As Graphics, ByVal pitch As Double)
        Dim w As Single = PictureBox2.Width / 10
        g.DrawLine(Pens.White, -w, pitch_to_pix(-pitch + 5), w, pitch_to_pix(-pitch + 5))
        g.DrawLine(Pens.White, -w * 5 / 3, pitch_to_pix(-pitch), w * 5 / 3, pitch_to_pix(-pitch))
        g.DrawString(pitch, PictureBox2.Font, Brushes.White, -w * 75 / 30, pitch_to_pix(-pitch) - 5)
        g.DrawString(pitch, PictureBox2.Font, Brushes.White, w * 2, pitch_to_pix(-pitch) - 5)
    End Sub

    Private Sub drawrollline(ByVal g As Graphics, ByVal a As Single)
        Dim w2 As Single = PictureBox2.Width / 2
        g.RotateTransform(a + 90)
        g.TranslateTransform(-w2 + 10, 0)
        g.DrawLine(Pens.White, 0, 0, 20, 0)
        g.TranslateTransform(10, 5)
        g.RotateTransform(-a - 90)
        g.DrawString("" & (a) & "°", New System.Drawing.Font("sans-serif", 9), Brushes.White, 0, 0)
        g.RotateTransform(+90 + a)
        g.TranslateTransform(-10, -5)
        g.TranslateTransform(+w2 - 10, 0)
        g.RotateTransform(-a - 90)
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim start_t As DateTime = Now
        Me.start_time = start_t
        Button2.Enabled = True
        Button1.Enabled = False
        Button3.Enabled = True

        Timer1.Start()
        Timer1.Interval = 100

        Label1.Visible = True
        Label2.Visible = True
        Label3.Visible = True
        Label7.Visible = True
        Label8.Visible = True
        Label9.Visible = True
        Label17.Visible = True
        Label18.Visible = True
        Label19.Visible = True
        Label10.Visible = True
        Label11.Visible = True
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Button1.Enabled = True
        Button2.Enabled = False
        Button3.Enabled = False

        Timer1.Stop()

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click

        Me.kk = 1

        file1 = My.Computer.FileSystem.OpenTextFileWriter("~\kalibr.txt", False)
        file1.WriteLine("  axis x   " & "   axis y   " & "   axis z " & vbCrLf)
       
        Label25.Visible = True
        Label26.Visible = True
        Label27.Visible = True
        Label28.Visible = True
        Label29.Visible = True
        Label30.Visible = True

        'calibration
        Dim dane_kalibracji()() As Double = New Double(60)() {}
        For i As Integer = 0 To 5
            MessageBox.Show("Sensor position  " + (i + 1).ToString)
            For j As Integer = 0 To 9
                Dim data() As Double = reader.ReadSingleSample()
                dane_kalibracji(i * 10 + j) = data
                file1.WriteLine(data(0).ToString("0.0000000") & " " & data(1).ToString("0.0000000") & " " & data(2).ToString("0.0000000"))
                System.Threading.Thread.Sleep(50)
            Next
            file1.WriteLine(" ")
        Next
        file1.Close()

        Label25.Visible = False
        Label26.Visible = False
        Label27.Visible = False
        Label28.Visible = False
        Label29.Visible = False
        Label30.Visible = False

        'sorting
        Dim p As Double
        For i As Integer = 0 To 2
            For j As Integer = 0 To 59
                For k As Integer = 0 To 59
                    If dane_kalibracji(j)(i) > dane_kalibracji(k)(i) Then
                        p = dane_kalibracji(j)(i)
                        dane_kalibracji(j)(i) = dane_kalibracji(k)(i)
                        dane_kalibracji(k)(i) = p
                    End If
                Next
            Next
        Next

        '10 mix/min values
        Dim max_x() As Double = New Double(10) {}
        Dim max_y() As Double = New Double(10) {}
        Dim max_z() As Double = New Double(10) {}
        Dim min_x() As Double = New Double(10) {}
        Dim min_y() As Double = New Double(10) {}
        Dim min_z() As Double = New Double(10) {}

        For i As Integer = 0 To 9
            max_x(i) = dane_kalibracji(i)(0)
            max_y(i) = dane_kalibracji(i)(1)
            max_z(i) = dane_kalibracji(i)(2)
            min_x(i) = dane_kalibracji(59 - i)(0)
            min_y(i) = dane_kalibracji(59 - i)(1)
            min_z(i) = dane_kalibracji(59 - i)(2)
        Next

        'average from 10 max/min values
        Dim suma_max_x As Double = 0
        Dim suma_max_y As Double = 0
        Dim suma_max_z As Double = 0
        Dim suma_min_x As Double = 0
        Dim suma_min_y As Double = 0
        Dim suma_min_z As Double = 0

        For i As Integer = 0 To 9
            suma_max_x += max_x(i)
            suma_max_y += max_y(i)
            suma_max_z += max_z(i)
            suma_min_x += min_x(i)
            suma_min_y += min_y(i)
            suma_min_z += min_z(i)
        Next

        Dim max_x_sr As Double = suma_max_x / 10
        Dim max_y_sr As Double = suma_max_y / 10
        Dim max_z_sr As Double = suma_max_z / 10
        Dim min_x_sr As Double = suma_min_x / 10
        Dim min_y_sr As Double = suma_min_y / 10
        Dim min_z_sr As Double = suma_min_z / 10

        Me.kal1 = (max_x_sr + min_x_sr) / 2
        Me.kal2 = (max_y_sr + min_y_sr) / 2
        Me.kal3 = (max_z_sr + min_z_sr) / 2

    End Sub
End Class
