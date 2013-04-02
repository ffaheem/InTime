Imports System
Imports System.Threading
Imports System.Windows.Controls
Imports Microsoft.Phone.Controls
Imports Microsoft.Phone.Shell
Imports System.Globalization
Imports System.Windows.Threading


Partial Public Class MainPage
    Inherits PhoneApplicationPage


    Dim nameofmonths(13) As String
    Dim linela(400) As Line
    Dim monthlabel(13) As TextBlock
    Dim daylabel(32) As TextBlock
    Dim resoltuion As Double
    Dim countdays As Integer
    Dim linewi(50) As Line
    Dim hourlabel(25) As TextBlock
    Dim mainhorline As New Line
    Private WithEvents timer As DispatcherTimer = New DispatcherTimer()
    Dim speed As New Point
    Dim deceleration As New Double
    ' Constructor
    Public Sub New()
        InitializeComponent()

        SupportedOrientations = SupportedPageOrientation.Landscape

        initsetup()

        newHourLine()

        mainline()

        newTimelineMonth(10, 2013)
    End Sub
   

    Private Sub initsetup()
        nameofmonths(1) = "Jan"
        nameofmonths(2) = "Feb"
        nameofmonths(3) = "Mar"
        nameofmonths(4) = "Apr"
        nameofmonths(5) = "May"
        nameofmonths(6) = "Jun"
        nameofmonths(7) = "Jul"
        nameofmonths(8) = "Aug"
        nameofmonths(9) = "Sept"
        nameofmonths(10) = "Oct"
        nameofmonths(11) = "Nov"
        nameofmonths(12) = "Dec"

        deceleration = 10

        timer.Interval = New TimeSpan(5000)

        resoltuion = 2400 ' space between lines

        'initializing ui components
        For x As Integer = 1 To 12
            monthlabel(x) = New TextBlock
        Next

        For x As Integer = 1 To 400
            linela(x) = New Line
        Next

        For x As Integer = 1 To 48
            linewi(x) = New Line
        Next

        For x As Integer = 1 To 24
            hourlabel(x) = New TextBlock
        Next

        For x As Integer = 1 To 31
            daylabel(x) = New TextBlock
        Next

    End Sub

    Private Sub mainline()
        'creating horizontal timeline
        With mainhorline
            .X1 = 0
            .X2 = 800
            .Y1 = 200
            .Y2 = 200
            .Stroke = New SolidColorBrush(Colors.Black)
        End With
        r2.Children.Add(mainhorline)
    End Sub

    Private Sub newHourLine()
        Dim lenght As Double


        For x As Integer = 1 To 48
            With linewi(x)
                If x Mod 2 = 0 Then
                    lenght = 20
                    With hourlabel(x / 2)
                        .Text = (x / 2).ToString & "H"
                        .FontSize = 20
                        .Foreground = New SolidColorBrush(Colors.Gray)
                        .Margin = New Thickness(lenght, x * 50, 0, 0)
                    End With
                    r2.Children.Add(hourlabel(x / 2))
                Else
                    lenght = 10
                End If

                .Stroke = New SolidColorBrush(Colors.Gray)
                .X1 = lenght
                .X2 = 0
                .Y1 = x * 50
                .Y2 = x * 50

            End With
            r2.Children.Add(linewi(x))
        Next
    End Sub




    Private Sub newTimeLine(year As Integer)

        Dim greg As New GregorianCalendar
        Dim lenght As Double

        Dim daysinmonth As Integer

        countdays = 1



        'start adding vertical lines

        For x As Integer = 1 To 12
            daysinmonth = greg.GetDaysInMonth(year, x, GregorianCalendar.ADEra)
            'add month label
            With monthlabel(x)
                .Text =  nameofmonths(x)
                .Foreground = New SolidColorBrush(Colors.Gray)
                .FontSize = 36
                .Margin = New Thickness(countdays / greg.GetDaysInYear(year) * resoltuion, 0, 0, 0)
            End With
            r2.Children.Add(monthlabel(x))

            For y As Integer = 1 To daysinmonth

                With linela(countdays)
                    If y = 1 Then
                        lenght = 100

                    ElseIf y Mod 7 = 0 Then
                        lenght = 30
                    Else
                        lenght = 10
                    End If

                    .Stroke = New SolidColorBrush(Colors.Gray)
                    .X1 = countdays / greg.GetDaysInYear(year) * resoltuion
                    .X2 = countdays / greg.GetDaysInYear(year) * resoltuion
                    .Y1 = mainhorline.Y1 + lenght
                    .Y2 = mainhorline.Y1 - lenght

                End With
                r2.Children.Add(linela(countdays))

                countdays = countdays + 1
            Next

        Next

    End Sub

    Private Sub newTimelineMonth(month As Integer, year As Integer)

        Dim greg As New GregorianCalendar
        Dim lenght As Double

        countdays = greg.GetDaysInMonth(year, month, GregorianCalendar.ADEra)

        For y As Integer = 1 To countdays

            With linela(y)
                If y = 1 Then
                    lenght = 100
                ElseIf y Mod 7 = 0 Then
                    lenght = 30
                Else
                    lenght = 10
                End If

                .Stroke = New SolidColorBrush(Colors.Gray)
                .X1 = y / countdays * resoltuion - 60
                .X2 = y / countdays * resoltuion - 60
                .Y1 = mainhorline.Y1 + lenght
                .Y2 = mainhorline.Y1 - lenght

            End With
            r2.Children.Add(linela(y))

            Dim z As String
            z = greg.GetDayOfWeek(New Date(year, month, y)).ToString

            With daylabel(y)
                .Text = y.ToString + " " + z.Substring(0, 3)
                .Foreground = New SolidColorBrush(Colors.Gray)
                .FontSize = 16
                .Margin = New Thickness(y / countdays * resoltuion - 80, mainhorline.Y1 - lenght - 20, 0, 0)
            End With
            r2.Children.Add(daylabel(y))


        Next

    End Sub

    Private Sub zoominout(zoomscale As Double)
        If (linela(countdays - 1).X1 - linela(1).X1) < 8000 And (linela(countdays - 1).X1 - linela(1).X1) > 1000 Then 'setting zoom limits

            For y As Integer = 1 To countdays

                linela(y).X1 = linela(y).X1 * zoomscale
                linela(y).X2 = linela(y).X2 * zoomscale

            Next
            For x As Integer = 1 To 12
                monthlabel(x).Margin = New Thickness(monthlabel(x).Margin.Left * zoomscale, 0, 0, 0)
            Next
            For z As Integer = 1 To 31
                daylabel(z).Margin = New Thickness((daylabel(z).Margin.Left) * zoomscale, daylabel(z).Margin.Top, 0, 0)
            Next
        ElseIf (linela(countdays - 1).X1 - linela(1).X1) > 8000 Then 'if zoom exceeds this limit a bounce back effect will take place
            zoomscale = 0.9
            For y As Integer = 1 To countdays

                linela(y).X1 = linela(y).X1 * zoomscale
                linela(y).X2 = linela(y).X2 * zoomscale

            Next
            For x As Integer = 1 To 12
                monthlabel(x).Margin = New Thickness(monthlabel(x).Margin.Left * zoomscale, 0, 0, 0)
            Next
            For z As Integer = 1 To 31
                daylabel(z).Margin = New Thickness((daylabel(z).Margin.Left) * zoomscale, daylabel(z).Margin.Top, 0, 0)
            Next

        ElseIf (linela(countdays - 1).X1 - linela(1).X1) < 1000 Then 'same here
            zoomscale = 1.1
            For y As Integer = 1 To countdays

                linela(y).X1 = linela(y).X1 * zoomscale
                linela(y).X2 = linela(y).X2 * zoomscale

            Next
            For x As Integer = 1 To 12
                monthlabel(x).Margin = New Thickness(monthlabel(x).Margin.Left * zoomscale, 0, 0, 0)
            Next
            For z As Integer = 1 To 31
                daylabel(z).Margin = New Thickness((daylabel(z).Margin.Left) * zoomscale, daylabel(z).Margin.Top, 0, 0)
            Next
        End If
    End Sub


    Private Sub scrollhorizontal(swipe As Double)
        If (linela(1).X1 + swipe) > 20 Then
            ' if line reaches end
        ElseIf (linela(countdays - 1).X1 + swipe) < 780 Then
            'same here
        Else
            For y As Integer = 1 To countdays
                linela(y).X1 = linela(y).X1 + swipe
                linela(y).X2 = linela(y).X2 + swipe
            Next
            For x As Integer = 1 To 12
                monthlabel(x).Margin = New Thickness(monthlabel(x).Margin.Left + swipe, 0, 0, 0)
            Next
            For z As Integer = 1 To 31
                daylabel(z).Margin = New Thickness(daylabel(z).Margin.Left + swipe, daylabel(z).Margin.Top, 0, 0)
            Next
        End If
    End Sub

    Private Sub scrollvertical(swipe As Double)
        If (linewi(1).Y1 + swipe) > 10 Then
            ' if line reaches end
        ElseIf (linewi(48).Y1 + swipe) < 380 Then
            'same here
        Else
            For y As Integer = 1 To 48
                linewi(y).Y1 = linewi(y).Y1 + swipe
                linewi(y).Y2 = linewi(y).Y2 + swipe
            Next
            For x As Integer = 1 To 24
                hourlabel(x).Margin = New Thickness(20, hourlabel(x).Margin.Top + swipe, 0, 0)
            Next
        End If
    End Sub

    

    Private Sub r1_ManipulationDelta(sender As Object, e As ManipulationDeltaEventArgs) Handles r1.ManipulationDelta


        Dim zoomscale As Double
        zoomscale = 1

        Try
            zoomscale = e.PinchManipulation.DeltaScale ' if pinch manipulation then reorute to zoom method
            zoominout(zoomscale)
        Catch ex As Exception

        End Try

        If zoomscale = 1 Then
            scrollhorizontal(e.DeltaManipulation.Translation.X)
            scrollvertical(e.DeltaManipulation.Translation.Y)
        End If




    End Sub

    Private Sub r1_ManipulationCompleted(sender As Object, e As ManipulationCompletedEventArgs) Handles r1.ManipulationCompleted
        speed = e.FinalVelocities().LinearVelocity

        timer.Start()
    End Sub

    Private Sub TimerClick(ByVal sender As System.Object, ByVal e As EventArgs) Handles timer.Tick

        If speed.X > 0 Or speed.Y > 0 Then
            scrollhorizontal(speed.X / 60)
            scrollvertical(speed.Y / 60)
            speed.X = speed.X / 1.05
            speed.Y = speed.Y / 1.05
            If speed.X < 1 Then
                speed.X = 0
            End If
            If speed.Y < 1 Then
                speed.Y = 0
            End If
        Else
            scrollhorizontal(speed.X / 60)
            scrollvertical(speed.Y / 60)
            speed.X = speed.X / 1.05
            speed.Y = speed.Y / 1.05
            If speed.X > -1 Then
                speed.X = 0
            End If
            If speed.Y > -1 Then
                speed.Y = 0
            End If

        End If
        If speed.Y = 0 And speed.X = 0 Then
            timer.Stop()
        End If

    End Sub

    Private Sub r1_ManipulationStarted(sender As Object, e As ManipulationStartedEventArgs) Handles r1.ManipulationStarted

        timer.Stop()
        speed.X = 0
        speed.Y = 0

    End Sub
End Class