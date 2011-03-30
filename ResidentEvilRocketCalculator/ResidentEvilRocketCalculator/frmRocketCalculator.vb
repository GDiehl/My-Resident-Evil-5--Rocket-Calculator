'Application Name:  Resident Evil 5 Rocket Calculator
'Author:            Greg Diehl
'Date:              3/30/2011
'Purpose:           This is a simple calculator that allows the user to input their completed time
'                   for each level and then calculate how much time they need to shave off in order
'                   to unlock the infinite rockets.  If they are under the 5 hours required, then it 
'                   will post a message and give the total playing time.


Option Strict On
Imports System.IO
Public Class frmRocketCalculator
    Dim _intTotalSeconds As Integer = 0
    Const _intFiveHours As Integer = (5 * 60) * 60
    Private Sub frmRocketCalculator_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call LoadFile(Me)
    End Sub
    Private Sub btnCalculate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCalculate.Click

        'Clear any totals or messages from previous Calculate clicks
        Call Clear()

        'Submit the text in the Hours text boxes to the HoursToSeconds Function
        'and then add those seconds to the _intTotalSeconds variable

        _intTotalSeconds += HoursToSeconds(txtHours1_1.Text)
        _intTotalSeconds += HoursToSeconds(txtHours1_2.Text)
        _intTotalSeconds += HoursToSeconds(txtHours2_1.Text)
        _intTotalSeconds += HoursToSeconds(txtHours2_2.Text)
        _intTotalSeconds += HoursToSeconds(txtHours2_3.Text)
        _intTotalSeconds += HoursToSeconds(txtHours3_1.Text)
        _intTotalSeconds += HoursToSeconds(txtHours3_2.Text)
        _intTotalSeconds += HoursToSeconds(txtHours3_3.Text)
        _intTotalSeconds += HoursToSeconds(txtHours4_1.Text)
        _intTotalSeconds += HoursToSeconds(txtHours4_2.Text)
        _intTotalSeconds += HoursToSeconds(txtHours5_1.Text)
        _intTotalSeconds += HoursToSeconds(txtHours5_2.Text)
        _intTotalSeconds += HoursToSeconds(txtHours5_3.Text)
        _intTotalSeconds += HoursToSeconds(txtHours6_1.Text)
        _intTotalSeconds += HoursToSeconds(txtHours6_2.Text)
        _intTotalSeconds += HoursToSeconds(txtHours6_3.Text)

        'Submit the text in the Minutes text boxes to the MinutesToSeconds Function
        'and then add those seconds to the _intTotalSeconds variable
        _intTotalSeconds += MinutesToSeconds(txtMinutes1_1.Text)
        _intTotalSeconds += MinutesToSeconds(txtMinutes1_2.Text)
        _intTotalSeconds += MinutesToSeconds(txtMinutes2_1.Text)
        _intTotalSeconds += MinutesToSeconds(txtMinutes2_2.Text)
        _intTotalSeconds += MinutesToSeconds(txtMinutes2_3.Text)
        _intTotalSeconds += MinutesToSeconds(txtMinutes3_1.Text)
        _intTotalSeconds += MinutesToSeconds(txtMinutes3_2.Text)
        _intTotalSeconds += MinutesToSeconds(txtMinutes3_3.Text)
        _intTotalSeconds += MinutesToSeconds(txtMinutes4_1.Text)
        _intTotalSeconds += MinutesToSeconds(txtMinutes4_2.Text)
        _intTotalSeconds += MinutesToSeconds(txtMinutes5_1.Text)
        _intTotalSeconds += MinutesToSeconds(txtMinutes5_2.Text)
        _intTotalSeconds += MinutesToSeconds(txtMinutes5_3.Text)
        _intTotalSeconds += MinutesToSeconds(txtMinutes6_1.Text)
        _intTotalSeconds += MinutesToSeconds(txtMinutes6_2.Text)
        _intTotalSeconds += MinutesToSeconds(txtMinutes6_3.Text)

        'Submit the text in the Seconds text boxes to the ConvertSeconds Function
        'and then add those seconds to the _intTotalSeconds variable
        _intTotalSeconds += ConvertSeconds(txtSeconds1_1.Text)
        _intTotalSeconds += ConvertSeconds(txtSeconds1_2.Text)
        _intTotalSeconds += ConvertSeconds(txtSeconds2_1.Text)
        _intTotalSeconds += ConvertSeconds(txtSeconds2_2.Text)
        _intTotalSeconds += ConvertSeconds(txtSeconds2_3.Text)
        _intTotalSeconds += ConvertSeconds(txtSeconds3_1.Text)
        _intTotalSeconds += ConvertSeconds(txtSeconds3_2.Text)
        _intTotalSeconds += ConvertSeconds(txtSeconds3_3.Text)
        _intTotalSeconds += ConvertSeconds(txtSeconds4_1.Text)
        _intTotalSeconds += ConvertSeconds(txtSeconds4_2.Text)
        _intTotalSeconds += ConvertSeconds(txtSeconds5_1.Text)
        _intTotalSeconds += ConvertSeconds(txtSeconds5_2.Text)
        _intTotalSeconds += ConvertSeconds(txtSeconds5_3.Text)
        _intTotalSeconds += ConvertSeconds(txtSeconds6_1.Text)
        _intTotalSeconds += ConvertSeconds(txtSeconds6_2.Text)
        _intTotalSeconds += ConvertSeconds(txtSeconds6_3.Text)

        'Check to see if the total play time is under 5 hours
        'Unhide the appropriate Message lables and the total time
        If _intTotalSeconds <= _intFiveHours Then
            Me.lblMessage.Visible = True
            Me.lblTotalTime.Text = ConvertToTime(_intTotalSeconds)
            Me.lblTotalTimeHeading.Visible = True
            Me.lblTotalTime.Visible = True
        Else
            Dim intGoalTime As Integer = 0

            Me.lblMessage.Text = "Time To Cut:"
            Me.lblMessage.Visible = True
            intGoalTime = (_intTotalSeconds - _intFiveHours)
            Me.lblGoalTime.Text = ConvertToTime(intGoalTime)
            Me.lblGoalTime.Visible = True
            Me.lblTotalTime.Text = ConvertToTime(_intTotalSeconds)
            Me.lblTotalTimeHeading.Visible = True
            Me.lblTotalTime.Visible = True
        End If

    End Sub
    Private Function HoursToSeconds(ByVal Hours As String) As Integer
        'This function checks to ensure that what was entered into the Hours text box
        'was a number and then converts the string to an integer.  Then is converts the 
        'Hours to seconds.

        Dim Total As Integer

        Try
            If Hours = "" Or Hours = " " Then
                Hours = "0"
            ElseIf Hours = "M" Or Hours = "m" Then
                MsgBox("Hello, Michelle!", , "Why, hello there.")
                Hours = "0"
            End If
            If IsNumeric(Hours) Then
                If CInt(Hours) < 0 Then
                    MsgBox("No negative numbers!", , "Error")
                End If
                Total = (CInt(Hours) * 60) * 60
            End If


            Return Total
        Catch ex As Exception
            MsgBox("Please enter a valid number!", , "Error")
        End Try
    End Function
    Private Function MinutesToSeconds(ByVal Minutes As String) As Integer
        'This function checks to ensure that what was entered into the Minutes text box
        'was a number and then converts the string to an integer.  Then is converts the 
        'minutes to seconds.

        Dim Total As Integer

        Try
            If Minutes = "" Or Minutes = " " Then
                Minutes = "0"
            ElseIf Minutes = "M" Or Minutes = "m" Then
                MsgBox("Hello, Michelle!", , "Why, hello there.")

            End If
            If IsNumeric(Minutes) Then
                If CInt(Minutes) < 0 Then
                    MsgBox("No negative numbers!", , "Error")
                End If
                Total = CInt(Minutes) * 60
            End If

            Return Total
        Catch ex As Exception
            MsgBox("Please enter a valid number!", , "Error")
        End Try
    End Function
    Private Function ConvertSeconds(ByVal Seconds As String) As Integer
        'This function checks to ensure that what was entered into the Seconds text box
        'was a number and then converts the string to an integer.

        Try
            If Seconds = "" Or Seconds = " " Then
                Seconds = "0"
            ElseIf Seconds = "M" Or Seconds = "m" Then
                MsgBox("Hello, Michelle!", , "Why, hello there.")
                Seconds = "0"

            End If
            If IsNumeric(Seconds) Then
                Convert.ToInt32(Seconds)
            End If


            Return CInt(Seconds)
        Catch ex As Exception
            MsgBox("Please enter a valid number!", , "Error")
        End Try
    End Function
    Private Function ConvertToTime(ByVal intTotalSeconds As Integer) As String
        'Converts the Seconds back to hours:minutes:seconds and returns the result as a string

        Dim intHours As Integer
        Dim intMinutes As Integer
        Dim intSeconds As Integer
        Dim strTime As String

        intHours = (intTotalSeconds \ 60) \ 60
        intMinutes = (intTotalSeconds \ 60) Mod 60
        intSeconds = intTotalSeconds Mod 60

        strTime = intHours & ":" & intMinutes & ":" & intSeconds

        Return strTime
    End Function
    Private Sub btnClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClear.Click
        'The clear fuction resets all messages, hides the messages, resets the Seconds counter (using 
        'the Clear() sub-routine), and resets all of the text boxes to null.

        Call Clear()
        txtHours1_1.Text = ""
        txtHours1_2.Text = ""
        txtHours2_1.Text = ""
        txtHours2_2.Text = ""
        txtHours2_3.Text = ""
        txtHours3_1.Text = ""
        txtHours3_2.Text = ""
        txtHours3_3.Text = ""
        txtHours4_1.Text = ""
        txtHours4_2.Text = ""
        txtHours5_1.Text = ""
        txtHours5_2.Text = ""
        txtHours5_3.Text = ""
        txtHours6_1.Text = ""
        txtHours6_2.Text = ""
        txtHours6_3.Text = ""

        txtMinutes1_1.Text = ""
        txtMinutes1_2.Text = ""
        txtMinutes2_1.Text = ""
        txtMinutes2_2.Text = ""
        txtMinutes2_3.Text = ""
        txtMinutes3_1.Text = ""
        txtMinutes3_2.Text = ""
        txtMinutes3_3.Text = ""
        txtMinutes4_1.Text = ""
        txtMinutes4_2.Text = ""
        txtMinutes5_1.Text = ""
        txtMinutes5_2.Text = ""
        txtMinutes5_3.Text = ""
        txtMinutes6_1.Text = ""
        txtMinutes6_2.Text = ""
        txtMinutes6_3.Text = ""

        txtSeconds1_1.Text = ""
        txtSeconds1_2.Text = ""
        txtSeconds2_1.Text = ""
        txtSeconds2_2.Text = ""
        txtSeconds2_3.Text = ""
        txtSeconds3_1.Text = ""
        txtSeconds3_2.Text = ""
        txtSeconds3_3.Text = ""
        txtSeconds4_1.Text = ""
        txtSeconds4_2.Text = ""
        txtSeconds5_1.Text = ""
        txtSeconds5_2.Text = ""
        txtSeconds5_3.Text = ""
        txtSeconds6_1.Text = ""
        txtSeconds6_2.Text = ""
        txtSeconds6_3.Text = ""
    End Sub
    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        'Exits application
        Call SaveFile(Me)
        Me.Close()
    End Sub
    Public Sub Clear()
        'Resets the Total Seconds counter, resets defualt Message, and hides all lables.

        _intTotalSeconds = 0
        Me.lblMessage.Visible = False
        Me.lblGoalTime.Visible = False
        Me.lblTotalTimeHeading.Visible = False
        Me.lblTotalTime.Visible = False
        Me.lblMessage.Text = "You Have" & vbCrLf & _
                            "Unlocked" & vbCrLf & _
                            "Rockets"
    End Sub
    Private Sub LoadFile(ByVal Page As Control)
        If File.Exists("Data.re5") Then
            Dim fsFillData As New FileStream("Data.re5", FileMode.Open, FileAccess.Read)
            Dim srFillData As New StreamReader(fsFillData)
            Dim FillData(47) As String
            Dim intCounter As Integer = 0

            Do Until intCounter > FillData.Length - 1
                FillData(intCounter) = srFillData.ReadLine.ToString
                intCounter += 1
            Loop

            intCounter = 0

            For Each cntrl As Control In Page.Controls
                If TypeOf cntrl Is TextBox Then
                    cntrl.Text = FillData(intCounter)
                    intCounter += 1
                End If
            Next
            srFillData.Close()
            fsFillData.Close()
        End If
    End Sub
    Private Sub SaveFile(ByVal Page As Control)
        Dim fsSaveData As New FileStream("Data.re5", FileMode.OpenOrCreate, FileAccess.Write)
        Dim swSaveData As New StreamWriter(fsSaveData)
        Dim FillValue(47) As String
        Dim intCounter As Integer = 0

        For Each cntrl As Control In Page.Controls
            If TypeOf cntrl Is TextBox Then
                FillValue(intCounter) = cntrl.Text
                intCounter += 1
            End If
        Next

        intCounter = 0

        Do Until intCounter > (FillValue.Length - 1)
            If FillValue(intCounter) = "" Then
                FillValue(intCounter) = " "
            End If
            swSaveData.WriteLine(FillValue(intCounter))
            intCounter += 1
        Loop

        swSaveData.Close()
        fsSaveData.Close()
    End Sub
End Class
