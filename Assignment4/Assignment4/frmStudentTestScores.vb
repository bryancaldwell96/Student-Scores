Option Strict On

Imports System.IO

Public Class frmStudentTestScores

    'dim variables
    Dim strStudent1 As String
    Dim strStudent2 As String
    Dim strStudent3 As String
    Dim strStudent4 As String
    Dim strStudent5 As String
    Dim strStudent6 As String

    Dim dblStudent1TestScore(4) As Double
    Dim dblStudent2TestScore(4) As Double
    Dim dblStudent3TestScore(4) As Double
    Dim dblStudent4TestScore(4) As Double
    Dim dblStudent5TestScore(4) As Double
    Dim dblStudent6TestScore(4) As Double

    Dim dblStudent1Average As Double
    Dim dblStudent2Average As Double
    Dim dblStudent3Average As Double
    Dim dblStudent4Average As Double
    Dim dblStudent5Average As Double
    Dim dblStudent6Average As Double

    Dim dblCumulativeTestScores1 As Double
    Dim dblCumulativeTestScores2 As Double
    Dim dblCumulativeTestScores3 As Double
    Dim dblCumulativeTestScores4 As Double
    Dim dblCumulativeTestScores5 As Double
    Dim dblCumulativeTestScores6 As Double

    Dim StudentTestScoresFile As StreamWriter

    Public Function StudentTestScoreCheck(ByVal TestScore As Double) As Boolean

        'test score input validation, check to see if it's a number
        If IsNumeric(TestScore) Then
            'if test score is greater than 0 check to see if it's less than 100
            If TestScore >= 0 Then
                'if test score is a number, greater than 0 and less than 100 than return true
                If TestScore <= 100 Then

                    Return True

                Else
                    'if test score is not less than 100 return false
                    MessageBox.Show("Please enter only numbers between 0 and 100 for Student's test scores.")
                    Return False

                End If

            Else
                'if test score is not greater than 0 return false
                MessageBox.Show("Please enter only numbers between 0 and 100 for Student's test scores.")
                Return False

            End If

        Else
            'if test score is not numeric return false
            MessageBox.Show("Please enter only numbers for Student's test scores.")
            Return False

        End If

    End Function

    Public Function StudentAverages(ByVal Average As Double, TestScore() As Double, CumulativeTestScores As Double) As Double

        CumulativeTestScores = TestScore(0) + TestScore(1) + TestScore(2) + TestScore(3) + TestScore(4)
        Average = CumulativeTestScores / 5

        Return Average

    End Function

    Public Function RunTestScoreCheck() As Boolean

        'set student test score values if check doesn't fail
        If StudentTestScoreCheck(CDbl(txtTestStudent1Score1.Text)) = False Then
            Return False
        Else
            dblStudent1TestScore(0) = CDbl(txtTestStudent1Score1.Text)
        End If
        If StudentTestScoreCheck(CDbl(txtTestStudent1Score2.Text)) = False Then
            Return False
        Else
            dblStudent1TestScore(1) = CDbl(txtTestStudent1Score2.Text)
        End If
        If StudentTestScoreCheck(CDbl(txtTestStudent1Score3.Text)) = False Then
            Return False
        Else
            dblStudent1TestScore(2) = CDbl(txtTestStudent1Score3.Text)
        End If
        If StudentTestScoreCheck(CDbl(txtTestStudent1Score4.Text)) = False Then
            Return False
        Else
            dblStudent1TestScore(3) = CDbl(txtTestStudent1Score4.Text)
        End If
        If StudentTestScoreCheck(CDbl(txtTestStudent1Score5.Text)) = False Then
            Return False
        Else
            dblStudent1TestScore(4) = CDbl(txtTestStudent1Score5.Text)
        End If

        If StudentTestScoreCheck(CDbl(txtTestStudent2Score1.Text)) = False Then
            Return False
        Else
            dblStudent2TestScore(0) = CDbl(txtTestStudent2Score1.Text)
        End If
        If StudentTestScoreCheck(CDbl(txtTestStudent2Score2.Text)) = False Then
            Return False
        Else
            dblStudent2TestScore(1) = CDbl(txtTestStudent2Score2.Text)
        End If
        If StudentTestScoreCheck(CDbl(txtTestStudent2Score3.Text)) = False Then
            Return False
        Else
            dblStudent2TestScore(2) = CDbl(txtTestStudent2Score3.Text)
        End If
        If StudentTestScoreCheck(CDbl(txtTestStudent2Score4.Text)) = False Then
            Return False
        Else
            dblStudent2TestScore(3) = CDbl(txtTestStudent2Score4.Text)
        End If
        If StudentTestScoreCheck(CDbl(txtTestStudent2Score5.Text)) = False Then
            Return False
        Else
            dblStudent2TestScore(4) = CDbl(txtTestStudent2Score5.Text)
        End If

        If StudentTestScoreCheck(CDbl(txtTestStudent3Score1.Text)) = False Then
            Return False
        Else
            dblStudent3TestScore(0) = CDbl(txtTestStudent3Score1.Text)
        End If
        If StudentTestScoreCheck(CDbl(txtTestStudent3Score2.Text)) = False Then
            Return False
        Else
            dblStudent3TestScore(1) = CDbl(txtTestStudent3Score2.Text)
        End If
        If StudentTestScoreCheck(CDbl(txtTestStudent3Score3.Text)) = False Then
            Return False
        Else
            dblStudent3TestScore(2) = CDbl(txtTestStudent3Score3.Text)
        End If
        If StudentTestScoreCheck(CDbl(txtTestStudent3Score4.Text)) = False Then
            Return False
        Else
            dblStudent3TestScore(3) = CDbl(txtTestStudent3Score4.Text)
        End If
        If StudentTestScoreCheck(CDbl(txtTestStudent3Score5.Text)) = False Then
            Return False
        Else
            dblStudent3TestScore(4) = CDbl(txtTestStudent3Score5.Text)
        End If

        If StudentTestScoreCheck(CDbl(txtTestStudent4Score1.Text)) = False Then
            Return False
        Else
            dblStudent4TestScore(0) = CDbl(txtTestStudent4Score1.Text)
        End If
        If StudentTestScoreCheck(CDbl(txtTestStudent4Score2.Text)) = False Then
            Return False
        Else
            dblStudent4TestScore(1) = CDbl(txtTestStudent4Score2.Text)
        End If
        If StudentTestScoreCheck(CDbl(txtTestStudent4Score3.Text)) = False Then
            Return False
        Else
            dblStudent4TestScore(2) = CDbl(txtTestStudent4Score3.Text)
        End If
        If StudentTestScoreCheck(CDbl(txtTestStudent4Score4.Text)) = False Then
            Return False
        Else
            dblStudent4TestScore(3) = CDbl(txtTestStudent4Score4.Text)
        End If
        If StudentTestScoreCheck(CDbl(txtTestStudent4Score5.Text)) = False Then
            Return False
        Else
            dblStudent4TestScore(4) = CDbl(txtTestStudent4Score5.Text)
        End If

        If StudentTestScoreCheck(CDbl(txtTestStudent5Score1.Text)) = False Then
            Return False
        Else
            dblStudent5TestScore(0) = CDbl(txtTestStudent5Score1.Text)
        End If
        If StudentTestScoreCheck(CDbl(txtTestStudent5Score2.Text)) = False Then
            Return False
        Else
            dblStudent5TestScore(1) = CDbl(txtTestStudent5Score2.Text)
        End If
        If StudentTestScoreCheck(CDbl(txtTestStudent5Score3.Text)) = False Then
            Return False
        Else
            dblStudent5TestScore(2) = CDbl(txtTestStudent5Score3.Text)
        End If
        If StudentTestScoreCheck(CDbl(txtTestStudent5Score4.Text)) = False Then
            Return False
        Else
            dblStudent5TestScore(3) = CDbl(txtTestStudent5Score4.Text)
        End If
        If StudentTestScoreCheck(CDbl(txtTestStudent5Score5.Text)) = False Then
            Return False
        Else
            dblStudent5TestScore(4) = CDbl(txtTestStudent5Score5.Text)
        End If

        If StudentTestScoreCheck(CDbl(txtTestStudent6Score1.Text)) = False Then
            Return False
        Else
            dblStudent6TestScore(0) = CDbl(txtTestStudent6Score1.Text)
        End If
        If StudentTestScoreCheck(CDbl(txtTestStudent5Score2.Text)) = False Then
            Return False
        Else
            dblStudent6TestScore(1) = CDbl(txtTestStudent6Score2.Text)
        End If
        If StudentTestScoreCheck(CDbl(txtTestStudent5Score3.Text)) = False Then
            Return False
        Else
            dblStudent6TestScore(2) = CDbl(txtTestStudent6Score3.Text)
        End If
        If StudentTestScoreCheck(CDbl(txtTestStudent5Score4.Text)) = False Then
            Return False
        Else
            dblStudent6TestScore(3) = CDbl(txtTestStudent6Score4.Text)
        End If
        If StudentTestScoreCheck(CDbl(txtTestStudent5Score5.Text)) = False Then
            Return False
        Else
            dblStudent6TestScore(4) = CDbl(txtTestStudent6Score5.Text)
        End If

        Return True

    End Function

    Private Sub btnCalculate_Click(sender As Object, e As EventArgs) Handles btnCalculate.Click

        Try

            'set student name values
            strStudent1 = txtStudent1.Text
            strStudent2 = txtStudent2.Text
            strStudent3 = txtStudent3.Text
            strStudent4 = txtStudent4.Text
            strStudent5 = txtStudent5.Text
            strStudent6 = txtStudent6.Text

            If RunTestScoreCheck() = False Then
                Exit Sub
            End If

            'get averages
            lblAverageStudent1.Text = CStr(StudentAverages(dblStudent1Average, dblStudent1TestScore, dblCumulativeTestScores1))
            lblAverageStudent2.Text = CStr(StudentAverages(dblStudent2Average, dblStudent2TestScore, dblCumulativeTestScores2))
            lblAverageStudent3.Text = CStr(StudentAverages(dblStudent3Average, dblStudent3TestScore, dblCumulativeTestScores3))
            lblAverageStudent4.Text = CStr(StudentAverages(dblStudent4Average, dblStudent4TestScore, dblCumulativeTestScores4))
            lblAverageStudent5.Text = CStr(StudentAverages(dblStudent5Average, dblStudent5TestScore, dblCumulativeTestScores5))
            lblAverageStudent6.Text = CStr(StudentAverages(dblStudent6Average, dblStudent6TestScore, dblCumulativeTestScores6))

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub btnReport_Click(sender As Object, e As EventArgs) Handles btnReport.Click

        Try

            If File.Exists("StudentTestScores.txt") Then
            StudentTestScoresFile = File.AppendText("StudentTestScores.txt")
        Else
            StudentTestScoresFile = File.CreateText("StudentTestScores.txt")
        End If

            If RunTestScoreCheck() = False Then
                Exit Sub
            End If

            'set student name values
            strStudent1 = txtStudent1.Text
            strStudent2 = txtStudent2.Text
            strStudent3 = txtStudent3.Text
            strStudent4 = txtStudent4.Text
            strStudent5 = txtStudent5.Text
            strStudent6 = txtStudent6.Text

            StudentTestScoresFile.Write(strStudent1)
            StudentTestScoresFile.Write(vbTab + CStr(dblStudent1TestScore(0)))
            StudentTestScoresFile.Write(vbTab + CStr(dblStudent1TestScore(1)))
            StudentTestScoresFile.Write(vbTab + CStr(dblStudent1TestScore(2)))
            StudentTestScoresFile.Write(vbTab + CStr(dblStudent1TestScore(3)))
            StudentTestScoresFile.Write(vbTab + CStr(dblStudent1TestScore(4)))
            StudentTestScoresFile.WriteLine(vbTab + CStr(StudentAverages(dblStudent1Average, dblStudent1TestScore, dblCumulativeTestScores1)))

            StudentTestScoresFile.Write(strStudent2)
            StudentTestScoresFile.Write(vbTab + CStr(dblStudent2TestScore(0)))
            StudentTestScoresFile.Write(vbTab + CStr(dblStudent2TestScore(1)))
            StudentTestScoresFile.Write(vbTab + CStr(dblStudent2TestScore(2)))
            StudentTestScoresFile.Write(vbTab + CStr(dblStudent2TestScore(3)))
            StudentTestScoresFile.Write(vbTab + CStr(dblStudent2TestScore(4)))
            StudentTestScoresFile.WriteLine(vbTab + CStr(StudentAverages(dblStudent2Average, dblStudent2TestScore, dblCumulativeTestScores2)))

            StudentTestScoresFile.Write(strStudent3)
            StudentTestScoresFile.Write(vbTab + CStr(dblStudent3TestScore(0)))
            StudentTestScoresFile.Write(vbTab + CStr(dblStudent3TestScore(1)))
            StudentTestScoresFile.Write(vbTab + CStr(dblStudent3TestScore(2)))
            StudentTestScoresFile.Write(vbTab + CStr(dblStudent3TestScore(3)))
            StudentTestScoresFile.Write(vbTab + CStr(dblStudent3TestScore(4)))
            StudentTestScoresFile.WriteLine(vbTab + CStr(StudentAverages(dblStudent3Average, dblStudent3TestScore, dblCumulativeTestScores3)))

            StudentTestScoresFile.Write(strStudent4)
            StudentTestScoresFile.Write(vbTab + CStr(dblStudent4TestScore(0)))
            StudentTestScoresFile.Write(vbTab + CStr(dblStudent4TestScore(1)))
            StudentTestScoresFile.Write(vbTab + CStr(dblStudent4TestScore(2)))
            StudentTestScoresFile.Write(vbTab + CStr(dblStudent4TestScore(3)))
            StudentTestScoresFile.Write(vbTab + CStr(dblStudent4TestScore(4)))
            StudentTestScoresFile.WriteLine(vbTab + CStr(StudentAverages(dblStudent4Average, dblStudent4TestScore, dblCumulativeTestScores4)))

            StudentTestScoresFile.Write(strStudent5)
            StudentTestScoresFile.Write(vbTab + CStr(dblStudent5TestScore(0)))
            StudentTestScoresFile.Write(vbTab + CStr(dblStudent5TestScore(1)))
            StudentTestScoresFile.Write(vbTab + CStr(dblStudent5TestScore(2)))
            StudentTestScoresFile.Write(vbTab + CStr(dblStudent5TestScore(3)))
            StudentTestScoresFile.Write(vbTab + CStr(dblStudent5TestScore(4)))
            StudentTestScoresFile.WriteLine(vbTab + CStr(StudentAverages(dblStudent5Average, dblStudent5TestScore, dblCumulativeTestScores5)))

            StudentTestScoresFile.Write(strStudent6)
            StudentTestScoresFile.Write(vbTab + CStr(dblStudent6TestScore(0)))
            StudentTestScoresFile.Write(vbTab + CStr(dblStudent6TestScore(1)))
            StudentTestScoresFile.Write(vbTab + CStr(dblStudent6TestScore(2)))
            StudentTestScoresFile.Write(vbTab + CStr(dblStudent6TestScore(3)))
            StudentTestScoresFile.Write(vbTab + CStr(dblStudent6TestScore(4)))
            StudentTestScoresFile.WriteLine(vbTab + CStr(StudentAverages(dblStudent6Average, dblStudent6TestScore, dblCumulativeTestScores6)))

            StudentTestScoresFile.Close()


        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        Close()
    End Sub

End Class
