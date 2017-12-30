Option Strict On

Public Class StateCapitalsGame
    '=============================================================================================================
    ' Author:	Holly Eaton
    ' 
    ' Program:	State Capitals Game
    ' 
    ' Dev Env:	Visual Studio
    ' 
    ' Description:
    '   Purpose:    Project that is a U.S. capitals study game.
    '  
    '   Input:      The user will select an answer from one of four radio buttons. 
    '
    '   Process:    The user will be presented with a multiple choice question about one of the states.
    '                   The application will check to see if the answer selected is correct or incorrect.
    '                   A running score total will keep track of the users success.
    '  
    '   Output:     User feedback:
    '                   If correct congratulations! –appears w/ happy sound.
    '                   If incorrect a message appears with the correct answer and unhappy sound.
    '                   The students running score total and average is displayed so they can track their success. 
    ' 
    '===============================================================================================================
    '   Class level variables use in multiple subroutines
    '       strStatesArray
    '       strCapitalsArray
    '       intScore                                
    '       strCorrectCapital
    '       intTotalQuestions   
    '       sngPercentCorrect
    '
    '   Local variables:
    '       Local to doNewQuestion:
    '       RandomClassInstance
    '       intStateRandomNumber        
    '       intFirstRandomNumber                                 
    '       intSecondRandomNumber                                 
    '	    intThirdRandomNumber                                 
    '	    intForthRandomNumber
    '       intReplaceRandomAnswer
    '
    '   Local to allRadioButtons_CheckedChanged:
    '	    selectedRadioButton
    '
    '   Local to doCheckAnswer:
    '	    blnIsItCorrect
    '
    '================================================================================================================
    '	Pseudocode:
    '       Determine which radiobutton was selected

    '       If radio button is checked call doCheckAnswer
    '           If correct/True:    friendly beep
    '                               message box with congratulations etc.
    '                               increment the number of correct responses
    '                               Call doNewQuestion
    '
    '           If incorrect/False: you missed it beep
    '                               message box telling user their response was incorrect and show the correct answer.
    '                               Call doNewQuestion
    '                               (only if keeping track of total questions) increment the number of total responses
    '
    '==================================================================================================================
    '	Calculations in pseudocode:
    '       Set option strict on at top
    '       intScore += 1            
    '       intTotalQuestions += 1
    '       sngPercentCorrect = (intScore/intTotalQuestions)*100
    '
    '==================================================================================================================
    '	Functions / Subroutines:
    '       doNewQuestion()  *Subroutine*
    '           Instantiate a new instance of the random class *Dim RandomClassInstance As New Random*
    '           Generate a random number between 0 and 49
    '               intStateRandomNumber = RandomClassInstance.Next(0, 50)
    '          
    '           Use the random number to pull out a state from the states array and the matching capital from the capitals array.
    '           Update the lblStateQuestion label question
    '               lblStateQuestion.Text = "What is the capital of" & strStatesArray(intStateRandomNumber) & "?"
    '
    '           Save the correct answer, the capital name, in strCorrectCapital
    '               strCorrectCapital = strCapitalsArray(intStateRandomNumber)
    '          
    '           Update each of the four answers with randomly selected capital names (can put while loop to stop duplicates)
    '               intFirstRandomNumber = RandomClassInstance.Next(0, 50)     
    '               rad1.Text = strCapitalsArray(intFirstRandomNumber)
    '               While intFirstRandomNumber = intSecondRandomNumber
    '                     intSecondRandomNumber = RandomClassInstance.Next(0, 50)
    '               End While     
    '               rad2.Text = strCapitalsArray(intSecondRandomNumber)
    '       Repeat 2 more times for 3rd and 4th answers.
    '          
    '           Replace one of the four answers with the correct answer.
    '               intReplaceRandomAnswer = RandomClassInstance.Next(0, 4)
    '               use a case statement and that number...
    '               case intReplaceRandomAnswer
    '                   rad1 or 2 or 3 or 4 .Text = strCorrectCapital
    '          
    '           Uncheck all four answers.
    '               rad#.Checked = False    (repeat four times for all the buttons)
    '
    '  
    '   doCheckAnswer(pass variable name in multiple radio button handler ie. selectedRadioButton)  returns true/correct false/incorrect *Function*
    '       If selectedRadioButton.Text = strCorrectCapital
    '           return true
    '       else
    '           return false
    '
    '==================================================================================================================  
    '==================================================================================================================
    '==================================================================================================================
    '  
    'Class level variables use in multiple subroutines
    Private intScore As Integer
    Private intTotalQuestions As Integer
    Private strCorrectCapital As String
    Private sngPercentCorrect As Single
    Private strStatesArray() As String = {"Alabama", "Alaska", "Arizona", "Arkansas", "California", "Colorado", "Connecticut", "Delaware", "Florida", "Georgia", "Hawaii", "Idaho", "Illinois", "Indiana", "Iowa", "Kansas", "Kentucky", "Louisiana", "Maine", "Maryland", "Massachusetts", "Michigan", "Minnesota", "Mississippi", "Missouri", "Montana", "Nebraska", "Nevada", "New Hampshire", "New Jersey", "New Mexico", "New York", "North Carolina", "North Dakota", "Ohio", "Oklahoma", "Oregon", "Pennsylvania", "Rhode Island", "South Carolina", "South Dakota", "Tennessee", "Texas", "Utah", "Vermont", "Virginia", "Washington", "West Virginia", "Wisconsin", "Wyoming"}
    Private strCapitalsArray() As String = {"Montgomery", "Juneau", "Phoenix", "Little Rock", "Sacramento", "Denver", "Hartford", "Dover", "Tallahassee", "Atlanta", "Honolulu", "Boise", "Springfield", "Indianapolis", "Des Moines", "Topeka", "Frankfort", "Baton Rouge", "Augusta", "Annapolis", "Boston", "Lansing", "St. Paul", "Jackson", "Jefferson City", "Helena", "Lincoln", "Carson City", "Concord", "Trenton", "Santa Fe", "Albany", "Raleigh", "Bismarck", "Columbus", "Oklahoma City", "Salem", "Harrisburg", "Providence", "Columbia", "Pierre", "Nashville", "Austin", "Salt Lake City", "Montpelier", "Richmond", "Olympia", "Charleston", "Madison", "Cheyenne"}

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'things that happen when the application window first appears
        MessageBox.Show("Welcome to the State Capitals Game. " & vbCrLf & "Just click on a radiobutton to select your answer.")

        doNewQuestion()
    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub

    Private Sub doNewQuestion()
        'Local variables:
        Dim intStateRandomNumber As Integer
        Dim intFirstRandomNumber As Integer
        Dim intSecondRandomNumber As Integer
        Dim intThirdRandomNumber As Integer
        Dim intForthRandomNumber As Integer
        Dim intReplaceRandomAnswer As Integer

        'Instantiate a new instance of the random class
        Dim RandomClassInstance As New Random

        'Use the random number to pull out a state from the states array and the matching capital from the capitals array.
        'Generate a random number between 0 and 49
        intStateRandomNumber = RandomClassInstance.Next(0, 50)

        'Update the lblStateQuestion label question
        lblStateQuestion.Text = "What is the capital of " & strStatesArray(intStateRandomNumber) & "?"
        'Save the correct answer, the capital name, in strCorrectCapital
        strCorrectCapital = strCapitalsArray(intStateRandomNumber)

        'Update each of the four answers with randomly selected capital names
        intFirstRandomNumber = RandomClassInstance.Next(0, 50)
        rad1.Text = strCapitalsArray(intFirstRandomNumber)
        While intFirstRandomNumber = intSecondRandomNumber
            intSecondRandomNumber = RandomClassInstance.Next(0, 50)
        End While
        rad2.Text = strCapitalsArray(intSecondRandomNumber)
        While intFirstRandomNumber = intThirdRandomNumber Or intSecondRandomNumber = intThirdRandomNumber
            intThirdRandomNumber = RandomClassInstance.Next(0, 50)
        End While
        rad3.Text = strCapitalsArray(intThirdRandomNumber)
        While intFirstRandomNumber = intForthRandomNumber Or intSecondRandomNumber = intForthRandomNumber Or intThirdRandomNumber = intForthRandomNumber
            intForthRandomNumber = RandomClassInstance.Next(0, 50)
        End While
        rad4.Text = strCapitalsArray(intForthRandomNumber)

        'Replace one of the four answers with the correct answer.
        intReplaceRandomAnswer = RandomClassInstance.Next(0, 4)

        'use a case statement and that number...
        Select Case intReplaceRandomAnswer
            Case 0
                rad1.Text = strCorrectCapital
            Case 1
                rad2.Text = strCorrectCapital
            Case 2
                rad3.Text = strCorrectCapital
            Case 3
                rad4.Text = strCorrectCapital

        End Select

        'Uncheck all four answers.
        rad1.Checked = False
        rad2.Checked = False
        rad3.Checked = False
        rad4.Checked = False

    End Sub

    Private Sub allRadioButtons_CheckedChanged(sender As Object, e As EventArgs) Handles rad1.CheckedChanged, rad2.CheckedChanged, rad3.CheckedChanged, rad4.CheckedChanged
        'detect which button has fired the event
        Dim selectedRadioButton As RadioButton
        selectedRadioButton = CType(sender, RadioButton)

        If selectedRadioButton.Checked Then
            Select Case doCheckAnswer(selectedRadioButton)
                Case True
                    'stuff to do if true/correct
                    'friendly Beep
                    Beep()
                    'or My.Computer.Audio.PlaySystemSound(Media.SystemSounds.Exclamation)

                    'message box with congratulations etc.
                    MessageBox.Show("That's it!  Good Job!")

                    'increment the number of correct responses and total questions
                    intScore += 1
                    intTotalQuestions += 1
                    sngPercentCorrect = CSng(intScore / intTotalQuestions)

                    lblScore.Text = intScore.ToString & " of " & intTotalQuestions.ToString & " quesions correct."
                    lblPercent.Text = sngPercentCorrect.ToString("P") & " correct"

                    doNewQuestion()
                Case False
                    'stuff to do if false/incorrect

                    'you missed it beep
                    My.Computer.Audio.PlaySystemSound(Media.SystemSounds.Asterisk)

                    'message box telling user their response was incorrect and show the correct answer.
                    MessageBox.Show("Sorry that is incorrect " & vbCrLf & "The correct answer is " & strCorrectCapital)

                    'increment the number of total questions
                    intTotalQuestions += 1
                    sngPercentCorrect = CSng(intScore / intTotalQuestions)

                    lblScore.Text = intScore.ToString & " of " & intTotalQuestions.ToString & " quesions correct."
                    lblPercent.Text = sngPercentCorrect.ToString("P") & " correct"

                    doNewQuestion()
            End Select
        End If

    End Sub

    Private Function doCheckAnswer(selectedRadioButton As RadioButton) As Boolean
        Dim blnIsItCorrect As Boolean

        If selectedRadioButton.Text = strCorrectCapital Then
            blnIsItCorrect = True
        Else
            blnIsItCorrect = False
        End If

        Return blnIsItCorrect
    End Function

    Private Sub btnNewPlayer_Click(sender As Object, e As EventArgs) Handles btnNewPlayer.Click
        'reset the environment for the new player
        lblScore.Text = " "
        lblPercent.Text = " "
        intScore = 0
        intTotalQuestions = 0

        MessageBox.Show("New Player, Welcome to the State Capitals Game. " & vbCrLf & "Just click on a radiobutton to select your answer.")

        doNewQuestion()

    End Sub
End Class
