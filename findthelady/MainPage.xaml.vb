Partial Public Class MainPage
    Inherits PhoneApplicationPage

    Public Sub New()
        InitializeComponent()
        VisualStateManager.GoToState(Me, "Welcome", False)
    End Sub

    Dim Score As Integer = 0
    Dim Card1Visible As Boolean = False
    Dim Card2Visible As Boolean = True
    Dim Card3Visible As Boolean = False

    Dim Speed As Integer = 750
    Dim NestedLevel As Integer = 2
    Dim baseNestedLevel As Integer = 2
    Dim baseSpeed As Integer = 750

    Dim Slot1Grid, Slot2Grid, Slot3Grid As Grid

    Dim Pos1, Pos2, Pos3 As Point

    Dim IsInGame As Boolean = False
    Dim IsInExpert As Boolean = False
    Dim IsInShuffling As Boolean = False

    Private Sub btnStartGame_Click(sender As Object, e As System.Windows.RoutedEventArgs) Handles btnStartGame.Click
        IsInExpert = False
        baseSpeed = 750
        baseNestedLevel = 2
        StartGame()
    End Sub

    Private Sub btnStartGame_Expert_Click(sender As Object, e As System.Windows.RoutedEventArgs) Handles btnStartGame_Expert.Click
        IsInExpert = True
        baseSpeed = 350
        baseNestedLevel = 5
        StartGame()
    End Sub

    Sub StartGame()
        NestedLevel = baseNestedLevel
        Speed = baseSpeed
        IsInGame = True
        anim_ShuffleButtonShow.Begin()
        txtScore.Text = 0
        VisualStateManager.GoToState(Me, "InGame", True)
        ShowLadyCard()
    End Sub

    Dim RND As New Random
    Dim AllMovements As New List(Of Entities.MoveVector)

    Private Sub btnRoll_Click(sender As Object, e As System.Windows.RoutedEventArgs) Handles btnRoll.Click
        anim_ShuffleButtonRemove.Begin()
        VisualStateManager.GoToState(Me, "Shuffling", True)
        IsInShuffling = True
        HideAll()

        AllMovements.Clear()
        Dim LastStartPos As Integer = 0
        For index = 1 To NestedLevel
            Dim Movement As New Entities.MoveVector
            Do
                Movement.StartPos = RND.Next(0, 3)
            Loop Until LastStartPos <> Movement.StartPos
            LastStartPos = Movement.StartPos
            Do
                Movement.EndPos = RND.Next(0, 3)
            Loop Until Movement.StartPos <> Movement.EndPos
            AllMovements.Add(Movement)
        Next

        ShuffleCards()
    End Sub

    Private Sub ShuffleCards()
        If AllMovements.Count > 0 Then
            MoveCardPosition(AllMovements(0).StartPos, AllMovements(0).EndPos, Sub()
                                                                                   AllMovements.RemoveAt(0)
                                                                                   ShuffleCards()
                                                                               End Sub, True)
        Else
            VisualStateManager.GoToState(Me, "Ready", True)
            IsInShuffling = False
        End If
    End Sub

    Private Sub HideAll()
        If Card1Visible Then
            anim_Card1Hide.Begin()
            Card1Visible = False
        End If
        If Card2Visible Then
            anim_Card2Hide.Begin()
            Card2Visible = False
        End If
        If Card3Visible Then
            anim_Card3Hide.Begin()
            Card3Visible = False
        End If
    End Sub

    Private Sub ShowAll()
        If Card1Visible = False Then
            anim_Card1Show.Begin()
            Card1Visible = True
        End If
        If Card2Visible = False Then
            anim_Card2Show.Begin()
            Card2Visible = True
        End If
        If Card3Visible = False Then
            anim_Card3Show.Begin()
            Card3Visible = True
        End If
    End Sub

    Private Sub ShowLadyCard()
        If Card1Grid.Tag Then
            anim_Card1Show.Begin()
            Card1Visible = True
        End If
        If Card2Grid.Tag Then
            anim_Card2Show.Begin()
            Card2Visible = True
        End If
        If Card3Grid.Tag Then
            anim_Card3Show.Begin()
            Card3Visible = True
        End If
    End Sub

    Private Sub MainPage_BackKeyPress(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles MyBase.BackKeyPress
        If IsInGame Then
            GoToWelcomeScreen
            e.Cancel = True
        ElseIf IsInHowTo Then
            GoToWelcomeScreen()
            e.Cancel = True
        End If
    End Sub

    Private Sub MainPage_Loaded(sender As Object, e As System.Windows.RoutedEventArgs) Handles MyBase.Loaded

        'Dummy list of sample success messages.
        SuccessMessages.Add("YUUUUPIIIII!")
        SuccessMessages.Add("OHLALA!")
        SuccessMessages.Add("YAAAAAAY!")
        SuccessMessages.Add("YEEHAAA!")
        SuccessMessages.Add("VOUW!")
        SuccessMessages.Add("YOU ROCK!")
        SuccessMessages.Add("WAAAAHAAAA!")
        ReArrangeHighScoreList()

        Card2Grid.Tag = True
        Card3Grid.Tag = False
        Card1Grid.Tag = False

        Pos1 = New Point(Canvas.GetLeft(Card1Grid), Canvas.GetTop(Card1Grid))
        Pos2 = New Point(Canvas.GetLeft(Card2Grid), Canvas.GetTop(Card2Grid))
        Pos3 = New Point(Canvas.GetLeft(Card3Grid), Canvas.GetTop(Card3Grid))

        Slot1Grid = Card1Grid
        Slot2Grid = Card2Grid
        Slot3Grid = Card3Grid
    End Sub

    Private Sub MoveCardPosition(CurrentPos As Integer, TargetPos As Integer, CompletedAction As Action, Optional Recursive As Boolean = False)
        Dim CardToMove As Grid
        Dim TargetLeft As Double
        Dim TargetTop As Double

        Select Case CurrentPos
            Case 0
                CardToMove = Slot1Grid
            Case 1
                CardToMove = Slot2Grid
            Case 2
                CardToMove = Slot3Grid
        End Select
        Select Case TargetPos
            Case 0
                TargetLeft = Pos1.X
                TargetTop = Pos1.Y
            Case 1
                TargetLeft = Pos2.X
                TargetTop = Pos2.Y
            Case 2
                TargetLeft = Pos3.X
                TargetTop = Pos3.Y
        End Select

        Dim SB As New Storyboard
        Dim DB As New DoubleAnimationUsingKeyFrames
        Dim KF As New EasingDoubleKeyFrame
        KF.Value = TargetLeft
        KF.KeyTime = TimeSpan.FromMilliseconds(Speed)
        DB.KeyFrames.Add(KF)
        Storyboard.SetTargetProperty(DB, New PropertyPath(Canvas.LeftProperty))
        Storyboard.SetTarget(DB, CardToMove)
        SB.Children.Add(DB)

        DB = New DoubleAnimationUsingKeyFrames
        KF = New EasingDoubleKeyFrame
        KF.Value = TargetTop
        KF.KeyTime = TimeSpan.FromMilliseconds(Speed)
        DB.KeyFrames.Add(KF)
        KF = New EasingDoubleKeyFrame
        KF.Value = TargetTop - RND.Next(-30, 30)
        KF.KeyTime = TimeSpan.FromMilliseconds(Speed / 2)
        DB.KeyFrames.Add(KF)
        Storyboard.SetTargetProperty(DB, New PropertyPath(Canvas.TopProperty))
        Storyboard.SetTarget(DB, CardToMove)
        SB.Children.Add(DB)

        If Recursive Then
            MoveCardPosition(TargetPos, CurrentPos, Nothing)

            AddHandler SB.Completed, Sub()
                                         Select Case TargetPos
                                             Case 0
                                                 Select Case CurrentPos
                                                     Case 1
                                                         Slot2Grid = Slot1Grid
                                                     Case 2
                                                         Slot3Grid = Slot1Grid
                                                 End Select
                                                 Slot1Grid = CardToMove
                                             Case 1
                                                 Select Case CurrentPos
                                                     Case 0
                                                         Slot1Grid = Slot2Grid
                                                     Case 2
                                                         Slot3Grid = Slot2Grid
                                                 End Select
                                                 Slot2Grid = CardToMove
                                             Case 2
                                                 Select Case CurrentPos
                                                     Case 0
                                                         Slot1Grid = Slot3Grid
                                                     Case 1
                                                         Slot2Grid = Slot3Grid
                                                 End Select
                                                 Slot3Grid = CardToMove
                                         End Select
                                     End Sub
        End If

        AddHandler SB.Completed, Sub()
                                     If CompletedAction IsNot Nothing Then
                                         CompletedAction()
                                     End If
                                 End Sub
        SB.Begin()
    End Sub

    Private Sub Card1Grid_MouseLeftButtonDown(sender As Object, e As System.Windows.Input.MouseButtonEventArgs) Handles Card1Grid.MouseLeftButtonDown, Card2Grid.MouseLeftButtonDown, Card3Grid.MouseLeftButtonDown
        If IsInShuffling = False And btnRoll.Opacity = 0 Then
            Dim SendrGrid As Grid = sender
            If SendrGrid Is Card1Grid Then
                anim_Card1Show.Begin()
                Card1Visible = True
            End If
            If SendrGrid Is Card2Grid Then
                anim_Card2Show.Begin()
                Card2Visible = True
            End If
            If SendrGrid Is Card3Grid Then
                anim_Card3Show.Begin()
                Card3Visible = True
            End If

            Dim Score As Integer = txtScore.Text
            If CBool(SendrGrid.Tag) = True Then
                Score += 1
                If NestedLevel < 20 Then
                    NestedLevel = baseNestedLevel + (Score / 10)
                End If
                If Speed > 250 Then
                    Speed = baseSpeed - (Score * 10)
                End If

                If IsInExpert = False Then
                    If System.IO.IsolatedStorage.IsolatedStorageSettings.ApplicationSettings.Contains("Score") Then
                        If Score > CInt(System.IO.IsolatedStorage.IsolatedStorageSettings.ApplicationSettings("Score")) Then
                            System.IO.IsolatedStorage.IsolatedStorageSettings.ApplicationSettings("Score") = Score
                        End If
                    Else
                        System.IO.IsolatedStorage.IsolatedStorageSettings.ApplicationSettings.Add("Score", Score)
                    End If
                Else
                    If System.IO.IsolatedStorage.IsolatedStorageSettings.ApplicationSettings.Contains("Score_Expert") Then
                        If Score > CInt(System.IO.IsolatedStorage.IsolatedStorageSettings.ApplicationSettings("Score_Expert")) Then
                            System.IO.IsolatedStorage.IsolatedStorageSettings.ApplicationSettings("Score_Expert") = Score
                        End If
                        System.IO.IsolatedStorage.IsolatedStorageSettings.ApplicationSettings("Score_Expert") = Score
                    Else
                        System.IO.IsolatedStorage.IsolatedStorageSettings.ApplicationSettings.Add("Score_Expert", Score)
                    End If
                End If
                txtSuccessMessage.Text = SuccessMessages(RND.Next(0, SuccessMessages.Count))

                VisualStateManager.GoToState(Me, "Right", True)
            Else
                Score -= 1
                ShowAll()
                VisualStateManager.GoToState(Me, "Wrong", True)
                If Score < 0 Then
                    GoToWelcomeScreen()
                End If
            End If
            txtScore.Text = Score
            anim_ShuffleButtonShow.Begin()
        End If
    End Sub

    Sub GoToWelcomeScreen()
        HideAll()
        ReArrangeHighScoreList()
        VisualStateManager.GoToState(Me, "Welcome", True)
        IsInGame = False
        IsInHowTo = False
    End Sub

    Sub ReArrangeHighScoreList()
        Dim NormalScore As Integer
        If System.IO.IsolatedStorage.IsolatedStorageSettings.ApplicationSettings.Contains("Score") Then
            NormalScore = System.IO.IsolatedStorage.IsolatedStorageSettings.ApplicationSettings("Score")
            txtHighestScore_Amateur.Text = NormalScore & " in Amateur Level"
            txtHighestScore_Amateur.Visibility = Windows.Visibility.Visible
        Else
            txtHighestScore_Amateur.Visibility = Windows.Visibility.Collapsed
        End If

        Dim ExpertScore As Integer
        If System.IO.IsolatedStorage.IsolatedStorageSettings.ApplicationSettings.Contains("Score_Expert") Then
            ExpertScore = System.IO.IsolatedStorage.IsolatedStorageSettings.ApplicationSettings("Score_Expert")
            txtHighestScore_Expert.Text = ExpertScore & " in Expert Level"
            txtHighestScore_Expert.Visibility = Windows.Visibility.Visible
        Else
            txtHighestScore_Expert.Visibility = Windows.Visibility.Collapsed
        End If

        If txtHighestScore_Amateur.Visibility = Windows.Visibility.Collapsed And txtHighestScore_Expert.Visibility = Windows.Visibility.Collapsed Then
            ScoreList.Visibility = Windows.Visibility.Collapsed
        Else
            ScoreList.Visibility = Windows.Visibility.Visible
        End If
    End Sub

    Dim IsInHowTo As Boolean = False

    Dim SuccessMessages As New List(Of String)

    Private Sub btnHowto_Click(sender As Object, e As System.Windows.RoutedEventArgs) Handles btnHowto.Click
        IsInHowTo = True
        VisualStateManager.GoToState(Me, "How_To", True)
    End Sub
End Class
