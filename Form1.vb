

Public Class Form1
    Private textCOLOR As Color = Color.PaleGreen
    Private cardBACK As String
    Private btnPos As Integer = 0

    Private pHand() As String = {"03c", "03h", "14s", "03s", "03d"}
    Private dHand() As String = {"03c", "03h", "14s", "03s", "03d"}
    Private pKeep() As Boolean = {True, True, True, True, True}
    Private dKeep() As Boolean = {True, True, True, True, True}
    Private hand As Integer = 0
    Private Res As Boolean = False

    Private dResults As New Result
    Private pResult As New Result


    Structure Result
        Dim primary As Integer
        Dim secondary As Integer
        Dim returnVal As Boolean
        Dim rank As Integer
        Dim handValue As Decimal
    End Structure
    'Buttons & Form Loading-------------------------------------------------------------------------------------------
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        cardBACK = "back"
        lblWIN.Visible = False
        lblLOSE.Visible = False
        lblLOSE.Image = ImageList2.Images("broken.png")
        'lblDraw.Image = ImageList3.Images("DRAW.jpg")
        pCardBack()
        dCardBack()
    End Sub

    Private Sub btnNew_Click(sender As Object, e As EventArgs) Handles btnNew.Click, mnuN.Click
        Dim dec As New Deck
        Dim temp(4) As String
        If btnNew.ForeColor = textCOLOR Then
            btnNew.ForeColor = Color.Black
            btnDis.ForeColor = textCOLOR
            YW.Visible = True
            YL.Visible = True
            btnPos = 1
            YWYL()

            ' Get hands from draw function
            pHand = dec.Draw()
            dHand = dec.Draw()


            Array.Sort(pHand)
            Array.Sort(dHand)

            Array.Reverse(pHand)
            Array.Reverse(dHand)


            'Display hand hDisplay(cards)
            hDISPLAY(pHand, 0)
            btnNew.FlatStyle = FlatStyle.Flat
            btnDis.FlatStyle = FlatStyle.Standard
            dCardBack()
        End If
    End Sub

    Private Sub btnDis_Click(sender As Object, e As EventArgs) Handles btnDis.Click
        Dim dec As New Deck
        If btnDis.ForeColor = textCOLOR Then

            pCard1.BorderStyle = BorderStyle.None
            pCard2.BorderStyle = BorderStyle.None
            pCard3.BorderStyle = BorderStyle.None
            pCard4.BorderStyle = BorderStyle.None
            pCard5.BorderStyle = BorderStyle.None
            btnDis.ForeColor = Color.Black
            btnCall.ForeColor = textCOLOR
            btnPos = 2
            YWYL()

            btnDis.FlatStyle = FlatStyle.Flat
            btnCall.FlatStyle = FlatStyle.Standard
            pHand = dec.Discard(pHand, pKeep)
            Array.Sort(pHand)
            Array.Reverse(pHand)
            hDISPLAY(pHand, 0)

            DiscardTimer.Enabled = True
            hand = 0
            DISCARD(0)
        End If
    End Sub

    Private Sub btnCall_Click(sender As Object, e As EventArgs) Handles btnCall.Click
        Dim dec As New Deck
        If btnCall.ForeColor = textCOLOR Then



            dResults = Rank(dHand)
            dKeep = dealerDiscard(dHand, dResults)
            dHand = dec.DealerDiscard(dHand, dKeep)
            DiscardTimer.Enabled = True
            hand = 1
            DISCARD(1)


            Array.Sort(dHand)
            Array.Reverse(dHand)

            btnCall.ForeColor = Color.Black
            btnNew.ForeColor = textCOLOR
            btnPos = 0
            YWYL()
            btnCall.FlatStyle = FlatStyle.Flat
            btnNew.FlatStyle = FlatStyle.Standard

            DiscardTimer.Enabled = True
            DISCARD(1)

         

            EndDes.Enabled = True
        dec.KCIP()
        RESET()
        End If

    End Sub
    '-----------------------------------------------------------------------------------------------







    'Menu Options ---------------------------------------------------------------------------------
    Private Sub PokerFaceToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PokerFaceToolStripMenuItem.Click

        If btnNew.ForeColor = textCOLOR Then
            DealerHand.ForeColor = Color.Black
            MenuStrip1.BackgroundImage = My.Resources.Wood
            MenuStrip2.BackgroundImage = My.Resources.Wood
            btnNew.BackgroundImage = My.Resources.Wood
            btnDis.BackgroundImage = My.Resources.Wood
            btnCall.BackgroundImage = My.Resources.Wood


            Me.BackgroundImage = My.Resources.Black_B
            PlayerHand.BackColor = Color.Silver
            DealerHand.BackColor = Color.Silver
            textCOLOR = Color.Snow
            btnNew.ForeColor = textCOLOR
            mnuF.ForeColor = textCOLOR
            mnuO.ForeColor = textCOLOR
            mnuT.ForeColor = textCOLOR
            pCard1.Image = My.Resources.PF
            pCard2.Image = My.Resources.PF
            pCard3.Image = My.Resources.PF
            pCard4.Image = My.Resources.PF
            pCard5.Image = My.Resources.PF


            dCard1.Image = My.Resources.PF
            dCard2.Image = My.Resources.PF
            dCard3.Image = My.Resources.PF
            dCard4.Image = My.Resources.PF
            dCard5.Image = My.Resources.PF

            My.Computer.Audio.Play(My.Resources.POKERFACE, _
        AudioPlayMode.BackgroundLoop)
        End If

    End Sub

    Private Sub MoneyToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MoneyToolStripMenuItem.Click
        If btnNew.ForeColor = textCOLOR Then
            My.Computer.Audio.Stop()
            DealerHand.ForeColor = Color.Black
            MenuStrip1.BackgroundImage = My.Resources.gold
            MenuStrip2.BackgroundImage = My.Resources.gold
            btnNew.BackgroundImage = My.Resources.gold
            btnDis.BackgroundImage = My.Resources.gold
            btnCall.BackgroundImage = My.Resources.gold

            mnuF.ForeColor = Color.Black
            mnuT.ForeColor = Color.Black
            mnuO.ForeColor = Color.Black
            btnNew.ForeColor = Color.Snow
            textCOLOR = Color.Snow
            btnDis.ForeColor = Color.Black
            btnCall.ForeColor = Color.Black
            Me.BackgroundImage = My.Resources.moneyField
            DealerHand.BackColor = Color.Black
            DealerHand.ForeColor = Color.Silver
            PlayerHand.BackColor = Color.Gold
            cardBACK = "benjamin1"
            pCardBack()
            cardBACK = "banker11"
            dCardBack()
        End If

    End Sub

    Private Sub DefaultToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DefaultToolStripMenuItem.Click
        If btnNew.ForeColor = textCOLOR Then
            My.Computer.Audio.Stop()
            MenuStrip1.BackgroundImage = My.Resources.Wood
            MenuStrip2.BackgroundImage = My.Resources.Wood
            btnNew.BackgroundImage = My.Resources.Wood
            btnDis.BackgroundImage = My.Resources.Wood
            btnCall.BackgroundImage = My.Resources.Wood
            Me.BackgroundImage = My.Resources.PTF
            PlayerHand.BackColor = Color.DarkSeaGreen
            DealerHand.BackColor = Color.DarkSeaGreen
            textCOLOR = Color.PaleGreen
            btnNew.ForeColor = textCOLOR
            mnuF.ForeColor = textCOLOR
            mnuO.ForeColor = textCOLOR
            mnuT.ForeColor = textCOLOR
            btnNew.ForeColor = textCOLOR
            cardBACK = "back"
            pCardBack()
            dCardBack()
        End If
    End Sub

    Private Sub DefaultToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles DefaultToolStripMenuItem1.Click
        cardBACK = "back"
        pCardBack()
        dCardBack()
    End Sub

    Private Sub RedToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RedToolStripMenuItem.Click
        cardBACK = "back2"
        pCardBack()
        dCardBack()
    End Sub

    Private Sub DiscardAnimationToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DiscardAnimationToolStripMenuItem.Click
        If cardBACK = "back" Then
            DiscardTimer.Enabled = True


            pCardBack()
            dCardBack()
        ElseIf cardBACK = "back2" Then
            DiscardTimer.Enabled = True

        End If
    End Sub

    Private Sub HandDemoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles HandDemoToolStripMenuItem.Click
        hDISPLAY(pHand, 0)
    End Sub

    Private Sub PlayerNameToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PlayerName.Click
        PlayerHand.Text = InputBox("New Name:", "Player Name")
    End Sub

    Private Sub DealerNameToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DealerName.Click
        DealerHand.Text = InputBox("New Name:", "Dealer Name")
    End Sub

    '-------------------------------------------------------------------------------------------








    'Display-------------------------------------------------------------------------------------
    Private Function hDISPLAY(ByVal cards() As String, ByVal PorD As Integer)
        If PorD = 0 Then
            pCard1.Image = ImageList1.Images(cards(0) & ".gif")
            pCard2.Image = ImageList1.Images(cards(1) & ".gif")
            pCard3.Image = ImageList1.Images(cards(2) & ".gif")
            pCard4.Image = ImageList1.Images(cards(3) & ".gif")
            pCard5.Image = ImageList1.Images(cards(4) & ".gif")
        Else
            dCard1.Image = ImageList1.Images(cards(0) & ".gif")
            dCard2.Image = ImageList1.Images(cards(1) & ".gif")
            dCard3.Image = ImageList1.Images(cards(2) & ".gif")
            dCard4.Image = ImageList1.Images(cards(3) & ".gif")
            dCard5.Image = ImageList1.Images(cards(4) & ".gif")
        End If
        Return 0
    End Function

    Private Sub pCardBack()
        pCard1.Image = ImageList1.Images(cardBACK & ".gif")
        pCard2.Image = ImageList1.Images(cardBACK & ".gif")
        pCard3.Image = ImageList1.Images(cardBACK & ".gif")
        pCard4.Image = ImageList1.Images(cardBACK & ".gif")
        pCard5.Image = ImageList1.Images(cardBACK & ".gif")
    End Sub

    Private Sub dCardBack()
        dCard1.Image = ImageList1.Images(cardBACK & ".gif")
        dCard2.Image = ImageList1.Images(cardBACK & ".gif")
        dCard3.Image = ImageList1.Images(cardBACK & ".gif")
        dCard4.Image = ImageList1.Images(cardBACK & ".gif")
        dCard5.Image = ImageList1.Images(cardBACK & ".gif")
    End Sub
    '-------------------------------------------------------------------------------------






    'Win/Lose/Draw---------------------------------------------------------------------------------
    Private Sub YOUWINToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles YW.Click
        lblWIN.Visible = True
        WinOrLose.Enabled = True
    End Sub

    Private Sub WOULOSEToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles YL.Click
        lblLOSE.Visible = True
        WinOrLose.Enabled = True
    End Sub

    Private Sub YD_Click(sender As Object, e As EventArgs) Handles YD.Click
        lblDraw.Visible = True
        WinOrLose.Enabled = True
    End Sub
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        lblWIN.Visible = False
        lblLOSE.Visible = False
        lblDraw.Visible = False


        Timer1.Enabled = False
    End Sub

    Private Sub WinOrLose_Tick(sender As Object, e As EventArgs) Handles WinOrLose.Tick
        Static count As Integer = 0
        If lblWIN.Visible = True Then
            If count = 0 Then
                lblWIN.ForeColor = Color.DarkSeaGreen
                count = 1
            Else
                lblWIN.ForeColor = Color.Green
                count = 0
            End If

        ElseIf lblLOSE.Visible = True Then
            If count = 0 Then
                lblLOSE.ForeColor = Color.DarkRed
                count = 1
            Else
                lblLOSE.ForeColor = Color.Red
                count = 0
            End If
        ElseIf lblDraw.Visible = True Then
            If count = 0 Then
                lblDraw.ForeColor = Color.DarkGoldenrod
                count = 1
            Else
                lblDraw.ForeColor = Color.Gold
                count = 0
            End If
        End If
    End Sub

    Private Sub YWYL()
        If btnPos = 0 OrElse btnPos = 2 Then
            YW.Visible = True
            YL.Visible = True
        Else
            YW.Visible = False
            YL.Visible = False
            lblDraw.Visible = False
            lblWIN.Visible = False
            lblLOSE.Visible = False
            WinOrLose.Enabled = False
        End If
    End Sub
    '-----------------------------------------------------------------------------------------------







    'Discard--------------------------------------------------------------------------------------
    Private Sub DiscardTimer_Tick(sender As Object, e As EventArgs) Handles DiscardTimer.Tick
        Static count As Integer = 0

        If count = 0 Then
            count = 1
        Else
            count = 0
            DiscardTimer.Enabled = False
            If hand = 0 Then
                hDISPLAY(pHand, 0)
            Else
                hDISPLAY(dHand, 1)

            End If
        End If

    End Sub

    Private Function DISCARD(ByVal PorD As Integer)
        If PorD = 0 Then
            If pKeep(0) = False Then
                pCard1.Image = Nothing
            End If
            If pKeep(1) = False Then
                pCard2.Image = Nothing
            End If
            If pKeep(2) = False Then
                pCard3.Image = Nothing
            End If
            If pKeep(3) = False Then
                pCard4.Image = Nothing
            End If
            If pKeep(4) = False Then
                pCard5.Image = Nothing
            End If
        Else
            If dKeep(0) = False Then
                dCard1.Image = Nothing
            End If
            If dKeep(1) = False Then
                dCard2.Image = Nothing
            End If
            If dKeep(2) = False Then
                dCard3.Image = Nothing
            End If
            If dKeep(3) = False Then
                dCard4.Image = Nothing
            End If
            If dKeep(4) = False Then
                dCard5.Image = Nothing
            End If
        End If
        Return 0
    End Function
    '---------------------------------------------------------------------------------------------






    'Player Discard--------------------------------------------------------------------------------
    Private numDiscard As Integer = 0
    Private Sub pCard1_Click(sender As Object, e As EventArgs) Handles pCard1.Click
        If btnDis.ForeColor = textCOLOR AndAlso numDiscard <= 2 Then
            If pKeep(0) = True Then
                If numDiscard <= 2 Then
                    pKeep(0) = False
                    pCard1.BorderStyle = BorderStyle.Fixed3D
                    numDiscard = numDiscard + 1
                End If
            Else
                pKeep(0) = True
                pCard1.BorderStyle = BorderStyle.None
                numDiscard = numDiscard - 1
            End If
        End If
    End Sub

    Private Sub pCard2_Click(sender As Object, e As EventArgs) Handles pCard2.Click
        If btnDis.ForeColor = textCOLOR Then
            If pKeep(1) = True Then
                If numDiscard <= 2 Then
                    pKeep(1) = False
                    pCard2.BorderStyle = BorderStyle.Fixed3D
                    numDiscard = numDiscard + 1
                End If
            Else
                pKeep(1) = True
                pCard2.BorderStyle = BorderStyle.None
                numDiscard = numDiscard - 1
            End If
        End If
    End Sub

    Private Sub pCard3_Click(sender As Object, e As EventArgs) Handles pCard3.Click
        If btnDis.ForeColor = textCOLOR AndAlso numDiscard <= 2 Then
            If pKeep(2) = True Then
                If numDiscard <= 2 Then
                    pKeep(2) = False
                    pCard3.BorderStyle = BorderStyle.Fixed3D
                    numDiscard = numDiscard + 1
                End If
            Else
                pKeep(2) = True
                pCard3.BorderStyle = BorderStyle.None
                numDiscard = numDiscard - 1
            End If
        End If
    End Sub

    Private Sub pCard4_Click(sender As Object, e As EventArgs) Handles pCard4.Click
        If btnDis.ForeColor = textCOLOR Then
            If pKeep(3) = True Then
                If numDiscard <= 2 Then
                    pKeep(3) = False
                    pCard4.BorderStyle = BorderStyle.Fixed3D
                    numDiscard = numDiscard + 1
                End If
            Else
                pKeep(3) = True
                pCard4.BorderStyle = BorderStyle.None
                numDiscard = numDiscard - 1
            End If
        End If
    End Sub

    Private Sub pCard5_Click(sender As Object, e As EventArgs) Handles pCard5.Click
        If btnDis.ForeColor = textCOLOR Then
            If pKeep(4) = True Then
                If numDiscard <= 2 Then
                    pKeep(4) = False
                    pCard5.BorderStyle = BorderStyle.Fixed3D
                    numDiscard = numDiscard + 1
                End If
            Else
                pKeep(4) = True
                pCard5.BorderStyle = BorderStyle.None
                numDiscard = numDiscard - 1
            End If
        End If
    End Sub

    '----------------------------------------------------------------------------------------------







    'Dealer AI-----------------------------------------------------------------------------------


    Private Sub setReplace(hand() As String, primary As Integer, secondary As Integer, ByRef keep() As Boolean)
        Dim replaces As Integer
        For i As Integer = 0 To 4 Step 1

            If secondary <> Val(hand(i).Substring(1)) AndAlso primary <> Val(hand(i).Substring(1)) AndAlso replaces < 3 Then
                keep(i) = False
                replaces += 1
            Else
                keep(i) = True
            End If
        Next i
    End Sub

    Private Function dealerDiscard(hand() As String, hand_data As Result) As Boolean()
        Dim redraw(4) As Boolean
        Select Case hand_data.rank
            Case 4 To 9
                Return redraw
            Case 2
                setReplace(hand, hand_data.primary, hand_data.secondary, redraw)
                Return redraw
            Case 1, 3
                setReplace(hand, hand_data.primary, -1, redraw)
                Return redraw
            Case Else
                Return {True, True, False, False, False}
        End Select
    End Function


    '---------------------------------------------------------------------------------------------


    'Rank-----------------------------------------------------------------------------------------

    Private Function Rank(ByVal hand() As String)
        Dim myresult As New Result

        myresult = isRoyal(hand)
        If myresult.rank = 9 Then
            'nothing
        Else
            myresult = isStraightFlush(hand)
            If myresult.rank = 8 Then
                'nothing
            Else
                myresult = isFourOfAKind(hand)
                If myresult.rank = 7 Then
                    'nothing
                Else
                    myresult = isFullHouse(hand)
                    If myresult.rank = 6 Then
                        'nothing 
                    Else
                        myresult = isFlush(hand)
                        If myresult.rank = 5 Then
                            'nothing
                        Else
                            myresult = isStraight(hand)
                            If myresult.rank = 4 Then
                                'nothing
                            Else
                                myresult = isThreeOfAKind(hand)
                                If myresult.rank = 3 Then
                                    'nothing
                                Else
                                    myresult = isTwoPair(hand)
                                    If myresult.rank = 2 Then
                                        'nothing 
                                    Else
                                        myresult = IsOnePair(hand)
                                        If myresult.rank = 1 Then
                                            'nothing
                                        Else
                                            isNoPair(hand)
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If

        'MsgBox(myresult.rank.ToString)
        Return myresult
    End Function
    Private Function isRoyal(hand() As String) As Result
        Dim myResult As Result
        myResult.returnVal = False

        If hand(0) Like "14?" And hand(1) Like "13?" And hand(2) Like "12?" And hand(3) Like "11?" And hand(4) Like "10?" Then
            If hand(0)(2) Like hand(1)(2) AndAlso hand(1)(2) Like hand(2)(2) AndAlso hand(2)(2) Like hand(3)(2) AndAlso pHand(3)(2) Like hand(4)(2) Then
                myResult.returnVal = True
                myResult.rank = 9
                Integer.TryParse(hand(0).Remove(2), myResult.primary)  'converting the string values to ints for high card
            End If
        End If

        Return myResult 'return the structure
    End Function

    Private Function isStraightFlush(hand() As String) As Result
        Dim myResult As Result
        myResult.returnVal = False

        Dim straight As Integer 'variable to hold value from loop
        Dim firstNum As Integer 'variable to hold the value of first array index converted to int
        Integer.TryParse(hand(0).Remove(2), firstNum) 'removes suit from first index array and converts to int

        'Adds 1 to the straight for each number in con
        For i As Integer = 0 To 4
            If hand(i).Contains((firstNum - i).ToString) Then
                straight += 1
            End If
        Next i

        ' if flush
        If hand(0)(2) Like hand(1)(2) AndAlso hand(1)(2) Like hand(2)(2) AndAlso hand(2)(2) Like hand(3)(2) AndAlso hand(3)(2) Like hand(4)(2) Then

            'if straight
            If straight = 5 Then
                myResult.returnVal = True
                myResult.rank = 8
                Integer.TryParse(hand(0).Remove(2), myResult.primary)  'converting the string values to ints for high and low cards

                'if ace is the low card
            ElseIf hand(0).Contains("14") AndAlso hand(1).Contains("05") AndAlso hand(2).Contains("04") AndAlso hand(3).Contains("03") AndAlso hand(4).Contains("02") Then
                myResult.returnVal = True
                myResult.rank = 8
                Integer.TryParse(hand(4).Remove(2), myResult.primary) 'converting the string values to ints for high card

            End If
        End If

        Return myResult 'return the structure
    End Function

    Private Function isFourOfAKind(hand() As String) As Result
        Dim myResult As Result
        myResult.returnVal = False

        If hand(0).Substring(0, 2) = hand(1).Substring(0, 2) AndAlso
           hand(0).Substring(0, 2) = hand(2).Substring(0, 2) AndAlso
           hand(0).Substring(0, 2) = hand(3).Substring(0, 2) Then
            myResult.returnVal = True
            myResult.rank = 7
            Integer.TryParse(hand(0).Substring(0, 2), myResult.primary)

        ElseIf hand(1).Substring(0, 2) = hand(2).Substring(0, 2) AndAlso
               hand(1).Substring(0, 2) = hand(3).Substring(0, 2) AndAlso
               hand(1).Substring(0, 2) = hand(4).Substring(0, 2) Then
            myResult.returnVal = True
            myResult.rank = 7
            Integer.TryParse(hand(1).Substring(0, 2), myResult.primary)

        End If
        Return myResult
    End Function

    Private Function isFullHouse(hand() As String) As Result
        ' This function determines if hand() is a Full House hand
        ' Assumptions:
        '   Hand is sorted in descending card value order
        '   Hand is not GT a Full House hand.

        Dim primary As Integer
        Dim secondary As Integer
        Dim trash As Integer = 0
        Dim handValue As Double
        Dim isHand As Boolean = False

        Dim myResult As Result
        myResult.handValue = 0.0
        myResult.returnVal = False

        If hand(0).Substring(0, 2) = hand(1).Substring(0, 2) Then
            ' First two cards match, check for third card match
            If hand(0).Substring(0, 2) = hand(2).Substring(0, 2) Then
                ' First three cards match.  Set as primary
                Integer.TryParse(hand(0).Substring(0, 2), primary)

                If hand(3).Substring(0, 2) = hand(4).Substring(0, 2) Then
                    ' Last two cards match.  This is a full house
                    ' Set as secondary
                    Integer.TryParse(hand(3).Substring(0, 2), secondary)
                    isHand = True
                End If
            Else
                ' Only first two cards match.  Set as secondary
                Integer.TryParse(hand(0).Substring(0, 2), secondary)

                If hand(3).Substring(0, 2) = hand(2).Substring(0, 2) AndAlso _
                   hand(3).Substring(0, 2) = hand(4).Substring(0, 2) Then
                    ' Last three cards match.  This is a full house
                    ' set as primary
                    Integer.TryParse(hand(3).Substring(0, 2), primary)
                    isHand = True
                End If
            End If
        End If

        If isHand Then
            handValue = 6 + primary * 0.01 + secondary * 0.0001

            ' load the result structure
            myResult.primary = primary
            myResult.secondary = secondary
            myResult.handValue = handValue
            myResult.returnVal = True
            myResult.rank = 6
        End If

        Return myResult
    End Function

    Private Function isFlush(hand() As String) As Result
        Dim myResult As Result
        Dim suit(4) As String

        suit(0) = hand(0).Chars(2)
        suit(1) = hand(1).Chars(2)
        suit(2) = hand(2).Chars(2)
        suit(3) = hand(3).Chars(2)
        suit(4) = hand(4).Chars(2)

        myResult.returnVal = False

        If suit(0) = suit(1) And suit(0) = suit(2) And suit(0) = suit(3) And suit(0) = suit(4) Then
            myResult.returnVal = True
            myResult.rank = 5
        End If

        myResult.primary = 0
        myResult.secondary = 0

        Return myResult
    End Function

    Function isStraight(hand() As String) As Result
        Dim myResult As Result
        myResult.returnVal = False
        Dim areMultipleNumbersInList As Boolean = False
        Dim i As Integer
        For i = 0 To 3
            If hand(i).Substring(0, 2) = hand(i + 1).Substring(0, 2) Then
                areMultipleNumbersInList = True
            End If
        Next
        Dim max As Integer
        Dim min As Integer
        Integer.TryParse(hand(0).Substring(0, 2), max)
        Integer.TryParse(hand(4).Substring(0, 2), min)
        If (max - min = 4 AndAlso Not areMultipleNumbersInList) Then
            myResult.returnVal = True
            myResult.rank = 4
            myResult.primary = max
        End If
        If hand(0).Contains("14") AndAlso hand(1).Contains("05") AndAlso hand(2).Contains("04") AndAlso hand(3).Contains("03") AndAlso hand(4).Contains("02") Then
            myResult.returnVal = True
            myResult.rank = 4
            Integer.TryParse(hand(4).Remove(2), myResult.primary)
        End If
        Return myResult
    End Function

    Private Function isThreeOfAKind(hand() As String) As Result
        ' This function determines if hand() is Three of a Kind
        ' Assumptions:
        '   Hand is sorted in descending card value order
        '   Hand is not GT Three of a Kind.

        Dim primary As Integer = 0
        Dim secondary As Integer = 0
        Dim trash As Integer = 0
        Dim handValue As Double
        Dim isHand As Boolean = False

        Dim myResult As Result
        myResult.handValue = 0.0
        myResult.returnVal = False

        If hand(0).Substring(0, 2) = hand(1).Substring(0, 2) AndAlso _
            hand(0).Substring(0, 2) = hand(2).Substring(0, 2) Then
            ' Hand is Three of a Kind w cards 4 & 5 as trash
            Integer.TryParse(hand(0).Substring(0, 2), primary)
            Integer.TryParse(hand(3).Substring(0, 2), secondary)
            Integer.TryParse(hand(4).Substring(0, 2), trash)
            isHand = True
        ElseIf hand(1).Substring(0, 2) = hand(2).Substring(0, 2) AndAlso _
            hand(1).Substring(0, 2) = hand(3).Substring(0, 2) Then
            ' Hand is Three of a Kind w cards 1 & 5 as trash
            Integer.TryParse(hand(1).Substring(0, 2), primary)
            Integer.TryParse(hand(0).Substring(0, 2), secondary)
            Integer.TryParse(hand(4).Substring(0, 2), trash)
            isHand = True
        ElseIf hand(2).Substring(0, 2) = hand(3).Substring(0, 2) AndAlso _
            hand(2).Substring(0, 2) = hand(4).Substring(0, 2) Then
            ' Hand is Three of a Kind w cards 1 & 2 as trash
            Integer.TryParse(hand(2).Substring(0, 2), primary)
            Integer.TryParse(hand(0).Substring(0, 2), secondary)
            Integer.TryParse(hand(1).Substring(0, 2), trash)
            isHand = True
        End If

        If isHand Then
            handValue = 3 + primary * 0.01 + secondary * 0.0001 + trash * 0.000001

            ' load the result structure
            myResult.primary = primary
            myResult.secondary = secondary
            myResult.handValue = handValue
            myResult.returnVal = True
            myResult.rank = 3
        End If

        Return myResult
    End Function

    Private Function isTwoPair(hand() As String) As Result
        ' This function determines if phand() is a two pair hand
        ' Assumptions:
        '   Hand is sorted in descending card value order
        '   Hand is not GT a two pair hand.

        Dim primary As Integer
        Dim secondary As Integer
        Dim trash As Integer
        Dim handValue As Double
        Dim isHand As Boolean = False

        Dim myResult As Result
        myResult.handValue = 0.0
        myResult.returnVal = False

        If hand(0).Substring(0, 2) = hand(1).Substring(0, 2) Then
            ' First two cards are primary pair
            Integer.TryParse(hand(0).Substring(0, 2), primary)

            If hand(3).Substring(0, 2) = hand(2).Substring(0, 2) OrElse _
                hand(3).Substring(0, 2) = hand(4).Substring(0, 2) Then
                ' Found Secondary pair
                Integer.TryParse(hand(3).Substring(0, 2), secondary)

                If hand(3).Substring(0, 2) = hand(2).Substring(0, 2) Then
                    ' Last card is trash card
                    Integer.TryParse(hand(4).Substring(0, 2), trash)
                Else
                    ' Third card is trash card
                    Integer.TryParse(hand(2).Substring(0, 2), trash)
                End If

                isHand = True
            End If
        ElseIf hand(1).Substring(0, 2) = hand(2).Substring(0, 2) Then
            ' Second and third cards are primary pair
            Integer.TryParse(hand(1).Substring(0, 2), primary)

            If hand(3).Substring(0, 2) = hand(4).Substring(0, 2) Then
                ' Found secondary pair
                Integer.TryParse(hand(3).Substring(0, 2), secondary)

                ' First card is trash card
                Integer.TryParse(hand(0).Substring(0, 2), trash)

                isHand = True
            End If
        End If

        If isHand Then
            handValue = 2 + primary * 0.01 + secondary * 0.0001 + trash * 0.000001

            ' load the result structure
            myResult.primary = primary
            myResult.secondary = secondary
            myResult.handValue = handValue
            myResult.rank = 2
            myResult.returnVal = True
        End If

        Return myResult

        ''Contains A Count Each Card Type
        'Dim pdCounts(12) As Integer
        'Dim myResult As Result
        'myResult.returnVal = False

        'Dim dIndex As Integer
        'Dim value As Integer

        'For dIndex = 0 To 4
        '    Integer.TryParse(hand(dIndex).Substring(0, 2), value)
        '    pdCounts(value - 2) += 1
        'Next dIndex

        'For dIndex = 0 To 12
        '    If pdCounts(dIndex) > 1 Then
        '        For dIndex2 = 0 To 12
        '            If pdCounts(dIndex2) > 1 And dIndex2 <> dIndex Then
        '                myResult.returnVal = True
        '                If dIndex > dIndex2 Then
        '                    myResult.primary = dIndex + 2
        '                    myResult.secondary = dIndex2 + 2
        '                Else
        '                    myResult.primary = dIndex2 + 2
        '                    myResult.secondary = dIndex + 2
        '                End If
        '                myResult.rank = 2
        '            End If
        '        Next dIndex2
        '    End If
        'Next dIndex
        'Return myResult
    End Function

    Public Function IsOnePair(hand() As String) As Result
        Dim myResult As Result

        Dim inthand(4) As Integer

        'change to strings with suits to ints with no suits
        For i As Integer = 0 To 4
            Integer.TryParse(hand(i).Remove(2), inthand(i)) ' remove suit
        Next i

        'test every card with the card next to it, only needed once because its sorted
        For i As Integer = 0 To 3
            If inthand(i) = inthand(i + 1) Then
                myResult.returnVal = True
                myResult.primary = inthand(i)
                myResult.secondary = inthand(i + 1)
                myResult.rank = 1
                Return myResult
            End If
        Next i

        Return myResult
    End Function

    Private Function isNoPair(hand() As String) As Result
        ' This function determines if hand() is a No Pair hand
        ' Assumptions:
        '   Hand is sorted in descending card value order
        '   Hand is not GT No Pair.

        Dim primary As Integer = 0
        Dim secondary As Integer = 0
        Dim trash As Integer = 0
        Dim trash1 As Integer = 0
        Dim trash2 As Integer = 0
        Dim handValue As Double

        Dim myResult As Result
        myResult.handValue = 0.0
        myResult.returnVal = False

        Integer.TryParse(hand(0).Substring(0, 2), primary)
        Integer.TryParse(hand(1).Substring(0, 2), secondary)
        Integer.TryParse(hand(2).Substring(0, 2), trash)
        Integer.TryParse(hand(3).Substring(0, 2), trash1)
        Integer.TryParse(hand(4).Substring(0, 2), trash2)

        handValue = primary * 0.01 + secondary * 0.0001 _
            + trash * 0.000001 + trash1 * 0.00000001 + trash2 * 0.0000000001

        ' load the result structure
        myResult.primary = primary
        myResult.secondary = secondary
        myResult.handValue = handValue
        myResult.returnVal = True
        myResult.rank = 0

        Return myResult
    End Function







    '-----------------------------------------------------------------------------------------------






    Private Function DetermineWinner(result1 As Result, result2 As Result, hand1() As String, hand2() As String) As Integer
        If result1.rank > result2.rank Then
            Return 1
        ElseIf result1.rank < result2.rank Then
            Return 2
        ElseIf result1.primary > result2.primary Then
            Return 1
        ElseIf result1.primary < result2.primary Then
            Return 2
        ElseIf result1.secondary > result2.secondary Then
            Return 1
        ElseIf result1.secondary < result2.secondary Then
            Return 2
        Else
            Dim i As Integer
            For i = 0 To 4
                If hand1(i) > hand2(i) Then
                    Return 1
                ElseIf hand1(i) < hand2(i) Then
                    Return 2
                End If
            Next
        End If
        Return 0 ' Tie
    End Function





    'Reset-----------------------------------------------------------------------------------------
    Private Sub RESET()
        pKeep = {True, True, True, True, True}
        dKeep = {True, True, True, True, True}
        numDiscard = 0

        pCard1.BorderStyle = BorderStyle.None
        pCard2.BorderStyle = BorderStyle.None
        pCard3.BorderStyle = BorderStyle.None
        pCard4.BorderStyle = BorderStyle.None
        pCard5.BorderStyle = BorderStyle.None
    End Sub
    '----------------------------------------------------------------------------------------------

    Private Sub EndDes_Tick(sender As Object, e As EventArgs) Handles EndDes.Tick
        Static count As Integer = 0
        If count = 0 Then
            count = 1
        Else
            count = 0

            Dim score As Integer
            pResult = Rank(pHand)
            score = DetermineWinner(pResult, dResults, pHand, dHand)

            If score = 0 Then
                lblDraw.Visible = True
                WinOrLose.Enabled = True
            End If

            If score = 1 Then
                lblWIN.Visible = True
                WinOrLose.Enabled = True
            End If

            If score = 2 Then
                lblLOSE.Visible = True
                WinOrLose.Enabled = True
            End If

            EndDes.Enabled = False
        End If
    End Sub

    Private Sub PlayerHandToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PlayerHandToolStripMenuItem.Click
        For i = 0 To 4
            pHand(i) = InputBox("Card" + (i + 1).ToString, "PlayerHand")
        Next
        hDISPLAY(pHand, 0)
    End Sub

    Private Sub DealerHandToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DealerHandToolStripMenuItem.Click
        For i = 0 To 4
            dHand(i) = InputBox("Card" + (i + 1).ToString, "DealerHand")
        Next
        hDISPLAY(dHand, 1)
    End Sub
End Class



Public Class Deck

    Public Cards = {"14s", "02s", "03s", "04s", "05s",
                    "06s", "07s", "08s", "09s", "10s",
                    "11s", "12s", "13s", "14d", "02d",
                    "03d", "04d", "05d", "06d", "07d",
                    "08d", "09d", "10d", "11d", "12d",
                    "13d", "13c", "12c", "11c", "10c",
                    "09c", "08c", "07c", "06c", "05c",
                    "04c", "03c", "02c", "14c", "13h",
                    "12h", "11h", "10h", "09h", "08h",
                    "07h", "06h", "05h", "04h", "03h",
                    "02h", "14h"}
    Private ranGen As New Random
    Shared CIP As String = String.Empty


    Public Function Draw()

        Dim temp As String
        Dim hand(0 To 4) As String

        For i = 0 To 4 Step 1
            Randomize()
            temp = Cards(ranGen.Next(0, 51))
            If CIP.Contains(temp) Then
                i = i - 1
            Else
                hand(i) = temp
                CIP = CIP + temp
            End If
        Next

        Return hand
    End Function

    Public Function Discard(ByVal phand, ByVal dishand)
        Dim temp As String

        For i = 0 To 4 Step 1
            If (dishand(i) = False) Then

                For i2 = 0 To 1 Step 1
                    Randomize()
                    temp = Cards(ranGen.Next(0, 51))
                    If CIP.Contains(temp) Then
                        i2 = i2 - 1
                    Else
                        CIP = CIP + temp
                        phand(i) = temp
                    End If
                Next
            End If
        Next
        Return phand
    End Function

    Public Function DealerDiscard(ByRef dealerh() As String, ByVal keep() As Boolean)
        Dim temp As String

        For i = 0 To 4 Step 1
            If (keep(i) = False) Then

                For i2 = 0 To 1 Step 1
                    Randomize()
                    temp = Cards(ranGen.Next(0, 51))
                    If CIP.Contains(temp) Then
                        i2 = i2 - 1
                    Else
                        CIP = CIP + temp
                        dealerh(i) = temp
                    End If
                Next
            End If
        Next
        Return dealerh
    End Function

    Public Sub KCIP()
        CIP = String.Empty
    End Sub



End Class
