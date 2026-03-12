Imports System.Data.SqlTypes
Imports System.Diagnostics.Eventing
Imports System.Reflection
Imports System.Transactions
Imports System.Xml.Schema

Module Program
    Sub Main()
        Dim index As Integer
        Dim valid(51) As Boolean
        For index = 0 To 51
            valid(index) = False
        Next
        Dim Card(51) As Cards
        readData(Card)
        Dim BlackJack As Boolean = False
        Dim RemainingCards As Integer = 0
        Dim Player(51) As Cards
        Dim Dealer(51) As Cards
        Dim answer As String = ""
        Dim x As Integer
        Dim count As Integer
        Dim DealerTotal As Integer
        Dim Total As Integer
        Dim MyMoney As Integer
        Dim MyBet As Integer
        Dim ChooseInsurance As Boolean
        Dim InsuranceCost As Integer
        Dim acecard As Boolean
        Dim dealerace As Boolean
        Dim AddMoney As String = ""
        MoneyAdded(MyMoney)
        Do Until answer = "n"
            If MyMoney = 0 Or AddMoney.ToLower = "y" Then
                MoneyAdded(MyMoney)
            End If
            AddMoney = ""
            Total = 0
            DealerTotal = 0
            If RemainingCards > 290 Then
                ResetValues(valid, Player, Dealer, RemainingCards)
                count = 0
                x = 0
                Console.WriteLine("")
                Console.WriteLine("Deck has been reset!")
                Console.WriteLine("")
            End If
            MyBet = 0
            Do Until MyBet <= MyMoney And MyBet > 0
                Console.WriteLine("How much money do you want to bet?")
                MyBet = Console.ReadLine()
                If MyBet > MyMoney Then
                    NotEnough(MyBet, MyMoney, answer)
                End If
            Loop
            StartGame(Player, Dealer, Card, RemainingCards, valid, x,
                      count, DealerTotal, Total, BlackJack, MyBet, ChooseInsurance, acecard)
            If Total = 21 And DealerTotal <> 21 Then
                MyMoney = MyMoney + (MyBet * 1.5)
                Console.WriteLine("-------------")
                Console.WriteLine("You win via BlackJack!")
                Console.WriteLine("New money total: " & MyMoney)
                Console.WriteLine("-------------")
            ElseIf ChooseInsurance = True And Total <> 21 Then
                InsuranceCost = Insurance(MyBet, MyMoney)
            End If
            If DealerTotal = 21 And Total <> 21 Then
                Console.WriteLine("-------------")
                Console.WriteLine("Dealer's cards: ")
                Dealer(count - 1).GetCard()
                Console.WriteLine("")
                Dealer(count).GetCard()
                MyMoney = MyMoney - MyBet
                Console.WriteLine("")
                Console.WriteLine("-------------")
                Console.WriteLine("Dealer wins via BlackJack!")
                Console.WriteLine("New money total: " & MyMoney)
                Console.WriteLine("-------------")
            ElseIf Total = 21 And DealerTotal = 21 Then
                Console.WriteLine("-------------")
                Console.WriteLine("Dealer's cards: ")
                Dealer(count - 1).GetCard()
                Console.WriteLine("")
                Dealer(count).GetCard()
                Console.WriteLine("")
                Console.WriteLine("-------------")
                Console.WriteLine("Tie!")
                Console.WriteLine("Money total: " & MyMoney)
                Console.WriteLine("-------------")
            End If
            If Total < 21 And Total <> 21 And DealerTotal <> 21 Then
                If (((ConvertNumber(Player(x - 1).GetNumber()) = ConvertNumber(Player(x).GetNumber())) Or
            Player(x - 1).GetNumber() = Player(x).GetNumber())) Then
                    Split(Player, Dealer, Card, RemainingCards, valid, x, count, DealerTotal, MyMoney, MyBet, answer)
                    If answer = "n" Then
                        DoubleBet(MyBet, MyMoney, valid, Card, Player, RemainingCards, x, Total, acecard, answer, DealerTotal, Dealer, count)
                        If answer = "n" Then
                            PlayGame(Player, Dealer, Card, RemainingCards, valid, x, count, DealerTotal, Total, MyMoney, MyBet, acecard)
                            DealersDraw(DealerTotal, Total, Dealer, count, valid, Card, RemainingCards, dealerace)
                        Else
                            DealersDraw(DealerTotal, Total, Dealer, count, valid, Card, RemainingCards, dealerace)
                        End If
                        Winner(Total, DealerTotal, MyMoney, MyBet)
                    End If
                Else
                    DoubleBet(MyBet, MyMoney, valid, Card, Player, RemainingCards, x, Total, acecard, answer, DealerTotal, Dealer, count)
                    If answer = "n" Then
                        PlayGame(Player, Dealer, Card, RemainingCards, valid, x, count, DealerTotal, Total, MyMoney, MyBet, acecard)
                        DealersDraw(DealerTotal, Total, Dealer, count, valid, Card, RemainingCards, dealerace)
                    Else
                        DealersDraw(DealerTotal, Total, Dealer, count, valid, Card, RemainingCards, dealerace)
                    End If
                    Winner(Total, DealerTotal, MyMoney, MyBet)
                End If

            End If
            x = x + 1
            count = count + 1
            Console.WriteLine("Do you want to play again?[y/n] ")
            answer = Console.ReadLine.ToLower
            Console.WriteLine("")
            dealerace = False
            acecard = False
            If answer = "y" Then
                Console.WriteLine("Would you like to add more money?[y/n]")
                AddMoney = Console.ReadLine
            End If
        Loop
    End Sub
    Sub StartGame(ByRef Player() As Cards, ByRef Dealer() As Cards, ByVal Card() As Cards,
                  ByRef RemainingCards As Integer, ByRef valid() As Boolean, ByRef x As Integer,
                  ByRef count As Integer, ByRef DealerTotal As Integer, ByRef Total As Integer, ByRef BlackJack As Boolean,
                  ByRef MoneyBet As Integer, ByRef ChooseInsurance As Boolean, ByRef acecard As Boolean)
        Dim cardnum As Integer
        Dim NewCard As Boolean = True
        Dim tempValue As Integer
        BlackJack = False
        acecard = False
        ChooseInsurance = False
        cardnum = DrawCard(valid, RemainingCards)
        Dealer(count) = Card(cardnum)
        Console.WriteLine("---------------")
        Console.WriteLine("Dealer's card: ")
        Console.WriteLine("  ")
        Dealer(count).GetCard()
        If Dealer(count).GetNumber() = "A" Then
            tempValue = ChooseValue(DealerTotal)
            ChooseInsurance = True
        Else
            tempValue = ConvertNumber(Dealer(count).GetNumber())
        End If
        CardValue(tempValue, DealerTotal, acecard)
        Console.WriteLine("---------------")
        count = count + 1
        cardnum = DrawCard(valid, RemainingCards)
        Dealer(count) = Card(cardnum)
        If Dealer(count).GetNumber() = "A" Then
            tempValue = ChooseValue(DealerTotal)
        Else
            tempValue = ConvertNumber(Dealer(count).GetNumber())
        End If
        CardValue(tempValue, DealerTotal, acecard)
        If BlackJack = True Then
            Console.WriteLine("Blackjack!")
        End If
        Console.WriteLine("Your cards: ")
        Console.WriteLine("  ")
        cardnum = DrawCard(valid, RemainingCards)
        Player(x) = Card(cardnum)

        Player(x).GetCard()
        Console.WriteLine("  ")
        x = x + 1
        cardnum = DrawCard(valid, RemainingCards)
        Player(x) = Card(cardnum)
        Player(x).GetCard()
        Console.WriteLine("")
        If Player(x - 1).GetNumber() = "A" Then
            tempValue = ChooseValue(Total)
            acecard = True
            CardValue(tempValue, Total, acecard)
        Else
            tempValue = ConvertNumber(Player(x - 1).GetNumber())
            CardValue(tempValue, Total, acecard)
        End If
        If Player(x).GetNumber() = "A" Then
            tempValue = ChooseValue(Total)
            acecard = True
        Else
            tempValue = ConvertNumber(Player(x).GetNumber())
        End If
        CardValue(tempValue, Total, acecard)
    End Sub
    Sub PlayGame(ByRef Player() As Cards, ByRef Dealer() As Cards, ByVal Card() As Cards,
                 ByRef RemainingCards As Integer, ByRef valid() As Boolean,
                 ByRef x As Integer, ByRef count As Integer, ByRef DealerTotal As Integer, ByRef Total As Integer,
                 ByRef MyMoney As Integer, ByRef MyBet As Integer, ByRef acecard As Boolean)
        Dim tempValue As Integer
        Dim BlackJack As Boolean = False
        Dim dealercard1 As Integer
        Dim dealerAce As Boolean
        Console.WriteLine("---------------")
        Console.WriteLine("Your total: " & Total)
        If Dealer(count - 1).GetNumber = "A" Then
            dealercard1 = 11
            dealerAce = True
        Else
            dealercard1 = ConvertNumber(Dealer(count - 1).GetNumber())
        End If
        Console.WriteLine("Dealer's total: " & dealercard1)
        Console.WriteLine("---------------")
        Console.WriteLine("")
        Dim NewCard As Boolean = True
        Dim cardnum As Integer
        Do Until NewCard = False Or Burned(Total) = True Or Total = 21
            NewCard = AddCard()
            Console.WriteLine("  ")
            If NewCard = True Then
                x = x + 1
                cardnum = DrawCard(valid, RemainingCards)
                Player(x) = Card(cardnum)
                Console.WriteLine("----------------")
                Player(x).GetCard()
                If Player(x).GetNumber = "A" Then
                    tempValue = ChooseValue(Total)
                    acecard = True
                    Console.WriteLine("-----------------")
                Else
                    tempValue = ConvertNumber(Player(x).GetNumber())
                End If
                CardValue(tempValue, Total, acecard)
                Console.WriteLine("----------------")
                Console.WriteLine("Your total: " & Total)
                Console.WriteLine("----------------")
                Console.WriteLine("")
            End If
        Loop

    End Sub
    Sub readData(ByRef Card() As Cards)
        Dim MyFile As IO.StreamReader
        Dim FileName As String = "C:\Users\combi\source\repos\Blackjack\Blackjack\DeckOfCards.txt"
        Dim number As String
        Dim colour As String
        Dim type As String
        Dim x As Integer = 0
        Try
            MyFile = IO.File.OpenText(FileName)
            While MyFile.Peek <> -1
                colour = MyFile.ReadLine
                type = MyFile.ReadLine
                number = MyFile.ReadLine.ToString()
                Card(x) = New Cards(colour, type, number)
                x = x + 1
            End While
        Catch ex As Exception
            Console.WriteLine("Invalid file!")
        End Try
    End Sub
    Function ConvertNumber(number As String) As Integer
        Dim ValueOfCard As Integer

        If number = "J" Then
            ValueOfCard = 10
        ElseIf number = "K" Then
            ValueOfCard = 10
        ElseIf number = "Q" Then
            ValueOfCard = 10
        Else
            ValueOfCard = Val(number)
        End If
        Return ValueOfCard

    End Function
    Function DrawCard(ByRef Taken() As Boolean, ByRef RemainingCards As Integer) As Integer
        Dim cardnum As Integer
        Dim generator As New Random()
        cardnum = generator.Next(0, 51)
        If Taken(cardnum) = True Then
            Do Until Taken(cardnum) = False
                cardnum = generator.Next(0, 51)
            Loop
        End If
        Taken(cardnum) = True
        RemainingCards = RemainingCards + 1
        Return cardnum
    End Function
    Function AddCard() As Boolean
        Console.WriteLine("Do you want to draw another card? [y/n] ")
        If Console.ReadLine().ToLower() = "y" Then
            Return True
        Else
            Return False
        End If
    End Function
    Sub Ace(ByRef total As Integer)
        If total > 21 Then
            total = total - 10
        End If
    End Sub
    Function ChooseValue(Total As Integer)
        Dim cardNum As Integer
        If Total <= 10 Then
            cardNum = 11
        Else
            cardNum = 1
        End If
        Return cardNum
    End Function
    Sub CardValue(ByVal CardValue As Integer, ByRef TotalNumber As Integer, AceCard As Boolean)
        TotalNumber = TotalNumber + CardValue
        If AceCard = True Then
            Ace(TotalNumber)
        End If
    End Sub
    Function Burned(Total As Integer) As Boolean
        If Total > 21 Then
            Return True
        Else
            Return False
        End If

    End Function
    Sub ResetValues(ByRef valid() As Boolean, ByRef player() As Cards, ByRef dealer() As Cards, ByRef remainingcards As Integer)
        Dim x As Integer
        For x = 0 To 311
            valid(x) = False
            player(x) = New Cards("", "", "")
            dealer(x) = New Cards("", "", "")
        Next
        remainingcards = 0
    End Sub
    Sub NotEnough(ByRef MyBet As Integer, ByRef MyMoney As Integer, ByRef answer As String)
        If MyBet > MyMoney - MyBet Then
            Console.WriteLine("")
            Console.WriteLine("Not enough money to support this action!")
            Console.WriteLine("")
            Console.WriteLine("-------------------")
            Console.WriteLine("(Currenct money balace: " & (MyMoney - MyBet) & ")")
            Console.WriteLine("(Current bet: " & MyBet & ")")
            Console.WriteLine("-------------------")
            Console.WriteLine("")
            Console.WriteLine("Would you like to add more money?[y/n]")
            answer = Console.ReadLine().ToLower
            If answer = "y" Then
                MoneyAdded(MyMoney)
                Do Until MyBet * 2 >= MyMoney
                    MoneyAdded(MyMoney)
                Loop
            End If
        End If
    End Sub
    Sub Split(ByVal Player() As Cards, ByVal Dealer() As Cards, ByVal Card() As Cards,
                  ByRef RemainingCards As Integer, ByRef valid() As Boolean, ByRef x As Integer,
                  ByRef count As Integer, ByRef DealerTotal As Integer, ByRef MyMoney As Integer, ByRef MyBet As Integer, ByRef answer As String)
        Dim Total(1) As Integer
        Dim cardNum As Integer
        Dim tempValue As Integer
        Dim card1(1) As Integer
        Dim NewCard As Boolean = True
        Dim dealerace As Boolean
        Dim counter As Integer = 0
        Dim sum As Integer = 0
        Dim BlackJack(1) As Boolean
        BlackJack(0) = False
        BlackJack(1) = False
        Dim index As Integer
        Dim currentBet As Integer
        Dim acecard(1) As Boolean
        acecard(0) = False
        acecard(1) = False
        currentBet = MyBet
        index = x
        Dim Ans As String = "y"
        If Player(x - 1).GetNumber = "A" Then
            tempValue = ChooseValue(sum)
            acecard(0) = True
            acecard(1) = True
        Else
            tempValue = ConvertNumber(Player(x).GetNumber)
        End If
        CardValue(tempValue, sum, acecard(0))
        If Player(x - 1).GetNumber = "A" Then
            tempValue = ChooseValue(sum)
            acecard(0) = True
        Else
            tempValue = ConvertNumber(Player(x - 1).GetNumber)
        End If
        CardValue(tempValue, sum, acecard(1))
        x = x + 1
        Console.WriteLine("Do you want to split?[y/n] ")
        answer = Console.ReadLine().ToLower
        If answer = "y" Then
            NotEnough(MyBet, MyMoney, Ans)
            If Ans = "y" Then
                MyMoney = MyMoney - MyBet
                Do Until counter = 2
                    MyBet = currentBet
                    If Player(index).GetNumber = "A" Then
                        card1(counter) = ChooseValue(Total(counter))
                        acecard(counter) = True
                    Else
                        card1(counter) = ConvertNumber(Player(index).GetNumber)
                    End If
                    CardValue(card1(counter), Total(counter), acecard(counter))
                    cardNum = DrawCard(valid, RemainingCards)
                    Player(x) = Card(cardNum)
                    If Player(x).GetNumber() = "A" Then
                        tempValue = ChooseValue(Total(counter))
                        acecard(counter) = True
                    Else
                        tempValue = ConvertNumber(Player(x).GetNumber())
                    End If
                    CardValue(tempValue, Total(counter), acecard(counter))
                    Console.WriteLine("------------")
                    Player(x).GetCard()
                    Console.WriteLine("------------")
                    Console.WriteLine("Your total: " & Total(counter))
                    Console.WriteLine("------------")
                    Console.WriteLine("")
                    If Total(counter) = 21 And DealerTotal <> 21 Then
                        MyMoney = MyMoney + (MyBet * 1.5)
                        BlackJack(counter) = True
                    Else
                        DoubleBet(MyBet, MyMoney, valid, Card, Player, RemainingCards, x, Total(counter), acecard(counter), answer, DealerTotal, Dealer, count)
                        If answer = "n" Then
                            PlayGame(Player, Dealer, Card, RemainingCards, valid, x, count, DealerTotal, Total(counter), MyMoney, MyBet, acecard(counter))
                        Else
                            MyMoney = MyMoney - MyBet
                        End If
                    End If
                    counter = counter + 1
                    index = index - 1
                Loop
                x = x + 1
                DealersDraw(DealerTotal, Total(0), Dealer, count, valid, Card, RemainingCards, dealerace)
                Console.WriteLine("--------------")
                Console.WriteLine("First pair: ")
                Console.WriteLine("Your total: " & Total(0))
                Console.WriteLine("--------------")
                Console.WriteLine("")
                Console.WriteLine("--------------")
                Console.WriteLine("Second pair: ")
                Console.WriteLine("Your total: " & Total(1))
                Console.WriteLine("--------------")
                Console.WriteLine("")
                Console.WriteLine("--------------")
                Console.WriteLine("Dealer's total: " & DealerTotal)
                Console.WriteLine("--------------")
                Console.WriteLine("")
                Console.WriteLine("--------------")
                MyMoney = MyMoney + MyBet
                Console.WriteLine("Verdict for pair 1:")
                If BlackJack(0) = True Then
                    Console.WriteLine("You win via BlackJack!")
                    Console.WriteLine("--------------")
                Else
                    Winner(Total(0), DealerTotal, MyMoney, MyBet)
                End If
                Console.WriteLine("")
                Console.WriteLine("--------------")
                Console.WriteLine("Verdict for second pair: ")
                If BlackJack(1) = True Then
                    Console.WriteLine("You win via BlackJack!")
                    Console.WriteLine("--------------")
                Else
                    Winner(Total(1), DealerTotal, MyMoney, MyBet)
                End If
                Console.WriteLine("")
                answer = "y"
            Else
                answer = "n"
            End If
        Else
            answer = "n"
        End If
    End Sub
    Sub Winner(Total As Integer, DealerTotal As Integer, ByRef TotalMoney As Integer, ByVal MoneyBet As Integer)
        Console.WriteLine("")
        Console.WriteLine("-----------------")
        Console.WriteLine("Your total: " & Total)
        Console.WriteLine("Dealer's total: " & DealerTotal)
        Console.WriteLine("-----------------")
        Console.WriteLine("")
        If DealerTotal > Total And DealerTotal <= 21 Then
            TotalMoney = TotalMoney - MoneyBet
            Console.WriteLine("Dealer wins!")
            Console.WriteLine("")
            Console.WriteLine("--------------")
            Console.WriteLine("New total amount of money: " & TotalMoney)
            Console.WriteLine("--------------")
        ElseIf Total > DealerTotal And Total <= 21 Then
            Console.WriteLine("You win!")
            TotalMoney = TotalMoney + MoneyBet
            Console.WriteLine("")
            Console.WriteLine("--------------")
            Console.WriteLine("New total amount of money: " & TotalMoney)
            Console.WriteLine("--------------")
        ElseIf Total = DealerTotal Then
            Console.WriteLine("Tie")
            Console.WriteLine("")
            Console.WriteLine("--------------")
            Console.WriteLine("New total amount of money: " & TotalMoney)
            Console.WriteLine("--------------")
        ElseIf DealerTotal > 21 And Total <= 21 Then
            Console.WriteLine("You win!")
            TotalMoney = MoneyBet + TotalMoney
            Console.WriteLine("")
            Console.WriteLine("--------------")
            Console.WriteLine("New total amount of money: " & TotalMoney)
            Console.WriteLine("--------------")
        ElseIf Total > 21 And DealerTotal <= 21 Then
            TotalMoney = TotalMoney - MoneyBet
            Console.WriteLine("Dealer wins!")
            Console.WriteLine("")
            Console.WriteLine("--------------")
            Console.WriteLine("New total amount of money: " & TotalMoney)
            Console.WriteLine("--------------")
        End If
    End Sub
    Sub DealersDraw(ByRef DealerTotal As Integer, Total As Integer, ByRef Dealer() As Cards, ByRef count As Integer,
                    ByRef valid() As Boolean, Card() As Cards, ByRef RemainingCards As Integer, ByVal dealerace As Boolean)
        Console.WriteLine("----------------")
        Console.WriteLine("Dealer's cards: ")
        Console.WriteLine("")
        Dealer(count - 1).GetCard()
        Dim cardnum As Integer
        Dim tempValue As Integer
        Dealer(count).GetCard()
        If Dealer(count - 1).GetNumber() = "A" Or Dealer(count).GetNumber() = "A" Then
            dealerace = True
        End If
        If DealerTotal <= 16 And Burned(Total) = False Then
            Do Until DealerTotal > 16 Or Burned(DealerTotal) = True
                cardnum = DrawCard(valid, RemainingCards)
                count = count + 1
                Dealer(count) = Card(cardnum)
                If Dealer(count).GetNumber() = "A" Then
                    If DealerTotal > 10 Then
                        tempValue = 1
                        dealerace = True
                    Else
                        tempValue = 11
                    End If
                Else
                    tempValue = ConvertNumber(Dealer(count).GetNumber())
                End If
                CardValue(tempValue, DealerTotal, dealerace)
                Dealer(count).GetCard()
                Console.WriteLine("----------------")
                Console.WriteLine("Dealer's total: " & DealerTotal)
                Console.WriteLine("----------------")
                Console.WriteLine("")
            Loop
        End If
    End Sub
    Function Insurance(ByRef MoneyBet As Integer, ByRef TotalMoney As Integer) As Integer
        Dim answer As String
        Dim InsuranceCost As Integer
        Console.WriteLine("Would you like insurace?[y/n]")
        answer = Console.ReadLine().ToLower
        If answer = "y" Then
            InsuranceCost = MoneyBet / 2
            TotalMoney = TotalMoney - InsuranceCost
        Else
            InsuranceCost = MoneyBet
        End If
        Return InsuranceCost
    End Function
    Sub DoubleBet(ByRef MoneyBet As Integer, ByRef totalMoney As Integer, ByRef valid() As Boolean, card() As Cards,
                  ByRef player() As Cards, ByRef RemainingCards As Integer, ByRef x As Integer, ByRef total As Integer,
                   ByRef acecard As Boolean, ByRef answer As String, ByRef dealertotal As Integer, Dealer() As Cards,
                  count As Integer)
        Dim cardnum As Integer
        Dim tempValue As Integer
        If Dealer(count - 1).GetNumber = "A" Then
            tempValue = 11
        Else
            tempValue = ConvertNumber(Dealer(count - 1).GetNumber())
        End If
        Console.WriteLine("-----------------")
        Console.WriteLine("Your total: " & total)
        Console.WriteLine("Dealer's total: " & tempValue)
        Console.WriteLine("-----------------")
        Console.WriteLine("")
        Console.WriteLine("Would you like to double your bet?[y/n]")
        answer = Console.ReadLine().ToLower()
        If answer = "y" Then
            If MoneyBet > totalMoney - MoneyBet Then
                NotEnough(MoneyBet, totalMoney, answer)
            End If
            Console.WriteLine("")
            If answer = "y" Then
                If totalMoney >= MoneyBet * 2 Then
                    MoneyBet = MoneyBet * 2
                    Console.WriteLine("New bet amount: " & MoneyBet)
                End If
                cardnum = DrawCard(valid, RemainingCards)
                x = x + 1
                player(x) = card(cardnum)
                If player(x).GetNumber() = "A" Then
                    tempValue = ChooseValue(total)
                    acecard = True
                Else
                    tempValue = ConvertNumber(player(x).GetNumber)
                End If
                CardValue(tempValue, total, acecard)
                Console.WriteLine("-----------------")
                Console.WriteLine("New card: ")
                Console.WriteLine("")
                player(x).GetCard()
                Console.WriteLine("-----------------")
                Console.WriteLine("Your total: " & total)
                Console.WriteLine("Dealer's total: " & dealertotal)
                Console.WriteLine("-----------------")
            Else
                answer = "n"
            End If
        End If
    End Sub
    Sub MoneyAdded(ByRef totalMoney As Integer)
        Dim NewMoney As Integer
        Do Until NewMoney >= 5
            Console.WriteLine("How much money do you want to add? (min 5 euros)")
            NewMoney = Console.ReadLine()
        Loop
        totalMoney = totalMoney + NewMoney
    End Sub
End Module


