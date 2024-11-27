' This application test getMetaData COM function.
' Use with Betting Assistant version 1.3.2.8l and later.

Module Module1

    Private Bx As BettingAssistantCom.Application.ComClass

    Sub Main()
        test()
    End Sub

    Sub waitForMarketChange(currentName As String)
        While Bx.marketName = currentName
            Threading.Thread.Sleep(50)
        End While
        Dim prices As Object, pricesFound As Boolean
        Do
            ' after market loaded wait for prices to load which are loaded in a background thread
            prices = Bx.getPrices
            If IsNothing(prices) Then
                Threading.Thread.Sleep(50)
                Continue Do
            End If
            If prices(0).marketId <> Bx.marketId Then
                Threading.Thread.Sleep(50)
            Else
                pricesFound = True
            End If
        Loop Until pricesFound
    End Sub

    Sub test()
        If IsNothing(Bx) Then Bx = New BettingAssistantCom.Application.ComClass
        Dim flagKeepGoing As Boolean = True
        Dim strStoreMarketName As String = ""
        Dim strStoreLastMarketName As String
        Bx.clearQuickPick(0)
        Bx.refreshMarkets()
        Bx.loadQuickPickList(1)
        strStoreMarketName = Bx.marketName
        Bx.openLastQuickPickMarket() 'this is so we know what the last market is in the picklist so we stop the loop
        waitForMarketChange(strStoreMarketName)
        strStoreLastMarketName = Bx.marketName
        strStoreMarketName = Bx.marketName
        Bx.openFirstQuickPickMarket()
        Do While flagKeepGoing
            waitForMarketChange(strStoreMarketName)
            strStoreMarketName = Bx.marketName
            Console.WriteLine($"Retrieving meta data for {strStoreMarketName}")
            Dim prices As Object = Bx.getPrices
            Dim _str As String = ""
            Dim _sel As String = prices(0).Selection
            Dim _met As Object = Bx.getMetaData(_sel)
            If Not IsNothing(_met) AndAlso IsNumeric(_met.foreCastPrice) Then
                prices = Bx.getPrices
            Else
                If IsNothing(_met) Then Console.WriteLine("ERROR: Meta Data is Nothing on first selection.")
            End If
            Try
                '========================
                REM COLLECT the Meta Data
                '========================
                For Each item In prices
                    Dim _meta As Object = Bx.getMetaData(item.Selection)
                    If IsNothing(_meta) Then
                        Console.WriteLine("ERROR: Meta Data is Nothing")
                        Console.ReadKey()
                        Continue For
                    End If
                    If _meta.jockey = "" Or _meta.saddleCloth = "" Then
                        Console.WriteLine($"Missing meta data: Horse: {_meta.selection}, Saddlecloth: {_meta.saddleCloth},  Jockey: {_meta.jockey}, Trainer: {_meta.trainer}, Owner: {_meta.owner}, Age/weight: {_meta.ageWeight}, Jockey claim: {_meta.jockeyClaim}, Bred: {_meta.bred}, Dam: {_meta.dam}, Sire: {_meta.sire}, DamSire: {_meta.damSire}, Colour/Sex: {_meta.colourSex}, Form: {_meta.form}, Official rating: {_meta.officialRating}, Forecast price: {_meta.forecastPrice}, Days since last run: {_meta.daysSinceLastRun}, Wearing: {_meta.wearing}, Stall draw: {_meta.stallDraw}")
                    End If
                Next
            Catch ex As System.Exception
                Console.WriteLine("ERROR: In price Loop")
                Console.WriteLine(ex.ToString)
            End Try
            If strStoreMarketName = strStoreLastMarketName Then
                Bx.openFirstQuickPickMarket()
                waitForMarketChange(strStoreMarketName)
                flagKeepGoing = False
            Else
                Bx.openNextQuickPickMarket()
            End If
        Loop
        'Print reports to .csv
        '.......
    End Sub
End Module
