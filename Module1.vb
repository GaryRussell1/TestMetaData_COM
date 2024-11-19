Module Module1

    Private Bx As BettingAssistantCom.Application.ComClass
    Sub Main()
        test()
    End Sub

    Sub test()
        Bx = New BettingAssistantCom.Application.ComClass
        Dim flagKeepGoing As Boolean = True
        Dim strStoreMarketName As String
        Dim strStoreLastMarketName As String
        Bx.clearQuickPick(0)
        Bx.refreshMarkets()
        Bx.loadQuickPickList(1)
        Bx.openLastQuickPickMarket() 'this is so we know what the last market is in the picklist so we stop the loop
        Threading.Thread.Sleep(2000) ' milliseconds
        strStoreLastMarketName = Bx.marketName
        Bx.openFirstQuickPickMarket()
        Threading.Thread.Sleep(2000)
        Dim strList As New List(Of String)
        Dim strList2 As New List(Of String)
        Dim strList3 As New List(Of String)
        Dim strList4 As New List(Of String)
        Dim storePrices As New Dictionary(Of String, Double)
        Dim storeForecast As New Dictionary(Of String, Double)
        Do While flagKeepGoing
            strStoreMarketName = Bx.marketName
            Console.WriteLine(strStoreMarketName)
            Dim prices As Object = Bx.getPrices
            Dim _str As String = ""
            Dim _sel As String = prices(0).Selection
            Dim _met As Object = Bx.getMetaData(_sel)
            If Not IsNothing(_met) AndAlso IsNumeric(_met.foreCastPrice) Then
                Threading.Thread.Sleep(50)
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
                        Continue For
                    End If
                    Console.WriteLine($"Horse: {_meta.selection}, Saddlecloth: {_meta.saddleCloth},  Jockey: {_meta.jockey}, Trainer: {_meta.trainer}, Owner: {_meta.owner}, Age/weight: {_meta.ageWeight}, Jockey claim: {_meta.jockeyClaim}, Bred: {_meta.bred}, Dam: {_meta.dam}, Sire: {_meta.sire}, DamSire: {_meta.damSire}, Colour/Sex: {_meta.colourSex}, Form: {_meta.form}, Official rating: {_meta.officialRating}, Forecast price: {_meta.forecastPrice}, Days since last run: {_meta.daysSinceLastRun}, Wearing: {_meta.wearing}, Stall draw: {_meta.stallDraw}")
                    _str = $"{strStoreMarketName}|{item.Selection}|{_meta.foreCastPrice}|{_meta.form}|{_meta.daysSinceLastRun}|{_meta.officialRating}"
                    If _meta.form <> "" And CInt(_meta.daysSinceLastRun) <= 66 Then
                        strList.Add(_str)
                    Else
                        strList4.Add($"{_str.Replace("|", ",")},Empty Form/Days > 66")
                    End If
                    If Not storePrices.ContainsKey(item.Selection) Then
                        storePrices.Add(item.Selection, item.backOdds1)
                    End If
                    If Not storeForecast.ContainsKey(item.Selection) Then
                        storeForecast.Add(item.Selection, _meta.foreCastPrice)
                    End If
                    Dim _fp As Double
                    If IsNumeric(_meta.foreCastPrice) Then
                        _fp = Trim(_meta.foreCastPrice)
                        If item.backOdds1 >= CDbl(_fp) * 2 And CDbl(_fp) < 11 Then
                            strList2.Add($"{strStoreMarketName}, {item.Selection}, {_meta.form}, {_meta.foreCastPrice}, {item.backOdds1}, {_meta.daysSinceLastRun}")
                        End If
                    Else
                        strList4.Add($"{_str.Replace("|", ",")},Forecast Price non numeric")
                    End If
                Next
            Catch ex As System.Exception
                Console.WriteLine("ERROR: In price Loop")
                Console.WriteLine(ex.ToString)
            End Try
            If strStoreMarketName = strStoreLastMarketName Then
                Bx.openFirstQuickPickMarket()
                flagKeepGoing = False
            Else
                Bx.openNextQuickPickMarket()
                Threading.Thread.Sleep(3000)
            End If
        Loop
        'Print reports to .csv
        '.......
    End Sub
End Module
