Imports GV8APILib
Imports System
Imports System.IO
Imports System.Linq
Imports System.Xml
Imports System.Xml.Linq
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.DataTable


Public Class Form1
    'Dim mApi As GV8APILib.GlobalVision8API
    Private logger As NLog.Logger = NLog.LogManager.GetCurrentClassLogger()
    Dim isConnected As Boolean = False


    Dim sQuery As String = "<CRITERIA><TRADES><TRADEDATE From='%dtFrom%T00:00:00' Until='%dtTo%T23:59:59'/></TRADES></CRITERIA>"
    Dim sQryinstDef As String = "<CRITERIA><INSTDEFINITIONS></INSTDEFINITIONS></CRITERIA>"
    'Dim sQryInstAttr = "<CRITERIA><INSTATTRIBUTESDATA/></CRITERIA>"
    Dim sQryInstAttr As String = "<CRITERIA><INSTATTRIBUTESDATA/></CRITERIA>"
    'Dim sQryInstAttr = "<?xml version=""1.0"" standalone='yes'?><CRITERIA><INSTPROPERTIES/></CRITERIA>"
    Dim sQryGroup As String = "<CRITERIA><GROUPS/></CRITERIA>"
    Dim sQryTradeSummary As String = "<CRITERIA><INSTTRADESUMMARIES/></CRITERIA>"
    Dim sCompaniesQuery As String = "<CRITERIA> <COMPANIES>All</COMPANIES></CRITERIA>"
    Dim sqryOrderSummaries As String = "<CRITERIA><ORDERSUMMARIES/></CRITERIA>"

    Dim sqryInstOrderSummaries As String = "<CRITERIA> <INSTORDERSUMMARIES></INSTORDERSUMMARIES></CRITERIA>"

    Dim xmldTrades As XDocument, xmlInstCodes As XDocument

    Dim FAPP_PATH As String = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
    Dim TRADES_XML_FILE_NAME As String = FAPP_PATH & "\Trades.xml"
    Dim INST_XML_FILE_NAME As String = FAPP_PATH & "\Institutions.xml"
    Dim PRICES_XML_FILE_NAME As String = FAPP_PATH & "\Prices.xml"
    Dim TRADESUMMARY_XML_FILE_NAME As String = FAPP_PATH & "\TradeSummary.xml"
    Dim INST_ORDER_SUMMARY_XML_FILE_NAME As String = FAPP_PATH & "\InstOrderSummaries.xml"

    Dim INSDEF_XML_FILE_NAME As String = FAPP_PATH & "\InsDef.xml"
    Dim INSATT_XML_FILE_NAME As String = FAPP_PATH & "\InsAtt.xml"
    Dim GROUP_XML_FILE_NAME As String = FAPP_PATH & "\Group.xml"
    Dim ORDER_XML_FILE_NAME As String = FAPP_PATH & "\Order.xml"
    Dim COMPANIES_XML_FILE_NAME As String = FAPP_PATH & "\Companies.xml"

    Dim ORDER_SUMMARIES_XML_FILE_NAME As String = FAPP_PATH & "\OrderSummaries.xml"

    'Dim TRADES_CSV_FILE_NAME As String = FAPP_PATH & "\DailyTrades.csv"
    Dim myFileName As String = String.Format("\DailyTrades_{0}.csv", Now.ToString("MMddyyyy_hhmm"))
    Dim TRADES_CSV_FILE_NAME As String = FAPP_PATH & myFileName

    Dim myPricesFileName As String = String.Format("\DailyPrices_{0}.csv", Now.ToString("MMddyyyy_hhmm"))

    Dim DAILYPRICES_CSV_FILE_NAME As String = FAPP_PATH & myPricesFileName

    Dim REPLACE_STRING As String = "encoding=""UTF-16"""
    Dim REPLACE_STRING2 As String = "xmlns=""gv8api-trayport-com"""

    Dim BLOCK_TYPE As String = ""



    Private Sub Connect()
        Dim sResult As String
        Dim sTrades As String
        Dim sQuery As String
        'Dim reader As XmlReader

        Dim Login As New Login()
        Login.ShowDialog()

        If Jse.connected = "True" Then
            Label1.Text = "Connected"
            btnConnect.Enabled = False
            btnDisconnect.Enabled = True
            btnChangePass.Enabled = True
            btnTrades.Enabled = True
            btnOrders.Enabled = True
            isConnected = True
        End If


        If Jse.CodeFromApi = "303" Then
            ' Password needs changing
            ' MsgBox("")
        End If
    End Sub
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnConnect.Click
        Connect()
    End Sub

    Private Sub Disconnect_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDisconnect.Click
        If isConnected Then mApi.Logout()

        Label1.Text = "Not Connected"
        btnDisconnect.Enabled = False
        btnConnect.Enabled = True
        btnChangePass.Enabled = False
        btnTrades.Enabled = False
        btnOrders.Enabled = False
        TextBox2.Clear()
        TextBox1.Clear()
        isConnected = False

    End Sub

    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If Not IsNothing(mApi) Then
            mApi.Logout()
        End If

        Application.Exit()
    End Sub

    'Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    Dim dtFrom As DateTime
    '    dtFrom = DateTimePicker1.Value
    '    MsgBox(dtFrom.ToString("yyyy-MM-dd"))
    'End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        initForm()
    End Sub

    Private Sub GetXMLFromAPIGroups()
        Dim sGroupsQuery As String
        sGroupsQuery = mApi.QueryXMLRecordSet(sQryGroup)
        Dim objWriter As New System.IO.StreamWriter(GROUP_XML_FILE_NAME, True)
        sGroupsQuery = sGroupsQuery.Replace(REPLACE_STRING, "").Trim
        ' Strip out the namespace
        sGroupsQuery = sGroupsQuery.Replace(REPLACE_STRING2, "").Trim

        ' create XML file
        objWriter.Write(sGroupsQuery)
        objWriter.Close()
    End Sub
    Private Sub GetXMLAPITrades(ByVal sTrades As String, ByVal sReplace As String)
        sTrades = mApi.QueryXMLRecordSet(sReplace)
        Dim objWriter As New System.IO.StreamWriter(TRADES_XML_FILE_NAME, True)
        sTrades = sTrades.Replace(REPLACE_STRING, "").Trim
        ' create XML file
        objWriter.Write(sTrades)
        objWriter.Close()
    End Sub
    Private Function SetQueryParameters() As String
        Dim dtFrom As DateTime
        Dim dtTo As DateTime
        Dim sReplace As String
        dtFrom = DateTimePicker1.Value
        dtTo = DateTimePicker2.Value

        sReplace = sQuery.Replace("%dtFrom%", dtFrom.ToString("yyyy-MM-dd")).Trim
        sReplace = sReplace.Replace("%dtTo%", dtTo.ToString("yyyy-MM-dd")).Trim
        Return sReplace
    End Function
    Private Sub GetXMLFromAPIINstitutions(ByVal sInstQuery As String)
        Dim result As String
        result = mApi.QueryXMLRecordSet(sInstQuery)
        result = result.Replace(REPLACE_STRING, "").Trim

        Dim objWriter As New System.IO.StreamWriter(INST_XML_FILE_NAME, True)
        ' create XML file
        objWriter.Write(result)
        objWriter.Close()
    End Sub
    Private Sub GetXMLFromAPICompanies(ByVal sCompaniesQuery As String)
        Dim result As String

        If sCompaniesQuery.Length <= 0 Then
            MsgBox("Invalid companies query")
        End If

        result = mApi.QueryXMLRecordSet(sCompaniesQuery)
2:      Dim objWriter As New System.IO.StreamWriter(COMPANIES_XML_FILE_NAME, True)
        ' create XML file
        result = result.Replace(REPLACE_STRING, "").Trim
        ' Strip out the namespace
        result = result.Replace(REPLACE_STRING2, "").Trim
        objWriter.Write(result)
        objWriter.Close()
    End Sub
    Private Sub btnTrades_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnTrades.Click
        Dim dtFrom As DateTime, dtTo As DateTime
        Dim sTrades As String, sReplace As String, sCompaniesQuery As String, sInstQuery As String, sGroupsQuery As String

        initForm()

        sTrades = ""
        sReplace = SetQueryParameters()
        'get trades now
        Try
            GetXMLAPITrades(sTrades, sReplace)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        Try
            GetXMLFromAPIGroups()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        sCompaniesQuery = "<CRITERIA><COMPANIES>All</COMPANIES></CRITERIA>"
        'next get all companies
        Try
            GetXMLFromAPICompanies(sCompaniesQuery)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        ' Now get institutions details
        sInstQuery = "<CRITERIA> <INSTDEFINITIONS>All</INSTDEFINITIONS></CRITERIA>"
        Try
            GetXMLFromAPIINstitutions(sInstQuery)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        ParseXmlToCsv()
    End Sub


    Private Sub ParseXmlToCsv()
        Dim dsTrades As New DataSet
        Dim tTrades As DataTable
        Dim tTe8rm As DataTable
        Dim InstSpecifier As DataTable
        Dim tOrders As DataTable
        Dim tOrder As DataTable
        Dim sHeader As String, sBody As String
        Dim sDate As DateTime


        Dim dsInst As New DataSet
        Dim tInst As DataTable

        Dim dsCompanies As New DataSet
        Dim tCompanies As DataTable

        Dim dsGroups As New DataSet
        Dim tGroups As DataTable

        Dim GroupsXML As New XmlDocument

        GroupsXML.Load(GROUP_XML_FILE_NAME)

        ' create trades dataset and tables-------------
        dsTrades.Tables.Clear()
        dsTrades.ReadXml(TRADES_XML_FILE_NAME)

        ' Check if the datastore contain tables
        If ((dsTrades.Tables.Count < 5) Or (IsNothing(dsTrades))) Then
            MsgBox("Insufficient Data to generate report from Trades XML. Exiting.")
            Exit Sub
        End If

        ' create companies dataset and tables-------------
        dsCompanies.Tables.Clear()
        dsCompanies.ReadXml(COMPANIES_XML_FILE_NAME)

        ' Check if the datastore contain tables
        If ((dsCompanies.Tables.Count < 1) Or (IsNothing(dsCompanies))) Then
            MsgBox("Insufficient Data to generate report from Companies XML. Exiting.")
            Exit Sub
        End If



        tTrades = dsTrades.Tables(0)
        InstSpecifier = dsTrades.Tables(1)
        tTerm = dsTrades.Tables(2)
        tOrders = dsTrades.Tables(3)
        tOrder = dsTrades.Tables(4)

        ' Get companies data
        tCompanies = dsCompanies.Tables(0)

        'now get groups 
        dsGroups.Tables.Clear()
        dsGroups.ReadXml(GROUP_XML_FILE_NAME)

        ' Check if the datastore contain tables
        If ((dsGroups.Tables.Count < 1) Or (IsNothing(dsGroups))) Then
            MsgBox("Insufficient Data to generate report from Groups XML. Exiting.")
            Exit Sub
        End If
        tGroups = dsGroups.Tables(0)

        ' create Inst dataset and tables------------
        dsInst.Tables.Clear()
        dsInst.ReadXml(INST_XML_FILE_NAME)
        tInst = dsInst.Tables(0)

        ' Check if the datastore contain tables
        If ((dsInst.Tables.Count < 1) Or (IsNothing(dsInst))) Then
            MsgBox("Insufficient Data to generate report from Institutions XML. Exiting.")
            Exit Sub
        End If

        Dim results = From a In tTrades.AsEnumerable
                      Select a


        'create file
        If results.Count > 0 Then
            Dim objWriter As New System.IO.StreamWriter(TRADES_CSV_FILE_NAME, True)

            'write header first
            sHeader = "TDATE,SYMBOL,MARKET,VOLUME,TPRICE,TICKET,SETTDATE,BUYORD,BUYACCT,BBROKERNO,SELORD,SELLACCT,SBROKERNO,VALUE,CANCEL,BBREF,SBREF,BUYREF,SELLREF,CURRENCY"

            sBody = ""

            objWriter.WriteLine(sHeader)

            For Each item In results
                Dim sAction As String, sSymbol As String
                Dim BuyOrder As String, SelOrder As String
                Dim fPrice As Decimal, fVolume As Decimal
                Dim iTradeId As String, sAccountReference As String, sBrokerNumber As String
                Dim iTrade_id As Integer
                Dim iValue As Double
                Dim ManualDeal As String, PriceSetting As String
                Dim sGroupType As String
                Dim sMarketCode As String
                Dim sCurrencyCode As String
                Dim iCompanyId As String, iCompanyId2 As String
                Dim iCounterParty As String
                Dim sAggressorCompanyId As String
                Dim sInitiatorCompanyId As String

                ' get the institution id
                Dim iInstId As Integer


                'build string and write to csve
                sDate = item.Field(Of String)("DateTime")
                sBody = sDate.ToString("yyyyMMdd")

                sAggressorCompanyId = item.Field(Of String)("AggressorCompanyID")
                sInitiatorCompanyId = item.Field(Of String)("InitiatorCompanyID")

                'now check if symbol is in list 
                If (sAggressorCompanyId.Equals("11") Or sInitiatorCompanyId.Equals("11")) Then
                    ' skip 
                Else
                    Continue For
                End If

                fVolume = item.Field(Of String)("Volume")
                fPrice = item.Field(Of String)("Price")

                iValue = (Convert.ToDouble(fVolume) * Convert.ToDouble(fPrice))
                iTradeId = item.Field(Of String)("TradeID")
                iTrade_id = item.Field(Of Integer)("TRADE_Id")
                sAction = item.Field(Of String)("InitiatorAction")
                ManualDeal = item.Field(Of String)("ManualDeal")
                PriceSetting = item.Field(Of String)("PriceSetting")

                iCompanyId = item.Field(Of String)("AggressorCompanyID")
                iCompanyId2 = item.Field(Of String)("InitiatorCompanyID")
                'iCounterParty = item.Field(Of String)("InitiatorCompanyID")

                Dim Companyresult = From a In tCompanies.AsEnumerable
                                Where a("CompanyID") = iCompanyId
                                Select a

                'Buyer Broker number
                Dim sCompanyCode = Companyresult.First.Field(Of String)("CompanyCode")


                ' seller broker number
                Dim Companyresult2 = From a In tCompanies.AsEnumerable
                                Where a("CompanyID") = iCompanyId2
                                Select a

                Dim sCompanyCode2 = Companyresult2.First.Field(Of String)("CompanyCode")


                BuyOrder = ""
                SelOrder = ""

                If (ManualDeal = "True" And PriceSetting = "False") Then
                    BLOCK_TYPE = "BLOCK"
                Else
                    BLOCK_TYPE = "NON_BLOCK"
                End If

                Dim specresult = From a In InstSpecifier.AsEnumerable
                                  Where a("TRADE_Id") = iTrade_id
                                   Select a

                If specresult.Count > 0 Then

                    Dim InstCode = From a In tInst.AsEnumerable
                                Where a("InstID") = specresult.First.Field(Of String)(0)
                                Select a

                    sSymbol = InstCode.First.Field(Of String)(2)
                Else
                    sSymbol = ""
                End If

                iInstId = specresult.First.Field(Of String)(0)

                ' call function passing in iInstId
                sGroupType = GroupType.GroupCategory.GetGroupCategory(GroupsXML, iInstId)

                'sGroupType = "Regular"
                If ((sGroupType = "Regular") And (BLOCK_TYPE = "BLOCK")) Then
                    sMarketCode = "513"
                    sCurrencyCode = "JMD"
                ElseIf ((sGroupType = "Regular") And (BLOCK_TYPE = "NON_BLOCK")) Then
                    sMarketCode = "510"
                    sCurrencyCode = "JMD"
                ElseIf ((sGroupType = "Junior") And (BLOCK_TYPE = "BLOCK")) Then
                    sMarketCode = "518"
                    sCurrencyCode = "JMD"
                ElseIf ((sGroupType = "Junior") And (BLOCK_TYPE = "NON_BLOCK")) Then
                    sMarketCode = "519"
                    sCurrencyCode = "JMD"
                ElseIf ((sGroupType = "USD") And (BLOCK_TYPE = "BLOCK")) Then
                    sMarketCode = "513"
                    sCurrencyCode = "USD"
                Else
                    sMarketCode = "510"
                    sCurrencyCode = "USD"
                End If

                'logger.Info("FROM VB ---GroupType : {0} ----- MarketCode : {1} ----InstId {2}", sGroupType, sMarketCode, iInstId)

                '-----------------------
                sBody = sBody & ", " & sSymbol
                'Get Market and set
                sBody = sBody & ", " & sMarketCode
                'Set
                sBody = sBody & "," & fVolume.ToString & "," & fPrice.ToString & "," & iTradeId.ToString & ", ,"

                'get the orders resultset
                 Dim Orderresults = From a In tOrders.AsEnumerable
                                Where a("TRADE_Id") = iTrade_id
                                 Select a

                If Orderresults.Count > 0 Then
                    Dim order = From a In tOrder.AsEnumerable
                                      Where a("ORDERS_Id") = Orderresults.First.Field(Of Integer)("ORDERS_Id")
                                       Select a
                    BuyOrder = ""
                    SelOrder = ""

                    Dim ctr As Integer = 1
                    For Each orderitem In order
                        If ctr = 1 Then
                            If sAction = "Buy" Then
                                BuyOrder = Trim(orderitem.Field(Of String)(2))
                            Else
                                SelOrder = Trim(orderitem.Field(Of String)(2))
                            End If
                        Else
                            If sAction = "Sell" Then
                                SelOrder = Trim(orderitem.Field(Of String)(2))
                            Else
                                BuyOrder = Trim(orderitem.Field(Of String)(2))
                            End If
                        End If
                        ctr += 1

                    Next
                End If

                If BuyOrder = "0" Then
                    BuyOrder = " "
                End If

                If SelOrder = "0" Then
                    SelOrder = " "
                End If

                ' logger.Info("FROM VB ---BuyOrder : {0} ----- SellOrde : {1} ", BuyOrder, SelOrder)

                '--------------------------get broker numer------------
                If sAction = "Buy" Then
                    sBrokerNumber = item.Field(Of String)("AggressorBrokerID")
                Else
                    sBrokerNumber = item.Field(Of String)("InitiatorBrokerID")
                End If

                Dim sSellerAccNum = ""

                ' get account reference - 


                'get seller broker number
                Dim SellerBrokerNum = ""
                ' add persisitentOrderdId1 and 2 to body

                Dim sBuyAccountNumber As String

                ' get account_reference
                Dim term = From a In tTerm.AsEnumerable
                            Where a("TRADE_Id") = iTrade_id And
                            a("Label") = "Account Reference" And
                             Not DBNull.Value.Equals(a("Counterparty")) AndAlso
                            a("Counterparty") = "buyer"
                            Select a


                If term.Count > 0 Then
                    sBuyAccountNumber = term.First.Field(Of String)("TERM_Text")
                Else
                    sBuyAccountNumber = " "
                End If

                '-----------------now get sell account 
                Dim term1 = From b In tTerm.AsEnumerable
                            Where b("TRADE_Id") = iTrade_id And
                            b("Label") = "Account Reference" And
                            Not DBNull.Value.Equals(b("Counterparty")) AndAlso
                            b("Counterparty") = "seller"
                            Select b


                If term1.Count > 0 Then
                    sSellerAccNum = term1.First.Field(Of String)("TERM_Text")
                Else
                    sSellerAccNum = " "
                End If

                ' sHeader = "TDATE,SYMBOL,MARKET,VOLUME,TPRICE,TICKET,SETTDATE,BUYORD,BUYACCT,BBROKERNO,SELORD,SELLACCT,SBROKERNO,VALUE,CANCEL,BBREF,SBREF,BUYREF,SELLREF,CURRENCY"
                'change broker number from sBrokerNumber to iCompanyId
                sBody = sBody & BuyOrder & "," & sBuyAccountNumber & "," & sCompanyCode2 & "," _
                        & SelOrder & "," & sSellerAccNum & "," & sCompanyCode & "," & _
                        iValue & "," & "," & "," & "," & "," & "," & sCurrencyCode

                ' logger.Info(sBody)

                objWriter.WriteLine(sBody)
            Next

            objWriter.Close()
        End If

        MsgBox("Completed report generation. Report ""TradeData.csv"" is located in the MyDocuments folder!")
    End Sub


    Private Sub ResetDailyPricesXML()
        If System.IO.`File.Exists(TRADES_XML_FILE_NAME) = True Then
            System.IO.File.Delete(TRADES_XML_FILE_NAME)
        End If

        If System.IO.File.Exists(TRADESUMMARY_XML_FILE_NAME) = True Then
            System.IO.File.Delete(TRADESUMMARY_XML_FILE_NAME)
        End If

        If System.IO.File.Exists(INSATT_XML_FILE_NAME) = True Then
            System.IO.File.Delete(INSATT_XML_FILE_NAME)
        End If

        If System.IO.File.Exists(GROUP_XML_FILE_NAME) = True Then
            System.IO.File.Delete(GROUP_XML_FILE_NAME)
        End If

        If System.IO.File.Exists(ORDER_SUMMARIES_XML_FILE_NAME) = True Then
            System.IO.File.Delete(ORDER_SUMMARIES_XML_FILE_NAME)
        End If

        If System.IO.File.Exists(INSDEF_XML_FILE_NAME) = True Then
            System.IO.File.Delete(INSDEF_XML_FILE_NAME)
        End If
    End Sub
    Private Sub GetXMLFromAPITradeSummary()
        Dim sTradesSummary As String
        sTradesSummary = mApi.QueryXMLRecordSet(sQryTradeSummary)
        Dim objWriter5 As New System.IO.StreamWriter(TRADESUMMARY_XML_FILE_NAME, True)
        sTradesSummary = sTradesSummary.Replace(REPLACE_STRING, "").Trim
        sTradesSummary = sTradesSummary.Replace(REPLACE_STRING2, "").Trim
        objWriter5.Write(sTradesSummary)
        objWriter5.Close()
    End Sub
    Private Sub GetXMLFromAPIInstrumentDefinitions()
        Dim sInstDefinitions As String
        sInstDefinitions = mApi.QueryXMLRecordSet(sQryinstDef)
        Dim objWriter1 As New System.IO.StreamWriter(INSDEF_XML_FILE_NAME, True)
        sInstDefinitions = sInstDefinitions.Replace(REPLACE_STRING, "").Trim
        objWriter1.Write(sInstDefinitions)
        objWriter1.Close()
    End Sub
    Private Sub GetXMLFromAPIInstrumentAttributes()
        Dim sInsAttributes As String
        sInsAttributes = mApi.QueryXMLRecordSet(sQryInstAttr)
        Dim objWriter2 As New System.IO.StreamWriter(INSATT_XML_FILE_NAME, True)
        sInsAttributes = sInsAttributes.Replace(REPLACE_STRING, "").Trim
        objWriter2.Write(sInsAttributes)
        objWriter2.Close()
    End Sub
    Private Sub GetXMLFromAPIOrderSummaries()
        Dim sresult = mApi.QueryXMLRecordSet(sqryInstOrderSummaries)
        Dim objWriter As New System.IO.StreamWriter(ORDER_SUMMARIES_XML_FILE_NAME, True)
        sresult = sresult.Replace(REPLACE_STRING, "").Trim
        ' Strip out the namespace
        sresult = sresult.Replace(REPLACE_STRING2, "").Trim

        ' create XML file
        objWriter.Write(sresult)
        objWriter.Close()
    End Sub
    Private Sub Button2_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim dsTrades As New DataSet
        Dim tTrades As DataTable
        Dim tTerm As DataTable
        Dim tOrders As DataTable
        Dim tOrder As DataTable
        Dim sHeader As String, sBody As String
        Dim TEST_FILE_NAME As String

        ' NOTE strip out encoding UTF-16
        TEST_FILE_NAME = "C:\Users\andrew\Desktop\TestData.csv"

        dsTrades.ReadXml(TRADES_XML_FILE_NAME)
        ' create dataset
        tTrades = dsTrades.Tables(0)
        tTerm = dsTrades.Tables(2)
        tOrders = dsTrades.Tables(3)
        tOrder = dsTrades.Tables(4)

        Dim results = From a In tTrades.AsEnumerable
                        Where a("AggressorCompanyID") = 10 Or a("InitiatorCompanyID") = 10
                        Select a

        'create file
        If results.Count > 0 Then
            Dim objWriter As New System.IO.StreamWriter(TEST_FILE_NAME, True)

            'write header first
            sHeader = "TDATE,SYMBOL,MARKET,VOLUME,TPRICE,TICKET,SETTDATE,BUYORD,BUYACCT,BBROKERNO,SELORD,SELLACCT,SBROKERNO,VALUE,CANCEL,BBREF,SBREF,BUYREF,SELLREF,CURRENCY"
            sBody = ""
            objWriter.WriteLine(sHeader)
            For Each item In results
                Dim sAction As String
                Dim BuyOrder As String, SelOrder As String
                Dim fPrice As Decimal, fVolume As Decimal
                Dim iTradeId As String, sAccountReference As String, sBrokerNumber As String
                Dim iTrade_id As Integer
                Dim iValue As Double

                'build string and write to csve
                sBody = item.Field(Of String)("DateTime")
                'Get symbol here and set
                sBody = sBody & ", "
                'Get Market and set
                sBody = sBody & ", "
                fVolume = item.Field(Of String)("Volume")
                fPrice = item.Field(Of String)("Price")

                iValue = (Convert.ToDouble(fVolume) * Convert.ToDouble(fPrice))
                iTradeId = item.Field(Of String)("TradeID")
                iTrade_id = item.Field(Of Integer)("TRADE_Id")
                sAction = item.Field(Of String)("InitiatorAction")
                BuyOrder = ""
                SelOrder = ""
                'Set

                Dim sSellerAccNum = ""
                Dim SellerBrokerNum = ""
                'Dim sAccountReference = ""


                sBody = sBody & "," & fVolume.ToString & "," & fPrice.ToString & "," & iTradeId.ToString & ", ,"

                '--------------------------get broker numer------------
                If sAction = "Buy" Then
                    sBrokerNumber = item("AggressorBrokerID")
                Else
                    sBrokerNumber = item("InitiatorBrokerID")
                End If
                sBody = sBody & BuyOrder & "," & sAccountReference & "," & sBrokerNumber & "," _
                        & SelOrder & "," & sSellerAccNum & "," & SellerBrokerNum & "," & _
                        iValue & "," & sAction & "," & "," & "," & ","


                objWriter.WriteLine(sBody)
            Next

            objWriter.Close()
        End If

        MsgBox("Completed..!")
    End Sub

    Private Sub btnOrders_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOrders.Click

        ResetDailyPricesXML()
        'Get Daily Price report
        Dim dtFrom As DateTime, dtTo As DateTime
        Dim sReplace As String, sInsAttributes As String, sInstDefinitions As String, sTrades As String, sTradesSummary As String

        dtFrom = DateTimePicker1.Value
        dtTo = DateTimePicker2.Value

        sReplace = sQuery.Replace("%dtFrom%", dtFrom.ToString("yyyy-MM-dd")).Trim
        sReplace = sReplace.Replace("%dtTo%", dtTo.ToString("yyyy-MM-dd")).Trim

        ' get XML for instDefinitions
        Try
            '-------------Trades--------------
            sTrades = mApi.QueryXMLRecordSet(sReplace)
            Dim objWriter4 As New System.IO.StreamWriter(TRADES_XML_FILE_NAME, True)

            sTrades = sTrades.Replace(REPLACE_STRING, "").Trim
            sTrades = sTrades.Replace(REPLACE_STRING2, "").Trim

            objWriter4.Write(sTrades)
            objWriter4.Close()

            '-------------- TradeSummary--------------------
            GetXMLFromAPITradeSummary()
            '-----------InsDef------------
            GetXMLFromAPIInstrumentDefinitions()
            '--------InsAttr----
            GetXMLFromAPIInstrumentAttributes()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        '-------------------------Groups------------------------------------
        Try
            Dim sresult = mApi.QueryXMLRecordSet(sQryGroup)
            Dim objWriter As New System.IO.StreamWriter(GROUP_XML_FILE_NAME, True)
            sresult = sresult.Replace(REPLACE_STRING, "").Trim
            ' Strip out the namespace
            sresult = sresult.Replace(REPLACE_STRING2, "").Trim

            ' create XML file
            objWriter.Write(sresult)
            objWriter.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        '------------------------Order Summaries-----------------

        Try
            GetXMLFromAPIOrderSummaries()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try


        ' now create the CSV for Daily Prices
        CreateDailyPrices()
    End Sub

    Private Sub CreateDailyPrices()
        Dim dsTrades As New DataSet
        Dim tTrades As New DataTable
        Dim tTerm As New DataTable
        Dim tInstSpecifier As New DataTable
        Dim tOrders As New DataTable
        Dim tOrder As New DataTable

        Dim sHeader As String, sBody As String

        Dim dsInstDefinitions As New DataSet
        Dim tInstDefinitions As New DataTable

        Dim dsInsatrributes As New DataSet
        Dim tInsatrributes As New DataTable

        Dim dsInst As New DataSet
        Dim tInst As New DataTable

        Dim sTrades As String

        Dim dsTradeSummary As New DataSet
        Dim dtTradeSummary As New DataTable

        Dim dsInstOrderSummary As New DataSet
        Dim dtInstOrderSummary As New DataTable
        Dim dtInstSpecifier As New DataTable

        Dim dsExcel As New DataSet
        Dim dtExcel As New DataTable
        Dim MyConnection As System.Data.OleDb.OleDbConnection
        Dim MyCommand As System.Data.OleDb.OleDbDataAdapter

        MyConnection = New System.Data.OleDb.OleDbConnection("provider=Microsoft.ACE.OLEDB.12.0; " & _
 "data source='" & FAPP_PATH & "\HistoricalTradeSummary.xlsx" & " '; " & "Extended Properties=Excel 12.0;")

        Try
            MyCommand = New System.Data.OleDb.OleDbDataAdapter("select * from [sheet1$]", MyConnection)

            MyCommand.TableMappings.Add("Table", "HistoricTradeSummary")
            dsExcel.Tables.Clear()
            dsExcel = New System.Data.DataSet
            MyCommand.Fill(dsExcel)
            dtExcel = dsExcel.Tables(0)
        Catch ex As Exception

        End Try
        ' Select the data from Sheet1 of the workbook.

        Dim GroupsXML As New XmlDocument

        GroupsXML.Load(GROUP_XML_FILE_NAME)

        '' 1 - create trades dataset and tables-------------
        Try
            dsTrades.Tables.Clear()
            dsTrades.ReadXml(TRADES_XML_FILE_NAME)

            '' Check if the datastore contain tables
            If ((dsTrades.Tables.Count < 5) Or (IsNothing(dsTrades))) Then
                MsgBox("Insufficient Data to generate report from Trades XML. Exiting.")
                Exit Sub
            End If

            tTrades = dsTrades.Tables(0)
            tInstSpecifier = dsTrades.Tables(1)
            tTerm = dsTrades.Tables(2)
            tOrders = dsTrades.Tables(3)
            tOrder = dsTrades.Tables(4)

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        ' create InstAttribute dataset and tables-------------
        dsInsatrributes.Tables.Clear()
        dsInsatrributes.ReadXml(INSATT_XML_FILE_NAME)
        tInsatrributes = dsInsatrributes.Tables(0)

        ' create Inst dataset and tables------------
        dsInst.Tables.Clear()
        dsInst.ReadXml(INSDEF_XML_FILE_NAME)
        tInst = dsInst.Tables(0)

        ' Check if the datastore contain tables
        If ((dsInst.Tables.Count < 1) Or (IsNothing(dsInst))) Then
            MsgBox("Insufficient Data to generate report from Institutions XML. Exiting.")
            Exit Sub
        End If

        ' create TradeSummaries dataset and tables------------
        dsTradeSummary.Tables.Clear()
        dsTradeSummary.ReadXml(TRADESUMMARY_XML_FILE_NAME)
        dtTradeSummary = dsTradeSummary.Tables(0)

        ' Check if the datastore contain tables
        If ((dsTradeSummary.Tables.Count < 1) Or (IsNothing(dsTradeSummary))) Then
            MsgBox("Insufficient Data to generate report from Trade Summary XML. Exiting.")
            Exit Sub
        End If


        ' create OrderSummaries dataset and tables------------
        dsInstOrderSummary.Tables.Clear()
        dsInstOrderSummary.ReadXml(ORDER_SUMMARIES_XML_FILE_NAME)
        dtInstSpecifier = dsInstOrderSummary.Tables(1)
        dtInstOrderSummary = dsInstOrderSummary.Tables(0)


        ' Check if the datastore contain tables
        If ((dsInstOrderSummary.Tables.Count < 1) Or (IsNothing(dsInstOrderSummary))) Then
            MsgBox("Insufficient Data to generate report from Order Summary XML. Exiting.")
            Exit Sub
        End If

        'haeder
        sHeader = "SYMBOL_CODE,SYMBOL_NAME,SECTOR_CODE,SECTOR_DESCRIPTION,LABELED_STATE_CODE,MARKET_CODE,TDEX_CORP_ACTION,HIGH_PRICE_52,HIGH_PRICE_52_IND,LOW_PRICE_52,LOW_PRICE,LOW_PRICE_52_IND,PRE_DIVIDEND_AMOUNT,PRE_DIV_CURR,PRE_CORPORATE_ACTION,DIVIDEND_AMOUNT,DIV_CURR,CORPORATE_ACTION,HIGH_PRICE,CLOSE_PRICE,CLOSE_NET_CHANGE,CLOSE_PERCENT_CHANGE,LASTTRADEDPRICE,BID_PRICE,ASK_PRICE,TOTALQTYTRADEDTODAY,TRADE_VALUE,TRADES_COUNT"

        Dim Traderesults = From a In tTrades.AsEnumerable
                            Where a("AggressorCompanyID") = "10" OrElse
                              a("InitiatorCompanyID") = "10"
                            Select a

        If Traderesults.Count > 0 Then
            Dim objWriter As New System.IO.StreamWriter(DAILYPRICES_CSV_FILE_NAME, True)

            'write header first
            sBody = ""

            objWriter.WriteLine(sHeader)

            Dim Symbollist As New List(Of String)
            
            For Each item In Traderesults
                'logger.Info("Processing Record #:{0}", results.Count)
                Dim sAction As String, sSymbol As String
                Dim BuyOrder As String, SelOrder As String
                Dim fPrice As Decimal, fVolume As Decimal
                Dim iTradeId As String, sAccountReference As String, sBrokerNumber As String
                Dim iTrade_id As Integer
                Dim iValue As Double
                Dim ManualDeal As String, PriceSetting As String
                Dim sGroupType As String
                Dim sMarketCode As String
                Dim sCurrencyCode As String
                Dim iInstId As Integer
                Dim TradeDate As Date

                ManualDeal = item.Field(Of String)("ManualDeal")
                PriceSetting = item.Field(Of String)("PriceSetting")
                iTradeId = item.Field(Of String)("TradeID")
                iTrade_id = item.Field(Of Integer)("TRADE_Id")


                Dim sDate = item.Field(Of String)("DateTime")
                TradeDate = DateTime.Parse(sDate).ToString("dd-MM-yyyy")

                Dim specresult = From a In tInstSpecifier.AsEnumerable
                                      Where a("TRADE_Id") = iTrade_id
                                       Select a

                'get name
                Dim InstName = specresult.First.Field(Of String)(6)

                ' Get Symbol
                If specresult.Count > 0 Then
                    Dim InstCode = From a In tInst.AsEnumerable
                                Where a("InstID") = specresult.First.Field(Of String)(0)
                                Select a

                    sSymbol = InstCode.First.Field(Of String)(2)
                Else
                    sSymbol = ""
                End If

                'now check if symbol is in list 
                If Symbollist.Contains(sSymbol) Then
                    Continue For
                Else
                    Symbollist.Add(sSymbol)
                End If

                iInstId = specresult.First.Field(Of String)(0)

                sBody = ""

                ' start filtering here
                ' add symbol and instname,sect code, sect desc, labe state code
                sBody = sSymbol & "," & InstName & "," & "," & ","

                logger.Info("Institution Name #:{0}", InstName)

                'now get market code
                BuyOrder = ""
                SelOrder = ""

                If (ManualDeal = "True" And PriceSetting = "False") Then
                    BLOCK_TYPE = "BLOCK"
                Else
                    BLOCK_TYPE = "NON_BLOCK"
                End If

                ' call function passing in iInstId
                sGroupType = GroupType.GroupCategory.GetGroupCategory(GroupsXML, iInstId)

                'sGroupType = "Regular"
                If ((sGroupType = "Regular") And (BLOCK_TYPE = "BLOCK")) Then
                    sMarketCode = "513"
                    sCurrencyCode = "JMD"
                ElseIf ((sGroupType = "Regular") And (BLOCK_TYPE = "NON_BLOCK")) Then
                    sMarketCode = "510"
                    sCurrencyCode = "JMD"
                ElseIf ((sGroupType = "Junior") And (BLOCK_TYPE = "BLOCK")) Then
                    sMarketCode = "518"
                    sCurrencyCode = "JMD"
                ElseIf ((sGroupType = "Junior") And (BLOCK_TYPE = "NON_BLOCK")) Then
                    sMarketCode = "519"
                    sCurrencyCode = "JMD"
                ElseIf ((sGroupType = "USD") And (BLOCK_TYPE = "BLOCK")) Then
                    sMarketCode = "513"
                    sCurrencyCode = "USD"
                Else
                    sMarketCode = "510"
                    sCurrencyCode = "USD"
                End If

                ' now read excel for tradesummary a
                Dim dInstCode = Convert.ToDouble(Trim(iInstId))

                Dim ExcelResults = From a In dtExcel.AsEnumerable _
                                  Where a("SummaryDate") = TradeDate AndAlso
                                  a("InstId") = dInstCode
                                  Select a

                Dim LastTradedPrice As Double, HighPrice As Double

                If ExcelResults.Count > 0 Then

                    'If Not (ExcelResults.First.Field(Of String)("LastTradedPrice") Is System.DBNull) Then
                    If (ExcelResults.First.Field(Of String)("LastTradedPrice")) <> "NULL" Then
                        LastTradedPrice = ExcelResults.First.Field(Of String)("LastTradedPrice")
                    Else
                        LastTradedPrice = 0
                    End If

                    If ((ExcelResults.First.Field(Of String)("BestPriceTradedToday"))) <> "NULL" Then
                        HighPrice = ExcelResults.First.Field(Of String)("BestPriceTradedToday")
                    Else
                        HighPrice = 0
                    End If
                Else
                    LastTradedPrice = 0
                    HighPrice = 0
                End If

                Dim InstSpecResult = From a In dtInstSpecifier.AsEnumerable
                                      Where a("InstID") = iInstId.ToString
                                       Select a

                Dim bidPrice As String, askPrice As String


                If InstSpecResult.Count > 0 Then

                    Dim OrdSummaryResult = From a In dtInstOrderSummary.AsEnumerable
                               Where a("INSTORDERSUMMARY_Id") = InstSpecResult.First.Field(Of Integer)(9)
                               Select a


                    bidPrice = OrdSummaryResult.First.Field(Of String)(4)
                    askPrice = OrdSummaryResult.First.Field(Of String)(10)
                Else
                    bidPrice = ""
                    askPrice = ""
                End If

                ' add market code , tede_corp_action
                sBody = sBody & "," & sMarketCode & "," & "," & "," & "," & "," & "," & "," & "," & "," & "," & "," & "," & "," & HighPrice _
                        & "," & "," & "," & "," & LastTradedPrice & "," & bidPrice & "," & askPrice & "," & ","

                objWriter.WriteLine(sBody)
            Next

            objWriter.Close()
            MsgBox("Completed report generation. Report ""DailyPrices.csv"" is located in the MyDocuments folder!")
        End If

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim Login As New Login()

        Login.ShowDialog()

        MsgBox("user Name is : " & Jse.username & " --- password is " & Jse.OldPassword)

    End Sub

    Private Sub btnChangePass_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnChangePass.Click
        Dim ResetPass As New ResetPassword()
        ResetPass.ShowDialog()

        ' if password changed successfully then show message

    End Sub

    Private Sub Button2_Click_2(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim dsTrades As New DataSet
        Dim tTrades As DataTable
        Dim tTerm As DataTable
        Dim InstSpecifier As DataTable
        Dim tOrders As DataTable
        Dim tOrder As DataTable
        Dim sHeader As String, sBody As String
        Dim sDate As DateTime


        ' create trades dataset and tables-------------
        dsTrades.Tables.Clear()
        dsTrades.ReadXml(TRADES_XML_FILE_NAME)

        ' Check if the datastore contain tables
        If ((dsTrades.Tables.Count < 5) Or (IsNothing(dsTrades))) Then
            MsgBox("Insufficient Data to generate report from Trades XML. Exiting.")
            Exit Sub
        End If

        tTrades = dsTrades.Tables(0)
        InstSpecifier = dsTrades.Tables(1)
        tTerm = dsTrades.Tables(2)
        tOrders = dsTrades.Tables(3)
        tOrder = dsTrades.Tables(4)


        Dim results = From a In tTrades.AsEnumerable
                        Where a("AggressorCompanyID") = "10" OrElse
                              a("InitiatorCompanyID") = "10"
                        Select a

        'create file
        If results.Count > 0 Then
            Dim objWriter As New System.IO.StreamWriter(TRADES_CSV_FILE_NAME, True)
            'write header first
            sHeader = "TDATE,SYMBOL,MARKET,VOLUME,TPRICE,TICKET,SETTDATE,BUYORD,BUYACCT,BBROKERNO,SELORD,SELLACCT,SBROKERNO,VALUE,CANCEL,BBREF,SBREF,BUYREF,SELLREF,CURRENCY"
            sBody = ""

            objWriter.WriteLine(sHeader)

            For Each item In results
                Dim sAction As String, sSymbol As String
                Dim BuyOrder As String, SelOrder As String
                Dim fPrice As Decimal, fVolume As Decimal
                Dim iTradeId As String, sAccountReference As String, sBrokerNumber As String
                Dim iTrade_id As Integer
                Dim iValue As Double
                Dim ManualDeal As String, PriceSetting As String
                Dim sGroupType As String
                Dim sMarketCode As String
                Dim sCurrencyCode As String
                Dim iCompanyId As String, iCompanyId2 As String
                Dim iCounterParty As String

                ' get the institution id
                Dim iInstId As Integer

                'build string and write to csve
                sDate = item.Field(Of String)("DateTime")
                sBody = sDate.ToString("yyyyMMdd")

                fVolume = item.Field(Of String)("Volume")
                fPrice = item.Field(Of String)("Price")
                iValue = (Convert.ToDouble(fVolume) * Convert.ToDouble(fPrice))
                iTradeId = item.Field(Of String)("TradeID")
                iTrade_id = item.Field(Of Integer)("TRADE_Id")
                sAction = item.Field(Of String)("InitiatorAction")
                ManualDeal = item.Field(Of String)("ManualDeal")
                PriceSetting = item.Field(Of String)("PriceSetting")

                iCompanyId = item.Field(Of String)("AggressorCompanyID")
                iCompanyId2 = item.Field(Of String)("InitiatorCompanyID")
                'iCounterParty = item.Field(Of String)("InitiatorCompanyID")

                'Buyer Broker number
                Dim sCompanyCode = ""
                Dim sCompanyCode2 = ""

                BuyOrder = ""
                SelOrder = ""

                iInstId = ""

                sBody = sBody & "," & fVolume.ToString & "," & fPrice.ToString & "," & iTradeId.ToString & ", ,"

                BuyOrder = " "
                SelOrder = " "
                sBrokerNumber = ""
                Dim sSellerAccNum = ""

                sBody = sBody & BuyOrder & "," & "," & sCompanyCode2 & "," _
                        & SelOrder & "," & sSellerAccNum & "," & sCompanyCode & "," & _
                        iValue & "," & "," & "," & "," & "," & "," & sCurrencyCode
                objWriter.WriteLine(sBody)
            Next

            objWriter.Close()
        End If

        MsgBox("Completed report generation. Report ""TradeData.csv"" is located in the MyDocuments folder!")
    End Sub


    Private Sub initForm()

        If Not isConnected Then
            btnDisconnect.Enabled = False
            btnConnect.Enabled = True
            btnTrades.Enabled = False
            btnOrders.Enabled = False
        End If

        
        If System.IO.File.Exists(TRADES_XML_FILE_NAME) = True Then
            System.IO.File.Delete(TRADES_XML_FILE_NAME)
        End If

        ' if exists delete the xml files
        If System.IO.File.Exists(INST_XML_FILE_NAME) = True Then
            System.IO.File.Delete(INST_XML_FILE_NAME)
        End If

        If System.IO.File.Exists(PRICES_XML_FILE_NAME) = True Then
            System.IO.File.Delete(PRICES_XML_FILE_NAME)
        End If

        If System.IO.File.Exists(TRADES_CSV_FILE_NAME) = True Then
            System.IO.File.Delete(TRADES_CSV_FILE_NAME)
        End If

        If System.IO.File.Exists(DAILYPRICES_CSV_FILE_NAME) = True Then
            System.IO.File.Delete(DAILYPRICES_CSV_FILE_NAME)
        End If

        If System.IO.File.Exists(INSDEF_XML_FILE_NAME) = True Then
            System.IO.File.Delete(INSDEF_XML_FILE_NAME)
        End If

        If System.IO.File.Exists(INSATT_XML_FILE_NAME) = True Then
            System.IO.File.Delete(INSATT_XML_FILE_NAME)
        End If

        If System.IO.File.Exists(GROUP_XML_FILE_NAME) = True Then
            System.IO.File.Delete(GROUP_XML_FILE_NAME)
        End If

        If System.IO.File.Exists(ORDER_XML_FILE_NAME) = True Then
            System.IO.File.Delete(ORDER_XML_FILE_NAME)
        End If

        If System.IO.File.Exists(TRADESUMMARY_XML_FILE_NAME) = True Then
            System.IO.File.Delete(TRADESUMMARY_XML_FILE_NAME)
        End If

        If System.IO.File.Exists(COMPANIES_XML_FILE_NAME) = True Then
            System.IO.File.Delete(COMPANIES_XML_FILE_NAME)
        End If


    End Sub

End Class