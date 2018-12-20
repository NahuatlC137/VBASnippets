 
Sub Historical()

    Dim ConnectionString As String
    Dim AccountDataQuery As String
    Dim TrxnQuery As String
    Dim AccountId As String
   
        If ActiveCell.Value = "" Then
            MsgBox ("Cannot run GetAccountData module on an empty cell")
            Exit Sub
        Else
            AccountId = ActiveCell.Value
        End If
        
   ' Open a connection by referencing the ODBC driver.
   ConnectionString = "ODBC;DRIVER=SQL Server;" _
        & "SERVER=;" _
        & "UID=;" _
        & "Trusted_Connection=Yes;" _
        & "APP=Microsoft Office 2010;" _
        & "WSID=;DATABASE="
    
    AccountDataQuery = "SELECT" & Chr(13) _
        & "AccountId = '''' + CONVERT(VARCHAR, ADF.AccountId)" & Chr(13) _
        & ", [Customer Name] = BOR.FirstName + ' ' + BOR.LastName" & Chr(13) _
        & ", [AmortMethod] = AD.AmortizationMethod" & Chr(13) _
        & ", [PurchaseID] = AD.PurchaseCode" & Chr(13) _
        & ", [NextDue] = CONVERT(VARCHAR, ADF.PmtNextOrPastDueDate, 101)" & Chr(13) _
        & ", [PmtAmt] = AD.PmtSchedPmtAmt" & Chr(13) _
        & "FROM [SERVER].[edw].[ALPHADELTAFOXTROT] ADF" & Chr(13) _
        & "JOIN edw.CurrentDate CD ON ADF.CalendarDate = CD.CurrentDate" & Chr(13) _
        & "JOIN edw.AccountDim AD ON  ADF.AccountDimRowId = AD.AccountDimRowId" & Chr(13) _
        & "JOIN edw.CustomerDim BOR ON ADF.BorrowerDimRowId = BOR.CustomerDimRowId" & Chr(13) _
        & "WHERE ADF.AccountId = '" & AccountId & "'" & Chr(13) _

    TrxnQuery = "SELECT" & Chr(13) _
        & "[ProcessDate] = CONVERT(VARCHAR, ProcessDate, 101)" & Chr(13) _
        & ", [EffectiveDate] = CONVERT(VARCHAR, EffectiveDate, 101)" & Chr(13) _
        & ", [Source] = ReasonCode" & Chr(13) _
        & ", [Code] = Code" & Chr(13) _
        & ", LoanBalance = BalanceAmt" & Chr(13) _
        & ", TotalAmount = TransactionAmt" & Chr(13) _
        & ", PrincipalAmt = PrincipalAmt" & Chr(13) _
        & ", InterestAmt = InterestAmt" & Chr(13) _
        & ", OverPmtAmt" & Chr(13) _
        & ", LateFeesAmt" & Chr(13) _
        & ", NsfFeesAmt" & Chr(13) _
        & ", MiscFeesAmt" & Chr(13) _
        & ", GainLossAmt" & Chr(13) _
        & "FROM [SERVER].[edw].[ALPHADELTAFOXTROT] ATF" & Chr(13) _
        & "WHERE ATF.AccountId = '" & AccountId & "'" & Chr(13) _
        & "AND TransactionCode in ('PRNPAY','PRNPAY-R','INTPAY','INTPAY-R','PAYMENT','PAYMENT-R','PAYFEE','PAYFEE-R','PAYFEE','PAYFEE-R','PAYFEE','PAYFEE-R','PRNCRADJ','PRNDRADJ','WAIVEFEE','INTCRADJ','INTDRADJ','ASSESFEE','WAIVEFEE','ASSESFEE','WAIVEFEE','ASSESFEE','PAYOFF')" & Chr(13) _

    Workbooks.Add
    ActiveSheet.Name = AccountId
    
    'Paste AccountData
    With ActiveSheet.ListObjects.Add(SourceType:=0, Source:=Array(Array(ConnectionString)), Destination:=Range("$A$1")).QueryTable
        .CommandText = AccountDataQuery
        .BackgroundQuery = True
        .ListObject.DisplayName = "AccountData" & Int((100 * Rnd) + 1)
        .Refresh BackgroundQuery:=False
    End With
    
    'Paste TrxnQuery
    With ActiveSheet.ListObjects.Add(SourceType:=0, Source:=Array(Array(ConnectionString)), Destination:=Range("$A$4")).QueryTable
        .CommandText = TrxnQuery
        .BackgroundQuery = True
        .ListObject.DisplayName = "TrxnHistory" & Int((100 * Rnd) + 1)
        .Refresh BackgroundQuery:=False
    End With

    'Formatting
    With ActiveWindow
        If .FreezePanes Then .FreezePanes = False
            .SplitColumn = 6
            .SplitRow = 4
            .FreezePanes = True
            Columns.AutoFit
            Range("F:N").NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
    End With
    
End Sub


