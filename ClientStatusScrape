' Date: 10/10/2017
' Desc: Returns the client status and Annual Report Status from www.sosnc.gov/
' How to obtain a WebID:
    ' 1. Navigate to http://www.sosnc.gov/search/index/corp
    ' 2. Search for the name of the company.
    ' 3. Once in the company's profile, i.e. http://www.sosnc.gov/Search/profcorp/5480528
        ' 4. The WebId are the last 7 digits of the url
    
Sub ClientStatusScrape()

Dim WebId As Range
    Set WebId = Range("C6", Sheets("NC Corporate Annual Report").Range("C6").End(xlDown)).Offset(0, -1)

' Creates ie instance
Dim IE As Object
    Set IE = CreateObject("InternetExplorer.Application")
    IE.Visible = False
    
        ' Scrape loop
        For Each cell In WebId
        
            If Not IsEmpty(cell) Then
        
                Dim url As String
                    url = "https://www.sosnc.gov/Search/profcorp/" & cell.Value
                    IE.Navigate (url)
                    
                    Do While IE.Busy
                        DoEvents
                    Loop
                    
                Dim CurrentURL As String
                    CurrentURL = IE.LocationURL
                
                If CurrentURL <> "https://www.sosnc.gov/Home/Error" Then
                  
                    ' Data scrape main
                    Dim Status As String
                    Dim AnnualReportStatus As String
                    Dim HTMLdoc As HTMLDocument
                    
                        Set HTMLdoc = IE.Document
                        
                        Status = HTMLdoc.getElementsByTagName("td").Item(5).innerText
                        AnnualReportStatus = HTMLdoc.getElementsByTagName("td").Item(7).innerText
                        
                    cell.Offset(0, 2) = Status
                    cell.Offset(0, 3) = AnnualReportStatus
                    
                    ' Per request, inserts company url
                    cell.Hyperlinks.Add anchor:=cell.Offset(0, 4), Address:=CurrentURL, TextToDisplay:="Visit URL"
                      
                ' Error handling by url name
                ElseIf CurrentURL = "https://www.sosnc.gov/Home/Error" Then
                
                    cell.Offset(0, 2) = "Verify WebID"
                    cell.Offset(0, 3) = "Verify WebID"
                
                End If
            
            ElseIf IsEmpty(cell) Then

                cell.Offset(0, 2) = "Missing WebId"
                cell.Offset(0, 3) = "Missing WebId"
            
            End If
            
        Next

MsgBox ("Web Scrape from the Department of the Secretary of State is now completed.")

End Sub
