'from my code review profile
'https://codereview.stackexchange.com/questions/183757/web-scraping-data-from-us-postal-zip-codes

Option Explicit

'PreprocessingData Variables
    Public StringRange As Range
    Public cell As Range


Sub PreprocessingData()
    
        Range("A1").Value = "rawZipCodes"
        Range("B1").Value = "City"
        Range("C1").Value = "State"
        Range("D1").Value = "ZipCode"
        Range("E1").Value = "County"
        Range("F1").Value = "Population"
        Range("G1").Value = "Median Home Value"
    
    With Sheets("rawZipCodes")
        .Range("A2", .Range("A2").End(xlDown)).Copy Destination:=Worksheets("ZipCodes").Range("A2")
    End With
    
    Set StringRange = Range("A2", Range("A2").End(xlDown))
        
        For Each cell In StringRange
        
            cell.Offset(0, 1).Value = returnCity(cell)
            cell.Offset(0, 2).Value = returnState(cell)
            cell.Offset(0, 3).NumberFormat = "@" 'inserting zipcode as text
            cell.Offset(0, 3).Value = returnZipCode(cell)
                
        Next cell
        
    Range("A2").EntireColumn.Delete
   
    ' inserting table
    Dim ZipTable As Range
    Set ZipTable = Range("A1").CurrentRegion
        ListObjects.Add(xlSrcRange, ZipTable, , xlYes).Name = "ZipTable"
    
    'resizing columns
    Sheets("ZipCodes").UsedRange.Columns.AutoFit
    
End Sub

Sub ZipCodeScrape()

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Const BASE_URL = "https://www.unitedstateszipcodes.org/"
    Dim HTMLDoc As MSHTML.HTMLDocument
    
    Dim ZipCodeRange As Range
    Set ZipCodeRange = Range("C2", Range("C2").End(xlDown))
    
    For Each cell In ZipCodeRange
        Set HTMLDoc = getDocument(BASE_URL & cell.Value)
        If Not HTMLDoc Is Nothing Then
            cell.Offset(0, 1).Value = getTDByTH(HTMLDoc, "County:")
            cell.Offset(0, 2).Value = getTDByTH(HTMLDoc, "Population")
            cell.Offset(0, 3).Value = getTDByTH(HTMLDoc, "Median Home Value")
        End If
    Next cell
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

End Sub

Private Function getDocument(URL As String) As MSHTML.HTMLDocument

    Dim HTMLDoc As MSHTML.HTMLDocument
    
    With New MSXML2.XMLHTTP60
        .Open "GET", URL, False
        .send
        
            If .readyState = 4 And .Status = 200 Then
                Set HTMLDoc = New MSHTML.HTMLDocument
                    HTMLDoc.body.innerHTML = .responseText
                Set getDocument = HTMLDoc
            Else
                Debug.Print _
                "URL Not Responding: "; URL, vbNewLine; _
                "Ready State: "; .readyState, vbNewLine; _
                "HTTP request status: "; .Status
            End If
    End With

End Function

Private Function getTDByTH(HTMLDoc As MSHTML.HTMLDocument, Heading As String) As String
    Dim post As Object
    For Each post In HTMLDoc.getElementsByTagName("TH")
        If post.innerText = Heading Then
            getTDByTH = post.ParentNode.getElementsByTagName("TD")(0).innerText
            Exit For
        End If
    Next post
End Function

Private Function returnZipCode(ByVal variableString As String)

    Dim SpaceIndex As Byte
        SpaceIndex = Application.Find(" ", variableString)
    
    Dim ZipCodeStartIndex As Byte
        ZipCodeStartIndex = SpaceIndex + 1
    
    returnZipCode = Mid(variableString, ZipCodeStartIndex, 5)

End Function

Private Function returnCity(ByVal variableString As String)

    Dim SpaceIndex As Byte
    SpaceIndex = Application.Find(" ", variableString)
    
    Dim commaIndex As Byte
    commaIndex = Application.Find(",", variableString)
    
    Dim cityStartIndex As Byte
    cityStartIndex = SpaceIndex + 6
    
    returnCity = Mid(variableString, cityStartIndex, commaIndex - cityStartIndex)

End Function
