 'OpenTicketReformatting v1.-
 'Formats Open Ticket to the correct upload version for USTAPP
 'Requirements: Create a folder in your desktop called "OpenTickets"
 '10-18-2017

Sub OpenTicketReformatting()

    Rows("2:2").Insert Shift:=xlDown, copyorigin:=xlFormatFromLeftOrAbove
    Rows("3:3").Insert Shift:=xlDown, copyorigin:=xlFormatFromLeftOrAbove

    Columns("F:G").MergeCells = False
    Columns("G:G").Delete

    Columns("D:D").Insert Shift:=xlLeft
    Columns("D:D").Insert Shift:=xlLeft
        Range("D4").Value = "GP"
        Range("E4").Value = "BR"

    Columns("D:E").NumberFormat = "@"

    'splits column C
    Dim GPBRrange As Range

        Set GPBRrange = Range("C5", Range("C4").End(xlDown))

            For Each cell In GPBRrange

                ' GP value
                cell.Offset(0, 1).Value = Left(cell.Value, 2)
                ' BR value
                cell.Offset(0, 2).Value = Right(cell.Value, 2)

            Next

            ' IsEmpty VIN verification
            For Each cell In GPBRrange

                If IsEmpty(cell.Offset(0, 19)) = True Then
                    MsgBox ("There are empty VIN cells.")
                    Exit Sub
                End If

            Next

    Columns("C:C").Delete
    
    ' Header comparison here
    ' To verify that future workbooks come through with the correct headers

    Dim headerArray As Variant
    Dim headerRange As Range

        headerArray = Array("State", "City", "GP", "BR", "Ticket", "Truck Type", "Cust#", _
                        "Customer Name", "Start Date", "Days Open", "Estimated Return Date", _
                        "Daily Rate", "Weekly Rate", "Monthly Rate", "Unlimited Miles", _
                        "Cents per Mile", "Free Per Day", "Free Per Week", "Free per Month", _
                        "Unit1", "Unit1 VIN ", "Unit1 Make", "Unit1 Model", "First Name", _
                        "Last Name", "CMS ID")
                        
        Set headerRange = Range("A4", Range("A4").End(xlToRight))

    Dim i As Integer

        i = 0

    'header comparison main
    For Each cell In headerRange

        If cell.Value <> headerArray(i) Then
        
            MsgBox "Workbook headers do not match import requirements.", vbCritical + vbOKCancel, "Header Mismatch!"
            Exit Sub
            
        End If
        
        i = i + 1

    Next

    ' formatting workbook and worksheet name
    Dim FileName As String
    Dim FileDate As String
    
        FileDate = Replace(Right(Range("A1").Value, 10), "/", "-")
        FileName = "OpenTicket_" & FileDate
    
    Dim WsDate As String
        WsDate = Left(FileDate, 5)
        Sheets(1).Name = "OT " & WsDate
  
    ' searching for Desktop and saving
    Dim SaveAddress As String
    Dim WSHShell As Object
    
        Set WSHShell = CreateObject("WScript.Shell")
        SaveAddress = WSHShell.SpecialFolders("Desktop") & "\OpenTickets"
        ThisWorkbook.SaveAs FileName:=SaveAddress & "\" & FileName

End Sub
