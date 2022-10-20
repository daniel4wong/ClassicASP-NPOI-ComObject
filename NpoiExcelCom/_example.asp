<%@ Language=VBScript %>
<h1>Example</h1>
<%
    Function writeHeaders_XLSX(filename)
        '.xlsx
        Response.ContentType = "aapplication/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        Response.Charset = "utf-8"
        Response.AddHeader "Content-Disposition", "attachment; filename="+filename
    End Function

    Function readFile(finename)
        Dim allText : allText = ""
        Set fso = Server.CreateObject("Scripting.FileSystemObject") 
        Set fs = fso.OpenTextFile(Server.MapPath(finename), 1, true) 
        Do Until fs.AtEndOfStream 
            allText = allText & fs.ReadLine
        Loop 
        fs.close: Set fs = nothing

        readFile = allText
    End Function
    
    On Error Resume Next

    Set excelHelper = Server.CreateObject("NpoiExcelCom.ExcelHelper")
  
    If Err.Number <> 0 and InStr(UCase(Err.Description), "800401F3") > 1 Then
        Response.Write("Cannot load NpoiExcelCom (800401F3)")
        Response.Write(htmlText)
        Err.Clear
    ElseIf Err.Number <> 0 and InStr(UCase(Err.Description), "80004027") > 1 Then
        Response.Write("Cannot load NpoiExcelCom (80004027)")
        Response.Write(htmlText)
        Err.Clear
    ElseIf Err.Number <> 0 Then
        Response.Write("Cannot load NpoiExcelCom (Unexpected Error)")
        Response.Write(htmlText)
        Err.Clear
    Else
        Dim filename : filename = "Excel_example" & ".xlsx"
        Dim htmlText : htmlText = readFile("_exampleTable.html")
        'Response.Write(htmlText)

        excelHelper.TestFile "_example.xlsx"
        Response.Write("<br />" & filename)
        Response.Write("<br />" & excelHelper.Health & " - " & excelHelper.FilePath)

        Response.Clear
        excelHelper.ConvertHtmlToExcel "Data", htmlText
        writeHeaders_XLSX filename
        Response.BinaryWrite(excelHelper.ExcelBinary)
    End If
%>