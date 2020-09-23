<div align="center">

## Paging recordsets from a database \(acess\) like Google\.\.\.


</div>

### Description

This code pages a database recordsets in a very cool way, like the good seaches on the web. The ideia I designed was to show all pages, but not all of them (imagine in a case of 1000 pages). So, it shows only ten pages placed in the point you are only, and a link for the last and first ones... It helps who does not have the patience to see all pages, and also for the ones who does.
 
### More Info
 
The most interesting of this code is that if you are, for example, in a database with 1000 pages, and you are in the middle of it, it will link you the first page "1", and also link the last "1000"... Also, will show your current location with 5 pages back and 5 pages foward... Isn`t great?

Be carefull to use the Database correctly... and I suppose that at this far you already know how to display your database recordsets... If not, look for examples in this web site.

You also can see this code in action in my personal page : http://www.bmbks.org/personal/index.asp?Idioma=1 (portuguese) - But there is an option to english and japanese (but the huge database is in portugues, sorry. The other languages contain the same idea but the database is still small)


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Bruno Martins Braga](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/bruno-martins-braga.md)
**Level**          |Intermediate
**User Rating**    |4.6 (23 globes from 5 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[Databases](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases__4-5.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/bruno-martins-braga-paging-recordsets-from-a-database-acess-like-google__4-8571/archive/master.zip)

### API Declarations

The original code was not mine, but I change it completely from the inicial idea.


### Source Code

```
<%
'Remember to set correctly your database location, and the contents as well. There are other
'codes in this site to explain how to show a recordset. I supposed this step we can foward...
Dim objConn
Set objConn = Server.CreateObject("ADODB.Connection")
objConn.Open "DBQ=" & Server.MapPath("database.mdb") & ";Driver={Microsoft Access Driver (*.mdb)}","username","password"
Dim TotalPages, NumPerPage, CurrentPage, Number
Dim strQuery, ObjRs, Count
'Defines the inicial Value as "5"
If request("NumPerPage") <> 0 then NumPerPage = Int(request("NumPerPage")) else NumPerPage = 5
strQuery = "SELECT * FROM TableName order by data desc"
Set ObjRs = Server.CreateObject("ADODB.Recordset")
If Request.QueryString("page") = "" then
 CurrentPage = 1 'We're on the first page
Else
 CurrentPage = CInt(Request.QueryString("page"))
End If
objRS.Open strQuery, objConn, 1, 1 'Opened as Read-Only
Number = objRS.RecordCount
If Not objRS.EOF Then
 	objRS.MoveFirst
	objRS.PageSize = NumPerPage
 TotalPages = objRS.PageCount
 objRS.AbsolutePage = CurrentPage
End If
Dim ScriptName
ScriptName = request.servervariables("ScriptName")
%>
<%
While Not objRS.EOF and Count < objRS.PageSize
count = count + 1
%>
Put here your RecordSet Display (usually a table with the ASP code together)
<%
objRS.MoveNext
Wend
objRs.close
Set objRs = Nothing
%>
<%
'Print the recent Data
response.write "Showing page <b>" & CurrentPage & "</b> of <b>" & TotalPages & "</b>: Total of <b>" & Number & "</b> written posts..."
%>
<%
'Creating the paging numbers
Dim ini, fim, a
'Display PREV page link, if appropriate
If Not CurrentPage = 1 Then
	Response.Write "&lt;a href='" & ScriptName & "?NumPerPage=" & NumPerPage & "&page=" & CurrentPage - 1 & "'><font size=1 face=Verdana><b>..</b></font>&lt;/a>&nbsp;&nbsp;"
if CurrentPage > 5 and TotalPages > 10 then
 Response.write("&lt;a href=" & ScriptName & "?NumPerPage=" & NumPerPage & "&page=1><font size=1 face=Verdana><b>1</b></font>&lt;/a>" & "<font size=1 face=Verdana><b>&nbsp;...&nbsp;</b> </font>" )
end if
if TotalPages > 10 then
	if CurrentPage > 5 then
		if TotalPages > (CurrentPage + 5) then
			ini = (CurrentPage - 4)
			fim = (CurrentPage + 5)
		else
			ini = (TotalPages - 9)
			fim = TotalPages
		end if
	else
		ini = 1
		fim = 10
	end if
else
ini=1
fim = TotalPages
end if
For a = ini to fim
 If a = Cint(request("page")) then
 Response.write( "<font face=Verdana color=#FF0000 size=3><b>" & a & "</b></font>&nbsp;&nbsp;")
 Else
 Response.write("&lt;a href=" & ScriptName & "?NumPerPage=" & NumPerPage & "&page=" & a &"><font size=1 face=Verdana><b>" & a & "</b></font>&lt;/a>" & "&nbsp;&nbsp;" )
 End if
Next
Else
	if TotalPages = 1 then
			Response.write ""
		Else
			Response.Write "<font face=Verdana color=#FF0000 size=3><b>1</b></font>&nbsp;&nbsp;"
	End if
	if TotalPages > 10 then
	fim = 10
	else
	fim = TotalPages
	end if
	For a = 2 to fim
 If a = Cint(request("page")) then
 Response.write( "<font face=Verdana color=#FF0000 size=3><b>" & a & "</b></font>&nbsp;&nbsp;")
 Else
 Response.write("&lt;a href=" & ScriptName & "?NumPerPage=" & NumPerPage & "&page=" & a &"><font size=1 face=Verdana><b>" & a & "</b></font>&lt;/a>" & "&nbsp;&nbsp;" )
 End if
Next
End If
if CurrentPage < TotalPages - 5 and TotalPages > 10 then
 Response.write("<font size=1 face=Verdana><b>...&nbsp;</b></font>&lt;a href=" & ScriptName & "?NumPerPage=" & NumPerPage & "&page=" & TotalPages &"><font size=1 face=Verdana><b>" & TotalPages & "</b></font>&lt;/a>" & "&nbsp;&nbsp;" )
end if
'Display NEXT page link, if appropriate
If Not CurrentPage = TotalPages Then
	Response.Write "&lt;a href='" & ScriptName & "?NumPerPage=" & NumPerPage & "&page=" & CurrentPage + 1 & "'><font size=1 face=Verdana><b>..</b></font>&lt;/a>"
Else
	Response.Write ""
End If
%>
```

