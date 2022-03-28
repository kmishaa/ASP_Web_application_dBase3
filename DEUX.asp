<%@ Language=VBScript %>
<HTML>
<HEAD>
  <META HTTP-EQUIV="Content-Type" CONTENT="text/html;CHARSET=windows-1251">
  <TITLE>ЛР3.2 Красовицкий Михаил</TITLE>
 </HEAD>
 <%
Set Conn=Server.CreateObject("ADODB.Connection")


Conn.Open "System DNS Foxpro for LR3 DB"


Set RS=Server.CreateObject("ADODB.Recordset")
RS.Open "SELECT TOP 20 product_id, english_name, quantity_in_unit, unit_cost, units_in_stock FROM [products.dbf] order by product_id", Conn, 3, 3
 %>

<h1 align=center style="color:green">Таблица: PRODUCTS</h1>
<table align=center border=2 BORDERCOLOR = black BGCOLOR = maroon>
<tr>
<th BGCOLOR = green>№</th><th BGCOLOR = green>Название</th><th BGCOLOR = green>Вес</th><th BGCOLOR = green>Стоимость</th><th BGCOLOR = green>Количество</th>
</tr>

<%Do While Not RS.EOF %>
<tr>
<td BGCOLOR = white align=center><%=RS(0).value%></td>
<td BGCOLOR = white align=left><%=RS(1).value%></td>
<td BGCOLOR = white align=left><%=RS(2).value%></td>
<td BGCOLOR = white align=right><%=RS(3).value%></td>
<td BGCOLOR = white align=right><%=RS(4).value%></td>
</tr>

<%RS.MoveNext
Loop
RS.Close
Conn.Close
%>

