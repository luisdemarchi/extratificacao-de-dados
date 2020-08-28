<%Response.CacheControl = "no-cache"%><%
Function rndit()
	dia = day(date)
	if dia < 10 then
	dia = "0"&dia
end if
mes = month(date)
if mes < 10 then
	mes = "0"&mes
end if
ano = year(date)
hora = hour(time)
if hora < 10 then
	hora = "0"&hora
end if
minuto = minute(time)
if minuto < 10 then
	minuto = "0"&minuto
end if
segundo = second(time)
if segundo < 10 then
	segundo = "0"&segundo
end if
rndit = (""&ano&""&mes&""&dia&""&hora&""&minuto&""&segundo&"")
rndit = "plan"&rndit
End Function 
if request.Form("user") = "luis" and request.Form("senha") = "luisdemarchi123" then

	application("planSenha") = "ativa"
	application("PlanAtiva")=rndit()
	application("nomeDaPlanilha") = request.Form("arquivo")
	local = Server.MapPath("./planilhas/velhas/")&"\"&application("nomeDaPlanilha")
	local2 = Server.MapPath(".") & "\planilhas\"&application("PlanAtiva")&".xls"
	
	Set fso = Server.CreateObject("Scripting.FileSystemObject")
    fso.GetFile(local).Copy local2
	application("nomeDaPlanilha") = local2

	response.redirect("parte01_inicio.asp")

else
	response.redirect("index.asp?erro=senha")
end if
%>