<%Response.CacheControl = "no-cache"%><%
planAtiva = application("PlanAtiva")
planSenha = application("planSenha")
if not planAtiva <> "" and not planSenha <> "" then
	response.redirect("index.asp")
end if

'ABRE O BANCO MYSQL
set conexaoDataBase = Server.CreateObject("ADODB.Connection")
Set mysql = Server.CreateObject("ADODB.Connection")
mysql.open "Driver={MySQL ODBC 3.51 Driver}; SERVER=localhost; UID=root; PWD=luisrevolution2007; DATABASE="&planAtiva&";"

function protecao(texto)
	if texto <> "" then
		protecao = replace(texto,chr(39), "&#39;")
		protecao = replace(protecao,chr(34), "&quot;") 
		protecao = Replace(protecao, "<", "&lt;")
		protecao = replace(protecao,vbcrlf,"<br>")
	else
		protecao = null
	end if
end function

mysql.execute "DELETE FROM comparacoes"
sql = "SELECT id FROM campos WHERE opcao = 2"
Set camposPerg = mysql.execute(sql)
Do while not camposPerg.eof
	perguntas = camposPerg("id")
	sql = "SELECT id FROM campos WHERE opcao = 1"
	Set camposItem = mysql.execute(sql)
	Do while not camposItem.eof
		itemEscolhido = camposItem("id")
		mysql.execute "INSERT INTO comparacoes (item,pergunta) VALUES ("&itemEscolhido&","&perguntas&")"

		camposItem.movenext
	loop
	camposPerg.movenext
loop


mysql.close
response.redirect("parte05_inicio.asp")%>
