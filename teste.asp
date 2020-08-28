<%Response.CacheControl = "no-cache"%><%planAtiva = application("PlanAtiva")
planSenha = application("planSenha")
set conexaoDataBase = Server.CreateObject("ADODB.Connection")
Set mysql = Server.CreateObject("ADODB.Connection")
mysql.open "Driver={MySQL ODBC 3.51 Driver}; SERVER=localhost; UID=root; PWD=luisrevolution2007; DATABASE="&planAtiva&";"



sql = "SELECT id FROM campos WHERE opcao = 2"
Set camposPerg = mysql.execute(sql)
Do while not camposPerg.eof
	perguntas = camposPerg("id")
	sql = "SELECT nome,id FROM campos WHERE opcao = 1 and not valorMax = 0"
	Set camposItem = mysql.execute(sql)
	%><ol type="a"><%
	Do while not camposItem.eof
	nomeCampoCol = camposItem("nome")
	%><li type="a">Por <%=nomeCampoCol%> (em %)</li><%
		itemEscolhido = camposItem("id")
		sql = "SELECT tabela FROM tabelas WHERE ligacaoID IN (SELECT id FROM comparacoes WHERE item = "&itemEscolhido&" and pergunta = "&perguntas&")"
		Set tabelasBase = mysql.execute(sql)
		Do while not tabelasBase.eof
			%><%=tabelasBase("tabela")%>
			<%
			tabelasBase.movenext
		loop
		
		camposItem.movenext
	loop
	%></ol><%
'	act.WriteLine("<br style=page-break-before:always>")
	camposPerg.movenext
loop%>