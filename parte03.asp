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
totalDeCampos = int(protecao(request.form("total")))
cidade = protecao(request.form("cidade"))
dias = protecao(request.form("dias"))
application("pontosPesquisas")=protecao(request.form("erro"))
mysql.execute "UPDATE info SET cidade='"&cidade&"',dias='"&dias&"' WHERE id =1"


For i = 1 To (totalDeCampos)
	nome = protecao(request.form("nome"&i))
	id = int(protecao(request.form("id"&i)))
	valorMax = int(protecao(request.form("valorMax"&i)))
	valorMaxReg = int(protecao(request.form("valorMaxReg"&i)))
	opcao= int(protecao(request.form("opcao"&i)))
	mysql.execute "UPDATE campos SET nome='"&nome&"',valorMax="&valorMax&",opcao="&opcao&" WHERE id ="&id
	if valorMax < valorMaxReg then
		sql = "SELECT coluna FROM campos WHERE id = "&id
		Set coluna = mysql.execute(sql)
		coluna = coluna("coluna")

		sql = "SELECT "&coluna&",id FROM dados WHERE "&coluna&" > "&valorMax
		Set linhas = mysql.execute(sql)
		Do While not linhas.eof
			idLinhas = linhas("id")
			mysql.execute "UPDATE dados SET "&coluna&"="&valorMax&" WHERE id ="&idLinhas
			linhas.movenext
		loop
	end if
Next

mysql.close
response.redirect("parte04_inicio.asp")%>