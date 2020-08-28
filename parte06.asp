<%
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

For i = 1 To (totalDeCampos)
	id = protecao(request.form("id"&i))
	nome = protecao(request.form("nome"&i))
	extra = protecao(request.form("extra"&i))
	itemEscolhido = protecao(request.form("item"&i))
	
	sql = "SELECT coluna FROM campos WHERE id = "&id
	Set dados = mysql.execute(sql)
	Coluna = dados("coluna")
	sql = "SELECT count(id) AS valor FROM dados WHERE "&Coluna&" = "&itemEscolhido
	Set dados = mysql.execute(sql)
	quantidade = int(CDbl(dados("valor")))
	mysql.execute "INSERT INTO quantidade (campo,valor,quantidade,nome,extra) VALUES ("&id&","&itemEscolhido&","&quantidade&",'"&nome&"',"&extra&")"
Next

mysql.close
response.redirect("parte05_inicio2.asp")%>
