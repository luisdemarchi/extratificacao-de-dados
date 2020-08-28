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


sql = "SELECT id,nome,valorMax,nomeCad,especificacao,opcao FROM campos WHERE not valorMax = 0 and opcao = 2 ORDER BY id ASC"
Set coluna = mysql.execute(sql)

colunasTodas=null
%>
<style type="text/css">
<!--
.style1 {color: #FFFFFF}
-->
</style>

<form id="form" name="form" method="post" action="parte06_02.asp">
<table width="434" border="0" align="center" cellpadding="0" cellspacing="0">
<%cont=0
contCol=0
Do While Not (coluna.EOF)
	contCol = contCol + 1
	id = coluna("id")
	nome = coluna("nome")
	valorMax = coluna("valorMax")
	nomeCad = coluna("nomeCad")
	especificacao = coluna("especificacao")
	opcao = coluna("opcao")
	%><tr><th height="22" colspan="2" bgcolor="#0000FF"><%if opcao = 2 then%><%totaldePerguntas=totaldePerguntas+1%><span class="style1">Pergunta <%=nome%>:
  <input name="idPerg<%=totaldePerguntas%>" type="hidden" id="idPerg<%=totaldePerguntas%>" value="<%=id%>" /> 
      <input name="pergunta<%=totaldePerguntas%>" type="text" id="pergunta<%=totaldePerguntas%>" style="color:#0000CC;background-color:#CCCCCC;font-size:10px;height:18;width:300" value="ÍNDICE ESTIMULADO DE VOTOS PARA <%=ucase(nome)%>">
</span>
    <%else%><span class="style1">Item <%=nome%></span><%end if%></th></tr><%
	For i = 1 To (valorMax)
		cont = cont+1
		Select Case especificacao
		case 6
			Select Case (valorMax-i)
				case 0
					valorDoCampo = "Não Informou"
				case 1
					valorDoCampo = "Não Aprova"
				case 2
					valorDoCampo = "Aprova"
			End Select
		case 7
			Select Case (valorMax-i)
				case 0
					valorDoCampo = "Indiferente"
				case 1
					valorDoCampo = "Não"
				case 2
					valorDoCampo = "Sim"
			End Select
		case else
			Select Case (valorMax-i)
				case 0
					valorDoCampo = "Não Informou"
				case 1
					valorDoCampo = "Branco/Nulo"
				case 2
					valorDoCampo = "Indecisos"
				case else
					valorDoCampo = ""
			End Select
		End Select
		if cor = "#F0FFF4" then
			cor = ""
		else
			cor = "#F0FFF4"
		end if
		%><tr bgcolor="<%=cor%>"><td width="120" height="25"><%=(nomeCad&" "&i)%></td><td><input name="id<%=cont%>" type="hidden" id="id<%=cont%>" value="<%=id%>" />
  <input name="item<%=cont%>" type="hidden" id="item<%=cont%>" value="<%=i%>" />
  <input name="nome<%=cont%>" type="text" id="nome<%=cont%>" value="<%=valorDoCampo%>" size="30">
  <%if especificacao = 1 then%><select name="extra<%=cont%>" id="extra<%=cont%>">
  <option value="1">Cidade</option>
  <option value="2">Rural</option>
  </select><%else%><input name="extra<%=cont%>" id="extra<%=cont%>" type="hidden" value="0" /><%end if%>
  </td>
</tr><%
	Next
	%><tr><td colspan="2">&nbsp;</td>
</tr><%
	coluna.movenext
Loop%><tr><tr><td colspan="2" align="center"><input type="submit" name="Submit" value="Cadastrar" />
    <input name="total" type="hidden" id="total" value="<%=cont%>" />
    <input name="perguntas" type="hidden" id="perguntas" value="<%=totaldePerguntas%>" /></td>
</tr>
</table>
</form>
<%mysql.close%>