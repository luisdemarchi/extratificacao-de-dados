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


sql = "SELECT id,nome,coluna,valorMax,opcao,especificacao FROM campos ORDER BY id ASC"
Set coluna = mysql.execute(sql)

%>
<style type="text/css">
<!--
.style1 {color: #FFFFFF}
-->
</style>

<title>Renomeando colunas</title><form id="form" name="form" method="post" action="parte03.asp">
<table width="410" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <th colspan="2" bgcolor="#0000FF"><span class="style1">Dados da pesquisa </span><span class="style1"> </span></th>
    </tr>
  <tr>
    <td width="50%">Nome da Cidade: </td>
    <td width="50%"><input name="cidade" type="text" id="cidade" size="30" maxlength="200" /></td>
  </tr>
  <tr>
    <td>Dias da pesquisa: </td>
    <td><input name="dias" type="text" id="dias" value=" e  de <%=MonthName(month(date))%> de <%=year(date)%>" size="30" maxlength="200" /></td>
  </tr>
  <tr>
    <td bgcolor="#F0FFF4">Margem T&eacute;cnica de Erro:</td>
    <td bgcolor="#F0FFF4"><input name="erro" type="text" id="erro" value="4,0" size="10" maxlength="200" />
      pnts percentuais</td>
  </tr>
  <tr>
    <td bgcolor="#F0FFF4">Total de registros: </td>
    <td bgcolor="#F0FFF4"><%=application("totaldeLinhas")%> Pessoas </td>
  </tr>
</table>
<br />
<table width="410" border="0" align="center" cellpadding="0" cellspacing="0">
<tr>
  <th width="227" bgcolor="#0000FF"><span class="style1">Nome </span></th>
		<th width="93" bgcolor="#0000FF"><span class="style1">Op</span></th>
		<th width="90" bgcolor="#0000FF"><span class="style1">Valor M&aacute;x </span></th>
  </tr><%
Do While Not (coluna.EOF)
	cont = cont+1
	especificacao = coluna("especificacao")
	id = coluna("id")
	nome = coluna("nome")
	col = coluna("coluna")
	opcao = coluna("opcao")
	sql = "SELECT MAX("&col&") as valorMaximo FROM dados"
	Set verValorMax = mysql.execute(sql)
	valorMaximo = verValorMax("valorMaximo")
	variavelLinha = "<tr height='24'>"
	Select Case especificacao
		case 1
			nome = "Região"
		case 2
			if valorMaximo > 2 then
				
				valorMaximo = 2
			end if
			nome = "Sexo"
		case 3
			if valorMaximo > 4 then
				
				valorMaximo = 4
			end if
			nome = "Faixa Etária "
		case 4
			if valorMaximo > 3 then
				
				valorMaximo = 3
			end if
			nome = "Escolaridade"
		case 5
			if valorMaximo > 4 then
				valorMaximo = 4
			end if
			nome = "Renda"
	End Select 
	%><%=variavelLinha%>
	  <td><input name="nome<%=cont%>" type="text" id="nome<%=cont%>" value="<%=nome%>" size="30" maxlength="30">
      <input name="id<%=cont%>" type="hidden" id="id<%=cont%>" value="<%=id%>" /></td>
	  <td><select name="opcao<%=cont%>" id="opcao<%=cont%>">
        <option value="1" <%if opcao = 1 then%>selected="selected"<%end if%>>Item</option>
		<option value="2" <%if opcao = 2 then%>selected="selected"<%end if%>>Perg</option>
      </select></td>
	  <td <%if len(valorMaximo) > 3 then%>bgcolor="#999999"<%end if%>><%if not len(valorMaximo) > 3 then%><input name="valorMax<%=cont%>" type="text" id="valorMax"value="<%=valorMaximo%>" size="5" maxlength="3"><input name="valorMaxReg<%=cont%>" type="hidden" id="valorMaxReg<%=cont%>" value="<%=valorMaximo%>" />
      <%else%><input name="valorMax<%=cont%>" type="hidden" id="valorMax<%=cont%>" value="0" /><input name="valorMaxReg<%=cont%>" type="hidden" id="valorMaxReg<%=cont%>" value="0" /><%end if%>
      </td>
	</tr><%
	coluna.movenext
loop
%>  <td colspan="4">&nbsp;</td>
		</tr><tr>
		  <td colspan="3" align="center"><input type="submit" name="Submit" value="Continuar" /><input name="total" type="hidden" id="total" value="<%=cont%>" /></td>
    </tr>
</table>
</form>
<%mysql.close%>