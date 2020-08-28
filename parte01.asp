<%
planAtiva = application("PlanAtiva")
planSenha = application("planSenha")
nomeDaPlanilha = application("nomeDaPlanilha")

if not planAtiva <> "" and not planSenha <> "" then
	response.redirect("index.asp")
end if


'ABRE O ARQUIVO DO EXCEL e MYSQL
set conexaoDataBase = Server.CreateObject("ADODB.Connection")
conexaoDataBase.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
           "Data Source="&nomeDaPlanilha&";" & _
           "Extended Properties=""Excel 8.0;HDR=YES;"""
Set mysql = Server.CreateObject("ADODB.Connection")
mysql.open "Driver={MySQL ODBC 3.51 Driver}; SERVER=localhost; UID=root; PWD=luisrevolution2007;"

'SELECIONA A PLANILHA
verifiq = "Select * from [Plan1$]"
Set x = conexaoDataBase.execute(verifiq)

'CRIA NOME PARA BANCO DE DADOS
totalColunas = x.Fields.count

'CRIA UM NOVO BANCO DE DADOS
mysql.execute "CREATE DATABASE "&planAtiva&";"
mysql.execute "USE "&planAtiva&";"
colunasParaCadastrar = null
For i = 0 To (totalColunas-1)
	if colunasParaCadastrar <> "" then
		colunasParaCadastrar = colunasParaCadastrar&","
	end if
	if i = 0 then
		colunasParaCadastrar = colunasParaCadastrar &"col"&i&" varchar(30)"
	else
		colunasParaCadastrar = colunasParaCadastrar &"col"&i&" int(3)"
	end if
Next

'CRIA AS TABELAS DO BANCO
mysql.execute "CREATE TABLE dados (id bigint NOT NULL auto_increment Primary key,"&colunasParaCadastrar&");"
mysql.execute "CREATE TABLE campos (id int NOT NULL auto_increment Primary key,nome varchar(30),nomeCad varchar(30),coluna varchar(8),valorMax int(3),opcao int(1),especificacao int(2));"
mysql.execute "CREATE TABLE info (id int NOT NULL auto_increment Primary key,colunas int(10),linhas int(10),data date,hora time,usuario varchar(30),nome varchar(30),cidade varchar(200),dias varchar(200));"
mysql.execute "CREATE TABLE comparacoes (id bigint NOT NULL auto_increment Primary key,item int(3),pergunta int(3));"
mysql.execute "CREATE TABLE somatorias (id bigint NOT NULL auto_increment Primary key,item int(3),pergunta int(3),total int(5),resultado double, ligacaoID int(20));"
mysql.execute "CREATE TABLE quantidade (id bigint NOT NULL auto_increment Primary key,campo int(20),valor int(5),quantidade int(5),nome varchar(200),extra int(2));"
mysql.execute "CREATE TABLE tabelas (id bigint NOT NULL auto_increment Primary key,ligacaoID int(20),tabela text,linhas int(3));"
mysql.execute "CREATE TABLE perguntas (id bigint NOT NULL auto_increment Primary key,ligacaoID int(20),pergunta varchar(220));"

'Cadastra os nomes das colunas
tipodaOpcao = 1
For i = 0 To (totalColunas-1)
	nomeCampo = x.Fields(i).name
	nomeCad=nomeCampo
	if tipodaOpcao = 1 and (InStrRev(nomeCampo,"(x)") > 0) then
		tipodaOpcao = 2
		nomeCampo = replace(nomeCampo,"(x)","")
		nomeCad = replace(nomeCad,"(x)","")
	end if
	nomeCadcase = ucase(nomeCad)
	if (InStrRev(nomeCadcase,"CODIGO") > 0) then
		especificacao = 1
	else
		if (InStrRev(nomeCadcase,"SEXO") > 0) then
			especificacao = 2
		else
			if(InStrRev(nomeCadcase,"IDADE") > 0) then
				especificacao = 3
			else
				if(InStrRev(nomeCadcase,"ESCOLA") > 0) then
					especificacao = 4
				else
					if(InStrRev(nomeCadcase,"RENDA") > 0) then
						especificacao = 5
					else
						if(InStrRev(nomeCadcase,"ADM") > 0) then
							especificacao = 6
						else
							if(InStrRev(nomeCadcase,"AP") > 0) then
								especificacao = 7
							else
								especificacao = 0
							end if
						end if
					end if
				end if
			end if
		end if
	end if
	mysql.execute  "INSERT INTO campos (nome,coluna,nomeCad,opcao,especificacao) VALUES ('"&nomeCampo&"','col"&i&"','"&nomeCampo&"',"&tipodaOpcao&","&especificacao&")"
Next



'CADASTRA DADOS DA PLANILHA
Do While Not (x.EOF)
	On error Resume Next
	dadosCadColunas = null
	valoresCadColunas = null
	For i = 0 To (totalColunas-1)
		valorDaColuna = x.Fields(i).Value
   		if dadosCadColunas <> "" then
			dadosCadColunas = dadosCadColunas&","
			valoresCadColunas = valoresCadColunas&","
		end if
		dadosCadColunas = dadosCadColunas &"col"&i&""
		if VarType(valorDaColuna) = 8 then valorDaColuna = "'"&valorDaColuna&"'"
		if len(valorDaColuna&"luis") <= 4 then
			valorDaColuna = 1
		end if
		valoresCadColunas = valoresCadColunas &""&valorDaColuna&""
	Next
	if not InStrRev(valoresCadColunas,",,") > 0 then
		
		mysql.execute  "INSERT INTO dados ("&dadosCadColunas&") VALUES ("&valoresCadColunas&")"
		If Err.Number <> 0 then
			application("error_a") = "Erro ao tentar cadastrar dados na tabela!<br>Linha de codigo: <b>INSERT INTO dados ("&dadosCadColunas&") VALUES ("&valoresCadColunas&")</b>"
			application("error_b") = "<font color=#FF0000>OLHE NA LINHA: "&(totaldeLinhas+2)&"</font><BR><BR>"
			application("error_b") = application("error_b")&"Erro: Numero "&Err.Number&"<br>"&Err.Description
			response.redirect("erro.asp")
		end if
	else
		application("error_a") = "ATENÇÃO.. TEM CELULAS EM BRANCO!!"
		application("error_b") = "OLHE NA LINHA: "&(totaldeLinhas+2)
		response.redirect("erro.asp")
	end if
	totaldeLinhas = totaldeLinhas+1
	x.movenext
	If Err.Number <> 0 then
		application("error_a") = "Erro na leitura do arquivo do Excel"
		application("error_b") = "<font color=#FF0000>OLHE NA LINHA: "&(totaldeLinhas+2)&"</font><BR><BR>"
		application("error_b") = application("error_b")&"Erro: Numero "&Err.Number&"<br>"&Err.Description
		response.redirect("erro.asp")
	end if
Loop
application("totaldeLinhas") = totaldeLinhas

'CADASTRA INFORMAÇÕES DA PLANILHA
dataX = Year(date) & "-" & Month(date) & "-" & Day(date)
mysql.execute  "INSERT INTO info (colunas,linhas,data,hora,usuario,nome) VALUES ("&totalColunas&","&totaldeLinhas&",'"&dataX&"','"&time&"','Luís','"&planAtiva&"')"


'FECHA BANCOS
conexaoDataBase.close
mysql.close

'Redireciona para proxima etapa
response.redirect("parte02.asp")
%>


