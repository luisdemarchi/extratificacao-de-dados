<%
planAtiva = application("PlanAtiva")
planSenha = application("planSenha")
if not planAtiva <> "" and not planSenha <> "" then
	response.redirect("index.asp")
end if

'ABRE O BANCO MYSQL
Set mysql = Server.CreateObject("ADODB.Connection")
mysql.open "Driver={MySQL ODBC 3.51 Driver}; SERVER=localhost; UID=root; PWD=luisrevolution2007; DATABASE="&planAtiva&";"
sql = "SELECT cidade FROM info"
Set info = mysql.execute(sql)
cidade = replace(info("cidade")," ","_")
cidade = replace(cidade,".","")
cidade = replace(cidade,"ã","a")
cidade = replace(cidade,"é","e")
cidade = replace(cidade,"ç","c")
cidade = replace(cidade,",","")&"-"&day(date)&""&month(date)&""&year(date)
application("nomeCidade")=cidade

'''''''''CRIA nova planilha
    x = 0
''''''''''''''''


sql = "SELECT id,item,pergunta FROM comparacoes ORDER BY pergunta ASC, item ASC"
Set comparacoes = mysql.execute(sql)

function cabecalho(itemTotal)
	numeroInicio = InStrRev(itemTotal,",")
	numeroInicio = right(itemTotal,len(itemTotal)-numeroInicio)
	itemTotal= CDbl(itemTotal)
	numeroFinal = Fix(itemTotal)
	if int(numeroFinal) > int(numeroInicio) then
		totalDeColunas = numeroFinal-numeroInicio
		tamanhoTable2=null
		Select Case itemEspecie
		case 1
			tamanhoTable = 480
'			if totalDeColunas = 4 then
'				tamanhoTable2 = "width:79.0pt;"
'				tamanhoTable2 = "width:"&((tamanhoTable-130)/totalDeColunas)&".0pt;"
'			end if
		case 2
			tamanhoTable = 250
		case 3
			tamanhoTable = 310
		case 4
			tamanhoTable = 310
		case 5
			tamanhoTable = 500
		case else
			tamanhoTable = 400
		End Select 
		tamanhoTable2 = "width:"&((tamanhoTable-130)/totalDeColunas)&".0pt;"
		
		cabecalho = "<table class=MsoTableGrid border=1 cellspacing=0 cellpadding=0 style=""border-collapse:collapse;border:none;mso-border-alt:solid black .5pt; mso-yfti-tbllook:"&tamanhoTable&";mso-padding-alt:0cm 5.4pt 0cm 5.4pt;mso-border-insideh:.5pt solid black;mso-border-insidev:0.5pt solid black"">"
		novaCont=0
		cabecalho = cabecalho&"<tr style=""mso-yfti-irow:0;mso-yfti-firstrow:yes""><td width=288 valign=top style=""width:130.0pt;border-top:none;border-left: none;border-bottom:solid black 1.0pt;border-right:solid black 1.0pt; mso-border-bottom-alt:solid black .5pt;mso-border-right-alt:solid black .5pt; padding:0cm 5.4pt 0cm 5.4pt"">&nbsp;</td>"
		For x = numeroInicio To (numeroFinal)
			sql = "SELECT nome FROM quantidade WHERE campo = "&itemId&" and valor = "&x
			Set dados = mysql.execute(sql)
			nomeColuna = dados("nome")
			cabecalho = cabecalho&"<td valign=top style="&tamanhoTable2&"border:solid black 1.0pt; border-left:none;mso-border-left-alt:solid black .5pt;mso-border-alt:solid black .5pt;padding:0cm 5.4pt 0cm 5.4pt""><p class=MsoNormal align=center style=""text-align:center""><i><span style=""font-size:14.0pt;font-family:Arial"">"&nomeColuna&"</span></i></p></td>"
			novaCont = novaCont+1
		Next
		cabecalho = cabecalho&"</tr>"
		cabecalho = cabecalho&""&listarItens(itemTotal)
	end if

end function

function listarItens(itemTotal)
	numeroInicio = InStrRev(itemTotal,",")
	numeroInicio = right(itemTotal,len(itemTotal)-numeroInicio)
	itemTotal= CDbl(itemTotal)
	numeroFinal = Fix(itemTotal)
	
	For i = 1 To (pergTotal)
		novaCont = 0
		sql = "SELECT nome FROM quantidade WHERE campo = "&pergId&" and valor = "&i
		Set dados = mysql.execute(sql)
		nomeLinha = dados("nome")
		espacamento = (len(nomeLinha)*11)
		listarItens = listarItens&"<tr style=""mso-yfti-irow:1;mso-yfti-lastrow:yes""><td width=288 valign=top style=""width:130.0pt;border:solid black 1.0pt; border-top:none;mso-border-top-alt:solid black .5pt;mso-border-alt:solid black .5pt; padding:0cm 5.4pt 0cm 5.4pt""><p class=MsoNormal style=""text-align:justify""><span style=""font-size:16.0pt;font-family:Arial"">"&nomeLinha&"</span></p></td>"
		For x = numeroInicio To (numeroFinal)
			sql = "SELECT count(id) AS valor FROM dados WHERE "&itemColuna&" = "&x&" and "&pergColuna&" = "&i
			Set dados = mysql.execute(sql)
			valor = int(CDbl(dados("valor")))
			
		
			sql = "SELECT quantidade FROM quantidade WHERE campo = "&itemId&" and valor = "&x
			Set dados = mysql.execute(sql)
			totalDePessoas = dados("quantidade")+0.0049
			if valor >= totalDePessoas then
				inteiroNum = Fix(porcentagem)
				acharVirgula = InStrRev(porcentagem,",")
				if acharVirgula > 0 then
					virgulaNum=left(right(porcentagem,len(porcentagem)-acharVirgula),2)
				else
					virgulaNum = "0"
				end if
				porcentagem= inteiroNum&"."&virgulaNum
				mysql.execute "INSERT INTO somatorias (item,pergunta,total,resultado,ligacaoID) VALUES ("&x&","&i&","&valor&","&porcentagem&","&idLigacao&")"			

	
		
				if len(inteiroNum) < 2 then
					inteiroNum = "0"&inteiroNum
				end if
				if len(virgulaNum) < 2 then
					virgulaNum = virgulaNum&"0"
				end if
				porcentagem= inteiroNum&"."&virgulaNum
				listarItens = listarItens&"<td valign=top style="&tamanhoTable2&"border-top:none;border-left: none;border-bottom:solid black 1.0pt;border-right:solid black 1.0pt; mso-border-top-alt:solid black .5pt;mso-border-left-alt:solid black .5pt; mso-border-alt:solid black .5pt;padding:0cm 5.4pt 0cm 5.4pt""><p class=MsoNormal align=right style=""text-align:right""><span style=""font-size:14.0pt;font-family:Arial"">"&porcentagem&"</span></p></td>"
				novaCont = novaCont+1
			end if
		Next
		listarItens = listarItens&"</tr>"
	Next
	listarItens = listarItens&"</table>"
end function


mysql.execute "DELETE FROM somatorias"
mysql.execute "DELETE FROM tabelas"
Do While Not (comparacoes.EOF)
	itemEscolhido = comparacoes("item")
	pergunta = comparacoes("pergunta")
	idLigacao = comparacoes("id")
	sql = "SELECT id,coluna,valorMax,nome FROM campos WHERE id = "&pergunta
	Set dados = mysql.execute(sql)
	pergColuna = dados("coluna")
	pergTotal = dados("valorMax")
	pergNome = dados("nome")
	pergId = dados("id")
	sql = "SELECT id,coluna,valorMax,nome,especificacao FROM campos WHERE id = "&itemEscolhido
	Set dados = mysql.execute(sql)
	itemColuna = dados("coluna")
	itemNome = dados("nome")
	itemTotal = dados("valorMax")
	itemId = dados("id")
	itemEspecie = dados("especificacao")
	
	totalValor=0
	


	if not itemTotal > 4 then
		mostrarTabela = cabecalho(itemTotal&",1")
		if mostrarTabela <> "" then
			mysql.execute "INSERT INTO tabelas (ligacaoID,tabela) VALUES ("&idLigacao&",'"&mostrarTabela&"')"
		end if
	else

		quantidadeTotal=int(itemTotal/4+0.9)
		inicioDoLaco = 1
		fimDoLaco = 4
		
		For w = 0 To (quantidadeTotal-1)
			if w > 0 then
				inicioDoLaco = inicioDoLaco+4
				fimDoLaco = fimDoLaco+4
				if fimDoLaco > itemTotal then
					fimDoLaco = itemTotal
				end if
			end if
			mostrarTabela = cabecalho(fimDoLaco&","&inicioDoLaco)
			if mostrarTabela <> "" then
				mysql.execute "INSERT INTO tabelas (ligacaoID,tabela) VALUES ("&idLigacao&",'"&mostrarTabela&"')"
			end if
		next
				
	end if
	
	comparacoes.movenext
loop
mysql.close
response.redirect("parte08_inicio.asp")
%>