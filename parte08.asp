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

sql = "SELECT cidade,dias FROM info ORDER BY id ASC"
Set informacoes = mysql.execute(sql)
cidade = informacoes("cidade")
dias= informacoes("dias")
arquivo_doc = application("nomeCidade")&".doc"

set fso = createobject("scripting.filesystemobject")
Set act = fso.CreateTextFile(server.mappath("arquivosWord/"&arquivo_doc), true)

%><!--#include file="./word/corpo.asp" -->

<!--#include file="./word/pag01.asp" -->
<!--#include file="./word/pag02.asp" -->
<!--#include file="./word/pag03.asp" -->
<!--#include file="./word/pag04.asp" -->
<!--#include file="./word/pag05.asp" -->
<!--#include file="./word/pag06.asp" --><%

act.WriteLine("</body>")
act.WriteLine("</html>")


response.redirect("final.asp")
%>