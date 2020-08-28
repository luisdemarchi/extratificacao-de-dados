<%Response.CacheControl = "no-cache"%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Coloque sua senha</title>
</head>

<body>
     <form method="POST" action="verificar.asp">
          ><br>
     
  <table width="300" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td>User:</td>
      <td><input name="user" type="text" id="user" /></td>
    </tr>
    <tr>
      <td>Password:</td>
      <td><input name="senha" type="password" id="senha" /></td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>Planilha:</td>
      <td><input name="arquivo" type="text" id="arquivo" /></td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td colspan="2" align="center"><input type=submit value="Entrar"></td>
    </tr>
  </table>
</form></body>
</html>
