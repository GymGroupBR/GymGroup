<%@LANGUAGE="VBSCRIPT"%>

<!doctype html>

<html lang="pt-br">

<head>
	<meta charset="UTF-8">
	<meta http-equiv="X-UA-Compatible" content="IE=edge">
	<meta name="viewport" content="width=device-width, initial-scale=1.0">
	<meta name="description" content="Seu Grupo Saudável">
	<meta name="keywords" content="Grupo, Saúde, Exercício, Corrida, Musculação, Fitness, Ginástica, Caminhada, Crossfit, Físico, Academia">
	<link rel="stylesheet" href="../css/gymgroup.css">
	<title>GYM GROUP ::: Redefinir Senha</title>
</head>

<body>
<table width="800" align="center">
	<tr>
		<td width="250"><img src="../imgs/logo.png" width="118" height="90" alt=""/></td>
	  <td align="center"><h1>REDEFINIR SENHA</h1></td>
  </tr>
</table>
<p align="center"><img src="../imgs/linha.png" width="800" height="15"></p>
<form name="frmCadastro" method="Post" action="alterarSenhaEnviar.asp">
<table width="800" align="center" border="0" cellspacing="0" cellpadding="0">
<tr>
    <td width="100"><p class="rotulo">e-Mail (Usu&aacute;rio)</p></td>
    <td width="700"><input class="inputRed" name="txtEmail" type="text" id="txtEmail" size="70" maxlength="70"></td>
  </tr>
  <tr>
    <td><p class="rotulo">Senha</p></td>
    <td><input class="inputRed" name="txtSenha01" type="password" id="txtSenha01" size="15" maxlength="8">
      <span class="selo"> 8 d&iacute;gitos alfa-num&eacute;ricos</span></td>
  </tr>
	<tr>
    <td><p class="rotulo">Repita Senha</p></td>
    <td><input class="inputRed" name="txtSenha02" type="password" id="txtSenha02" size="15" maxlength="8">
      <span class="selo"> 8 d&iacute;gitos alfa-num&eacute;ricos</span></td>
  </tr>
</table>
<p align="center"><img src="../imgs/linha.png" width="800" height="15"></p>
<div align="center">
  <input name="cmdEnviar" type="submit" id="cmdEnviar" value="Enviar">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<input name="cmdLimpar" type="reset" id="cmdLimpar" value="Limpar">
</div>
</form>
<p align="center"><img src="../imgs/linha.png" width="800" height="15"></p>
<div class="aviso" align="center">Todos os Direitos Reservados <%Response.Write("2024" & "-" & Year(Now))%> © GYM GROUP</div>
</body>
</html>
