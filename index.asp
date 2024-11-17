<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!doctype html>

<html lang="pt-br">
<head>
<meta charset="UTF-8">
	<meta http-equiv="X-UA-Compatible" content="IE=edge">
	<meta name="viewport" content="width=device-width, initial-scale=1.0">
	<meta name="description" content="Seu Grupo Saudável">
	<meta name="keywords" content="Grupo, Saúde, Exercício, Corrida, Musculação, Fitness, Ginástica, Caminhada, Crossfit, Físico, Academia">
	<link rel="stylesheet" href="css/gymgroup.css">
	<title>::: GYM GROUP :::</title>
</head>

<body>
	<form name="frmEntrar" method="post" action="pgns/entrar.asp" >
		<table width="350" align="center">
			<tr>
				<td align="center"><img src="imgs/logo.png" width="235" height="179" alt=""/></td>
			</tr>
		</table>
		<p align="center"><img src="imgs/linha.png" width="350" height="15"></p>
	  <table width="350" align="center">
   			<tr>
			  <td width="20%" height="50" valign="middle"><label for="email">e-Mail:</label></td>
                <td valign="middle"><input name="txtEmail" type="email" id="txtEmail" size="30"></td>
		  </tr>
			<tr>
			  <td width="20%" height="50" valign="middle"><label for="password">Senha:</label></td>
              <td valign="middle"><input type="password" name="txtSenha" id="txtSenha" size="20"></td>
   		  </tr>
		  <tr>
			  <td colspan="2"><hr></td>
   		  </tr>
	  </table>
		
		<table width="350" height="50" align="center">
			<tr>
				<td align="center" valign="middle" width="50%"><label><input type="submit" value="Entrar"></label></td>
				<td align="center" valign="middle" width="50%"><label><input type="reset" value="Limpar"></label></td>
		  </tr>
	  </table>
	</form>
	<p align="center"><img src="imgs/linha.png" width="350" height="15"></p>
	<table width="350" align="center">
	  		<tr>
				<td align="center"><label>Não Tem Cadastro?</label><br>
		        <a href="pgns/cadastro.asp">Clique Aqui</a></td>
				<td align="center"><label>Esqueceu a Senha?</label><br>
		        <a href="pgns/alterarSenha.asp">Clique Aqui</a></td>
      		</tr>
	</table>
<p align="center"><img src="imgs/linha.png" width="350" height="15"></p>
<div class="aviso" align="center">GYM GROUP <%Response.Write("2024" & "-" & Year(Now))%> © Todos os Direitos Reservados</div>
</body>
</html>
