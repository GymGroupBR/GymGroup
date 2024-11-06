<%@LANGUAGE="VBSCRIPT"%>
<%
Dim sCnn
Dim adoCnn
Dim sSQL
Dim rsCadastro
Dim Regra
Dim MSG
Dim Email
Dim Senha01
Dim Senha02

SET adoCnn = Server.CreateObject("ADODB.Connection")
SET rsCadastro = Server.CreateObject("ADODB.RecordSet")
sCnn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("../../db/dbGG.mdb") 
adoCnn.Open(sCnn)

Email = Request.Form("txtEmail")
Senha01 = Request.Form("txtSenha01")
Senha02 = Request.Form("txtSenha02")

sSQL = "SELECT * FROM Cadastro WHERE cadastroEmail='" & Email & "'"
rsCadastro.Open sSQL, adoCnn, 1, 3
IF rsCadastro.RecordCount = 0 THEN
	Regra = 0
   	MSG = "Este e-Mail não consta em nosso sistema!"
ELSE
	IF Senha01 = EMPTY OR Senha02 = EMPTY THEN
   		Regra = 1
   		MSG = "Você deve inserir as Senhas!"
   	ELSEIF Senha01 = Senha02 THEN
   		rsCadastro("cadastroSenha") = Senha01
   		rsCadastro.Update
		Regra = 3
   		MSG = "Senha Alterada com Sucesso!"
	ELSE
   		Regra = 2
   		MSG = "As Senhas devem ser iguais!"
	END IF
END IF
rsCadastro.Close
SET rsCadastro = NOTHING
adoCnn.Close
SET adoCnn = Nothing
%>
<!doctype html>

<html lang="pt-br">

<head>
	<meta charset="UTF-8">
	<meta http-equiv="X-UA-Compatible" content="IE=edge">
	<meta name="viewport" content="width=device-width, initial-scale=1.0">
	<meta name="description" content="Seu Grupo Saudável">
	<meta name="keywords" content="Grupo, Saúde, Exercício, Corrida, Musculação, Fitness, Ginástica, Caminhada, Crossfit, Físico, Academia">
	<link rel="stylesheet" href="../css/gymgroup.css">
	<title>GYM GROUP ::: Cadastro</title>
</head>

<body>
<table width="800" align="center">
	<tr>
		<td width="250"><img src="../imgs/logo.png" width="118" height="90" alt=""/></td>
	  <td align="center"><h1>REDEFINIR SENHA</h1></td>
  </tr>
</table>
<p align="center"><img src="../imgs/linha.png" width="800" height="15"></p>
<table width="800" align="center" border="0" cellspacing="0" cellpadding="0">
	<%IF Regra = 0 THEN%>
	<tr>
		<td colspan="4" align="center"><p class="aviso"><%Response.Write(MSG)%><br>
			Deseja Realizar seu Cadastro?</p></td>
	</tr>
	<tr align="center">
		<td width="30%"></td>
		<td width="20%"><a class="aviso" href="cadastro.asp" onMouseOver="MM_displayStatusMsg('..:: GYM GROUP :: Redefinir Senha ::..');return document.MM_returnValue">Sim</a></td>
		<td width="20%"><a class="aviso" href="javascript:history.back()" onMouseOver="MM_displayStatusMsg('..:: GYM GROUP :: Redefinir Senha ::..');return document.MM_returnValue">Não</a></td>
		<td width="30%"></td>
	</tr>
	<%ELSEIF Regra = 1 OR Regra = 02 THEN%>
	<tr>
		<td align="center"><p class="aviso"><%Response.Write(MSG)%></p></td>
	</tr>
	<tr align="center">
		<td><a class="aviso" href="javascript:history.back()" onMouseOver="MM_displayStatusMsg('..:: GYM GROUP :: Redefinir Senha ::..');return document.MM_returnValue">Voltar</a></td>
	</tr>
	<%ELSEIF Regra = 3 THEN%>
	<tr>
		<td align="center"><p class="aviso"><%Response.Write(MSG)%></p></td>
	</tr>
	<tr align="center">
		<td><a class="aviso" href="../index.asp" onMouseOver="MM_displayStatusMsg('..:: GYM GROUP :: Redefinir Senha ::..');return document.MM_returnValue">Realizar Login</a></td>
	</tr>
	<%END IF%>
</table>
<p align="center"><img src="../imgs/linha.png" width="800" height="15"></p>
<div class="aviso" align="center">Todos os Direitos Reservados <%Response.Write("2024" & "-" & Year(Now))%> © GYM GROUP</div>
</form>
</body>
</html>
