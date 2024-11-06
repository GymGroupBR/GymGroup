<%@LANGUAGE="VBSCRIPT"%>
<%
Dim sCnn
Dim adoCnn
Dim sSQL
Dim rsCadastro

SET adoCnn = Server.CreateObject("ADODB.Connection")
SET rsCadastro = Server.CreateObject("ADODB.RecordSet")
sCnn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("../../db/dbGG.mdb") 
adoCnn.Open(sCnn)

Tipo = Ucase(Request.Form("cboTipo"))
Email = Request.Form("txtEmail")
Senha = Request.Form("txtSenha")
Nome = Ucase(Request.Form("txtNome"))
Nascimento = Request.Form("txtNascimento")
Deficiencia = Ucase(Request.Form("cboDeficiencia"))
Celular = Request.Form("txtCelular")
Endereco = Ucase(Request.Form("txtEndereco"))
No = Ucase(Request.Form("txtNo"))
IF Request.Form("txtComplemento") = EMPTY THEN
	Complemento = "-"
ELSE
   Complemento = Ucase(Request.Form("txtComplemento"))
END IF
Bairro = Ucase(Request.Form("txtBairro"))
Uf = Ucase(Request.Form("cboUf"))
Cidade = Ucase(Request.Form("txtCidade"))
Cep = Request.Form("txtCep")
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
	<title>GYM GROUP ::: Redefinir Senha</title>
</head>

<body>
<table width="800" align="center">
	<tr>
		<td width="250"><img src="../imgs/logo.png" width="118" height="90" alt=""/></td>
	  <td align="center"><h1>CADASTRO</h1></td>
  </tr>
</table>
<p align="center"><img src="../imgs/linha.png" width="800" height="15"></p>
<%IF Len(Senha) < 8 THEN%>
<div class="aviso" align="center">A Senha deve conter 8 digitos!<br>
<br>
<a href="javascript:history.back()" onMouseOver="MM_displayStatusMsg('..:: GYM GROUP :: Cadastro ::..');return document.MM_returnValue">Clique Aqui, Reveja seu Cadastro</a></div>
<%ELSE
IF Senha = Empty OR Nome = Empty OR Nascimento = Empty OR Email = Empty OR Cidade = Empty OR Cep = Empty OR Bairro = Empty THEN%>
<div class="aviso" align="center">Verifique os campos vermelhos!<br>
  Eles s&atilde;o de preenchimento obrigat&oacute;rio! <br>
<br>
<a href="javascript:history.back()" onMouseOver="MM_displayStatusMsg('..:: GYM GROUP :: Cadastro ::..');return document.MM_returnValue">Clique Aqui, Reveja seu Cadastro</a></div>
<%ELSE%>
<%sSQL = "SELECT * FROM Cadastro WHERE cadastroEmail='" & Email & "'"%>
<%rsCadastro.Open sSQL, adoCnn, 1, 1%>
<%IF rsCadastro.RecordCount >= 1 THEN%>
<div align="center" class="aviso">Apenas é permitido um cadastro por e-Mail.<br>
Este e-Mail já consta em nosso sistema.<br>
<br>
<a href="javascript:history.back()" onMouseOver="MM_displayStatusMsg('..:: GYM GROUP :: Cadastro ::..');return document.MM_returnValue">Clique Aqui, Reveja seu Cadastro</a>
<p align="center"><img src="../imgs/linha.png" width="800" height="15"></p>
Esqueceu sua Senha?<br>
<br>
<a href="alterarSenha.asp" onMouseOver="MM_displayStatusMsg('..:: GYM GROUP :: Cadastro ::..');return document.MM_returnValue">Clique Aqui, Redefina sua Senha</a></div>
<%rsCadastro.Close%>
<%SET rsCadastro = NOTHING%>
<%ELSE%>
<%
SET rsCadastro = Server.CreateObject("ADODB.RecordSet")
rsCadastro.Open "Cadastro", adoCnn, 1, 3
rsCadastro.AddNew
rsCadastro("cadastroEmail") = Email
rsCadastro("cadastroSenha") = Senha
rsCadastro("cadastroTipo") = Tipo
rsCadastro("cadastroNome") = Nome
rsCadastro("cadastroNascimento") = Nascimento
rsCadastro("cadastroDeficiencia") = Deficiencia
rsCadastro("cadastroCelular") = Celular
rsCadastro("cadastroEndereco") = Endereco
rsCadastro("cadastroNo") = No
rsCadastro("cadastroComplemento") = Complemento
rsCadastro("cadastroBairro") = Bairro
rsCadastro("cadastroUf") = Uf
rsCadastro("cadastroCidade") = Cidade
rsCadastro("cadastroCep") = Cep

rsCadastro.Update
rsCadastro.Close
SET rsCadastro = Nothing
adoCnn.Close
SET adoCnn = Nothing
%>
<p align="center">
<font face="Segoe, 'Segoe UI', 'DejaVu Sans', 'Trebuchet MS', Verdana, 'sans-serif'" size="2"><b><%Response.Write(Nome)%></b>,<br>
Seja Bem-Vindo a GYM GROUP!<br>
<br>
Seu cadastro já está disponível em nosso banco de dados.<br>
Confira seus dados abaixo, para eventuais altera&ccedil;&otilde;es, basta entrar com seu usuário (que é seu e-Mail) e sua senha.<br>
<br>
Equipe GYM GROUP</font><br>
<br>
<font face="Segoe, 'Segoe UI', 'DejaVu Sans', 'Trebuchet MS', Verdana, 'sans-serif'" size="2" color="#4260AC"><b>Usuário e Senha</b></font><br>
<font face="Segoe, 'Segoe UI', 'DejaVu Sans', 'Trebuchet MS', Verdana, 'sans-serif'" size="1">Tipo:&nbsp;<b><%Response.Write(Tipo)%></b></font><br>
<font face="Segoe, 'Segoe UI', 'DejaVu Sans', 'Trebuchet MS', Verdana, 'sans-serif'" size="1">e-Mail:&nbsp;<b><%Response.Write(Email)%></b></font><br>
<font face="Segoe, 'Segoe UI', 'DejaVu Sans', 'Trebuchet MS', Verdana, 'sans-serif'" size="1">Senha:&nbsp;<b><%Response.Write(Senha)%></b></font><br>
<br>
<font face="Segoe, 'Segoe UI', 'DejaVu Sans', 'Trebuchet MS', Verdana, 'sans-serif'" size="2" color="#4260AC"><b>Informações Pessoais</b></font><br>
<font face="Segoe, 'Segoe UI', 'DejaVu Sans', 'Trebuchet MS', Verdana, 'sans-serif'" size="1">Nome:&nbsp;<b><%Response.Write(Nome)%></b></font><br>
<font face="Segoe, 'Segoe UI', 'DejaVu Sans', 'Trebuchet MS', Verdana, 'sans-serif'" size="1">Data Nascimento:&nbsp;<b><%Response.Write(Nascimento)%></b></font><br>
<font face="Segoe, 'Segoe UI', 'DejaVu Sans', 'Trebuchet MS', Verdana, 'sans-serif'" size="1">Possui Deficiência:&nbsp;<b><%Response.Write(Deficiencia)%></b></font><br>
<font face="Segoe, 'Segoe UI', 'DejaVu Sans', 'Trebuchet MS', Verdana, 'sans-serif'" size="1">Telefone Celular:&nbsp;<b><%Response.Write(Celular)%></b></font><br>
<font face="Segoe, 'Segoe UI', 'DejaVu Sans', 'Trebuchet MS', Verdana, 'sans-serif'" size="1">Endereço:&nbsp;<b><%Response.Write(Endereco)%></b></font><br>
<font face="Segoe, 'Segoe UI', 'DejaVu Sans', 'Trebuchet MS', Verdana, 'sans-serif'" size="1">Nº:&nbsp;<b><%Response.Write(No)%></b></font><br>
<font face="Segoe, 'Segoe UI', 'DejaVu Sans', 'Trebuchet MS', Verdana, 'sans-serif'" size="1">Complemento:&nbsp;<b><%Response.Write(Complemento)%></b></font><br>
<font face="Segoe, 'Segoe UI', 'DejaVu Sans', 'Trebuchet MS', Verdana, 'sans-serif'" size="1">Bairro:&nbsp;<b><%Response.Write(Bairro)%></b></font><br>
<font face="Segoe, 'Segoe UI', 'DejaVu Sans', 'Trebuchet MS', Verdana, 'sans-serif'" size="1">Estado:&nbsp;<b><%Response.Write(Uf)%></b></font><br>
<font face="Segoe, 'Segoe UI', 'DejaVu Sans', 'Trebuchet MS', Verdana, 'sans-serif'" size="1">Cidade:&nbsp;<b><%Response.Write(Cidade)%></b></font><br>
<font face="Segoe, 'Segoe UI', 'DejaVu Sans', 'Trebuchet MS', Verdana, 'sans-serif'" size="1" color="#FFFFFF">CEP:&nbsp;<b><%Response.Write(Cep)%></b></font><br>
<br>
<a href="../index.asp">Realizar Login</a>
	</p>
<%END IF%>
<%END IF%>
<%END IF%>
<p align="center"><img src="../imgs/linha.png" width="800" height="15"></p>
<div class="aviso" align="center">Todos os Direitos Reservados <%Response.Write("2024" & "-" & Year(Now))%> © GYM GROUP</div>
</body>
</html>
