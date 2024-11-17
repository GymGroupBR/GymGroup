<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Dim sCnn
Dim adoCnn
Dim sSQL
Dim rsCadastro
Dim iRegra

SET adoCnn = Server.CreateObject("ADODB.Connection")
SET rsCadastro = Server.CreateObject("ADODB.RecordSet")
'--sCnn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("../../db/dbGG.mdb")--'
sCnn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & "D:\web\localuser\gymgroup\banco\dbGG.mdb"
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
	<title>GYM GROUP :: Cadastro</title>
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
<a href="javascript:history.back()">Clique Aqui, Reveja seu Cadastro</a></div>
<%ELSEIF Senha = Empty OR Nome = Empty OR Nascimento = Empty OR Email = Empty OR Cidade = Empty OR Cep = Empty OR Bairro = Empty THEN%>
<div class="aviso" align="center">Verifique os campos vermelhos!<br>
  Eles s&atilde;o de preenchimento obrigat&oacute;rio! <br>
<br>
<a href="javascript:history.back()">Clique Aqui, Reveja seu Cadastro</a></div>
<%ELSE%>
<%sSQL = "SELECT * FROM Cadastro WHERE cadastroEmail='" & Email & "'"%>
<%rsCadastro.Open sSQL, adoCnn, 1, 1%>
<%IF rsCadastro.RecordCount = 1 THEN%>
<div align="center" class="aviso">Apenas é permitido um cadastro por e-Mail.<br>
Este e-Mail já consta em nosso sistema.<br>
<br>
<a href="../index.asp">Clique Aqui, Realize o Login</a>
<p align="center"><img src="../imgs/linha.png" width="800" height="15"></p>
Esqueceu sua Senha?<br>
<br>
<a href="alterarSenha.asp">Clique Aqui, Redefina sua Senha</a></div>
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

IF Tipo = "INDIVIDUAL" THEN
	sSQL = "SELECT * FROM Individual"
   	rsCadastro.Open sSQL, adoCnn, 1, 3
   	rsCadastro.AddNew
   	rsCadastro("individualEmail") = Email
   	rsCadastro("individualAltura") = CDbl(0)
   	rsCadastro("individualPeso") = CDbl(0)
   	rsCadastro.Update
   	rsCadastro.Close
ELSEIF Tipo = "PROFISSIONAL" THEN
	sSQL = "SELECT * FROM Profissional"
   	rsCadastro.Open sSQL, adoCnn, 1, 3
   	rsCadastro.AddNew
   	rsCadastro("profissionalEmail") = Email
   	rsCadastro("profissionalCref") = "NAO"
    rsCadastro("profissionalCrefNo") = "-"
   	rsCadastro("profissionalEspecialidades") = "-"
   	rsCadastro("profissionalCondominios") = 0
   	rsCadastro("profissionalResidencias") = 0
   	rsCadastro("profissionalAcademias") = 0
   	rsCadastro("profissionalEmpresas") = 0
   	rsCadastro("profissionalEscolas") = 0
   	rsCadastro("profissionalHospitais") = 0
   	rsCadastro("profissionalParques") = 0
    rsCadastro.Update
   	rsCadastro.Close
END IF
SET rsCadastro = Nothing
adoCnn.Close
SET adoCnn = Nothing
%>
<p align="center">
<font class="rotulo"><%Response.Write(Nome)%></font>,<br>
<font class="selo">Seja Bem-Vindo a GYM GROUP!<br>
<br>
Seu cadastro já está disponível em nosso banco de dados.<br>
Confira seus dados abaixo, para eventuais altera&ccedil;&otilde;es, basta entrar com seu usuário (que é seu e-Mail) e sua senha.<br>
<br>
Equipe GYM GROUP</font><br>
<br>
<font class="rotulo"><b>Usuário e Senha</b></font><br>
<font class="selo">Tipo:&nbsp;<b><%Response.Write(Tipo)%></b></font><br>
<font class="selo">e-Mail:&nbsp;<b><%Response.Write(Email)%></b></font><br>
<font class="selo">Senha:&nbsp;<b><%Response.Write(Senha)%></b></font><br>
<br>
<font class="rotulo"><b>Informações Pessoais</b></font><br>
<font class="selo">Nome:&nbsp;<b><%Response.Write(Nome)%></b></font><br>
<font class="selo">Data Nascimento:&nbsp;<b><%Response.Write(Nascimento)%></b></font><br>
<font class="selo">Possui Deficiência:&nbsp;<b><%Response.Write(Deficiencia)%></b></font><br>
<font class="selo">Telefone Celular:&nbsp;<b><%Response.Write(Celular)%></b></font><br>
<font class="selo">Endereço:&nbsp;<b><%Response.Write(Endereco)%></b></font><br>
<font class="selo">Nº:&nbsp;<b><%Response.Write(No)%></b></font><br>
<font class="selo">Complemento:&nbsp;<b><%Response.Write(Complemento)%></b></font><br>
<font class="selo">Bairro:&nbsp;<b><%Response.Write(Bairro)%></b></font><br>
<font class="selo">Estado:&nbsp;<b><%Response.Write(Uf)%></b></font><br>
<font class="selo">Cidade:&nbsp;<b><%Response.Write(Cidade)%></b></font><br>
<font class="selo">CEP:&nbsp;<b><%Response.Write(Cep)%></b></font><br>
<br>
<a href="../index.asp">Realizar Login</a>
	</p>
<%END IF%>
<%END IF%>
<p align="center"><img src="../imgs/linha.png" width="800" height="15"></p>
<div class="aviso" align="center">GYM GROUP <%Response.Write("2024" & "-" & Year(Now))%> © Todos os Direitos Reservados</div>
</body>
</html>
