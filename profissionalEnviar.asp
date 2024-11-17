<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Dim sCnn
Dim adoCnn
Dim sSQL
Dim rsCadastro
Dim Email
Dim sCdstr
Dim Nome, Deficiencia, Nascimento, Celular, Endereco, No, Complemento, Bairro, Estado, Cidade, Cep
Dim Cref, CrefNo
Dim Especialidades, Condominios, Academias, Residencias, Empresas, Escolas, Hospitais, Parques

Email = Request.QueryString("sEmail")

SET adoCnn = Server.CreateObject("ADODB.Connection")
SET rsCadastro = Server.CreateObject("ADODB.RecordSet")
sCnn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("../../db/dbGG.mdb")
'--sCnn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & "D:\web\localuser\gymgroup\banco\dbGG.mdb--'
adoCnn.Open(sCnn)

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

Cref = Request.Form("cboCREF")
CrefNo = Request.Form("txtCREF")

Especialidades = Ucase(Request.Form("txtEspecial")) 
Condominios = Request.Form("chkCondominios")
Academias = Request.Form("chkAcademias")
Residencias = Request.Form("chkResidencias")
Empresas = Request.Form("chkEmpresas")
Escolas = Request.Form("chkEscolas")
Hospitais = Request.Form("chkHospitais")
Parques = Request.Form("chkParques")
   
sSQL = "SELECT * FROM Cadastro WHERE cadastroEmail='" & Email & "'"
rsCadastro.Open sSQL, adoCnn, 1, 3
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

	sSQL = "SELECT * FROM Profissional WHERE profissionalEmail='" & Email & "'"
	rsCadastro.Open sSQL, adoCnn, 1, 3
	rsCadastro("profissionalCref") = Cref
   	IF Cref = "SIM" THEN
   		rsCadastro("profissionalCrefNo") = CrefNo
   	ELSE
   		rsCadastro("profissionalCrefNo") = "-"
   	END IF
   	IF Especialidades = EMPTY THEN
   		rsCadastro("profissionalEspecialidades") = "-"
   	ELSE
   		rsCadastro("profissionalEspecialidades") = Especialidades
   	END IF
   	IF Condominios = "on" THEN
   		rsCadastro("profissionalCondominios") = 1
   	ELSE
   		rsCadastro("profissionalCondominios") = 0
   	END IF
   	IF Academias = "on" THEN
   		rsCadastro("profissionalAcademias") = 1
   	ELSE
   		rsCadastro("profissionalAcademias") = 0
   	END IF
   	IF Residencias = "on" THEN
   		rsCadastro("profissionalResidencias") = 1
   	ELSE
   		rsCadastro("profissionalResidencias") = 0
   	END IF
   	IF Empresas = "on" THEN
   		rsCadastro("profissionalEmpresas") = 1
   	ELSE
   		rsCadastro("profissionalEmpresas") = 0
   	END IF
   	IF Escolas = "on" THEN
   		rsCadastro("profissionalEscolas") = 1
   	ELSE
   		rsCadastro("profissionalEscolas") = 0
   	END IF
   	IF Hospitais = "on" THEN
   		rsCadastro("profissionalHospitais") = 1
   	ELSE
   		rsCadastro("profissionalHospitais") = 0
   	END IF
   	IF Parques = "on" THEN
   		rsCadastro("profissionalParques") = 1
   	ELSE
   		rsCadastro("profissionalParques") = 0
   	END IF
   	rsCadastro.Update
   	rsCadastro.Close
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
	<title>GYM GROUP :: Cadastro :: Detalhes</title>
</head>

<body>
<table width="800" align="center">
  <tr>
	<td width="230"><img src="../imgs/logo.png" width="118" height="90" alt=""/></td>
	  <td width="450" align="center"><h1>CADASTRO</h1></td>
	  <td width="70" align="right"><a href="javascript:history.back()">VOLTAR</a></td>
	  <td width="50" align="right"><a href="../index.asp">SAIR</a></td>
  </tr>
</table>
<p align="center"><img src="../imgs/linha.png" width="800" height="15"></p>
<p align="center">
	<font class="rotulo"><%Response.Write(Nome)%></font><br>
	<font class="aviso">Seu cadastro foi atualizado com êxito.</font><br>
<br>
<font class="rotulo">Informações Pessoais</font><br>
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
<font class="selo">Possui CREF?:&nbsp;<b><%Response.Write(Cref)%></b></font><br>
<font class="selo">Nº CREF:&nbsp;<b><%Response.Write(CrefNo)%></b></font><br>
<img src="../imgs/linha.png" width="800" height="15"><br>
<font class="rotulo">Especialidades</font><br>
<font class="selo"><b><%Response.Write(Especialidades)%></b></font><br>
<img src="../imgs/linha.png" width="800" height="15"><br>
<font class="rotulo">Áreas de Atuação</font><br>
<%IF Condominios = "on" THEN%>
	<font class="selo"><b>CONDOMÍNIOS / GRUPOS</b></font><br>
<%END IF%>
<%IF Academias = "on" THEN%>
	<font class="selo"><b>ACADEMIAS / CENTROS ESPORTIVOS</b></font><br>
<%END IF%>
<%IF Residencias = "on" THEN%>
	<font class="selo"><b>RESIDÊNCIAS / PARTICULAR</b></font><br>
<%END IF%>
<%IF Empresas = "on" THEN%>
	<font class="selo"><b>EMPRESAS / INSTITUIÇÕES</b></font><br>
<%END IF%>
<%IF Escolas = "on" THEN%>
	<font class="selo"><b>ESCOLAS / COLÉGIOS / UNIVERSIDADES</b></font><br>
<%END IF%>
<%IF Hospitais = "on" THEN%>
	<font class="selo"><b>HOSPITAIS / CLÍNICAS MÉDICAS</b></font><br>
<%END IF%>
<%IF Parques = "on" THEN%>
	<font class="selo"><b>PARQUES / AR LIVRE / LOCAIS PÚBLICOS</b></font><br>
<%END IF%>
</p>
<p align="center"><img src="../imgs/linha.png" width="800" height="15"></p>
<div class="aviso" align="center">GYM GROUP <%Response.Write("2024" & "-" & Year(Now))%> © Todos os Direitos Reservados</div>
<%
SET rsCadastro = NOTHING
adoCnn.Close
SET adoCnn = NOTHING
%>
</body>
</html>
