<%@LANGUAGE="VBSCRIPT"%>
<%
Dim adoCnn
Dim adoEntrar
Dim rsEntrar
Dim rsCadastro
Dim sCnn
Dim sSQL
Dim Email
Dim Senha
Dim Tipo
Dim Nome
Dim Cadastros
Dim MSG
Dim Regra

Email = Request.Form("txtEmail")
Senha = Request.Form("txtSenha")
   
Set adoCnn = Server.CreateObject("ADODB.Connection")
Set rsCadastro = Server.CreateObject("ADODB.RecordSet")
sCnn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("../../db/dbGG.mdb")
adoCnn.Open(sCnn)

IF Email = EMPTY AND Senha = EMPTY THEN
	Regra = 0
	MSG = "Você precisa inserir seu e-Mail e Senha!"
ELSEIF NOT ISEMPTY(Email) AND Senha = EMPTY THEN
	Regra = 0
	MSG = "Você precisa inserir sua Senha!"
ELSEIF Email = EMPTY AND NOT ISEMPTY(Senha) THEN
	Regra = 0
	MSG = "Você precisa inserir seu e-Mail!"
ELSE
	sSQL = "SELECT * FROM Cadastro WHERE cadastroEmail='" & Email & "'"
	rsCadastro.Open sSQL, adoCnn, 1, 1
   	IF rsCadastro.RecordCount = 0 THEN
   		Regra = 1
   		MSG = "e-Mail Não Cadastrado!"
   	ELSE
   		IF rsCadastro("cadastroSenha") = Senha THEN
   			Regra = 3
   			Tipo = rsCadastro("cadastroTipo")
			Nome = rscadastro("cadastroNome")
   		ELSE
   			Regra = 2
			MSG = "Senha Inválida!"
   		END IF
   	END IF
	rscadastro.Close
END IF
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
	<title>GYM GROUP ::: Pesquisar</title>
</head>

<body>
<table width="800" align="center">
	<tr>
		<td align="center"><img src="../imgs/logo.png" width="235" height="179" alt=""></td>
	</tr>
</table>
<p align="center"><img src="../imgs/linha.png" width="800" height="15"></p>
<%IF Regra = 0 OR Regra = 1THEN%>
<table width="800" align="center">
	<tr>
		<td align="center"><p><%Response.Write(MSG)%></p><br>
			<a href="javascript:history.back()" onMouseOver="MM_displayStatusMsg('..:: GYM GROUP ::..');return document.MM_returnValue">Voltar</a>
		</td>
	</tr>
</table>
<%ELSEIF Regra = 2 THEN%>
<table width="800" align="center">
	<tr>
		<td align="center"><p><%Response.Write(MSG)%></p><br>
			<a href="javascript:history.back()" onMouseOver="MM_displayStatusMsg('..:: GYM GROUP ::..');return document.MM_returnValue">Voltar</a>
		</td>
	</tr>
</table>
<%ELSEIF Regra = 3 THEN%>
<table width="800" align="center" border="0" cellspacing="0" cellpadding="0">
  <tr height="20">
    <td width="90%" bgcolor="#4260AC"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#FFCC00"><b>&nbsp;Usuário: <font color="#FFFFFF"><%=Response.Write(Nome)%> (<%=Response.Write(Tipo)%>)</font></b></font></td>
	<td width="10%" bgcolor="#4260AC" align="right"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b><a href="../index.asp" onMouseOver="MM_displayStatusMsg('..:: GYM GROUP :: Sair ::..');return document.MM_returnValue">SAIR</a>&nbsp;</b></font></td>
  </tr>
</table>
<br>
<form name="frmPesquisar" method="Post" action="pesquisar.asp?Nome=<%=Nome%>&Senha=<%=Senha%>&Tipo=<%=Tipo%>&Email=<%=Email%>">
<table width="800" align="center" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td class="subttl" align="center">REALIZAR PESQUISA</td>
  </tr>
</table>
<table width="800" align="center" border="0" cellspacing="0" cellpadding="0">
	<tr>
	  <td width="50"><p class="rotulo">Estado</p>
		<td><select name="cboUf" id="cboUf" class="inputWithe">
  <option selected></option>
  <option>AC</option>
  <option>AL</option>
  <option>AM</option>
  <option>AP</option>
  <option>BA</option>
  <option>CE</option>
  <option>DF</option>
  <option>ES</option>
  <option>GO</option>
  <option>MA</option>
  <option>MG</option>
  <option>MS</option>
  <option>MT</option>
  <option>PA</option>
  <option>PB</option>
  <option>PE</option>
  <option>PI</option>
  <option>PR</option>
  <option>RJ</option>
  <option>RN</option>
  <option>RO</option>
  <option>RR</option>
  <option>RS</option>
  <option>SC</option>
  <option>SE</option>
  <option>SP</option>
  <option>TO</option>
</select></td>
	</tr>
	<tr>
		<td><p class="rotulo">Cidade</p></td>
		<td><input name="txtCidade" type="text" id="txtCidade" size="70" maxlength="100" class="inputWithe"></td>
	</tr>
	<tr>
		<td><p class="rotulo">Bairro</p></td>
		<td><input name="txtBairro" type="text" id="txtBairro"  size="70" maxlength="100" class="inputWithe"></td>
	</tr>
</table>
<br>
<table width="800" align="center" border="0" cellspacing="0" cellpadding="0">
	<tr align="center">
		<td width="25%"></td>
		<td width="25%"><input name="cmdPesquisar" type="submit" id="cmdPesquisar" value="Pesquisar" class="inputBotao"></td>
		<td width="25%"><input name="cmdLimpar" type="reset" id="cmdLimpar" value="  Limpar  " class="inputBotao"></td>
		<td width="25%"></td>
	</tr>
</table>
</form>
<%END IF%>
<p align="center"><img src="../imgs/linha.png" width="800" height="15"></p>
<div class="aviso" align="center">Todos os Direitos Reservados <%Response.Write("2024" & "-" & Year(Now))%> © GYM GROUP</div>
</body>
</html>
