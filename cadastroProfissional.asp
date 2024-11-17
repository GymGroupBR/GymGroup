<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Dim adoCnn
Dim rsCadastro
Dim sCnn, sSQL

Email = Request.QueryString("sEmail")

Set adoCnn = Server.CreateObject("ADODB.Connection")
Set rsCadastro = Server.CreateObject("ADODB.RecordSet")
sCnn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("../../db/dbGG.mdb")
'--sCnn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & "D:\web\localuser\gymgroup\banco\dbGG.mdb"--'
adoCnn.Open(sCnn)

sSQL = "SELECT * FROM Cadastro WHERE cadastroEmail='" & Email & "'"
rsCadastro.Open sSQL, adoCnn, 1, 1
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
		<td width="230"><img src="../imgs/logo.png" width="118" height="90" alt=""/></td>
	  <td width="450" align="center"><h1>CADASTRO</h1></td>
	  <td width="70" align="right"><a href="javascript:history.back()">VOLTAR</a></td>
	  <td width="50" align="right"><a href="../index.asp">SAIR</a></td>
  </tr>
</table>
<p align="center"><img src="../imgs/linha.png" width="800" height="15"></p>
<form name="frmEspecial" method="Post" action="profissionalEnviar.asp?sEmail=<%=Email%>">
<table width="800" align="center" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td class="subttl">&nbsp;Informa&ccedil;&otilde;es Pessoais</td>
  </tr>
</table>
<table width="800" align="center" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="130"><p class="rotulo">Nome Completo</p></td>
    <td width="670"><input class="inputWhite" name="txtNome" type="text" id="txtNome" size="100" maxlength="100" value="<%=rsCadastro("cadastroNome")%>"></td>
  </tr>
    <tr>
    <td><p class="rotulo">Data Nascimento</p></td>
    <td><input class="inputWithe" type="text" name="txtNascimento" id="txtNascimento" size="10" maxlength="10" value="<%=rsCadastro("cadastroNascimento")%>"> </td>
  </tr>
  <tr>
    <td><p class="rotulo">Possui Defici&ecirc;ncia</p></td>
    <td><select class="inputWithe" name="cboDeficiencia" id="cboDeficiencia" value="<%=rsCadastro("cadastroDeficiencia")%>">
      		<option>Nao</option>
			<option>Auditiva - Parcial</option>
      		<option>Auditiva - Total</option>
      		<option>Fisica - Membros Inferiores</option>
      		<option>Fisica - Membros Superiores</option>
      		<option>Fisica - Motora</option>
      		<option>Multipla</option>
      		<option>Visao - Parcial</option>
      		<option>Visao - Total</option>
      		<option>Reabilitado</option>
			<option>Outro</option>
		</select>
	</td>
  </tr>
  <tr>
    <td><p class="rotulo">Telefone Celular</p></td>
    <td><input class="inputWithe" name="txtCelular" type="text" id="txtCelular" size="14" maxlength="14" value="<%=rsCadastro("cadastroCelular")%>">
		<span class="selo">&nbsp;somente números (DDD e número do celular)</span></td>
  </tr>
  <tr>
    <td><p class="rotulo">Endere&ccedil;o</p></td>
    <td><input class="inputWithe" name="txtEndereco" type="text" id="txtEndereco" size="100" maxlength="100" value="<%=rsCadastro("cadastroEndereco")%>"></td>
  </tr>
	<tr>
    <td><p class="rotulo">Nº</p></td>
    <td><input class="inputWithe" name="txtNo" type="text" id="txtNo" size="10" maxlength="10" value="<%=rsCadastro("cadastroNo")%>"></td>
  </tr>
</tr>
	<tr>
    <td><p class="rotulo">Complemento</p></td>
    <td><input class="inputWithe" name="txtComplemento" type="text" id="txtComplemento" size="50" maxlength="50" value="<%=rsCadastro("cadastroComplemento")%>"></td>
  </tr>
  <tr>
    <td><p class="rotulo">Bairro</p></td>
    <td><input class="inputWithe" name="txtBairro" type="text" id="txtBairro" size="30" maxlength="30" value="<%=rsCadastro("cadastroBairro")%>"></td>
  </tr>
  <tr>
    <td><p class="rotulo">Estado</p></td>
    <td><select class="inputWithe" name="cboUf">
	  <option><%=rsCadastro("cadastroUf")%></option>
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
    <td><input class="inputWithe" name="txtCidade" type="text" id="txtCidade" size="40" maxlength="40" value="<%=rsCadastro("cadastroCidade")%>"></td>
  </tr>
  <tr>
    <td><p class="rotulo">CEP</p></td>
    <td><input class="inputWithe" name="txtCep" type="text" id="txtCep" size="11" maxlength="9" value="<%=rsCadastro("cadastroCep")%>"><span class="selo">&nbsp;exemplo: 12345-678</span></td>
  </tr>
</table>
<%
rsCadastro.Close
%>
<p align="center"><img src="../imgs/linha.png" width="800" height="15"></p>
<%
sSQL = "SELECT * FROM Profissional WHERE profissionalEmail='" & Email & "'"
rsCadastro.Open sSQL, adoCnn, 1, 1
%>
<table width="800" align="center" border="0" cellspacing="0" cellpadding="0">
  	<tr>
		<td width="100"><p class="rotulo">Possui CREF?</p></td>
		<td><select class="inputWithe" name="cboCREF" id="cboCREF" value="<%=rsCadastro("profissionalCref")%>">
        <option selected>SIM</option>
        <option>NAO</option>
		</select>
		</td>
	</tr>
	<tr>
    <td width="100"><p class="rotulo">Nº CREF</p></td>
    <td width="700"><input class="inputWithe" name="txtCREF" type="text" id="txtCREF" size="20" maxlength="20" value="<%=rsCadastro("profissionalCrefNo")%>"><span class="selo"> caso informou que NÃO possui CREF, deixar este campo em branco</span></td>
  </tr>
</table>
<br>
<table width="800" align="center" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td class="subttl">&nbsp;DESCREVA SUAS ESPECIALIDADES, SEPARANDO-AS POR PONTO-E-VÍRGULA ::: NÃO UTILIZE ACENTUAÇÕES NAS PALAVRAS :::</td>
  </tr>
</table>
<table width="800" align="center" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="800"><textarea class="inputWithe" name="txtEspecial" id="txtEspecial" cols="106" rows="10"><%=rsCadastro("profissionalEspecialidades")%></textarea></td>
  </tr>
</table>
<br>
<table width="800" align="center" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td class="subttl">&nbsp;INFORME SUAS ÁREAS DE ATUAÇÃO</td>
  </tr>
</table>
<table width="800" align="center" border="0" cellspacing="0" cellpadding="0">
  <tr>
	  	<%IF rsCadastro("profissionalCondominios") = 1 THEN%>
	  		<td ><input type="checkbox" name="chkCondominios" id="chkCondominios" checked><label class="rotulo">CONDOMÍNIOS / GRUPOS</label></input></td>
		<%ELSE%>
			<td ><input type="checkbox" name="chkCondominios" id="chkCondominios" nochecked><label class="rotulo">CONDOMÍNIOS / GRUPOS</label></input></td>	
		<%END IF%>
  </tr>
<tr>
    <%IF rsCadastro("profissionalAcademias") = 1 THEN%>
	<td><input type="checkbox" name="chkAcademias" id="chkAcademias" checked><label class="rotulo">ACADEMIAS / CENTROS ESPORTIVOS</label></input></td>
	<%ELSE%>
	<td><input type="checkbox" name="chkAcademias" id="chkAcademias" nochecked><label class="rotulo">ACADEMIAS / CENTROS ESPORTIVOS</label></input></td>
	<%END IF%>
  </tr>
<tr>
    <%IF rsCadastro("profissionalResidencias") = 1 THEN%>
		<td><input type="checkbox" name="chkResidencias" id="chkResidencias" checked><label class="rotulo">RESIDÊNCIAS / PARTICULAR</label></input></td>
	<%ELSE%>
		<td><input type="checkbox" name="chkResidencias" id="chkResidencias" nochecked><label class="rotulo">RESIDÊNCIAS / PARTICULAR</label></input></td>
	<%END IF%>
  </tr>
<tr>
    <%IF rsCadastro("profissionalEmpresas") = 1 THEN%>
		<td><input type="checkbox" name="chkEmpresas" id="chkEmpresas" checked><label class="rotulo">EMPRESAS / INSTITUIÇÕES</label></input></td>
	<%ELSE%>
		<td><input type="checkbox" name="chkEmpresas" id="chkEmpresas" nochecked><label class="rotulo">EMPRESAS / INSTITUIÇÕES</label></input></td>
	<%END IF%>
  </tr>
<tr>
    <%IF rsCadastro("profissionalEscolas") = 1 THEN%>
		<td><input type="checkbox" name="chkEscolas" id="chkEscolas" checked><label class="rotulo">ESCOLAS / COLÉGIOS / UNIVERSIDADES</label></input></td>
	<%ELSE%>
		<td><input type="checkbox" name="chkEscolas" id="chkEscolas" nochecked><label class="rotulo">ESCOLAS / COLÉGIOS / UNIVERSIDADES</label></input></td>
	<%END IF%>
  </tr>
	<tr>
    <%IF rsCadastro("profissionalHospitais") = 1 THEN%>
		<td><input type="checkbox" name="chkHospitais" id="chkHospitais" checked><label class="rotulo">HOSPITAIS / CLÍNICAS MÉDICAS</label></input></td>
	<%ELSE%>
		<td><input type="checkbox" name="chkHospitais" id="chkHospitais" nochecked><label class="rotulo">HOSPITAIS / CLÍNICAS MÉDICAS</label></input></td>
	<%END IF%>
  </tr>
<tr>
    <%IF rsCadastro("profissionalParques") = 1 THEN%>
		<td><input type="checkbox" name="chkParques" id="chkParques" checked><label class="rotulo">PARQUES / AR LIVRE / LOCAIS PÚBLICOS</label></td>
	<%ELSE%>
		<td><input type="checkbox" name="chkParques" id="chkParques" nochecked><label class="rotulo">PARQUES / AR LIVRE / LOCAIS PÚBLICOS</label></td>
	<%END IF%>
  </tr>
</table>
<p align="center"><img src="../imgs/linha.png" width="800" height="15"></p>
<div align="center"><input name="cmdEnviar" type="submit" id="cmdEnviar" value="Atualizar Cadastro"></div>
</form>
<p align="center"><img src="../imgs/linha.png" width="800" height="15"></p>
<div class="aviso" align="center">GYM GROUP <%Response.Write("2024" & "-" & Year(Now))%> © Todos os Direitos Reservados</div>
</body>
</html>
