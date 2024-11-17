<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Dim adoCnn
Dim rsCadastro
Dim sCnn, sSQL
Dim Nome, Deficiencia, Nascimento, Celular, Endereco, No, Complemento, Bairro, Estado, Cidade, Cep

Email = Request.QueryString("sEmail")

Set adoCnn = Server.CreateObject("ADODB.Connection")
Set rsCadastro = Server.CreateObject("ADODB.RecordSet")
sCnn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("../../db/dbGG.mdb")
'--sCnn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & "D:\web\localuser\gymgroup\banco\dbGG.mdb"--'
adoCnn.Open(sCnn)

sSQL = "SELECT * FROM Cadastro WHERE cadastroEmail='" & Email & "'"
rsCadastro.Open sSQL, adoCnn, 1, 1
Nome = rsCadastro("cadastroNome")
Nascimento = rsCadastro("cadastroNascimento")
Deficiencia = rsCadastro("cadastroDeficiencia")
Celular	= rsCadastro("cadastroCelular")
Endereco = rsCadastro("cadastroEndereco")
No = rsCadastro("cadastroNo")
Complemento = rsCadastro("cadastroComplemento")
Bairro = rsCadastro("cadastroBairro")
Uf = rsCadastro("cadastroUf") 
Cidade = rsCadastro("cadastroCidade")
Cep	= rsCadastro("cadastroCep")
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
		<td width="230"><img src="../imgs/logo.png" width="118" height="90" alt=""/></td>
	  <td width="450" align="center"><h1>CADASTRO</h1></td>
	  <td width="70" align="right"><a href="javascript:history.back()">VOLTAR</a></td>
	  <td width="50" align="right"><a href="../index.asp">SAIR</a></td>
  </tr>
</table>
<p align="center"><img src="../imgs/linha.png" width="800" height="15"></p>
<form name="frmIndividual" method="Post" action="individualEnviar.asp?sEmail=<%=Email%>&iCdstr=1">
<table width="800" align="center" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td class="subttl">&nbsp;Informa&ccedil;&otilde;es Pessoais</td>
  </tr>
</table>
<table width="800" align="center" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="130"><p class="rotulo">Nome Completo</p></td>
    <td width="670"><input class="inputWhite" name="txtNome" type="text" id="txtNome" size="100" maxlength="100" value="<%=Nome%>"></td>
  </tr>
    <tr>
    <td><p class="rotulo">Data Nascimento</p></td>
    <td><input class="inputWithe" type="text" name="txtNascimento" id="txtNascimento" size="10" maxlength="10" value="<%=Nascimento%>"> </td>
  </tr>
  <tr>
    <td><p class="rotulo">Possui Defici&ecirc;ncia</p></td>
    <td><select class="inputWithe" name="cboDeficiencia" id="cboDeficiencia" value="<%=Deficiencia%>">
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
    <td><input class="inputWithe" name="txtCelular" type="text" id="txtCelular" size="14" maxlength="14" value="<%=Celular%>">
		<span class="selo">&nbsp;somente números (DDD e número do celular)</span></td>
  </tr>
  <tr>
    <td><p class="rotulo">Endere&ccedil;o</p></td>
    <td><input class="inputWithe" name="txtEndereco" type="text" id="txtEndereco" size="100" maxlength="100" value="<%=Endereco%>"></td>
  </tr>
	<tr>
    <td><p class="rotulo">Nº</p></td>
    <td><input class="inputWithe" name="txtNo" type="text" id="txtNo" size="10" maxlength="10" value="<%=No%>"></td>
  </tr>
</tr>
	<tr>
    <td><p class="rotulo">Complemento</p></td>
    <td><input class="inputWithe" name="txtComplemento" type="text" id="txtComplemento" size="50" maxlength="50" value="<%=Complemento%>"></td>
  </tr>
  <tr>
    <td><p class="rotulo">Bairro</p></td>
    <td><input class="inputWithe" name="txtBairro" type="text" id="txtBairro" size="30" maxlength="30" value="<%=Bairro%>"></td>
  </tr>
  <tr>
    <td><p class="rotulo">Estado</p></td>
    <td><select class="inputWithe" name="cboUf">
	  <option><%=Uf%></option>
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
    <td><input class="inputWithe" name="txtCidade" type="text" id="txtCidade" size="40" maxlength="40" value="<%=Cidade%>"></td>
  </tr>
  <tr>
    <td><p class="rotulo">CEP</p></td>
    <td><input class="inputWithe" name="txtCep" type="text" id="txtCep" size="11" maxlength="9" value="<%=Cep%>"><span class="selo">&nbsp;exemplo: 12345-678</span></td>
  </tr>
</table>
<%
rsCadastro.Close
%>
<div align="center"><input name="cmdEnviar" type="submit" id="cmdEnviar" value="Atualizar Cadastro">
</div>
</form>
<p align="center"><img src="../imgs/linha.png" width="800" height="15"></p>
<form name="frmComplemento" method="Post" action="individualEnviar.asp?sEmail=<%=Email%>&iCdstr=2">
<%
sSQL = "SELECT * FROM Individual WHERE individualEmail='" & Email & "'"
rsCadastro.Open sSQL, adoCnn, 1, 1
%>
<table width="800" align="center" border="0" cellspacing="0" cellpadding="0">
  	<tr>
		<td width="100"><p class="rotulo">Altura (M)</p></td>
    <td width="700"><input class="inputWithe" name="txtAltura" type="text" id="txtAltura" size="15" maxlength="15" value="<%=Cdbl(rsCadastro("individualAltura"))%>">
    <span class="selo"> exemplo: 1,80</span></td>
	</tr>
	<tr>
    <td width="100"><p class="rotulo">Peso (KG)</p></td>
    <td width="700"><input class="inputWithe" name="txtPeso" type="text" id="txtPeso" size="15" maxlength="15" value="<%=Cdbl(rsCadastro("individualPeso"))%>">
    <span class="selo"> exemplo: 90,0</span></td>
  </tr>
</table>
<br>
<table width="800" align="center" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td class="subttl">&nbsp;CONTE MAIS SOBRE VOCÊ ::: NÃO UTILIZE ACENTUAÇÕES NAS PALAVRAS :::</td>
  </tr>
</table>
<table width="800" align="center" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="800"><textarea class="inputWithe" name="txtSobre" id="txtSobre" cols="106" rows="10"><%=rsCadastro("individualSobre")%></textarea></td>
  </tr>
</table>
<br>
<div align="center"><input name="cmdEnviar" type="submit" id="cmdEnviar" value="Complementar Cadastro">
</div>
</form>
<p align="center"><img src="../imgs/linha.png" width="800" height="15"></p>
<div class="aviso" align="center">GYM GROUP <%Response.Write("2024" & "-" & Year(Now))%> © Todos os Direitos Reservados</div>
<%
rsCadastro.Close
SET rsCadastro = NOTHING
adoCnn.Close
SET adoCnn = NOTHING
%>
</body>
</html>
