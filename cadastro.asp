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
	<title>GYM GROUP ::: Cadastro</title>
</head>

<body>
<table width="800" align="center">
	<tr>
		<td width="250"><img src="../imgs/logo.png" width="118" height="90" alt=""/></td>
	  <td align="center"><h1>CADASTRO</h1></td>
  </tr>
</table>
<p align="center"><img src="../imgs/linha.png" width="800" height="15"></p>
<div align="center">
  <label>ATEN&Ccedil;&Atilde;O<br>
Os campos em Vermelho s&atilde;o de Preencimento Obrigat&oacute;rio<br>
Não Utilizar Acentuações no Cadastro</label>
</div>
	<br>
<form name="frmCadastro" method="Post" action="cadastroEnviar.asp">
<table width="800" align="center" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td class="subttl">&nbsp;Usu&aacute;rio e Senha</td>
  </tr>
</table>
<table width="800" align="center" border="0" cellspacing="0" cellpadding="0">
  	<tr>
		<td width="110"><p class="rotulo">Tipo</p></td>
		<td><select class="inputRedUC" name="cboTipo" id="cboTipo">
        <option selected>PROFISSIONAL</option>
        <option>GRUPO</option>
		<option>INDIVIDUAL</option>
			</select>
		</td>
	</tr>
	<tr>
    <td width="100"><p class="rotulo">e-Mail (Usu&aacute;rio)</p></td>
    <td width="700"><input class="inputRed" name="txtEmail" type="text" id="txtEmail" size="70" maxlength="70"></td>
  </tr>
  <tr>
    <td><p class="rotulo">Senha</p></td>
    <td><input class="inputRed" name="txtSenha" type="password" id="txtSenha" size="15" maxlength="8">
      <span class="selo"> 8 d&iacute;gitos alfa-num&eacute;ricos</span></td>
  </tr>
</table>
<br>
<table width="800" align="center" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td class="subttl">&nbsp;Informa&ccedil;&otilde;es Pessoais</td>
  </tr>
</table>
<table width="800" align="center" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="110"><p class="rotulo">Nome Completo</p></td>
    <td width="690"><input class="inputRedUC" name="txtNome" type="text" id="txtNome" size="100" maxlength="100"></td>
  </tr>
    <tr>
    <td><p class="rotulo">Data Nascimento</p></td>
    <td><input class="inputRedUC" type="date" name="txtNascimento" id="txtNascimento"> </td>
  </tr>
  <tr>
    <td><p class="rotulo">Possui Defici&ecirc;ncia</p></td>
    <td><select class="inputWithe" name="cboDeficiencia" id="cboDeficiencia">
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
    <td><input class="inputWithe" name="txtCelular" type="text" id="txtCelular" size="14" maxlength="14">
		<span class="selo">&nbsp;somente números (DDD e número do celular)</span></td>
  </tr>
  <tr>
    <td><p class="rotulo">Endere&ccedil;o</p></td>
    <td><input class="inputWithe" name="txtEndereco" type="text" id="txtEndereco" size="100" maxlength="100"></td>
  </tr>
	<tr>
    <td><p class="rotulo">Nº</p></td>
    <td><input class="inputWithe" name="txtNo" type="text" id="txtNo" size="10" maxlength="10"></td>
  </tr>
</tr>
	<tr>
    <td><p class="rotulo">Complemento</p></td>
    <td><input class="inputWithe" name="txtComplemento" type="text" id="txtComplemento" size="50" maxlength="50"></td>
  </tr>
  <tr>
    <td><p class="rotulo">Bairro</p></td>
    <td><input class="inputRedUC" name="txtBairro" type="text" id="txtBairro" size="30" maxlength="30"></td>
  </tr>
  <tr>
    <td><p class="rotulo">Estado</p></td>
    <td><select class="inputRedUC" name="cboUf" id="cboUf">
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
    <td><input class="inputRedUC" name="txtCidade" type="text" id="txtCidade" size="40" maxlength="40"></td>
  </tr>
  <tr>
    <td><p class="rotulo">CEP</p></td>
    <td><input class="inputRed" name="txtCep" type="text" id="txtCep" size="11" maxlength="9"><span class="selo">&nbsp;exemplo: 12345-678</span></td>
  </tr>
</table>
<p align="center"><img src="../imgs/linha.png" width="800" height="15"></p>
<div align="center">
  <input name="cmdEnviar" type="submit" id="cmdEnviar" value="Enviar Cadastro">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<input name="cmdLimpar" type="reset" id="cmdLimpar" value="Limpar Cadastro">
</div>
</form>
<p align="center"><img src="../imgs/linha.png" width="800" height="15"></p>
<div class="aviso" align="center">GYM GROUP <%Response.Write("2024" & "-" & Year(Now))%> © Todos os Direitos Reservados</div>
</body>
</html>
