<!--include file="funcoes/funcoes.asp"-->

<%

if request.QueryString("erro") = "1" then
	msgErro = "Login e/ou senha inválidos."
else
	session("teste") = "teste"
	
end if



sub principal	%>
<html>
	<head>
	<title>Login</title>
	<link rel="stylesheet" type="text/css" href="./Layout/estilos.css">
	</head>
	
	<body>
		<form action="Funcoes/Logar.asp?redir=<%=request.QueryString("redir")%>" method="post" enctype="application/x-www-form-urlencoded" name="form1">
		  <table width="100%"  border="0" cellspacing="0" cellpadding="0">
			<tr>
			  <td class="titulo">Acesso</td>
			</tr>
			<tr>
			  <td><font class="cabecalho">Usu&aacute;rio</font><br>
			  <input name="usuario" type="text" class="caixa" id="usuario" size="16" maxlength="20"></td>
			</tr>
			<tr>
			  <td><font class="cabecalho">Senha</font><br>
			  <input name="senha" type="password" class="caixa" id="senha" size="16" maxlength="16"></td>
			</tr>
			<tr>
				<td><input name="OK" type="submit" id="OK" value="OK" class="botao"></td>
			</tr>
			<tr>
				<td class="erro"><%=msgErro%></td>
			</tr>
		  </table>
		</form>
	</body>

</html>	
<%end sub%>
<!--#include file="layout/layoutlogin.asp"-->

