<!--#include file="funcoes.asp"-->
<!--#include file="mudaModulo.asp"-->
<!--#include file="../dados/dbfuncoes.asp"-->

<%
'################################################
' ações realizadas na hora de logar no sistema
'################################################


dim usuario, senha, redir, rs

usuario = replace(request.Form("usuario"),"'", "''")
senha = replace(request.Form("senha"), "'", "''")
redir = request.QueryString("redir")

abreBanco

SQL = "Select * from usuarios where login='" & usuario & "' AND senha='" & senha & "' AND ativo=1"
set rs = Conn.execute(SQL)


if rs.eof then
	rs.close : set rs = nothing
	fechaBanco
	response.Redirect("../login.asp?erro=1")
else

	'carrega os parametros do usuario
	ativarUsuario rs
	rs.close : set rs = nothing
	fechaBanco
	
	if redir = "" then redir = "../Usuarios/default.asp"

	mudaModulo redir
	

	response.Redirect(redir)
	
end if



sub ativarUsuario(rs)
	dim SQL, rsM

	
	session("usuario_cpf") = rs("cpf")
	session("usuario_nome") = rs("nome")
	session("usuario_cpf") = rs("cpf")
	session("usuario_grupo") = rs("id_usuarioGrupo")

	'carregando os menus

	sql = "SELECT id, descricao, chamada, id_superior, id_modulo, paginaRedir " & _
			" from menu where id NOT IN(SELECT id_menuitem FROM restricoes where id_restricoesGrupo=" & session("usuario_grupo") & ") " & _
			" AND ativo=1 ORDER BY id_superior, ordem"


	
	set rsM = conn.execute(SQL)

	'cria uma copia do recordset em uma matriz
	'a primeira coluna indica a coluna no banco de dados, a segunda coluna é o registro
	'exemplo
	'pegar o campo "id" do segundo registro
	'  +------> id (o primeiro campo)
	'  |  +---> o segundo registro (começa do zero)
	'm(1, 0)
	m = rsM.getRows()
	
	rsM.close :	set rsM = nothing
	
	session("usuario_mnu") = m
end sub


%>
