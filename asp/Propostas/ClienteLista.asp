<!--#include file="../Dados/DBfuncoes.asp"-->
<!--#include file="../funcoes/funcoes.asp"-->
<!--#include file="../funcoes/gridgeral.asp"-->
<%call abreBanco%>

<%
dim acao


acao = request.Form("acao")

if acao = "edit"  then
	response.Redirect("clienteDetalhes.asp?tipo=0&idParceiro=" & request.Form("optLinha") & "&acao=" & acao)
elseif acao = "add" then
	response.Redirect("cadCliente.asp?redir=ClienteLista.asp&tipo=0&idParceiro=" & request.Form("optLinha") & "&acao=" & acao)
elseif acao = "del" then
	id = request.Form("optLinha")
	if isNumeric(id) and id<> "" then
		sql = "Update parceiro set ativo = 0 where id=" & id
		conn.execute(SQL)
	end if
elseif acao = "rec" then
	id = request.Form("optLinha")
	if isNumeric(id) and id<> "" then
		sql = "Update parceiro set ativo = 1 where id=" & id
		conn.execute(SQL)
	end if
end if

sub principal

	formLista
	
end sub

sub Impressao
end sub

sub auxiliar
	%>
	<form name="form1" action="" method="post" enctype="application/x-www-form-urlencoded">
		<input type="hidden" name="filtros" value="1">
		<%
		dim itemTexto(1), itemValor(1)
		
		itemTexto(0) = "Ativos"
		itemValor(0) = 1
		
		itemTexto(1) = "Deletados"
		itemValor(1) = 0
		
		%><label class="cabecalho">Situação</label><br><%
		Combo "situacao", itemTexto, itemValor, request.Form("situacao")
		
		%>
		<p>
		<input type="submit" name="OK" value="OK" class="botao">
		</p>
	</form>
	<%
end sub

sub FormLista
	
	dim situacao
	
	situacao = request.Form("situacao")
	'por padrao situação é sempre ativo
	if situacao = "" then situacao = "1"
			
	sql = "Select id, replace(nome,char(13),'<br>') as nome, DataCadastro from Parceiro where tipo = 0 and ativo = " & situacao

	set rs = conn.execute(sql)
	
	' configurando o grid
	dim colunas(4,2), botoes(3, 3)
	
	colunas(0,0) = "ID"
	colunas(1,0) = 0
	colunas(2,0) = "5%"
		
	colunas(0,1) = "Nome"
	colunas(1,1) = 1
	colunas(2,1) = "70%"
		
	colunas(0,2) = "Data de Cadastro"
	colunas(1,2) = 2
	colunas(2,2) = "25%"

	
	'botoes
	if situacao = 1 then
		botoes(0,0) = true
		botoes(0,1) = true
		botoes(0,2) = true		
		botoes(0,3) = false	
	else
		botoes(0,0) = false
		botoes(0,1) = false
		botoes(0,2) = false			
		botoes(0,3) = true		
	end if
	
	' fim da configuração do grid
	%>
	<form action="" method="post">
	<label class="titulo">Cliente</label>
	<br>
	<label class="subtitulo">Lista dos Clientes</label>
	<%
	
		monta_grid rs, "", "", "", 0, 0, Colunas, Botoes, ""
		rs.close : set rs = nothing
	
	
	%></form><%
end sub
%>

<!--#include file="../layout/layout.asp"-->
 