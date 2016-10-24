<!--#include file="../Dados/DBfuncoes.asp"-->
<!--#include file="../funcoes/funcoes.asp"-->
<!--#include file="../funcoes/gridgeral.asp"-->
<%call abreBanco%>

<%
dim acao, erro


acao = request.Form("acao")


sub principal

	if acao = "add" or acao = "edit" then
		FormAddEdit
	elseif acao = "del" then
		if request.form("optlinha") = "128" then erro = " O grupo master não pode ser excluído. "
		if erro = "" then
			sql = "update grupo set ativo=0 where id=" & request.Form("optLinha")
			conn.execute(SQL)
		end if
		
		formLista
	elseif acao = "rec" then
		sql = "update grupo set ativo=1 where id=" & request.Form("optLinha")
		conn.execute(SQL)
		formLista		
	else
		formLista
	end if
end sub

sub Impressao
end sub

sub auxiliar
	if acao <> "add" and acao <> "edit" then
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
	end if
end sub

sub FormLista
	dim sqlGrupos, rsGrupos
	dim situacao
	situacao = request.Form("situacao")
	'por padrao situação é sempre ativo
	if situacao = "" then situacao = "1"
			
	sqlGrupos = "Select * from grupo where ativo = " & situacao
	set rsGrupos = conn.execute(sqlGrupos)
	
	dim colunas(4,2), botoes(3, 3)
	
	colunas(0,0) = "ID"
	colunas(1,0) = 0
	colunas(2,0) = "10%"
		
	colunas(0,1) = "Descrição"
	colunas(1,1) = 1
	colunas(2,1) = "70%"
		
	colunas(0,2) = "Ativo"
	colunas(1,2) = 2
	colunas(2,2) = "20%"

	
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
	
<form action="cadgrupos.asp" method="post">
	<label class="titulo">Grupos</label>
	<br>
	<label class="subtitulo">Lista dos Grupos de Usuários</label>

	<%
	monta_grid rsGrupos, "", "", "", 0, 0, Colunas, Botoes, ""
	rsGrupos.close : set rsGrupos = nothing
	
%></form>	<label class="erro"><%=erro%></label><%
	
end sub
sub FormAddEdit
	dim sqlEdit, rsEdit, linhaEdit, descricao, erro
	
	'se cancelar, volta para a lista
	if request.Form("btnCancelar") <> "" then 
		response.Redirect("cadGrupos.asp")
	end if
	
	'caso editar, recupera o item selecionado	
	if request.Form("cadastroPostBack") = "" then
		if acao = "edit"  then
			id = request.form("optLinha")
			sqlEdit = "Select * from Grupo where id=" & id
			
			set rsEdit = conn.execute(sqlEdit)
			descricao = rsEdit("descricao")
			rsEdit.close : set rsEdit=nothing
		end if
	else
		descricao = request.Form("txtDescricao")
		id = request.Form("id")
		
		if descricao = "" then erro = "Descrição inválida. "
		
		if erro = "" then
			select case acao
			case "add"
				sql = "Insert into grupo (descricao) values ('" & descricao & "')"
				conn.execute(SQL)		
			case "edit"		
				sql = "Update grupo set descricao='" & descricao & "' where id=" & id
				conn.execute(SQL)		
			end select
			
			response.Redirect("cadGrupos.asp")
		end if		
	end if
	
	
	%>
	<label class="titulo">Cadastro de Grupos</label><br>

	
	<form name="FrmAddEdit" action="<%=request.ServerVariables("SCRIPT_NAME")%>" method="post">
		<input type="hidden" name="id" value="<%=id%>">
		<input type="hidden" name="acao" value="<%=acao%>">
		<input type="hidden" name="cadastroPostBack" value="1">

		<table cellpadding="0" cellspacing="0">
			<tr>
				<Td>
					<span class="cabecalho">Grupo de Usuários</span><BR>
					<input type="text" name="txtDescricao" value="<%=descricao%>" maxlength="50" size="30" class="caixa">
				</Td>
			</tr>
			
			<tr>
				<td>&nbsp;
									
				</td>
			</tr>		
			
			<Tr>
				<td>
									
					
					
					<input type="submit" name="btnOK" Value="OK" class="botao">
					&nbsp;
					<input type="submit" name="btnCancelar" value="Cancelar" class="botao">
					
				</td>
			</tr>
		
		
		</table>
		
			<p>
				&nbsp;<label class="erro"><%=erro%></label>
			</p>
		
	</form>
<%end sub%>

<!--#include file="../layout/layout.asp"-->
<%call FechaBanco%>