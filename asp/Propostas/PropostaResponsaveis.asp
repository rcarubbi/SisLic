<!--#include file="../Dados/DBfuncoes.asp"-->
<!--#include file="../funcoes/funcoes.asp"-->
<!--#include file="../funcoes/gridGeral.asp"-->
<%call abreBanco%>
<%
dim situacao, acao
	
	
acao = request.form("acao")
situacao = request.form("cboSituacao")
if situacao = "" then situacao = 1

sub principal
	if acao="add" or acao="edit" then
		call FormAddEdit
	elseif acao = "del" then
		sql = "update propostaResponsavel set ativo=0 where cpf=" & request.Form("optLinha") 
		conn.execute(SQL)
		formLista
	elseif acao = "rec" then
		sql = "update propostaResponsavel set ativo=1 where cpf=" & request.Form("optLinha") 
		conn.execute(SQL)
		formLista		
	else
		call FormLista
	end if
end sub


sub Impressao
end sub

sub auxiliar
	
		if acao <> "add" and acao <> "edit" then
			dim situacoesDescr, situacoesVal
			
			redim situacoesDescr(1)
			situacoesDescr(0) = "Ativos"
			situacoesDescr(1) = "Deletados"
			redim situacoesVal(1)
			situacoesVal(0) = 1
			situacoesVal(1) = 0
		
		%>
		<form name="frmAux" action="<%=request.ServerVariables("SCRIPT_NAME")%>" method="post">
			<label class="cabecalho">Situação</label><br>
			<%Combo "cboSituacao",SituacoesDescr,SituacoesVal,situacao%>
			<p>
				<input type="submit" value="OK" class="botao" name="CmdAux">
				</p>
			
		</form>
		<%
		end if
	
end sub


sub FormLista
	dim colunas(4,5), botoes(3, 3), sql, rsUsuarios
	
	colunas(0,0) = "CPF"
	colunas(1,0) = 0
	colunas(2,0) = "10%"

	colunas(0,1) = "R.G."
	colunas(1,1) = 1
	colunas(2,1) = "10%"
	
	colunas(0,2) = "Nome"
	colunas(1,2) = 2
	colunas(2,2) = "10%"

	
	
	colunas(0,3) = "Dt. Cadastro"
	colunas(1,3) = 3
	colunas(2,3) = "10%"
	
	colunas(0,4) = "Ativo"
	colunas(1,4) = 4
	colunas(2,4) = "10%"
	
	colunas(0,5) = "Cargo"
	colunas(1,5) = 5
	colunas(2,5) = "10%"
	
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
				
	sql = "Select r.cpf, r.rg, r.nome, r.DataCadastro, r.ativo, r.cargo from propostaResponsavel r where r.ativo = " & situacao
			
	set rsResponsaveis = conn.execute(sql)

	' configurando o grid

	
	' fim da configuração do grid
%>
<form action="PropostaResponsaveis.asp" method="post">

	<label class="titulo">Responsáveis</label>
	<br>
	<label class="subtitulo">Lista dos Responsáveis pelas propostas</label>
	<%
	monta_grid rsResponsaveis, "", "", "", 0, 0, Colunas, Botoes, ""
	rsResponsaveis.close : set rsResponsaveis = nothing



%></form><%
end sub

sub FormAddEdit
	
	
	dim sqlEdit, rsEdit, linhaEdit
	
	'se cancelar, volta para a lista
	if request.Form("btnCancelar") <> "" then 
		response.Redirect("PropostaResponsaveis.asp")
	end if
	
	'caso editar, recupera o item selecionado	
	if request.Form("cadastroPostBack") = "" then
		if acao = "edit"  then
			linhaEdit = request.form("optLinha")
			sqlEdit = "Select * from propostaResponsavel where cpf=" & linhaEdit 
			set rsEdit = conn.execute(sqlEdit)
			cpf = rsEdit("cpf")
			rg = rsEdit("rg")
			nome = rsEdit("nome")
			cargo = rsEdit("cargo")
			rsEdit.close : set rsEdit=nothing
		end if
	else
		dim cpf, rg, nome, cargo, erro
		
		cpf = trim(replace(replace(request.form("txtcpf"),".",""),"-",""))
		rg = replace(trim(request.form("txtrg")),"'","''")
		cargo = trim(replace(request.form("txtcargo"),"'","''"))
		nome = replace(trim(request.form("txtnome")),"'","''")
'		response.write login
'		response.End()
		
		if isempty(cpf) or cpf = "" then
			erro = "CPF Obrigatório. " 	
		elseif not valida_cpf(cpf) then 
			erro = "CPF Inválido. "
		end if
		
		if acao = "add" and cpf <> "" then	
			if verifica_existe(cpf,"cpf","propostaResponsavel") then
				erro = erro & "CPF já cadastrado. "
			end if
		end if
		if isempty(rg) or rg= "" then
			erro = erro & "R.G. Obrigatório. " 	
		end if
		
		if len(nome) < 5 then
			erro = erro & "O Nome deve ter mais de cinco caracteres. "
		end if
		
		
		if erro = "" then
			if acao = "add" then
				dim sqlInsert
				sqlInsert = "Insert into propostaResponsavel (cpf, rg, nome, cargo) values ( " & cpf & ",'" & rg & "','" & nome & "','" & cargo & "')" 
				'response.write sqlinsert
				'response.end
				
				conn.execute(sqlInsert)	
				response.Redirect("propostaResponsaveis.asp")
			elseif acao = "edit" then
				dim sqlUpdate
				sqlUpdate = "Update propostaResponsavel set nome='" & nome & "', rg='" & rg & "', cargo='" & cargo & "'"
			    sqlUpdate = sqlUpdate & " where cpf = " & cpf 
				conn.execute(sqlUpdate)	
				response.Redirect("propostaResponsaveis.asp")
			end if
		end if
		 
	end if
	%>
	<form name="FrmAddEdit" action="<%=request.ServerVariables("SCRIPT_NAME")%>" method="post">
		<input type="hidden" name="hidID" value="<%=linhaEdit%>">
		<input type="hidden" name="acao" value="<%=acao%>">
		<input type="hidden" name="cadastroPostBack" value="1">
		<label class="titulo">Cadastro de Responsáveis pelas propostas</label>
	<br><BR>
		<table cellpadding="0" cellspacing="0" width="250">
			<tr>
				<Td colspan="2">
					<span class="cabecalho">CPF</span><BR>
					<input type="text" name="txtCPF" value="<%=CPF%>"  <%if acao = "edit" then response.write "readonly"%> maxlength="11" size="16" class="caixa">
				</Td>
			</tr>
			<tr >
				<Td colspan="2">
					<span class="cabecalho">RG</span><BR>
					<input type="text" name="txtrg" value="<%=rg%>" maxlength="20" size="22" class="caixa">
				</Td>
			</tr>
			<tr>
				<Td colspan="2">
					<span class="cabecalho">Nome do Responsavel</span><BR>
					<input type="text" name="txtNome" value="<%=Nome%>" maxlength="150" size="34" class="caixa">
				</Td>
			</tr>
			
			<tr >
				<Td style="padding-right:5px;">
					<span class="cabecalho">Cargo</span><BR>
					<input type="text" name="txtcargo" value="<%=cargo%>" maxlength="50" size="25" class="caixa">
				</Td>
				
			
			<Tr>
				<td style="padding-top:5px;">
					<input type="submit" name="cmdOk" Value="OK" class="botao">
				</td>
			</tr>
			
		
		
		</table>
		<br>
		<div class="erro" style="width:50%" ><%=iif(erro<>"",erro,"&nbsp;")%></div>				

	</form>
<%end sub%>





<!--#include file="../layout/layout.asp"-->
<%call FechaBanco%>