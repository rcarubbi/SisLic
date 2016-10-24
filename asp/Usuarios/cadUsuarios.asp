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
		sql = "update usuarios set ativo=0 where cpf='" & request.Form("optLinha") & "'"
		conn.execute(SQL)
		formLista
	elseif acao = "rec" then
		sql = "update usuarios set ativo=1 where cpf='" & request.Form("optLinha") & "'"
		conn.execute(SQL)
		formLista		
	else
		call FormLista
	end if
end sub


sub Impressao
end sub

sub auxiliar
	if session("usuario_grupo") = "128" then
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
	end if
end sub


sub FormLista
	dim colunas(4,5), botoes(3, 3), sql, rsUsuarios
	
	colunas(0,0) = "CPF"
	colunas(1,0) = 0
	colunas(2,0) = "10%"

	colunas(0,1) = "Nome"
	colunas(1,1) = 1
	colunas(2,1) = "10%"
	
	colunas(0,2) = "Login"
	colunas(1,2) = 2
	colunas(2,2) = "10%"

	
	
	colunas(0,3) = "Dt. Cadastro"
	colunas(1,3) = 3
	colunas(2,3) = "10%"
	
	colunas(0,4) = "Ativo"
	colunas(1,4) = 4
	colunas(2,4) = "10%"
	
	colunas(0,5) = "Grupo"
	colunas(1,5) = 5
	colunas(2,5) = "10%"
	
	'botoes
	if situacao = 1 then
		if session("usuario_grupo") = "128" then		
			botoes(0,0) = true
			botoes(0,1) = true
			botoes(0,2) = true
			botoes(0,3) = false	
		else
			botoes(0,0) = false
			botoes(0,1) = true
		end if
	else
		botoes(0,0) = false
		botoes(0,1) = false
		botoes(0,2) = false	
		if session("usuario_grupo") = "128" then		
			botoes(0,3) = true		
		end if
	end if
				
	sql = "Select u.cpf, u.nome, u.login, u.DataCadastro, u.ativo, g.descricao from Usuarios u INNER JOIN grupo g on g.id=u.id_UsuarioGrupo where u.ativo = " & situacao
	if session("usuario_grupo") <> "128" then
		sql = sql & " and cpf = " & session("usuario_cpf")
	end if
	set rsUsuarios = conn.execute(sql)

	' configurando o grid

	
	' fim da configuração do grid
%>
<form action="cadUsuarios.asp" method="post">

	<label class="titulo">Usuários</label>
	<br>
	<label class="subtitulo">Lista dos Usuários</label>
	<%
	monta_grid rsUsuarios, "", "", "", 0, 0, Colunas, Botoes, ""
	rsUsuarios.close : set rsUsuarios = nothing



%></form><%
end sub

sub FormAddEdit
	
	
	dim sqlEdit, rsEdit, linhaEdit
	
	'se cancelar, volta para a lista
	if request.Form("btnCancelar") <> "" then 
		response.Redirect("cadusuarios.asp")
	end if
	
	'caso editar, recupera o item selecionado	
	if request.Form("cadastroPostBack") = "" then
		if acao = "edit"  then
			linhaEdit = request.form("optLinha")
			sqlEdit = "Select * from usuarios where cpf='" & linhaEdit & "'"
			set rsEdit = conn.execute(sqlEdit)
			cpf = rsEdit("cpf")
			nome = rsEdit("nome")
			login = rsEdit("login")
			senha = rsEdit("senha")
			confsenha = rsEdit("senha")
			grupo = rsEdit("id_usuariogrupo")
			
			rsEdit.close : set rsEdit=nothing
		end if
	else
		dim cpf, nome, login, senha, confsenha, grupo, erro
		
		cpf = trim(replace(replace(request.form("txtcpf"),".",""),"-",""))
		nome = replace(trim(request.form("txtnome")),"'","''")
		login = trim(replace(request.form("txtlogin"),"'","''"))
'		response.write login
'		response.End()
		senha = request.form("txtsenha")
		confsenha = request.form("txtConfSenha")
		if session("usuario_grupo") = "128" then
			grupo = request.form("cbo_grupo")
		end if 
		if isempty(cpf) or cpf = "" then
			erro = "CPF Obrigatório. " 	
		elseif not valida_cpf(cpf) then 
			erro = "CPF Inválido. "
		end if
		if acao = "add" and cpf <> "" then	
			if verifica_existe(cpf,"cpf","usuarios") then
				erro = erro & "CPF já cadastrado. "
			end if
		end if
		if len(nome) < 5 then
			erro = erro & "O Nome deve ter mais de cinco caracteres. "
		end if
		if len(login) < 4 then
			erro = erro & "O Login deve ter mais de quatros caracteres. "
		end if


		if acao = "add" then
			if  verifica_existe(login,"login","usuarios") then 
				erro = erro & "Este login já está sendo usado por outro usuário. "
			end if
		end if
		if len(senha) < 4 then
			erro = erro & "A senha deve possuir mais de 4 digitos. "
		elseif instr(senha," ") > 0 then
			erro = erro & "A senha não deve possuir espaços. "
		elseif senha <> confsenha then
			erro = erro & "Confirmação de senha inválida. "
		end if
		if session("usuario_grupo") = "128" then
			if grupo = "0" or grupo = "" or not isNumeric(grupo) then erro = erro & "Grupo obrigatório."
		end if			
		if erro = "" then
			if acao = "add" then
				dim sqlInsert
				sqlInsert = "Insert into usuarios values ( " & cpf & ",'" & nome & "','" & login & "','" & senha & "'," & formata_data_sql(now) & ",1," & grupo & ")" 
				'response.write sqlinsert
				'response.end
				
				conn.execute(sqlInsert)	
				response.Redirect("cadusuarios.asp")
			elseif acao = "edit" then
				dim sqlUpdate
				sqlUpdate = "Update usuarios set nome='" & nome & "', login='" & login & "', senha = '" & senha & "'"
				if session("usuario_grupo") = "128" then
					sqlUpdate = sqlUpdate & ", id_usuarioGrupo=" & grupo 
				end if
			    sqlUpdate = sqlUpdate & " where cpf = '" & cpf & "'"
				
				conn.execute(sqlUpdate)	
				response.Redirect("cadusuarios.asp")
			end if
		end if
		 
	end if
	%>
	<form name="FrmAddEdit" action="<%=request.ServerVariables("SCRIPT_NAME")%>" method="post">
		<input type="hidden" name="hidID" value="<%=linhaEdit%>">
		<input type="hidden" name="acao" value="<%=acao%>">
		<input type="hidden" name="cadastroPostBack" value="1">
		<label class="titulo">Cadastro de Usuários</label>
	<br><BR>
		<table cellpadding="0" cellspacing="0" width="250">
			<tr>
				<Td colspan="2">
					<span class="cabecalho">CPF</span><BR>
					<input type="text" name="txtCPF" value="<%=CPF%>"  <%if acao = "edit" then response.write "readonly"%> maxlength="11" size="16" class="caixa">
				</Td>
			</tr>
			<tr>
				<Td colspan="2">
					<span class="cabecalho">Nome do usuário</span><BR>
					<input type="text" name="txtNome" value="<%=Nome%>" <%if session("usuario_grupo") <> "128" then response.write "readonly"%> maxlength="150" size="34" class="caixa">
				</Td>
			</tr>
			<tr >
				<Td colspan="2">
					<span class="cabecalho">Login</span><BR>
					<input type="text" name="txtlogin" value="<%=login%>" maxlength="20" size="15" class="caixa">
				</Td>
			</tr>
			<tr >
				<Td style="padding-right:5px;">
					<span class="cabecalho">Senha</span><BR>
					<input type="password" name="txtsenha" value="<%=senha%>" maxlength="16" size="15" class="caixa">
				</Td>
				<Td>
					<span class="cabecalho">Confirmar Senha</span><BR>
					<input type="password" name="txtConfSenha" value="<%=confSenha%>" maxlength="16" size="15" class="caixa">
				</Td>
			</tr>
			<tr >
				<Td colspan="2">
					<span class="cabecalho">Grupo do Usuário</span><BR>
					<%if session("usuario_grupo") <> "128" then%>
						<%sql = "select descricao from grupo where id = " & grupo
						set rs = conn.execute(sql)
						if not rs.eof then
							grupoDesc = rs("descricao")
						end if
						rs.close : set rs = nothing
						%>
						<input type="text" name="txtGrupo" value="<%=GrupoDesc%>" readonly maxlength="20" size="15" class="caixa">
					<%else
						dim sqlGrupoUsu
						sqlGrupoUsu = "SELECT id, descricao from grupo where ativo = 1"
						ComboDB "cbo_Grupo","descricao","id",sqlGrupoUsu,grupo,"",""
					end if	
						%>
						
				</Td>
			</tr>
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