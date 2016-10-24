<!--#include file="../Dados/DBfuncoes.asp"-->
<!--#include file="../funcoes/funcoes.asp"-->
<!--#include file="../funcoes/gridgeral.asp"-->
<%call abreBanco%>
<%
dim acao

acao = request.Form("acao")

sub principal
	if acao = "edit" then
		FormEdit
	else
		FormLista
	end if
end sub

sub FormLista
	dim matrizParametros, SQL, rs, intIndice
	dim situacao
	
	situacao = request.Form("situacao")
	'por padrao situação é sempre ativo
	if situacao = "" then situacao = "1"
			
	SQL = "Select * from grupo where ativo = " & situacao
	set rs = conn.execute(SQL)
' configurando o grid
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

	

		botoes(0,0) = false
		botoes(0,1) = true
		botoes(0,2) = false		
		botoes(0,3) = false	

	
	' fim da configuração do grid
	%>
		
	<form action="cadrestricoes.asp" method="post">
	<label class="titulo">Restrições</label>
	<br>
	<label class="subtitulo">Lista de restrições para os grupos de usuários</label>
	<%
	
		monta_grid rs, "", "", "", 0, 0, Colunas, Botoes, ""
		rs.close : set rs = nothing
	%></form><%
end sub


sub Impressao
end sub

sub auxiliar
end sub


sub FormEdit
	dim id, grupoDescricao, i
	
	id = request.Form("optLinha")
	
	if request.Form("cadastroPostBack") = "" then
		'retornando a descrição do grupo ------------
		sql = "SELECT descricao from grupo where id=" & id
		set rs = conn.execute(SQL)
		if not rs.eof then
			grupoDescricao = rs("descricao")
		end if
		rs.close
		set rs = nothing
		'--------------------------------------------
	else
		'deleta todas as restricoes para o grupo e as recria
		
		sql = "delete from restricoes where id_RestricoesGrupo=" & id
		conn.execute(SQL)
		
		for each input in request.Form
			if inStr(input, "chk") > 0 then
				if isNumeric(request.Form(input)) then
					sql = "Insert into restricoes(id_restricoesGrupo, id_menuItem) Values (" & _
								id & ", " & request.Form(input) & ")"
					conn.execute(SQL)						
				end if
			end if
		next
		
		formLista
		exit sub
	
	end if	
	
%>
	<form method="post" enctype="application/x-www-form-urlencoded" name="form1">
		<input type="hidden" name="optLinha" value="<%=id%>">
		<input type="hidden" name="acao" value="<%=acao%>">
		<input type="hidden" name="cadastroPostBack" value="1">
		<p>
		<label class="titulo">Restrições para <%=grupoDescricao%></label><br>
		<label class="subtitulo">Selecione os itens que deseja restringir</label>
		</p>
		<%
		dim modText(3)
				
		modText(0) = "Propostas"
		modText(1) = "Comércio"
		modText(2) = "Financeiro"
		modText(3) = "Usuários"
		
		
		for i = 1 to 4
		
			%><p><label class="cabecalho"><%=modText(i-1)%></label><br><%
			dim sql, rs
			set rs = server.CreateObject("ADODB.recordset")
			
			'sql = "select id, descricao, chamada, id_superior, select case  from menu where id_modulo=" & i & " and ativo=1 order by id_superior, id"
			sql = "SELECT menu.id, menu.Descricao, menu.Chamada, menu.Id_superior, restricoes.id AS restricao_id " & _
						"FROM menu LEFT JOIN restricoes ON menu.id = restricoes.id_menuitem " & _
						" AND restricoes.id_restricoesGrupo = " & id & _
						" WHERE id_modulo=" & i 
						
						
						'response.Write(sql)
			rs.open sql, conn, 1
			if not rs.eof then
				m = rs.GetRows
				exibeMenu m, modulo, 0, 0
			end if
			rs.close
			set rs = nothing
			%></p><%
		next
			
			
		%>
		
		<br><br>
		<input type="submit" name="btnOK" Value="OK" class="botao">
		&nbsp;
		<input type="submit" name="btnCancelar" value="Cancelar" class="botao">
	</form>

<%
	

end sub



Function exibeMenu(m, modulo, id_superior, nivel)

	dim t, i, z
	
	t = 0
	
	
	for i = 0 to ubound(m,2)
	
		if cInt(m(3,i)) = cInt(id_superior) then
			
			
			'colocando um espaço na frente de acordo com o nivel
			for z = 0 to nivel -1 
			%>&nbsp;&nbsp;&nbsp;<%
			next
			%><input type="checkbox" name="chk_<%=m(0,i)%>" id="chk_<%=m(0,i)%>" value="<%=m(0,i)%>" <%
				if not isNull(m(4,i)) then
					response.Write("checked")
				end if
			%>>
			<label class="texto" for="chk_<%=m(0,i)%>"><%=m(1, i)%></label><br>
			<%
								
			exibeMenu m, modulo, m(0,i), nivel + 1
			
			
		end if	
		
	next
	

end function
%>




<!--#include file="../layout/layout.asp"-->