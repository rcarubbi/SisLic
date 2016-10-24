<!--#include file="../Dados/DBfuncoes.asp"-->
<!--#include file="../funcoes/funcoes.asp"-->
<%
call abreBanco

dim acao, id, tipo, erro

acao = request("acao")
id = request("id")
idParceiro = request("idParceiro")
tipo = request("tipo")
redir = request("redir")


sub Impressao
end sub

sub auxiliar
end sub


sub principal
	dim Nome, telefone, email, celular, rs, sql
	
	
	'caso a primeira vez que passa e ação igual a editar
	if acao = "edit" and request.Form("isPostBack") = "" then
		SQL = "Select nome, telefone, email, celular from contato where id=" & id
		'response.Write(sql)
		
		set rs = conn.execute(SQL)
		
		if not rs.eof then
			nome = rs("nome")
			telefone = rs("telefone")
			email = rs("email")
			celular = rs("celular")
		end if
		rs.close
		set rs = nothing
	elseif request.Form("isPostBack") <> "" then
		nome = request.Form("nome")
		telefone = request.Form("telefone")
		email = request.Form("email")
		celular = request.Form("celular")

		
		if nome = "" then erro = erro & "Nome inválido. "
		
		if erro = "" then
			select case acao
			case "add"
				
				SQL = "INSERT INTO contato " & _
						"(idParceiro,  " & _
						"Nome,  " & _						
						"telefone, " & _
						"email,  " & _
						"celular) " & _
						"VALUES " & _
						"(" & idParceiro & ", " & _						
						"'" & replace(nome, "'", "''") & "', " & _
						"'" & replace(telefone, "'", "''") & "', " & _
						"'" & replace(email, "'", "''") & "', " & _
						"'" & replace(celular, "'", "''") & "')"
						
				conn.execute(SQL)
				
				
			case "edit"
			
				SQL = "UPDATE contato " & _
						   "SET nome = '" & replace(nome, "'", "''") & "', " & _
							  "telefone = '" & replace(telefone, "'", "''") & "', " & _
							  "email = '" & replace(email, "'", "''") & "', " & _
							  "celular = '" & replace(celular, "'", "''") & "' " & _							  
						 "WHERE id=" & id
				conn.execute(SQL)
				 
				
			end select
			response.Redirect(redir)
			'response.Redirect(redir & "?tipo=" & tipo & "&idParceiro=" & idParceiro)
		end if
		
	end if
%>
	<form action="" method="post" enctype="application/x-www-form-urlencoded" name="form1" target="_self">
		<input type="hidden" name="acao" value="<%=acao%>">
		<input type="hidden" name="id" value="<%=id%>">
		<input type="hidden" name="idParceiro" value="<%=idParceiro%>">		
		<input type="hidden" name="tipo" value="<%=tipo%>">
		<input type="hidden" name="redir" value="<%=redir%>">		
		<input type="hidden" name="isPostBack" value="1">
		

	  <label class="titulo">Cadastro de Contatos</label><br>
	  <label class="subtitulo">Informe os dados do contato no cliente</label><br>
	  <br>

	 
	  <label class="cabecalho">Nome</label>
	  <br>
	  <input name="Nome" type="text" class="caixa" id="Nome" size="40" maxlength="255" value="<%=nome%>">
	  <br>
	  <label class="cabecalho">Telefone</label>
	  <br>
	  <input name="telefone" type="text" class="caixa" id="telefone" size="20" maxlength="20" value="<%=telefone%>">
	  <br>
	  <label class="cabecalho">Celular</label>
	  <br>
	  <input name="Celular" type="text" class="caixa" id="Celular" size="20" maxlength="20" value="<%=celular%>">
	  <br>
	  <label class="cabecalho">E-mail</label>
	  <br>
	  <input name="email" type="text" class="caixa" id="email" size="20" maxlength="100" value="<%=email%>">
	  <br>
	  <br>
	  <input name="btnOK" type="submit" class="botao" id="btnOK" value="OK">
	  <input name="btnCancelar" type="button" class="botao" id="btnCancelar" value="Cancelar" onClick="window.location.href='<%=redir%>'">
	  <P class="erro">
		&nbsp;<%=erro%>
	  </P>
	</form>
<%end sub


%>

<!--#include file="../layout/layout.asp"-->
<%call FechaBanco%>