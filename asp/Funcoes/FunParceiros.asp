<!--#include file="../Dados/DBfuncoes.asp"-->
<!--#include file="../funcoes/funcoes.asp"-->
<%
call abreBanco

dim acao, id, tipo, erro

acao = request("acao")
id = request("idParceiro")
tipo = request("tipo")
redir = request("redir")

sub Impressao
end sub

sub auxiliar
end sub


sub principal
	dim CNPJ, CNPJOriginal, Nome, Site, rs, sql
	
	
	'caso a primeira vez que passa e ação igual a editar
	if acao = "edit" and request.Form("isPostBack") = "" then
		SQL = "Select cnpj, nome, site from parceiro where id=" & id
		'response.Write(sql)
		
		set rs = conn.execute(SQL)
		
		if not rs.eof then
			cnpj = rs("cnpj")
			nome = rs("nome")
			site = rs("site")
			CNPJOriginal = cnpj
		end if
		rs.close
		set rs = nothing
	elseif request.Form("isPostBack") <> "" then
		cnpj = request.Form("cnpj")
		nome = request.Form("nome")
		site = request.Form("site")
		CNPJOriginal = request.Form("CNPJOriginal")
		
		
		if cnpj <> "" then
			if valida_cnpj(cnpj) = false then 
				erro = erro & "CNPJ inválido. "
			elseif acao = "add" or CNPJOriginal <> cnpj then
				if verifica_existe(cnpj, "cnpj", "parceiro") then 
					erro = erro & "Já existe outro "
					if tipo = "0" then
						erro = erro & "cliente "
					else
						erro = erro & "fornecedor "
					end if
					erro = erro & "com esse CNPJ. "
				end if
			end if
		end if
		
				
		if nome = "" then erro = erro & "Nome inválido. "
		
		if erro = "" then
			select case acao
			case "add"
				
				SQL = "INSERT INTO Parceiro " & _
						"(Cnpj,  " & _
						"Nome, " & _
						"Site,  " & _
						"Ativo, " & _
						"DataCadastro, " & _
						"Tipo) " & _
						"VALUES " & _
						"(" & iif(cnpj="","null",cnpj) & ", " & _
						"'" & replace(nome, "'", "''") & "', " & _
						"'" & replace(site, "'", "''") & "', " & _
						"1, " & _
						formata_data_sql(now) & ", " & _
						tipo & ")"
						
				conn.execute(SQL)
				
				sql = "SELECT max(id) as maxId FROM Parceiro"
				set rs = conn.execute(SQL)
				if not rs.eof then
					id = rs("maxId")
				end if
				rs.close
				set rs = nothing
				
				'se novo, obrigatoriamente cadastra um endereco

				
				response.Redirect("cadEndereco.asp?redir=" & server.URLEncode(redir) & "&tipo=" & tipo & "&idParceiro=" & id & "&acao=" & acao)


			case "edit"
			
				SQL = "UPDATE Parceiro " & _
						   "SET Cnpj = " & iif(cnpj="","null",cnpj) & ", " & _
							  "Nome = '" & replace(nome, "'", "''") & "', " & _
							  "Site = '" & replace(site, "'", "''") & "' " & _
						 "WHERE id=" & id
				conn.execute(SQL)
				 
				response.Redirect(redir) ' & "?tipo=" & tipo & "&idParceiro=" & id)
			end select
			
			
		end if
		
	end if
%>
	<form action="" method="post" enctype="application/x-www-form-urlencoded" name="form1" target="_self">
		<input type="hidden" name="acao" value="<%=acao%>">
		<input type="hidden" name="idParceiro" value="<%=id%>">
		<input type="hidden" name="tipo" value="<%=tipo%>">
		<input type="hidden" name="redir" value="<%=redir%>">		
		<input type="hidden" name="isPostBack" value="1">
		<input type="hidden" name="cnpjOriginal" value="<%=CnpjOriginal%>">
		
	  <%if tipo = "0" then 'só exibe o cnpj se for fornecedor%>
		  <label class="titulo">Cadastro de Cliente</label><br>
		  <label class="subtitulo">Informe os dados do cliente</label><br>
		  <br>
	 <%else%>
		  <label class="titulo">Cadastro de Fornecedores</label><br>
		  <label class="subtitulo">Informe os dados do fornecedor</label><br>
		  <br>
	 <%end if%>
	 
	  <label class="cabecalho">CNPJ</label><br>
	  <input name="Cnpj" type="text" class="caixa" id="Cnpj" size="19" maxlength="14" value="<%=cnpj%>">
	  <br>
	  <label class="cabecalho">Nome</label>
	  <br>
      <textarea name="Nome" cols="40" rows="5" class="caixa" id="Nome" onKeyUp="maxLengthTextArea(this, event, 1000);"><%=Nome%></textarea>
	   <BR>
	  <label class="cabecalho">Site</label>
	  <br>
	  <input name="Site" type="text" class="caixa" id="Site" size="40" maxlength="800" value="<%=site%>">
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