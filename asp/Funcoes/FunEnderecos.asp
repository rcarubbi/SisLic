<!-- #include file="../Dados/DBfuncoes.asp"-->
<!-- #include file="../funcoes/funcoes.asp"-->
<%
call abreBanco

dim acao, id, idParceiro, erro, tipo, SoEndereco

acao = request("acao") 'a ação a fazer, novo ou editar
id = request("id") 'o id do endereco
idParceiro = request("idParceiro") 'o id do parceiro
redir = request("redir") 'a pagina para redirecionar
tipo = request("tipo") 'o tipo do parceiro
soEndereco = request("soEndereco") 'se só cadastrao endereco e depois vai direto para o redir. se vazio, vai para o cadastro de contato

sub Impressao
end sub

sub auxiliar
end sub


sub principal
	dim Logradouro, numero, complemento, cep, bairro, cidade, estado, pais
	
	
	'caso a primeira vez que passa e ação igual a editar
	if acao = "edit" and request.Form("isPostBack") = "" then
		SQL = "Select * from endereco where id=" & id
		'response.Write(sql)
		
		set rs = conn.execute(SQL)
		
		if not rs.eof then
			logradouro = rs("logradouro")
			numero = rs("numero")
			complemento = rs("complemento")
			cep = rs("cep")
			bairro = rs("bairro")
			cidade = rs("cidade")
			estado = rs("estado")
			pais = rs("pais")											
		end if
		rs.close
		set rs = nothing
	elseif request.Form("isPostBack") <> "" then
		logradouro = request.Form("logradouro")
		numero = request.Form("numero")
		complemento = request.Form("complemento")
		cep = request.Form("cep")
		bairro = request.Form("bairro")
		cidade = request.Form("cidade")
		estado = request.Form("estado")
		pais = request.Form("pais")



		
		if logradouro = "" then erro = erro & "Logradouro inválido. "
		if estado = "" then erro = erro & "UF inválida. "
		
		if erro = "" then
			logradouro = replace(logradouro, "'", "''")
			numero = replace(numero, "'", "''")
			complemento = replace(complemento, "'", "''")
			cep = replace(cep, "'", "''")
			bairro = replace(bairro, "'", "''")
			cidade = replace(cidade, "'", "''")
			estado = replace(estado, "'", "''")
			pais = replace(pais, "'", "''")
			
			select case acao
			case "add"
				
				SQL = "INSERT INTO Endereco " & _
						"(idParceiro,  " & _
						"logradouro,  " & _						
						"numero, " & _
						"complemento,  " & _
						"cep, " & _
						"bairro, " & _
						"cidade, " & _
						"estado, " & _												
						"pais) " & _
						"VALUES " & _
						"(" & idParceiro & ", " & _
						"'" & logradouro & "', " & _
						"'" & numero & "', " & _
						"'" & complemento & "', " & _
						"'" & cep & "', " & _
						"'" & bairro & "', " & _
						"'" & cidade & "', " & _
						"'" & estado & "', " & _						
						"'" & pais & "')"
						
				conn.execute(SQL)
				

			case "edit"
			
				SQL = "UPDATE Endereco " & _
						   "SET logradouro = '" & logradouro & "', " & _
							  "numero = '" & numero & "', " & _
							  "complemento = '" & complemento & "', " & _
							  "cep = '" & cep & "', " & _
							  "bairro = '" & bairro & "', " & _
							  "cidade = '" & cidade & "', " & _
							  "pais = '" & pais & "' " & _
						 "WHERE id=" & id
				conn.execute(SQL)
				 
				
			end select
			
			
			if soEndereco = "" then
				'se novo, obrigatoriamente cadastra um contato
				response.Redirect("cadContato.asp?redir=" & server.URLEncode(Redir) & "&tipo=" & tipo & "&idParceiro=" & idParceiro & "&acao=" & acao)
		
			else
				'response.Redirect(redir & "?tipo=" & tipo & "&idParceiro=" & idParceiro)
				response.Redirect(redir)
			end if
			
			
		end if
		
	end if
%>
	<form action="" method="post" enctype="application/x-www-form-urlencoded" name="form1" target="_self">
		<input type="hidden" name="acao" value="<%=acao%>">
		<input type="hidden" name="id" value="<%=id%>">
		<input type="hidden" name="idParceiro" value="<%=idParceiro%>">		
		<input type="hidden" name="redir" value="<%=redir%>">		
		<input type="hidden" name="tipo" value="<%=tipo%>">				
		<input type="hidden" name="soEndereco" value="<%=soEndereco%>">						
		<input type="hidden" name="isPostBack" value="1">

	  <label class="titulo">Cadastro de Endereço</label><br>
	  <label class="subtitulo">Informe o Endereço</label><br>
	  <br>
	  <table width="100%"  border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td><label class="cabecalho">Logradouro</label>
            <br>
            <input name="Logradouro" type="text" class="caixa" id="Logradouro" size="40" maxlength="255" value="<%=logradouro%>"></td>
          <td><label class="cabecalho">Num.</label>
            <br>
            <input name="numero" type="text" class="caixa" id="numero2" size="5" maxlength="6" value="<%=numero%>"></td>
          <td><label class="cabecalho">Compl.</label>
            <br>
            <input name="complemento" type="text" class="caixa" id="complemento2" size="10" maxlength="100" value="<%=complemento%>"></td>
        </tr>
        <tr>
          <td><label class="cabecalho">CEP</label>
            <br>
            <input name="cep" type="text" class="caixa" id="cep2" size="10" maxlength="8" value="<%=cep%>"></td>
          <td><label class="cabecalho">Bairro</label>
            <br>
            <input name="bairro" type="text" class="caixa" id="bairro2" size="20" maxlength="100" value="<%=bairro%>"></td>
          <td><label class="cabecalho">Cidade</label>
            <br>
            <input name="cidade" type="text" class="caixa" id="cidade2" size="20" maxlength="100" value="<%=cidade%>"></td>
        </tr>
        <tr>
          <td><label class="cabecalho">UF</label>
            <br>
            <input name="estado" type="text" class="caixa" id="estado3" size="2" maxlength="2" value="<%=estado%>"></td>
          <td><label class="cabecalho">País</label>
            <br>
            <input name="pais" type="text" class="caixa" id="pais2" size="20" maxlength="50" value="<%=pais%>"></td>
          <td>&nbsp;</td>
        </tr>
      </table>
	  <br>
	  

	  
	  
	  
	  
	  <label class="cabecalho"></label>

	  <input name="btnOK" type="submit" class="botao" id="btnOK" value="OK">
	  <input name="btnCancelar" type="button" class="botao" id="btnCancelar" value="Cancelar" onClick="window.location.href='<%=redir%>'">
	  <P class="erro">
		&nbsp;<%=erro%>
	  </P>
	</form>
<%end sub


%>

<!-- #include file="../layout/layout.asp"-->
<%call FechaBanco%>