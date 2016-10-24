<!--#include file="../Dados/DBfuncoes.asp"-->
<!--#include file="funcoes.asp"-->

<%call abreBanco%>
<%

dim id, tipo, erro, acao
dim result

id = request("idParceiro")
tipo = request("tipo")
acao = request("acao")

select case acao
	case "delEndereco"
		result = deletaEndereco
		if result = 0 then 
			erro = "Erro ao deletar endereço"
		end if
	case "delContato"
		result = deletaContato
		if result = 0 then
			erro = "Erro ao deletar contato"
		end if
end select

sub principal
	
	if tipo = "0" then
	%>
		<label class="titulo">Detalhes do Cliente</label><br>
		<label class="subtitulo">Segue abaixo os dados do cliente</label>
	<%
	else
	%>
		<label class="titulo">Detalhes do Fornecedor</label><br>
		<label class="subtitulo">Segue abaixo os dados do fornecedor</label>
	<%
	end if
	
	
	'###################################################
	' exibindo os dados do cliente
	
	SQL = "SELECT * from parceiro where id=" & id
	set rs = conn.execute(SQL)
	
	%>
	<br><br>
	<%if tipo = "0" then%>
		<span class="subtitulo">Cliente</span>
	<%else%>
		<span class="subtitulo">Fornecedor</span>			
	<%end if%>
	<a href="<%=pagCliente%>?redir=<%=server.URLEncode(pagDetalhes & "?tipo=" & tipo & "&idParceiro=" & id)%>&tipo=<%=tipo%>&idParceiro=<%=id%>&acao=edit" class="texto">Editar</a> 
	
	<%
	if not rs.eof then
		%>
	
			<table width="100%"  border="1" cellspacing="3" cellpadding="0" style="border-collapse:collapse ">
	
			  <tr>
				<td><span class="cabecalho">CNPJ</span><br>
				  <span class="texto"><%=formataCNPJ(rs("cnpj"))%></span></td>
			  </tr>
			  <tr>
				<td><span class="cabecalho">Nome</span><br>
				 <span class="texto"><%=replace(rs("nome"),chr(13),"<BR>")%></span>
				 </td>
			  </tr>
			  <tr>
				<td><span class="cabecalho">Site</span><br>
				  <span class="texto"><%=rs("site")%></span></td>
			  </tr>
			  <tr>
				<td><span class="cabecalho">Data de Cadastro</span><br>
				  <span class="texto"><%=formata_data(rs("dataCadastro"))%></span></td>
			  </tr>
			</table>
	<%
	else
		%><span class="erro">Nenhum registro encontrado</span>
			<p><a href="#" class="texto" onClick="histoy.back(-1);">Voltar</a></p>
		<%
	end if
	
	rs.close
	set rs = nothing
	' fim dados cliente
	'##################################################
	%>		
	<br>
	
	<%
	'##################################################
	' endereco do cliente
	
	
	SQL = "Select * from endereco where idParceiro=" & id
	set rs = conn.execute(SQL)
	%>
	<span class="subtitulo">Endereços</span> - <a href="<%=pagEndereco%>?redir=<%=server.URLEncode(pagDetalhes & "?tipo=" & tipo & "&idParceiro=" & id)%>&tipo=<%=tipo%>&idParceiro=<%=id%>&id=&acao=add&soEndereco=1" class="texto">Novo</a> 
	<table width="100%"  border="1" cellspacing="3" cellpadding="0" style="border-collapse:collapse ">
		
	<%
	while not rs.eof
	%>
			<tr>
				<td>
					<table width="100%"  border="0" cellspacing="0" cellpadding="0">
					
							<tr>
							  <td><label class="cabecalho">Logradouro</label><br>
								<span class="texto"><%=rs("logradouro")%></span></td>
							  <td><label class="cabecalho">Num.</label><br>
								<span class="texto"><%=rs("numero")%></span></td>
							  <td><label class="cabecalho">Compl.</label><br>
								<span class="texto"><%=rs("complemento")%></span></td>
							</tr>
							<tr>
							  <td><label class="cabecalho">CEP</label><br>
								<span class="texto"><%=formataCep(rs("cep"))%></span></td>
							  <td><label class="cabecalho">Bairro</label><br>
								<span class="texto"><%=rs("bairro")%></span></td>
							  <td><label class="cabecalho">Cidade</label><br>
								<span class="texto"><%=rs("cidade")%></span></td>
							</tr>
							<tr>
							  <td><label class="cabecalho">UF</label><br>
								<span class="texto"><%=rs("estado")%></span></td>
							  <td colspan="2"><label class="cabecalho">País</label><br>
								<span class="texto"><%=rs("pais")%></span></td>
							</tr>
					</table>
				
				
				</td>
				
				<td width="100">
					<a href="<%=pagEndereco%>?redir=<%=server.URLEncode(pagDetalhes & "?tipo=" & tipo & "&idParceiro=" & id)%>&idParceiro=<%=id%>&id=<%=rs("id")%>&acao=edit&soEndereco=1" class="texto">Editar</a>
					<a href="<%=pagDetalhes%>?tipo=<%=tipo%>&idParceiro=<%=id%>&idEndereco=<%=rs("id")%>&acao=delEndereco" class="texto">Excluir</a>
				</td>
			</tr>
	
	<%	rs.movenext
	wend
	rs.close
	set rs = nothing
	%>	</table>
	<span class="erro"><%=erro%></span>
	<%
	'fim endereco do cliente
	'#########################################################################
	
	
	
	'#########################################################################
	'dados dos contatos
	sql = "Select * from contato where idParceiro=" & id
	set rs = conn.execute(SQL)
	
	
	%><br>
	<span class="subtitulo">Contatos</span> - <a href="<%=pagContato%>?redir=<%=server.URLEncode(pagDetalhes & "?tipo=" & tipo & "&idParceiro=" & id)%>&tipo=<%=tipo%>&idParceiro=<%=id%>&id=&acao=add" class="texto">Novo</a>
	  <table width="100%"  border="1" cellspacing="3" cellpadding="0" style="border-collapse:collapse ">
	
	<%
	while not rs.eof
	%>
	
	
			<tr>
			  <td>
				  <table width="100%"  border="0" cellspacing="0" cellpadding="0">
					<tr>
					  <td><span class="cabecalho">Nome</span><br>
							<span class="texto"><%=rs("nome")%></span></td>
					</tr>
					<tr>
					  <td><span class="cabecalho">Telefone</span><br>
							  <span class="texto"><%=rs("telefone")%></span></td>
					</tr>
					<tr>
					  <td><span class="cabecalho">Celular</span><br>
							<span class="texto"><%=rs("celular")%></span></td>
					</tr>
					<tr>
					  <td><span class="cabecalho">E-Mail</span><br>
							 <span class="texto"><%=rs("email")%></span> </td>
					</tr>
				  </table>            </td>
			  <td width="100"> <a href="<%=pagContato%>?redir=<%=server.URLEncode(pagDetalhes & "?tipo=" & tipo & "&idParceiro=" & id)%>&tipo=<%=tipo%>&idParceiro=<%=id%>&id=<%=rs("id")%>&acao=edit" class="texto">Editar</a> 
			  <a href="<%=pagDetalhes%>?tipo=<%=tipo%>&idParceiro=<%=id%>&idContato=<%=rs("id")%>&acao=delContato" class="texto">Excluir</a> </td>
			</tr>

		 
	
	<%	rs.movenext
	wend
	
	rs.close
	set rs = nothing
	%>
	</table>
	<p align="center"><input type="button" onclick="window.location.href='<%=iif(instr(request.ServerVariables("SCRIPT_NAME"),"cliente") > 0,"clienteLista.asp","fornecedorLista.asp") %>';" value="voltar" class="botao"></p>
	<span class="erro"><%=erro%></span>
	<%
	' fim dados do contato
	'##########################################################
end sub

sub Impressao
end sub

sub auxiliar
	
end sub


function DeletaEndereco
	dim idEnd
	idEnd = request.querystring("idEndereco")
	if not isNumeric(idEnd) then
		DeletaEndereco = 0 'erro
		exit function
	end if
	
	sql = "DELETE from endereco where id=" & idEnd
	conn.execute(SQL)
	
	DeletaEndereco = 1
end function
	
function DeletaContato
	dim idCont
	idCont = request.Querystring("idContato")
	if not isNumeric(idCont) and idCont <> "" then
		DeletaContato = 0
		exit function
	end if
	
	sql = "DELETE from contato where id=" & idCont

	conn.execute(SQL)

	DeletaContato = 1
end function

%>
    
<!--#include file="../layout/layout.asp"-->

