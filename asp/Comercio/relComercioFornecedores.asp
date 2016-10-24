<!--#include file="../Dados/DBfuncoes.asp"-->
<!--#include file="../funcoes/funcoes.asp"-->
<%call abreBanco%>

<%dim nomeFornecedor, Situacao, sql, rs, vcor, cor
nomeFornecedor = trim(replace(request("nomeFornecedor"),"'","''"))
Situacao = request("situacao")
sub principal
	
	if request("isPostBack") then
		
		
		sql = "SELECT cnpj, nome, site, dataCadastro FROM parceiro WHERE Ativo = " & situacao & _
				" AND tipo = 1"	
				if nomeFornecedor<>"" then sql = sql & " AND nome like '" & nomeFornecedor & "%'"
				
					
		set rs = conn.execute(sql)
		%>
			<label class="titulo">Relatório de Fornecedores</label><br>
			<label class="subtitulo">Situação: <%=iif(situacao=0,"Deletados","Ativos")%></label><br>
			<label class="subtitulo">Pesquisar por: <%=iif(nomeFornecedor="","Todos","""" & nomeFornecedor & """")%></label>
			<table style="border-collapse:collapse;" border="1" width="100%">
				<thead >
					<tr class="cabecalho" style="background-color:#f5f5f5;">
						<th width="20%">C.N.P.J.</th>
						<th width="30%">Nome</th>
						<th width="25%">Site</th>
						<th width="25%">Data de Cadastro</th>
					</tr>
				</thead>
				<tbody>
		<%while not rs.eof %>		
		<%if vcor then
			cor = "#ffffff"
		else
			cor= "#ffffcc"
		end if%>
				<tr bgcolor="<%=cor%>">
					<td class="texto"><%=formataCNPJ(rs("cnpj"))%></td>
					<td class="texto"><%=replace(rs("nome"),chr(13),"<BR>")%></td>
					<td class="texto"><%=rs("site")%></td>
					<td class="texto"><%=rs("dataCadastro")%></td>
				</tr>
			
			
		<%	vcor = not vcor
			rs.movenext
		wend	
		%></tbody>
		</table><%
	end if
end sub
%>

<%sub Impressao
	response.write request.ServerVariables("SCRIPT_NAME") & "?nomeFornecedor=" & nomeFornecedor & "&situacao=" & situacao & "&imp=1&isPostBack=1&time=" &time
end sub

sub auxiliar%>
<form name="filtros" method="post" action="">
<input type="hidden" value="1" name="isPostBack">
<table>
	<tr>
		<td><label class="cabecalho">Situação</label><br>
			<%dim itemTexto(1), itemValor(1)
			itemTexto(0) = "Ativos"
			itemValor(0) = 1
			itemTexto(1) = "Deletados"
			itemValor(1) = 0
			Combo "situacao", itemTexto, itemValor, request.Form("situacao")%>
		</td>
	</tr>
	<tr>
		<td><label class="cabecalho">Pesquisar por nome</label><br>
		<input type="text" size="20" value="<%=nomeFornecedor%>" name="nomeFornecedor" class="caixa">
		</td>
		
	</tr>
	<tr>
		<td><br>
		<input type="submit" value="Ok" name="cmdOk" class="botao">&nbsp;
		<input type="button" value="Imprimir" name="cmdOk"  onClick="window.print();" class="botao">
		</td>
	</tr>
</table>
</form>
<%end sub%>
<%if request.QueryString("imp") = "" then%>
<!--#include file="../layout/layout.asp"-->
<%else%>
<!--#include file="../layout/layout_print.asp"-->
<%end if%>

<%call fechaBanco%>