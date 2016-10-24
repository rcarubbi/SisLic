<!--#include file="../Dados/DBfuncoes.asp"-->
<!--#include file="../funcoes/funcoes.asp"-->
<!--#include file="../funcoes/funNumExtenso.asp"-->
<!--#include file="../funcoes/gridgeral.asp"-->
<%call abreBanco%>
<%
dim acao, idProcesso

acao = request.Form("acao")
idProcesso = request.QueryString("idprocesso")
if idProcesso = "" then response.Redirect("./")

if request("isPostBack") then
	' voltar para a tela anterior
	response.redirect("PropostaLista.asp")
end if

sub principal
	dim sql, rs, sqlitens, rsitens, cor ,vcor, cont, ImpRowsSpan, qtdRowSpan
	
	sql = "Select edi.numProcesso, edi.numEdital, edt.descricao as EditalTipo, replace(pro.outrasInf,char(13),'<BR>') as outrasInf, " & _
		  "replace(par.nome,char(13),'<br>') as cliente, replace(pro.InfCabecalho,char(13),'<BR>') as infCabecalho, convert(nvarchar,pro.dataProposta,103) as dataProposta, pr.cpf, pr.rg, pr.nome, pr.cargo " & _ 
		  " from proposta pro inner join edital edi on pro.idprocesso = edi.idprocesso inner join editalTipo edt on edi.ideditaltipo = edt.id inner join parceiro par on edi.idparceiro = par.id inner join propostaResponsavel pr on pro.cpfResponsavel = pr.cpf where edi.idprocesso =" & idProcesso
	
	set rs = conn.execute(Sql)
	
    sqlitens = "Select ipr.numItem, ipr.especificacao, ipr.unidade, ipr.quantidade, " & _
		  "ipr.valorUnitario, ipr.marca from propostaitem ipr Where idProcesso = " & idprocesso & " order by ipr.numitem"
	
	set rsitens = conn.execute(sqlitens)
	%><Table border="0" width="100%" cellpadding="0" cellpadding="0" height="100%">
	<thead style="display:table-header-group ">
	<tr><!-- cabecalho da proposta -->
		<td colspan="7" align="center" class="titulo" style="border-bottom-Style:solid;border-bottom-width:1px;border-bottom-color:#000000;">
		<%=application("par_2")%>
		</td>
	</tr>
	</thead>
	<tbody style="display:table-row; ">
	<tr><!-- corpo da proposta -->
	<td colspan="7" class="cabecalho" style="height:20px;">À <BR>
				    <%=rs("Cliente")%><BR><BR>
					REF.: <%=rs("EditalTipo")%> Nº. <%=rs("numEdital")%><BR>
					&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Processo Nº. <%=rs("numProcesso")%>
	</td>
	</tr>
	<%if not isnull(rs("infCabecalho")) and trim(rs("infCabecalho"))<>"" and not isempty(trim(rs("infCabecalho"))) then%>
	
	<tr>
			<td colspan="7"  valign="top" class="cabecalho">
				<BR>&nbsp;<%=rs("infCabecalho")%>
			</td></tr>
	<tr>
	<%end if%>	
	<td colspan="7" align="center" class="subtitulo">Proposta Comercial</td>
	</tr><tr>
		<td valign="top">
	<table border="1" width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;"><!-- Itens da proposta -->
	<tr>
	<td bgcolor="f5f5f5" class="cabecalho" align="center">Item</td>
	<td bgcolor="f5f5f5" class="cabecalho" align="center" width="40%">Descrição</td>
	<td bgcolor="f5f5f5" class="cabecalho" align="center">Unid.</td>
	<td bgcolor="f5f5f5" class="cabecalho" align="center">Qtde.</td>
	<td bgcolor="f5f5f5" class="cabecalho" align="center">Preço Unit.</td>
	<td bgcolor="f5f5f5" class="cabecalho" align="center">Preço Total</td>
	<td bgcolor="f5f5f5" class="cabecalho" align="center">Marca</td>	
	</tr>
	<%cont = 1%>
	<%while not rsitens.eof%>
		<%
		
		
		if cont = 1 then
			ImpRowSpan = true
			QtdRowspan = getQtdItensProp(rsitens("numItem"))
		else
			ImpRowSpan = false
		end if
		
		if cont = QtdRowspan then
				cont = 0
		end if
		
		%>
		
		<tr style="padding:5px;">
		<%if improwspan then%>
		<td bgcolor="<%=cor%>" class="texto" rowspan="<%=qtdRowspan%>"><%=rsitens("numItem")%></td>
		<%end if%>
		<td bgcolor="<%=cor%>" class="texto" width="40%"><%=rsitens("especificacao")%></td>
		<td bgcolor="<%=cor%>" class="texto" align="center"><%=rsitens("unidade")%></td>
		<td bgcolor="<%=cor%>" class="texto" align="center"><%=rsitens("quantidade")%></td>
		<td bgcolor="<%=cor%>" class="texto" ><div align="center"><%=formatcurrency(rsitens("ValorUnitario"))%></div><BR><div align="left">(<%=extenso(rsItens("valorunitario"))%>)</div></td>
		<td bgcolor="<%=cor%>" class="texto" align="center"><div align="center"><%=formatcurrency(rsitens("quantidade") * rsitens("valorUnitario"))%></div><BR><div align="left">(<%=extenso(rsitens("quantidade") * rsitens("valorUnitario"))%>)</div></td>
		<td bgcolor="<%=cor%>" class="texto" align="center"><%=rsitens("marca")%></td>	
		</tr>
				<%rsitens.movenext
				cont = cont +1
				
			wend%>
			<tr>
			<td colspan="7">
			<label class="cabecalho">&nbsp;Outras Informações</label><BR>
			<label class="texto">&nbsp;<%=rs("outrasInf")%></label>
			</td></tr>
			</table>
		</td>
		</tr>
				<tr><Td>&nbsp;</Td></tr>
		<tr><td class="cabecalho"><%="São Paulo, " & day(rs("dataProposta")) & " de " & monthname(month(rs("dataProposta"))) & " de " & year(rs("dataProposta"))%></td></tr>
				<tr><Td>&nbsp;</Td></tr>
		<tr><Td>&nbsp;</Td></tr>
		<tr><Td class="cabecalho">
		<div style="width:300px;border-top-style:solid;border-top-width:1px;border-top-color:#000000;">
		<div align="center" style="width:300px;">
		<%=rs("nome")%> - <%=rs("cargo")%><BR>CPF: <%=formataCPF(rs("cpf"))%><BR>R.G: <%=rs("rg")%>
		</div>
		</div>
		
		</Td></tr>
		<tr><Td>&nbsp;</Td></tr>
		</tbody>
	<tfoot style="display:table-footer-group ">
	<tr><td colspan="7" class="cabecalho" style="border-top-style:solid;border-top-width:1px;border-top-color:#000000;" align="center">
	<%=Application("par_6")%> - Fone/Fax: <%=Application("par_7")%> - CEP <%=Application("par_8")%>.<BR>
	C.N.P.J. <%=application("par_9")%> - Inscrição Estadual <%=application("par_10")%> - <%=application("par_11")%> - <%=application("par_12")%><BR>
	<%=Application("par_13")%> - Fone/Fax: <%=Application("par_14")%> - CEP <%=Application("par_15")%>.<BR>
	C.N.P.J. <%=application("par_16")%> - Inscrição Estadual <%=application("par_17")%> - <%=application("par_18")%> - <%=application("par_19")%>
	
	</td></tr>
	</tfoot>
	</Table><%
end sub

sub Impressao
	response.write request.ServerVariables("SCRIPT_NAME") & "?imp=1&idprocesso=" & idprocesso & "&data=" & now
end sub


sub auxiliar
%><input type="button" value="Imprimir" name="cmdOk"  onClick="window.print();" class="botao"><BR><BR>
<input type="button" value="Voltar" name="cmdVoltar"  onClick="window.location.href='propostaLista.asp';" class="botao">
<%
end sub

%>

<%if request.QueryString("imp") = "" then%>
	<!--#include file="../layout/layout.asp"-->
<%else%>
	<!--#include file="../layout/layout_print.asp"-->
<%end if%>
<%function getQtdItensProp(numItem)
	dim retorno, sql, rs
	sql = "Select count(numItem) as qtdItens from PropostaItem where numItem = '" & numItem & "' and idProcesso=" & idProcesso
	
	set rs = conn.execute(sql)
	if not rs.eof then
		retorno = rs("qtdItens")
	else
		retorno = 0
	end if
	rs.close  : set rs = nothing
	getQtdItensProp = retorno
end function
%>
 
