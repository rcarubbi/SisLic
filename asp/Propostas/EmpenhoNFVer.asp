<!--#include file="../Dados/DBfuncoes.asp"-->
<!--#include file="../funcoes/funcoes.asp"-->
<!--#include file="../funcoes/gridgeral.asp"-->
<%call abreBanco%>
<%
dim idProcesso, acao, numNF
dim sqldel
idProcesso = trim(replace(request.QueryString("idprocesso"),"'",""))
acao = trim(replace(request.QueryString("acao"),"'",""))

	

sub principal
	dim sql, rs
	sql = "SELECT empenhoNF.*, edital.NumProcesso from empenhoNF inner join edital on empenhoNF.idProcesso=edital.idProcesso where empenhoNF.idProcesso = " & idProcesso
	set rs = conn.execute(sql)
	
	%>
	<label class="titulo">Notas Fiscais</label><BR>
	<label class="subtitulo">Vizualizar Notas Fiscais</label><BR>
<%	
	if rs.eof then
		%><p align="center"><label class="erro">Não existem notas fiscais referentes a este processo.</label></p><%
	else
		%>
		<label class="subtitulo">ref. processo nº.: <%=rs("NumProcesso")%></label>
		<%
	end if
	while not rs.eof
		%>
		<table width="100%" cellpadding="0" cellspacing="0" border="1" style="border-collapse:collapse;">
		<tr>
		<td width="40px" rowspan="2" align="center"><a href="cancelamentoNF.asp?numNF=<%=rs("numNF")%>&idprocesso=<%=idProcesso%>" class="texto">Excluir</a></td>
		<TD width="30%">
		<label class="cabecalho">&nbsp;Núm NF.</label><BR>
		<label class="texto">&nbsp;<%=rs("numNF")%></label><BR>
		</TD>
		<Td width='30%'>
		<label class="cabecalho">&nbsp;Descrição</label><BR>
		<label class="texto">&nbsp;<%=rs("descricao")%></label><BR>
		</Td>
		<Td width="120px">
		<label class="cabecalho">&nbsp;Criada em</label><BR>
		<label class="texto">&nbsp;<%=formata_Data(rs("data"))%></label><BR>
		</Td>
		</tr>
		<tr>
		<td>
		<label class="cabecalho">&nbsp;Observações</label><BR>
		<label class="texto">&nbsp;<%=rs("observacoes")%></label><BR>
		</td>
		<td>
		<label class="cabecalho">&nbsp;Valor Total</label><BR>
		<label class="texto">&nbsp;<%=formatCurrency(rs("valorTotal"),2)%></label><BR>
		</td>
		<td>
		<label class="cabecalho">&nbsp;Valor dos Impostos</label><BR>
		<label class="texto">&nbsp;<%=formatCurrency(rs("impostos"),2)%></label><BR>
		</td>

		</tr>
		</table>

		<%
		rs.movenext
		if not rs.eof then response.write "<hr width='100%'>"
	wend

	
	rs.close : set rs = nothing
	%><p align="center"><input type="button" value="voltar" class="botao" onClick="window.location.href='empenhoLista.asp?situacao=1';"></p><%
end sub

sub Impressao
end sub


sub auxiliar
end sub

%><!--#include file="../layout/layout.asp"-->