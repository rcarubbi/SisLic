<!--#include file="../Dados/DBfuncoes.asp"-->
<!--#include file="../funcoes/funcoes.asp"-->
<%call abreBanco%>

<%
dim dtIni
dim dtFim

if request.QueryString("imp") = "" then
	dtIni = recebe_datador("dtIni",0)
	dtFim = recebe_datador("dtFim",0)

	
else
	'para não alterar o arquivo de funcoes
	'foi feita essa condicional
	'para não precisar fazer essa condicional, teria que alterar
	'o arquivo de funcoes, a funcao recebe_datador para receber tanto em form como em query
	dtIni = request.QueryString("dtIni")
	dtFim = request.QueryString("dtFim")

end if
if not isDate(dtIni) then
	dtini = dateadd("m",-1,date)
	dtfim = date
end if

if cdate(dtFim) < cdate(dtIni) then dtfim = dtini

sub principal
	dim sql
	dim rs
	dim cor
	dim tot
	'seleciona os empenho
	sql = "SELECT  Empenho.idProcesso, Edital.NumProcesso, Empenho.DataEmpenho, " & _
			"(SELECT sum(valorTotal) as saldoEmpenho from empenhoNF WHERE empenhoNF.idProcesso = empenho.idProcesso) as ValorTotalNota" & _
			", (SELECT convert(smallmoney, sum(valorTotal-impostos) - sum((valorTotal * " & replace(application("par_25"), ",", ".") & ") / 100)) as saldoEmpenho from empenhoNF WHERE " & _
			"empenhoNF.idProcesso = empenho.idProcesso) - " & _
			"isNull((SELECT sum(vlTotal - vlImpostos) as saldoPedido from " & _
			"pedido where idProcesso=empenho.idProcesso and recebido=1 and cancelado=0),0) as saldo " & _
			"FROM Empenho INNER JOIN " & _
			"Edital ON Empenho.idProcesso = Edital.IdProcesso " & _
			"WHERE empenho.dataEmpenho BETWEEN " & formata_data_sql(dtIni) & " AND " & formata_data_sql(dtFim & " 23:59:59") 
'response.Write(sql)
'response.End()
	set rs = conn.execute(SQL)
	
	%>
	<p>
	<label class="titulo">Saldo geral (por Empenho)</label><br>
	<label class="subtitulo">Periodo entre <%=(dtIni)%> e <%=(dtFim)%></label>
	</p>
	<table width="100%" cellpadding="3" cellspacing="0" border="1" style="border-collapse:collapse ">
		<tr bgcolor="#f5f5f5"> 
			<td class="cabecalho">Nº Processo</td>
			<td class="cabecalho">Data</td>
			<td align="right" class="cabecalho">Vl. Total</td>
			<td align="right" class="cabecalho">Saldo (créditos/débitos)</td>									
		</tr>
	<%
	cor =false
	tot = 0
	while not rs.eof
		if not isNull(rs("saldo")) then
			%>
			<tr bgcolor="<%if cor then response.Write("#ffffcc")%>">
				<td ><a href="EmpenhoSaldoDetalhe.asp?idProcesso=<%=rs("idProcesso")%>" class="texto"><%=rs("numProcesso")%></a></td>
				<td class="texto"><%=formata_campo(rs("dataEmpenho").value)%></td>
				<td align="right" class="texto"><%=formata_campo(rs("valortotalnota").value)%></td>
				<td align="right" class="texto"><%=formata_campo(rs("saldo").value)%></td>
			</tr>		
			<%
			tot = tot + rs("saldo")
			cor = not cor
		end if
		
		rs.movenext
	wend
	%>
		<tr bgcolor="#f5f5f5">
			<td align="right" colspan="4" class="cabecalho">Saldo Total: <%=formata_campo(tot)%></td>
		
		</tr>
	</table>
	<%
	rs.close
	set rs = nothing
	
end sub
%>

<%sub Impressao
	response.write request.ServerVariables("SCRIPT_NAME") & "?imp=1&dtIni=" & dtIni & "&dtFim=" & dtFim
end sub

sub auxiliar
%>
	<form name="form1" action="" method="post" enctype="application/x-www-form-urlencoded">
		<input type="hidden" name="filtros" value="1">
		
		<label class="cabecalho">Data inicial</label>
		<br>
		<%
		Datador "dtIni",0,1980, year(now),dtIni
		%>
		<br>
		<label class="cabecalho">Data final</label>
		<br>
		<%
		Datador "dtFim",0,1980, year(now),dtFim	
		%>
		<p>
		<input type="submit" name="OK" value="OK" class="botao"> &nbsp;&nbsp;
		<input type="button" onClick="window.print();" value="Imprimir" class="botao">

		</p>
	</form>
	<%
end sub

if request.QueryString("imp") = "" then%>
<!--#include file="../layout/layout.asp"-->
<%else%>
<!--#include file="../layout/layout_print.asp"-->
<%end if%>

<%call fechaBanco%>
