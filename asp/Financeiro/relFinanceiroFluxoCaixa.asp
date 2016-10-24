<!--#include file="../Dados/DBfuncoes.asp"-->
<!--#include file="../funcoes/funcoes.asp"-->
<%call abreBanco%>

<%dim dataInic, dataFinal, mesAnoAnterior, sql, rs, vcor, cor

dataInic = recebe_datador("dataInic",0)
dataFinal = recebe_datador("dataFinal",0)
if request("isPostBack") then
	if not isdate(dataInic) then
		erro = "Data Inicial Inválida. "
	end if
	if not isdate(dataFinal) then
		erro = erro & "Data Final Inválida. "
	end if
	if isdate(dataInic) and isdate(dataFinal) then
		if cdate(dataInic) > cdate(dataFinal) then
				erro = "A data inicial deve ser anterior a data final. "
		end if
	end if
end if
sub principal
	dim saldoAnterior, subtotal
	if request("isPostBack") then
		 
		if erro = "" then
		 
		 
			saldoAnterior = getSaldoAnterior(dataInic)
			
			sql = "Select f.dataLancamento, f.referente, f.valor, p.tipo, p.descricao as conta " & _
				  "from fluxoCaixa f inner join planocontas p on p.codconta = f.codplanoconta where f.cancelado = 0 " & _
				  "and f.dataLancamento between '" & dataInic & " 00:00:00' and '" & dataFinal & " 23:59:59'"
			set rs = conn.execute(sql)
			%>
				<label class="titulo">Relatório de Fluxo de Caixa</label><br>
				<label class="subtitulo">Período de <%=dataInic%> até <%=dataFinal%></label><BR>
				
				
				
				
				
				
				<table style="border-collapse:collapse;" border="1" width="100%">
					
					
					<thead>
					</tr>
						<tr class="cabecalho" style="background-color:#f5f5f5;">
							<th width="10%" align="center">Data</th>
							<th width="">Referente à</th>
							<th width="">Conta</th>
							<th width="" align="right">Valor Crédito</th>
							<th width="" align="right">Valor Débito</th>
							<th width="" align="right">Sub-Total</th>
						</tr>
					
					</thead>
					<tbody>
					<tr>
					<td colspan="5" align="right" bgcolor="<%if saldoAnterior >= 0 then response.write "#A6D9FF" else response.write "#FFA8A8" end if%>" class="cabecalho">
					Saldo Anterior:</td>
					
					<td align="right" bgcolor="<%if saldoAnterior >= 0 then response.write "#A6D9FF" else response.write "#FFA8A8" end if%>" class="cabecalho"><%=formatCurrency(saldoAnterior)%></td> 
				</td>
					</tr>
					<%subtotal = saldoAnterior%>
			<%while not rs.eof %>		
			<%if vcor then
				cor = "#ffffff"
			else
				cor= "#ffffcc"
			end if%>
					<tr bgcolor="<%=cor%>">
						<td class="texto" align="center"><%=Formata_data_abrev(rs("dataLancamento"))%></td>
						<td class="texto" style="padding-left:5px;"><%=rs("referente")%></td>
						<td class="texto" style="padding-left:5px;"><%=rs("conta")%></td>
						<td class="texto" align="right"><%if rs("tipo") = false then response.write formatcurrency(rs("valor"))%></td>
						<td class="texto" align="right"><%if rs("tipo") = true then response.write formatcurrency(rs("valor"))%></td>
						<%subtotal = subtotal + iif(rs("tipo")=0,rs("valor"),rs("valor")*(-1))%>
						<td class="texto" align="right"><%=formatCurrency(SubTotal)%></td>
					</tr>
				
				
			<%	vcor = not vcor
				rs.movenext
			wend	
			%>
			<tr>
					<td colspan="5" align="right" bgcolor="<%if SubTotal >= 0 then response.write "#A6D9FF" else response.write "#FFA8A8" end if%>" class="cabecalho">
					Saldo Final:</td>
					
					<td align="right" bgcolor="<%if SubTotal >= 0 then response.write "#A6D9FF" else response.write "#FFA8A8" end if%>" class="cabecalho"><%=formatCurrency(SubTotal)%></td> 
				</td>
					</tr>
			</tbody>
			</table><%
		end if
	end if
end sub
%>

<%sub Impressao
	response.write request.ServerVariables("SCRIPT_NAME") & "?imp=1&ispostback=1&dataInic=" & dataInic & "&dataFinal=" & dataFinal
end sub

sub auxiliar%>
<form name="filtros" method="post" action="">
<input type="hidden" value="1" name="isPostBack">
<table>
	<tr>
		<td><label class="subtitulo">Período</label><br>
			
			<label class="cabecalho">Data Inicial</label><br>
			<%Datador "DataInic",0,year(date)-15,year(date),DataInic%>
		</td>
	</tr>
	<tr>
		<td><label class="cabecalho">Data Final</label><br>
			<%Datador "DataFinal",0,year(date)-15,year(date),DataFinal%>
		</td>
		
	</tr>
	
	<tr>
		<td><br>
		<input type="submit" value="Ok" name="cmdOk" class="botao">&nbsp;
		<input type="button" value="Imprimir" name="cmdOk"  onClick="window.print();" class="botao">
		<p class="erro"> &nbsp;
			<%=erro%>
	 	 </p>
		</td>
	</tr>
</table>
</form>
<%end sub%>
<%
function getSaldoAnterior(dataReferente)
	dim retorno, sql, rs, mesAnterior, anoAnterior
	
	mesAnterior = month(dateadd("m",-1,datareferente))
	anoAnterior = year(dateadd("m",-1,datareferente))
	sql = "Select valor from fechamentoCaixa where mes=" & mesAnterior & " and ano=" & anoAnterior
	set rs = conn.execute(sql)
	if not rs.eof then
		if not isnull(rs("valor")) then
			retorno = rs("valor")
		else
			retorno = 0
		end if
	else
		retorno = 0
	end if
	rs.close : set rs = nothing
	getSaldoAnterior = retorno

end function

%>

<%if request.QueryString("imp") = "" then%>
<!--#include file="../layout/layout.asp"-->
<%else%>
<!--#include file="../layout/layout_print.asp"-->
<%end if%>

<%call fechaBanco%>