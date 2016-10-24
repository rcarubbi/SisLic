<!--#include file="../Dados/DBfuncoes.asp"-->
<!--#include file="../funcoes/funcoes.asp"-->
<%call abreBanco%>

<%dim dataEmpInic, dataEmpFinal, sql, rs, vcor, cor, erro
dataEmpInic = recebe_datador("dataEmpInic",0)
dataEmpFinal = recebe_datador("dataEmpFinal",0)
if not isDate(dataEmpInic) then
	dataEmpInic = dateadd("m",-1,date)
	dataEmpFinal = date
end if

clienteid = request("clienteid")
if request("isPostBack") then
	if not isdate(dataEmpInic) then
		erro = "Data Inicial Inválida. "
	end if
	if not isdate(dataEmpFinal) then
		erro = erro & "Data Final Inválida. "
	end if
	if isdate(dataEmpInic) and isdate(dataEmpFinal) then
		if cdate(dataEmpInic) > cdate(dataEmpFinal) then
				erro = "A data inicial deve ser anterior a data final. "
		end if
	end if
end if
sub principal
	
	if request("isPostBack") then
			
		if erro = "" then
		
			sql = "Select emp.valortotal, emp.prazo, emp.validade, Edi.numProcesso from empenho emp " & _
				  "Inner join proposta pro on pro.idProcesso = emp.idProcesso " & _
				  "Inner join edital edi on edi.idProcesso = pro.idProcesso " & _
				  "where emp.DataEmpenho between '" & dataEmpInic & " 00:00:00' and '" & dataEmpFinal & " 23:59:59'"
				
				
			set rs = conn.execute(sql)
			%>
				<label class="titulo">Relatório de Empenhos</label><br>
				<label class="subtitulo">Empenhos de <%=dataEmpInic%> até <%=dataEmpFinal%></label><BR>
				<table style="border-collapse:collapse;" border="1" width="50%">
					<thead>
						<tr class="cabecalho" style="background-color:#f5f5f5;">
							<th width="">Núm. do Processo</th>
							<th width="" align="right">Valor do Empenho</th>
							<th width="25%" align="right">Prazo (dias)</th>
							<th width="15%">Validade</th>
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
						<td class="texto"><%=rs("numProcesso")%></td>
						<td class="texto" align="right"><%=formata_Campo(ccur(rs("valortotal")))%></td>
						<td class="texto" align="right"><%=rs("prazo")%></td>
						<td class="texto" align="right"><%=rs("validade")%></td>
						
					</tr>
				
				
			<%	vcor = not vcor
				rs.movenext
			wend	
			%></tbody>
			</table><%
		end if
	end if
end sub
%>

<%sub Impressao
	response.write request.ServerVariables("SCRIPT_NAME") & "?imp=1&ispostback=1&dataEmpInic=" & dataEmpInic & "&dataEmpFinal=" & dataEmpFinal
end sub

sub auxiliar%>
<form name="filtros" method="post" action="">
<input type="hidden" value="1" name="isPostBack">
<table cellpadding="0" cellpadding="0">
	<tr>
		<td>
			<label class="subtitulo">Período</label><br>
			<label class="cabecalho">Data Inicial</label><br>
			<%Datador "DataEmpInic",0,year(date)-15,year(date),DataEmpInic%>
		</td>
	</tr>
	<tr>
		<td><label class="cabecalho">Data Final</label><br>
			<%Datador "DataEmpFinal",0,year(date)-15,year(date),DataEmpFinal%>
		</td>
		
	</tr>
	
	<tr>
		<td valign="top">
		
		<br>
		<input type="submit" value="Ok" name="cmdOk" class="botao">&nbsp;
		<input type="button" value="Imprimir" name="cmdOk"  onClick="window.print();" class="botao">
		<p class="erro"> &nbsp;
			<%=erro%>
	 	 </p>

		 </td>
	</tr>
	<tr>
		
			
		
	
</table>
</form>
<%end sub%>

<%if request.QueryString("imp") = "" then%>
<!--#include file="../layout/layout.asp"-->
<%else%>
<!--#include file="../layout/layout_print.asp"-->
<%end if%>

<%call fechaBanco%>