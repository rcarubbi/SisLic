<!--#include file="../Dados/DBfuncoes.asp"-->
<!--#include file="../funcoes/funcoes.asp"-->
<%call abreBanco%>

<%dim sql, rs, vcor, cor
  dim ano, mes, valores, contas, ValoresAux, erro, total
  redim totais(11)
  redim contas(3,-1)
 ' matriz contas
 ' 0,x --> codigo da conta
 ' 1,x --> codigo da conta mae
 ' 2,x --> descricao
 ' 3,x --> espacos de itendação
 	if request("isPostBack") then
		
		ano = request("ano")
		if ano = "" then
			erro = "Não há lancamentos no plano de contas. "
		end if
	end if
sub principal

	dim ehRaiz
	if request("isPostBack") then
		
		
		
		if erro = "" then
			sql = "select codconta, descricao from planocontas where ativo = 1 and codcontasuperior is null"
			set rs = conn.execute(sql)
			while not rs.eof 
				
				addConta rs("codconta"), 0
				rs.movenext
			wend
			
				' exibe o relatorio
				%>
					<label class="titulo">Relatório de Lançamentos<br>
					<label class="subtitulo">Ano: <%=ano%></label><BR>
	
				<table style="border-collapse:collapse;" border="1" width="100%">
					<thead>
						<tr class="cabecalho">
						<td>Conta</td>
						<%for mes = 1 to 12 %>
						<td width="7%"><%=left(monthName(mes),3)%></td>
						<%next%>
						</tr>
					</thead>
					<tbody>
						
						<%for x=0 to ubound(contas,2)
							redim Valores(11)
							if vcor then
								cor = "#ffffff"
							else
								cor= "#ffffcc"
							end if%>
							<tr class="texto" bgcolor="<%=cor%>">
								<td><%=contas(3,x) & contas(2,x)%></td>
								<%getValores contas(0,x)%>
								<%ehRaiz = ContaRaiz(contas(0,x))%>
								<%for mes = 1 to 12 %>
									<td style="<%=iif(valores(mes-1)>0,"color:#0000FF",iif(valores(mes-1)=0,"color:#000000","color:#990000"))%>" align="right"><%=formatNumber(valores(mes-1),0)%></td>
									<%if ehRaiz then totais(mes-1) = totais(mes-1) + valores(mes-1)	%>									
								<%next%>
							</tr>
					
					
						<%	vcor = not vcor
						next%>
						<tr>
							<td class="cabecalho">Total</td>
								<%for mes = 1 to 12 %>
									<td width="6%" style="<%=iif(totais(mes-1)>0,"color:#0000FF",iif(totais(mes-1)=0,"color:#000000","color:#990000"))%>" align="right" class="cabecalho"><%=formatNumber(totais(mes-1),0)%></td>
								<%next%>				
						</tr>
					</tbody>
				</table>
			<%
		end if
	end if
end sub
%>

<%sub Impressao
	response.write request.ServerVariables("SCRIPT_NAME") & "?imp=1&ispostback=1&ano=" & ano
end sub

sub auxiliar
dim anos, sqlanos, rsanos
redim anos(-1)
%>
<form name="filtros" method="post" action="">
<input type="hidden" value="1" name="isPostBack">
<table>
	<tr>
		<td><label class="subtitulo">Filtros</label><br><BR>
			<label class="cabecalho">Ano</label><br>
			<%
			sqlanos = "select distinct year(dataLancamento) as ano from fluxoCaixa fc inner join planocontas pc on fc.codplanoconta = pc.codconta where pc.ativo = 1"
			set rsanos = conn.execute(Sqlanos)
			while not rsanos.eof
				redim preserve anos(ubound(anos)+1)
				anos(ubound(anos)) = rsanos("ano")
				rsanos.movenext
			wend
			rsanos.close : set rsanos = nothing
			Combo "ano", Anos, Anos, ano%>
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

%>

<%if request.QueryString("imp") = "" then%>
<!--#include file="../layout/layout.asp"-->
<%else%>
<!--#include file="../layout/layout_print.asp"-->
<%end if%>

<%call fechaBanco%>


<%


function getFilhos(codconta)
	dim sql, rs, retorno
	redim retorno(-1)
	sql = "select codconta from planocontas where codcontasuperior = " & codconta
	set rs = conn.execute(sql)
	while not rs.eof 
		redim preserve retorno(ubound(retorno)+1)
		retorno(ubound(retorno)) = rs("codconta")
		rs.movenext
	Wend
	rs.close : set rs = nothing
	getFilhos = retorno
end function

sub addConta(codconta, nivel)
	dim filhos, conta
	
	redim preserve contas(3,ubound(contas,2)+1)
	conta = getConta(codconta)
	
	for x=0 to ubound(conta)
		contas(x,ubound(contas,2)) = conta(x)
	next
	for x=0 to nivel
		contas(3,ubound(contas,2)) = contas(3,ubound(contas,2)) & "&nbsp;"
	next
	filhos = getfilhos(codconta)
	
	if ubound(filhos)>-1 then
		for x = 0 to ubound(filhos)
			addConta filhos(x), (nivel + 2)
			
		next
	end if


end sub


function getConta(codconta)
	dim sql, rs, retorno
	redim retorno(2)
	sql = "Select codconta, codcontasuperior, descricao from planocontas where codconta = " & codconta
	set rs = conn.execute(sql)
	for x=0 to ubound(retorno)
		retorno(x) = rs.fields(x)
	next
	rs.close :set rs = nothing
	getConta = retorno
end function
function contaRaiz(codconta)
	dim rs, sql, retorno
	sql = "Select codconta from planocontas where codcontaSuperior is null and codconta = " & codconta
	set rs = conn.execute(Sql)
	if not rs.eof then
		retorno = true
	else
		retorno = false
	end if
	rs.close : set rs = nothing
	contaRaiz = retorno
end function
sub getValores(codconta)
	dim sql, rs, filhos
	
	sql = "Select mes1 "
	for x=2 to 12
		sql = sql & ", mes"&x
	next
	sql = sql & " from totalLancamentosmes where codconta = " & codconta & " and ano = " & ano
	
	set rs = conn.execute(Sql)
	if not rs.eof then
		for x =0 to ubound(valores)
			valores(x) = valores(x) + Nulo(rs.fields(x))
		next
	end if
	
	filhos = getFilhos(codconta)
	if ubound(filhos)>-1 then
		for x=0 to ubound(filhos)
			getValores filhos(x)
		next
	end if


end sub


	
	
	

%>