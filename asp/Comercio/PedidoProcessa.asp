<!--#include file="../Dados/DBfuncoes.asp"-->
<!--#include file="../funcoes/funcoes.asp"-->
<!--#include file="funPedido.asp"-->

<%call abreBanco%>
<%
dim numPedido,  nf, vlImposto, dtRecebimento, hrRecebimento, obs, tela
dim codPlanoContaImposto, codPlanoContaNota
numPedido = request("numPedido")



tela = request.Form("tela")
if tela = "" then tela = 0

if request.Form("isPostBack") = "1" and tela = 1 then

	nf = request.Form("nf")
	vlImposto = request.Form("vlImposto")
	codPlanoContaImposto = request.Form("codPlanoContaImposto")
	codPlanoContaNota = request.Form("codPlanoContaNota")
	dtRecebimento = recebe_datador("dtRecebimento", 0)
	hrRecebimento = recebe_datador("hrRecebimento", 1)	
	
	obs = request.Form("obs")	
	if not isNumeric(nf) or nf = "" then erro = erro & "NF. de compra inválida. "
	if not isNumeric(vlImposto) or vlImposto = "" then erro = erro & "Vl. do Imposto inválido. "
	if not isDate(dtRecebimento) then erro = erro & "Data de recebimento inválida. "
	if not isDate(hrRecebimento) then erro = erro & "Hora de recebimento inválida. "	
	if not isNumeric(codPlanoContaImposto) or codPlanoContaImposto = "0" then erro = erro & "Plano de contas para os Impostos inválido. "
	if not isNumeric(codPlanoContaNota) or codPlanoContaNota = "0" then erro = erro & "Plano de contas para a nota inválida. "
	
	if erro = "" then
		tela = 2
	end if	
elseif request.Form("isPostBack") = "1" and tela = 2 then
	nf = request.Form("nf")
	vlImposto = request.Form("vlImposto")
	codPlanoContaImposto = request.Form("codPlanoContaImposto")
	codPlanoContaNota = request.Form("codPlanoContaNota")
	dtRecebimento = request.Form("dtRecebimento")
	hrRecebimento = request.Form("hrRecebimento")	
	obs = request.Form("obs")	

	PedidoProcessa NumPedido, nf, vlImposto, dtRecebimento, hrRecebimento,obs, codPlanoContaImposto, codPlanoContaNota
	response.Redirect("PedidoLista.asp")
else
	tela = 1
	sql = "select * from pedido where numPedido=" & numPedido
	set rs = conn.execute(SQL)
	if not rs.eof then
		nf = rs("numNFCompra")
		vlImposto = rs("vlImpostos")
		dtRecebimento = day(rs("dtRecebimento")) & "/" & month(rs("dtRecebimento")) & "/" & year(rs("dtRecebimento"))
		hrRecebimento = hour(rs("dtRecebimento")) & "/" & minute(rs("dtRecebimento")) & "/" & second(rs("dtRecebimento"))		
		
		obs = rs("obs")
	end if
	
	dim rAux
	if application("par_4") <> "" and isNumeric(application("par_4")) then
		sql = "Select codConta from PlanoContas where ativo=1 and codConta=" & application("par_4")
		set rAux = conn.execute(SQL)
		if not rAux.eof then
			codPlanoContaNota = rAux("codConta")
		end if
		rAux.close
		set rAux = nothing
	end if

	if application("par_5") <> "" and isNumeric(application("par_5")) then
		sql = "Select codConta from PlanoContas where ativo=1 and codConta=" & application("par_5")
		
		set rAux = conn.execute(SQL)
		if not rAux.eof then
			codPlanoContaImposto = rAux("codConta")
		end if
		rAux.close
		set rAux = nothing
	end if
end if


sub Impressao
end sub

sub auxiliar

end sub

sub principal

%>
	<form name="form1" method="post" action="PedidoProcessa.asp">
		<input type="hidden" name="isPostBack" value="1">
		<input type="hidden" name="NumPedido" value="<%=numPedido%>">	
		<input type="hidden" name="tela" value="<%=tela%>">	
		<%
		select case tela
		case 1
			%>
			<font class="titulo">Pedido nº: <%=numPedido%></font><br>
			<font class="subtitulo">Recebimento do pedido.</font><br><br>
			<p>
			<font class="cabecalho">Data de Recebimento:</font><br>
			<%Datador "dtRecebimento",0,year(date)-10,year(date)+1,dtRecebimento%><br>
			
			<font class="cabecalho">Hora do Recebimento:</font><br>
			<%Datador "hrRecebimento",1,0,0,hrRecebimento%><br>			
			
			<font class="cabecalho">NF. de Compra:</font><br>
			<input name="nf" type="text" id="nf" size="20" maxlength="10" class="caixa" value="<%=nf%>">
			<br> 	
			<font class="cabecalho">Vl. Impostos:</font><br>
			<input name="vlImposto" type="text" id="vlImposto" size="10" maxlength="8" class="caixa" value="<%=vlImposto%>">
			<br> 				
			<font class="cabecalho">Plano de Contas p/ a Nota:</font><br>
			<%
			
			sql = "select * from planoContas where tipo = 1 and ativo = 1 and codConta not in" & _
					"(select isNull(codContaSuperior, 0) from planoContas)"
			
			ComboDB "codPlanoContaNota", "descricao", "codConta", sql,  codPlanoContaNota, "",""
			%>
			<br>
			<font class="cabecalho">Plano de Contas p/ os Impostos:</font><br>
			<%
			
			sql = "select * from planoContas where tipo = 0 and ativo = 1 and codConta not in" & _
					"(select isNull(codContaSuperior, 0) from planoContas)"
			
			ComboDB "codPlanoContaImposto", "descricao", "codConta", sql,  codPlanoContaImposto, "",""
			%><br>
			<font class="cabecalho">Obs:</font><br>
			<textarea name="obs" cols="40" rows="5" class="caixa" id="obs"><%=obs%></textarea>
			<br>
			</p>
			<input type="submit" name="btnOK" value="Próximo" class="botao">&nbsp;
			<input type="button" name="btnCancelar" value="Cancelar" class="botao" onClick="window.location.href='PedidoLista.asp'">			
		<%
		case 2
		%>
			<input type="hidden" name="nf" value="<%=nf%>">
			<input type="hidden" name="vlImposto" value="<%=vlImposto%>">
			<input type="hidden" name="dtRecebimento" value="<%=dtRecebimento%>">
			<input type="hidden" name="hrRecebimento" value="<%=hrRecebimento%>">			
			<input type="hidden" name="obs" value="<%=obs%>">				
			<input type="hidden" name="codPlanoContaImposto" value="<%=codPlanoContaImposto%>">				
			<input type="hidden" name="codPlanoContaNota" value="<%=codPlanoContaNota%>">										
			<font class="titulo">Pedido nº: <%=numPedido%></font><br>
			<font class="subtitulo">Comfirma o recebimento do pedido?</font><br><br>
			<table border="0" cellpadding="3">
				<tr>
					<td align="right"><font class="cabecalho">NF. de Compra:</font></td>
					<td align="left"><font class="texto"><%=nf%></font></td>					
				</tr>
				<tr>
					<td align="right"><font class="cabecalho">Vl. Impostos:</font></td>
					<td align="left"><font class="texto"><%=formata_campo(cCur(vlImposto))%></font></td>					
				</tr>
				<tr>
					<td align="right"><font class="cabecalho">Obs:</font></td>
					<td align="left"><font class="texto"><%=obs%></font></td>					
				</tr>								
				<tr>
					<td colspan="2" align="center">
					<br>
						<input type="submit" name="btnSim" value="Sim" class="botao">&nbsp;
						<input type="button" name="btnNao" value="Não" class="botao" onClick="history.back(-1);">			
					
					</td>
				</tr>
	  </table>
		<%
		end select
		%>
		<p class="erro">
			&nbsp;<%=erro%>
		</p>
	</form>

<%end sub


%>
<!--#include file="../layout/layout.asp"-->
<%call FechaBanco%>