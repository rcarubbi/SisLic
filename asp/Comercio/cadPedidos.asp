<!--#include file="../Dados/DBfuncoes.asp"-->
<!--#include file="../funcoes/funcoes.asp"-->
<!--#include file="../funcoes/GridGeral.asp"-->
<!--#include file="funPedido.asp"-->

<%call abreBanco%>
<%
dim fornecedor
dim itens
dim acaoItens
dim erro
dim dtPedido

acaoItens = request.Form("acaoItens")

itens = session("pedidoItens")
	
if request.QueryString("limpar") = "1" or not isArray(itens) then
	'colunas
	'0=indice
	'1=descricao
	'2=quantidade
	'3=vl. unitário
	'4=vl. total (apenas para exibição, não grava no banco)
	redim itens(5,-1)
	
	session("pedidoItens") = itens
	session("pedidoFornecedor") = ""
	session("pedidoIdProcesso") = ""
	session("pedidoObs") = application("par_24")
	session("pedidoDt") = ""
	session("pedidoPlanoContaImposto") = ""
	session("pedidoPlanoContaNota") = ""
	session("pedidoComprador") = application("par_21")
	session("pedidoPrazoEntrega") = application("par_22")
	session("pedidoFormaPagto") = application("par_23")
end if



if request.Form("isPostBack") = "1" then
	session("pedidoFornecedor") = request.Form("fornecedor")
	session("pedidoIdProcesso") = request.Form("processo")
	session("pedidoObs") = request.Form("obs")
	session("pedidoDt") = recebe_datador("dtPedido",0)
	session("pedidoPlanoContaImposto") = request.Form("planoContaImposto")
	session("pedidoPlanoContaNota") = request.Form("planoContaNota")
	session("pedidoComprador") = request.Form("comprador")
	session("pedidoPrazoEntrega") = request.Form("prazoEntrega")
	session("pedidoFormaPagto") = request.Form("formaPagto")

	if acaoItens <> "" then
		response.Redirect("cadPedidoItens.asp?indice=" & request.Form("optLinha") & "&acaoItens=" & acaoItens)
	end if
	
	if ubound(itens,2) = -1 then erro = erro & "Nenhum item foi adicionado na lista. "
	if session("pedidoFornecedor") = "0" or not isNumeric(session("pedidoFornecedor")) then erro = erro & "Fornecedor inválido. "
	if not isDate(session("pedidoDt")) then erro = erro & "Data do Pedido inválida!"
	
	if erro = "" then 'ok
		SalvaNovoPedido session("pedidoIdProcesso"), session("pedidoObs"), session("pedidoFornecedor"), session("pedidoDt"), session("pedidoComprador"), session("pedidoFormaPagto"), session("pedidoPrazoEntrega")
		response.Redirect("PedidoLista.asp")
	end if
	
end if


sub Impressao
end sub

sub auxiliar

end sub

sub principal

%>	<a id="topo"></a>
	<form name="form1" method="post" action="cadPedidos.asp">
		<input type="hidden" name="isPostBack" value="1">
		<font class="titulo">Pedido</font><br>
		<font class="subtitulo">Novo Pedido</font><br>
		
		
		<p><font class="cabecalho">N&uacute;m. Processo:</font><br>
		<%
		ComboDB "processo", "numProcesso", "idProcesso", "Select proposta.idProcesso, NumProcesso from proposta inner join edital on edital.idProcesso=proposta.idProcesso where proposta.situacao=1", session("pedidoIdProcesso"), "",""
		%>

		<br>
		<font class="cabecalho">Fornecedor:</font><br>
		<%
		ComboDB "fornecedor", "nome", "id", "Select * from parceiro where ativo=1 and tipo=1",  session("pedidoFornecedor"), "","355"
		%>
		<br>
		<font class="cabecalho">Data do Pedido:</font><br>
		<%datador "dtPedido", 0, year(now) - 10, year(now) + 1, session("pedidoDt")%>		
		<br> 		
		<font class="cabecalho">Comprador (Responsável): </font><br>
		<input name="comprador" type="text" class="caixa" value="<%=session("pedidoComprador")%>" maxlength="100">
		<br> 		
		<font class="cabecalho">Prazo de Entrega: </font><br>
		<input name="prazoEntrega" type="text" class="caixa" value="<%=session("pedidoPrazoEntrega")%>" maxlength="100">
		<br> 		
		<font class="cabecalho">Forma de Pagamento: </font><br>
		<input name="formaPagto" type="text" class="caixa" value="<%=session("pedidoFormaPagto")%>" maxlength="100">
		<br>
		<font class="cabecalho">Obs:</font><br>
		<textarea name="obs" cols="40" rows="5" class="caixa" id="obs"><%=session("pedidoObs")%></textarea>
		<br>
		</p>
		<hr noshade size="1" color="#CCCCCC">
		<%itensLista%>
		<br>
			<input type="submit" name="btnOK" value="OK" class="botao">&nbsp;
			<input type="button" name="btnCancelar" value="Cancelar" class="botao" onClick="window.location.href='PedidoLista.asp'">			
		<p class="erro">
			&nbsp;<%=erro%>
		</p>
	</form>

<%end sub



sub itensLista

	dim colunas(4,3), botoes(3, 3), sql, rs
	
	colunas(0,0) = "Descrição"
	colunas(1,0) = 1
	colunas(2,0) = "20%"

	colunas(0,1) = "Vl. Unitário"
	colunas(1,1) = 2
	colunas(2,1) = "20%"
	colunas(3,1) = true
	
	colunas(0,2) = "Qt."
	colunas(1,2) = 3
	colunas(2,2) = "20%"
	colunas(3,2) = true
	
	colunas(0,3) = "Vl. Total"
	colunas(1,3) = 4
	colunas(2,3) = "20%"
	colunas(3,3) = true
	
	botoes(0,0) = true
	botoes(0,1) = true
	botoes(0,2) = true		
	botoes(0,3) = false	

				

	' configurando o grid

	
	' fim da configuração do grid
%>


	<label class="subtitulo">Itens do Pedido</label>
	<%
	monta_grid itens, "acaoItens", "", "", 0, 0, Colunas, Botoes, ""
	
end sub

%>
<!--#include file="../layout/layout.asp"-->
<%call FechaBanco%>