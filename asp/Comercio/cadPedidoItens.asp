<!--#include file="../Dados/DBfuncoes.asp"-->
<!--#include file="../funcoes/funcoes.asp"-->
<!--#include file="../funcoes/gridGeral.asp"-->
<!--#include file="funPedido.asp"-->

<script language="javascript" src="../Funcoes/funcoes.js" type="text/javascript"></script>

<%call abreBanco%>
<%
dim acao, erro
dim descricao, qt, vlUnitario
dim itensPedido
dim indice

itens = session("pedidoItens")
indice = request("indice")
acao = request("acaoItens")

if acao = "del" then
	delItemPedido indice
	response.Redirect("cadPedidos.asp")
end if

if request.Form("isPostBack") = "1" then
	descricao = request.Form("descricao")
	qt = request.Form("qt")
	vlUnitario = request.Form("vlUnitario")
	
	if descricao = "" then erro = "Descrição inválida. "
	if qt = "" or Not IsNumeric(qt) then erro = erro & "Qt. inválida. "
	if vlUnitario = "" or not isNumeric(vlUnitario) then erro = erro & "Vl. Unitário inválido. "
	
	if erro = "" then
		if acao = "add" then 
			addItemPedido descricao, qt, vlUnitario, "null"
		elseif acao = "edit" then 
			editItemPedido indice, descricao, qt, vlUnitario, "null"
		end if

		response.Redirect("cadPedidos.asp#topo")
	end if
elseif acao = "edit" then
	descricao = itens(1, indice)
	vlUnitario = itens(2, indice)
	qt = itens(3, indice)
end if


sub principal
%>
<form name="form1" method="post" action="cadPedidoItens.asp">
	<input type="hidden" name="isPostBack" value="1">
	<input type="hidden" name="indice" value="<%=indice%>">
	<input type="hidden" name="acaoItens" value="<%=acao%>">
  <p><font class="titulo">Item do Pedido</font><br>
    <font class="subtitulo">Informa&ccedil;&otilde;es do item do pedido</font><br>
    <br>
    <label class="cabecalho">Descri&ccedil;&atilde;o:</label><br>
    <textarea name="descricao" cols="40" rows="5" class="caixa" id="descricao" onKeyUp="maxLengthTextArea(this, event, 1000);"><%=descricao%></textarea>
    <br>
    <label class="cabecalho">Qt:</label><br>
    <input name="qt" type="text" class="caixa" id="qt" size="10" maxlength="8" value="<%=qt%>">
    <br>	
    <label class="cabecalho">Vl. Unit&aacute;rio:</label><br>
    <input name="vlUnitario" type="text" class="caixa" id="vlUnitario" size="10" maxlength="8" value="<%=vlUnitario%>">
    <br><br>
    <input name="btnOK" type="submit" id="btnOK" value="OK" class="botao">
    <input name="btnCancelar" type="button" id="btnCancelar" value="Cancelar" class="botao" onClick="window.location.href='cadPedidos.asp#topo'">
	<p class="erro">
		&nbsp;<%=erro%>
	</p>
</p>
  </form>

<%
end sub


sub Impressao
end sub

sub auxiliar

end sub

%>

<!--#include file="../layout/layout.asp"-->
<%call FechaBanco%>