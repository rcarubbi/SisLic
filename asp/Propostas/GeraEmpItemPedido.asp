<!--#include file="../Dados/DBfuncoes.asp"-->
<!--#include file="../funcoes/funcoes.asp"-->
<!--#include file="../funcoes/gridGeral.asp"-->
<!--#include file="funGeraEmpenho.asp"-->

<script language="javascript" src="../Funcoes/funcoes.js" type="text/javascript"></script>

<%call abreBanco%>
<%
dim acao, erro
dim descricao, qtPedido, qtProduto, vlUnitario, idFornecedor
dim idPropostaItem, ID
dim idProcesso

acao = request("acao")
idPropostaItem = request("idPropostaItem")
ID = request("ID")
idProcesso = request("idProcesso")

if request.Form("isPostBack") = "1" then
	descricao = request.Form("descricao")
	qtProduto = request.Form("qtProduto")
	qtPedido = request.Form("qtPedido")	
	vlUnitario = request.Form("vlUnitario")
	idFornecedor = request.Form("idFornecedor")
	
	if descricao = "" then erro = "Descrição inválida. "
	if idFornecedor = "" or not isNumeric(idFornecedor) then erro = erro & "Fornecedor inválido. "	
	if qtProduto = "" or Not IsNumeric(qtProduto) then erro = erro & "Qt. Produto inválida. "
	if qtPedido = "" or Not IsNumeric(qtPedido) then erro = erro & "Qt. Pedido inválida. "	
	if vlUnitario = "" or not isNumeric(vlUnitario) then erro = erro & "Vl. Unitário inválido. "
	if erro = "" then 
		if cdbl(qtPedido) < cdbl(qtProduto) then erro = erro & "A quantidade do pedido não pode ser menor que a quantidade utilizada no produto. "
	end if
	if erro = "" then
		
		if acao = "add" then 
			tempGravaSubItensPedido idPropostaItem, idFornecedor, descricao, qtProduto, qtPedido, vlUnitario
		elseif acao = "edit" then 
			tempEditaSubItensPedido ID, idFornecedor, descricao, qtProduto, qtPedido, vlUnitario
		end if

		response.Redirect("GeraEmpItemLista.asp?ori=ItemPedido&idProcesso=" & idProcesso)
	end if
elseif acao = "edit" then
	set rs = tempGetSubItensPedidos("", ID)
	if not rs.eof then
		descricao = rs("descricao")
		vlUnitario = rs("valorUnitario")
		qtProduto = rs("qtProduto")
		qtPedido = rs("qtPedido")
		idPropostaItem = rs("idPropostaItem")
		idFornecedor = rs("idFornecedor")
	end if
	rs.close
	set rs = nothing
elseif acao = "del" then
	tempDeletaSubItemPedido ID
	response.Redirect("GeraEmpItemLista.asp?ori=ItemPedido&idProcesso=" & idProcesso)
end if


sub principal
%>
<form name="form1" method="post" action="GeraEmpItemPedido.asp">
	<input type="hidden" name="isPostBack" value="1">
	<input type="hidden" name="idPropostaItem" value="<%=idPropostaItem%>">
	<input type="hidden" name="acao" value="<%=acao%>">
	<input type="hidden" name="ID" value="<%=ID%>">
	<input type="hidden" name="idProcesso" value="<%=idProcesso%>">
  <p><font class="titulo">Item do Pedido</font><br>
    <font class="subtitulo">Informa&ccedil;&otilde;es do item do pedido</font><br>
    <br>
	<font class="cabecalho">Fornecedor:</font><br>
	<%
	ComboDB "idFornecedor", "nome", "id", "Select * from parceiro where ativo=1 and tipo=1",  idFornecedor, "",""
	%>
	<br>
    <label class="cabecalho">Descri&ccedil;&atilde;o:</label><br>
    <textarea name="descricao" cols="40" rows="5" class="caixa" id="descricao" onKeyUp="maxLengthTextArea(this, event, 1000);"><%=descricao%></textarea>

	<table cellpadding="0" cellspacing="0" border="0" width="200">
		<tr>
			<td> 
				<label class="cabecalho">Qt Produto:</label><br>
    			<input name="qtProduto" onBlur="qtPedido.value=this.value" type="text" class="caixa" id="qtProduto" size="10" maxlength="8" value="<%=qtProduto%>">
			</td>
			<td> 
				<label class="cabecalho">Qt Pedido:</label><br>
    			<input name="qtPedido" type="text" class="caixa" id="qtPedido" size="10" maxlength="8" value="<%=qtPedido%>">
			</td>
		</tr>
	</table>
<label class="cabecalho">Vl. Unit&aacute;rio:</label><br>
<input name="vlUnitario" type="text" class="caixa" id="vlUnitario" size="10" maxlength="8" value="<%=vlUnitario%>">   

        <br><br>
    <input name="btnOK" type="submit" id="btnOK" value="OK" class="botao">
    <input name="btnCancelar" type="button" id="btnCancelar" value="Cancelar" class="botao" onClick="window.location.href='GeraEmpItemLista.asp?ori=ItemPedido&idProcesso=<%=idprocesso%>'">
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