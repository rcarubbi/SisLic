<!--#include file="../Dados/DBfuncoes.asp"-->
<!--#include file="../funcoes/funcoes.asp"-->
<!--#include file="../funcoes/gridGeral.asp"-->
<!--#include file="funGeraEmpenho.asp"-->

<%call abreBanco%>
<%
dim acao, erro
dim idEstoque, Qt
dim idPropostaItem, ID
dim idProcesso

acao = request("acao")
idPropostaItem = request("idPropostaItem")
ID = request("ID")
idProcesso = request("idProcesso")


if request.Form("isPostBack") = "1" then
	qt = request.Form("qt")
	idEstoque = request.Form("idEstoque")
	
	if idEstoque = "" or not isNumeric(idEstoque) or idEstoque = "0" then erro = erro & "Produto inválido. "	
	if qt = "" or Not IsNumeric(qt) or qt = "0" then erro = erro & "Quantidade inválida. "
	
	if erro = "" then
		
		if acao = "add" then 
			erro = tempGravaSubItensEstoque(idPropostaItem, idEstoque, qt)
		elseif acao = "edit" then 
			erro = tempEditaSubItensEstoque(ID, idEstoque, qt)
		end if

		if erro = "" then response.Redirect("GeraEmpItemLista.asp?ori=ItemEstoque&idProcesso=" & idProcesso)
		
	end if
elseif acao = "edit" then
	set rs = tempGetSubItensEstoque("", ID)
	if not rs.eof then
		qt = rs("qt")
		idPropostaItem = rs("idPropostaItem")
		idEstoque = rs("idEstoque")
	end if
	rs.close
	set rs = nothing
elseif acao = "del" then
	tempDeletaSubItemEstoque ID
	response.Redirect("GeraEmpItemLista.asp?ori=ItemEstoque&idProcesso=" & idProcesso)
end if


sub principal
%>
<form name="form1" method="post" action="GeraEmpItemEstoque.asp">
	<input type="hidden" name="isPostBack" value="1">
	<input type="hidden" name="idPropostaItem" value="<%=idPropostaItem%>">
	<input type="hidden" name="acao" value="<%=acao%>">
	<input type="hidden" name="ID" value="<%=ID%>">
	<input type="hidden" name="idProcesso" value="<%=idProcesso%>">
	
  <p><font class="titulo">Baixa em estoque</font><br>
    <font class="subtitulo">Selecione o produto em estoque e a quantidade desejada.</font><br>
    <br>
	<font class="cabecalho">Produto:</font><br>
	<%
	dim sql
	'sql para retornar a descrição do produto mais a quantidade disponivel 
	'subtraindo pelo quantidade já utilizada desse mesmo item da tabela temporaria de 
	'estoque
	sql = "select " & _
	"Estoque.ID, " & _
	"'[' + str(" & _
				"Estoque.qt - isNull(tempEstoqueUsado.tot,0), " & _
				"len(" & _
					"Estoque.qt - isNull(tempEstoqueUsado.tot,0)" & _
					")" & _
			  ")" & _
	 "+ ']' + PedidoItem.Descricao as descrQT" & _
	" from " & _
	"PedidoItem INNER join " & _
	"(" & _
		"Estoque LEFT join" & _
		"(" & _
			"SELECT sum(qt) as tot, IDEstoque " & _
			"from tempEmpItemEstoque " & _
			"where sessionID='" & session.SessionID & "' " & _
			"group by IDEstoque" & _
		") as tempEstoqueUsado " & _
		"ON Estoque.ID = tempEstoqueUsado.IDEstoque" & _
	") " & _
	"ON Estoque.idPedidoItem = PedidoItem.id" & _
	" WHERE (Estoque.qt - isNull(tempEstoqueUsado.tot, 0)) > 0"
	if acao = "edit" then
		sql = sql & " or Estoque.ID=" & idEstoque
	end if

	ComboDB "idEstoque", "descrQT", "id", sql,  idEstoque, "",""
	%>
	<br>
	<label class="cabecalho">Quantidade:</label><br>
 	<input name="qt" type="text" class="caixa" id="qt" size="10" maxlength="8" value="<%=qt%>">
     <br><br>
    <input name="btnOK" type="submit" id="btnOK" value="OK" class="botao">
    <input name="btnCancelar" type="button" id="btnCancelar" value="Cancelar" class="botao" onClick="window.location.href='GeraEmpItemLista.asp?ori=ItemEstoque&idProcesso=<%=idProcesso%>'">
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