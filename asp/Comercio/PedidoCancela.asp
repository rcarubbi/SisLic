<!--#include file="../Dados/DBfuncoes.asp"-->
<!--#include file="../funcoes/funcoes.asp"-->
<!--#include file="funPedido.asp"-->

<%call abreBanco%>
<%
dim numPedido
numPedido = request("numPedido")

if request.Form("isPostBack") = "1" then
	PedidoCancela numPedido
	response.Redirect("PedidoLista.asp")
end if

sub Impressao
end sub

sub auxiliar

end sub

sub principal

%>
	<form name="form1" method="post" action="PedidoCancela.asp">
		<input type="hidden" name="isPostBack" value="1">
		<input type="hidden" name="NumPedido" value="<%=numPedido%>">	
				
			<font class="titulo">Cancelamento do Pedido n�: <%=numPedido%></font><br>
			<font class="subtitulo">Comfirma o cancelamento do pedido n�: <%=numPedido%>?</font><br><br>
			<font class="erro">Aten��o: Ap�s o cancelamento do pedido n�o � poss�vel recuper�-lo!</font>
			<p>
				<input type="submit" name="btnSim" value="Sim" class="botao">&nbsp;
				<input type="button" name="btnNao" value="N�o" class="botao" onClick="history.back(-1);">			
			</p>
	</form>

<%end sub


%>
<!--#include file="../layout/layout.asp"-->
<%call FechaBanco%>