<!--#include file="../Dados/DBfuncoes.asp"-->
<!--#include file="../funcoes/funcoes.asp"-->
<!--#include file="funPedido.asp"-->

<%call abreBanco%>
<%
dim numPedido, motivo
numPedido = request("numPedido")

if request.Form("isPostBack") <> "" then
	motivo = request.form("motivoCancelamento")
	PedidoCancelaRecebido numPedido, motivo
	response.Redirect("PedidoLista.asp")
end if

sub Impressao
end sub

sub auxiliar

end sub

sub principal

%>
	<form name="form1" method="post" action="PedidoCancelaRecebido.asp">
		<input type="hidden" name="isPostBack" value="1">
		<input type="hidden" name="NumPedido" value="<%=numPedido%>">	
				
			<font class="titulo">Cancelamento do Pedido nº: <%=numPedido%></font><br><BR>
			<label class="subtitulo">Informe o motivo do cancelamento</label>
		 	<br>
		  	<label class="cabecalho">Motivo:</label>
		  	<br>
  	      	<textarea class="caixa" cols="40" rows="3" name="motivoCancelamento" onKeyUp="maxLengthTextArea(this, event, 255);"><%=motivoCancelamento%></textarea>
		  	<br>
		   	<BR>
			<font class="subtitulo">Comfirma o cancelamento do pedido nº: <%=numPedido%>?</font><br>
			<font class="erro">Atenção: Após o cancelamento do pedido não é possível recuperá-lo!</font>
			<p>
				<input type="submit" name="btnSim" value="Sim" class="botao">&nbsp;
				<input type="button" name="btnNao" value="Não" class="botao" onClick="history.back(-1);">			
			</p>
	</form>

<%end sub


%>
<!--#include file="../layout/layout.asp"-->
<%call FechaBanco%>