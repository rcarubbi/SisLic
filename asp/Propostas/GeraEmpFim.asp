<!--#include file="../Dados/DBfuncoes.asp"-->
<!--#include file="../funcoes/funcoes.asp"-->
<%call abreBanco%>
<%
sub Principal
%>
	<p>
		<label class="titulo">Empenho</label>
		<br>
		<label class="subtitulo">Empenho gerado com sucesso!</label>
	</p>
	<input type="button" name="pedidos" value="Ver Empenho" class="botao" onClick="window.location.href='EmpenhoLista.asp'">&nbsp;
	<input type="button" name="pedidos" value="Voltar para as propostas" class="botao" onClick="window.location.href='propostaLista.asp'">	
<%
end sub

sub Impressao
end sub

sub auxiliar
	
end sub

%>
    
<!--#include file="../layout/layout.asp"-->

