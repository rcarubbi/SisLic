<!--#include file="../Dados/DBfuncoes.asp"-->
<!--#include file="../funcoes/funcoes.asp"-->
<!--#include file="../funcoes/gridGeral.asp"-->
<%call abreBanco%>

<%
dim acao
dim dtini, dtfim

acao = request.Form("acao")

if acao = "add" then
	response.Redirect("cadPedidos.asp?acao=add&limpar=1")
elseif acao = "can" then
	response.Redirect("pedidoCancela.asp?NumPedido=" & request.Form("optLinha"))
elseif acao = "canRec" then
	response.Redirect("pedidoCancelaRecebido.asp?NumPedido=" & request.Form("optLinha"))
elseif acao = "vis" then
	response.Redirect("pedidoVisualizar.asp?NumPedido=" & request.Form("optLinha"))
elseif acao = "proc" then
	response.Redirect("PedidoProcessa.asp?NumPedido=" & request.Form("optLinha"))
end if

dtIni = recebe_datador("dtIni",0)
dtFim = recebe_datador("dtFim",0)


if not isDate(dtIni) then
	dtini = dateadd("m",-1,date)
	dtfim = date
end if
if cdate(dtFim) < cdate(dtIni) then dtfim = dtini
sub principal

	formLista
	
end sub

sub Impressao
end sub

sub auxiliar
	%>

	<form name="form2" action="" method="post" enctype="application/x-www-form-urlencoded">
		<input type="hidden" name="filtros" value="1">
		<%
		dim itemTexto(2), itemValor(2)
		
		itemTexto(0) = "Pendentes"
		itemValor(0) = 0
		
		itemTexto(1) = "Recebidos"
		itemValor(1) = 1
		
		itemTexto(2) = "Cancelados"
		itemValor(2) = 2		
		
		%><label class="cabecalho">Situação</label><br><%
		Combo "situacao", itemTexto, itemValor, request.Form("situacao")
		
		%>
		<br>
		<label class="cabecalho">Data inicial</label>
		<br>
		<%
		Datador "dtIni",0,1980, year(now),dtIni
		%>
		<br>
		<label class="cabecalho">Data final</label>
		<br>
		<%
		Datador "dtFim",0,1980, year(now),dtFim	
		%>
		<p>
		<input type="submit" name="OK" value="OK" class="botao">
		</p>
	</form>

	<%
end sub

sub FormLista
	dim colunas(4,6), botoes(3, 7), sql, rs
	situacao = request.Form("situacao")
	if situacao =  "" then situacao = 0
		
	colunas(0,0) = "Num. Pedido"
	colunas(1,0) = 0
	colunas(2,0) = "15%"

	colunas(0,1) = "Num. Processo"
	colunas(1,1) = 1
	colunas(2,1) = "15%"
	
	colunas(0,2) = "Dt. Pedido"
	colunas(1,2) = 2
	colunas(2,2) = "15%"

	colunas(0,3) = "Dt. Recebimento"
	colunas(1,3) = 3
	colunas(2,3) = "15%"
	
	colunas(0,4) = "Vl. Total"
	colunas(1,4) = 4
	colunas(2,4) = "15%"
	
	colunas(0,5) = "Vl. Impostos"
	colunas(1,5) = 5
	colunas(2,5) = "20%"
	
	colunas(0,6) = "Vl. Fornecedor"
	colunas(1,6) = 6
	colunas(2,6) = "20%"
	
	'botoes
	botoes(0,0) = true 'novo
	botoes(0,1) = false 'editar
	botoes(0,2) = false  'deletar
	botoes(0,3) = false	 'recuperar
	
	select case situacao
	case 0 'pendente
		botoes(0,4) = true
		botoes(1,4) = "Cancelar"
		botoes(2,4) = "can"
		botoes(3,4) = true
			
		botoes(0,5) = true
		botoes(1,5) = "Receber"
		botoes(2,5) = "proc"
		botoes(3,5) = true
	case 1 'recebido
		botoes(0,4) = true
		botoes(1,4) = "Cancelar"
		botoes(2,4) = "canRec"
		botoes(3,4) = true
	case 2 'cancelado
	end select
	
	botoes(0,6) = true
	botoes(1,6) = "Visualizar"
	botoes(2,6) = "vis"
	botoes(3,6) = true


	
	sql = "SELECT     Pedido.NumPedido, Edital.NumProcesso, Pedido.DtPedido, Pedido.DtRecebimento, Pedido.vlTotal, Pedido.vlImpostos, Parceiro.Nome " & _
			"FROM      Edital RIGHT JOIN " & _
					"(Parceiro INNER JOIN Pedido ON Parceiro.Id = Pedido.idFornecedor) " & _
					"ON Edital.idProcesso = Pedido.idProcesso"
	select case situacao
	case 0
		sql = sql & " WHERE pedido.recebido=0 and pedido.cancelado=0"
	case 1
		sql = sql & " WHERE pedido.recebido=1 and pedido.cancelado=0"
	case 2
		sql = sql & " WHERE pedido.cancelado=1"
	end select
	sql = sql & " and pedido.dtpedido between " & formata_data_sql(dtIni) & " and " & formata_data_sql(dtFim & " 23:59:59")
	set rs = conn.execute(sql)

	
%>
<form action="PedidoLista.asp" method="post">

	<label class="titulo">Pedido</label>
	<br>
	<label class="subtitulo">Lista de Pedidos - 
	<%
		select case situacao
		case 0
		%>
		(Pendentes)
		<%
		case 1
		%>
		(Recebidos)
		<%		
		case 2
		%>
		(Cancelados)
		<%		
		end select
	%>
	</label>
	<%
	monta_grid rs, "", "", "", 0, 0, Colunas, Botoes, ""
	rs.close : set rs = nothing



%></form><%
end sub

%>
<!--#include file="../layout/layout.asp"-->
<%call FechaBanco%>
 