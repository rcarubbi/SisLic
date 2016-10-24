<%sub monta_grid(Dados, acaoNome, width, height, SelTipo, SelIndex, Colunas, Botoes, OutrosParametros)
	'dados = recordset ou matriz que vai conter as informações
	'width = largura do grid
	'heigth = altura do grid
	'SelTipo = o tipo de seleção (0 = option, 1 = check)
	'SelIndex = o indice do campo no banco de dados ou na matriz que vai ser o "ID"
	'Colunas:
		'duas dimensões
		'a primeira é a coluna, a segunda é os itens
		'coluna(0, x) = descrição da coluna
		'coluna(1, x) = field da coluna (o indice do campo no banco de dados ou na matriz)
		'coluna(2, x) = largura da coluna (em string % ou px)
		'coluna(3, x) = se soma
		'coluna(4, x) = se ordena
	'Botoes
		'duas dimensões
		'a primeira é a coluna, a segunda é os itens
		'botao(0, x) = exibe o botao
		'botao(1, x) = descrição do botão
		'botao(2, x) = acao do botao
		'botao(3, x) = se para sua ação ser executada, um item tem que estar ou não marcado		
		
	'outros Parametros = parametros que podem ser implementados no futuro, é uma matriz de uma dimensão
	
	if width = "" then width = "550px"
	if height = "" then height = "255px"
	if typename(SelTipo) <> "Integer" then SelTipo = 0
	if typename(SelIndex) <> "Integer" then SelIndex = 0
	if acaoNome = "" then acaoNome = "acao"
	
	'outros parametros (ainda não implementado)
	if isArray(outrosParametros) then
	
	end if
	
	
	
	'validando o tipo de dados
	if typename(dados) = "Recordset" then
		if not dados.eof  then
			mDados = dados.getRows
			
		end if
	elseif isArray(dados) then
		mDados = dados
	end if
	
	if not isArray(mdados) then
		redim mDados(2,-1)
	end if
	
	if not isArray(botoes) then
		redim botoes (3, 3)
	end if

	'só deixa a visibilidade padrao se não foi passando a matriz corretamente
	if typename(botoes(0,0)) <> "Boolean" then	botoes(0,0) = true
	botoes(1,0) = "Adicionar"
	botoes(2,0) = "add"
	botoes(3,0) = false

	if typename(botoes(0,0)) <> "Boolean" then	botoes(0,1) = true
	botoes(1,1) = "Editar"
	botoes(2,1) = "edit"
	botoes(3,1) = true

	if typename(botoes(0,0)) <> "Boolean" then	botoes(0,2) = true
	botoes(1,2) = "Deletar"
	botoes(2,2) = "del"
	botoes(3,2) = true

	if typename(botoes(0,0)) <> "Boolean" then	botoes(0,3) = true
	botoes(1,3) = "Recuperar"
	botoes(2,3) = "rec"
	botoes(3,3) = true
	

	
	
	dim vcor, cor, i, z, soma
	
	'matriz responsavel por pegar os totais
	redim soma (ubound(colunas,2))
	
	'valor padrao para as colunas
	for i = 0 to ubound(colunas,2)
		if colunas(0,i) = "" then colunas(0,i) = "indefinido"
		if typename(colunas(1,i)) <> "Integer" then colunas(1,i) = i
		if typename(colunas(3,i)) <> "Boolean" then colunas(3,i) = false
		if typename(colunas(4,i)) <> "Boolean" then colunas(4,i) = false
	next
	%>
	
	<BR>
	<center>
	
	<script language="javascript">
	
		function ExecutaAcao(acao, obrigatorio) {
			var ok = false;
			var chk = false;
		
	
			
			if (obrigatorio == false) {
				ok = true;
			} else if (VerificaSelecionado()) {
				if (acao == "del") {
					if (confirm("Tem certeza?")) {
						ok = true;
					}
				} else {
					ok = true;
				}
			}

			
			if (ok) {
				document.getElementById('<%=acaoNome%>').value = acao;
				return true;
			} else {
				return false;
			}
		}
		
		
		function VerificaSelecionado() {
			var i = document.getElementsByName("optLinha");
			var z = 0;
			
			
			for ( z = 0 ; z<i.length; z ++) {
				if (i(z).checked == true) {
					return true
				}
			}
			return false
			
		}
		
	
	
	</script>
	<!--
	<style>
	/* Locks the left column */
	td.locked, th.locked {
	
	position:relative;
	cursor: default;
	/*IE5+ only*/
	left: expression(document.getElementById("DataGrid").scrollLeft-1);
	}
	
	/* Locks table header */
	th {
	position:relative;
	cursor: default;
	/*IE5+ only*/
	left: expression(document.getElementById("DataGrid").scrollLeft-1);
	z-index: 10;
	}
	
	</style>
	-->
	<input type="hidden" name="<%=acaoNome%>" id="<%=acaoNome%>" value="">
	
	<div style="border-width:1px;width:<%=width%>;height:30px;" align="right">
		<table>
			<tr>
			<%
			for i = 0 to ubound(botoes,2) %>
				<%if botoes(0, i) = true then%>
					<td>
						<Input type="submit" name="cmd<%=i%>" class="botao" value="<%=botoes(1, i)%>" onClick="javascript:return ExecutaAcao('<%=botoes(2,i)%>', <%=iif(botoes(3,i)=true, "true", "false")%>);">
					</td>
				<%end if%>
			<%next%>
			</tr>
		</table>
	</div>
	<div id="DataGrid" style="overflow:auto;width:<%=width%>;height:<%=height%>;border:1px solid #CCCCCC;background-color:#cccccc;">
	
	<table align="center" width="100%" cellpadding="0" cellspacing="0" border="1" style="border-collapse:collapse;padding:0px 10px 0px 10px " bordercolor="#999999">
		<tr>
		
			<th  width="1%" bgcolor="#FFFFFF">&nbsp;</th>
			
			<%
			'response.Write(ubound(mdados,1))
			'response.End()
			for i = 0 to ubound(colunas,2)
			
				%>
					<th class="cabecalho" nowrap bgcolor="#FFFFFF" width="<%=colunas(2,i)%>" align="<%if ubound(mDados,2) > -1 then response.Write(setAlign(mDados(colunas(1,i),0)))%>">
						<%if colunas(4, i) = true then 'com link de ordenação (por enquanto nao faz automatico)%>
							<a href=""><%=colunas(0,i)%></a>
						<%else%>
							<%=colunas(0,i)%>
						<%end if%>
					</th>
				<%
			 next%>
			
		</tr>
	<%'response.write ubound(mdados,2)
	'response.End()
		for i = 0 to ubound(mDados,2)
			if vcor then
				cor = "#ffffff"
			else
				cor = "#eeeeee"
			end if
		%>
			<tr bgcolor=<%=cor%>>
			<%if selTipo = 0 then%>
				<td>
					<input type="radio" name="optLinha" id="i_<%=mDados(selIndex,i)%>" value="<%=mDados(selIndex,i)%>" <%if cStr(request.form("optLinha")) = cStr(mDados(selIndex,i)) then response.write "Checked"%>>
				</td>
			<%elseif selTipo = 1 then%>
				<td>
					<input type="checkbox" name="chk<%=mDados(selIndex,i)%>" id="i_<%=mDados(selIndex,i)%>" value="1">
				</td>		
			<%end if%>
			<%for z=0 to ubound(colunas,2)%>
			<label for="i_<%=mDados(selIndex,i)%>">
				<td  nowrap align="<%=setAlign(mDados(colunas(1,z),i))%>" class="texto" style="padding-left:5px; cursor:hand;">
						<%=formata_campo(mDados(colunas(1,z),i))%>
				</td>
				<% 
				if colunas(3, z) = true then 'se soma, armazena a soma na variavel soma
					soma(z) = soma(z) + cDbl(mDados(colunas(1,z),i))
 				end if
					
				%>
			</label>
			<%next%>
			</tr>
		<%
			vcor = not vcor
		next
		%>
		<tr bgcolor="#FFFFFF">
			<td>&nbsp;</td>		
		<%
		for z=0 to ubound(colunas,2)%>
			<td align="<%=SetAlign(soma(z))%>" class="cabecalho">
				<%
				if colunas(3,z) = true then 'se soma, exibe o resultado
					response.write(formata_campo(soma(z)))
				end if
				%>
				
				
			</td>		
		<%
		next
		%>
		</tr>
	</table>
	</div>
	</center>
	<BR>

<%end sub%>

<%

function setAlign(valor)
	dim retorno, tipo
	
	err.number = 0
	
	on error resume next
		
	tipo = typename(valor)
	
	if err.number > 0 then 'se o tipo não foi reconhecido
		tipo = ""
	end if
	
	select case tipo
		case "Integer" 'int
			retorno = "Right"
		case "String" 'string
			retorno = "Left"
		case "Boolean" 'boolean
			retorno = "center"
		case "Date" 'data hora
			retorno = "left"
		case "Double"
			retorno = "Right"
		case "Currency"
			retorno = "Right"
		case else
			retorno = "Right"
	end select
	setAlign=retorno
end function%>