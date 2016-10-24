<%
'deleta itens antigos deste usuario
sub tempInicializa
	sql = "DELETE from tempEmpItem where SessionID='" & session.SessionID & "'"
	
	conn.execute(SQL)
	
	sql = "DELETE from tempEmpItemEstoque where SessionID='" & session.SessionID & "'"
	
	conn.execute(SQL)
	
	sql = "DELETE from tempEmpItemPedido where SessionID='" & session.SessionID & "'"
	
	conn.execute(SQL)
end sub
'grava os itens selecionados em uma tabela temporaria,
'se não selecionou nenhum
'retorna "0" como erro
'se foi tudo ok
'retorna "1"
'o parametro itens tem que ter o seguinte formato "<item1>, <item2>, <item3>"
Function tempGravaItensGanhos(itens, idProcesso)
	dim sql, rs
	
	if trim(itens) = "" then
		tempGravaItensGanhos = "0"
		exit function
	end if
	
	itens = replace(itens, ", ", "','")
	itens = "'" & itens & "'"
	
	
	
	sql = "SELECT * from PropostaItem where id IN(" & itens & ")"

	set rs = conn.execute(SQL)
	
	while not rs.eof
	
		sql = "INSERT INTO tempEmpItem" & _
			   "(idPropostaItem" & _
			   ",Especificacao" & _
			   ",Unidade" & _
			   ",Quantidade" & _
			   ",ValorUnitario" & _
			   ",SessionID)" & _
		 "VALUES" & _
			   "(" & rs("id") & " " & _
			   ",'" & rs("Especificacao") & "' " & _
			   ",'" & rs("Unidade") & "' " & _
			   "," & replace(rs("Quantidade"), ",", ".") & " " & _
			   "," & replace(rs("ValorUnitario"), ",",".") & " " & _
			   ",'" & session.SessionID & "')"
			   
		conn.execute(SQL)
		
		rs.movenext
		
	wend
	
	rs.close
	set rs = nothing
	
	tempGravaItensGanhos = "1"
end function

'retorna um recordset contendo todos os itens ganhos selecionados que já foram enviado para a 
'tabela temporaria atrave da funcao tempGravaItensGanhos
function tempGetItensGanhos
	dim rs
	
	set rs = server.CreateObject("ADODB.recordset")

	SQL = "SELECT * from tempEmpItem where sessionID = '" & session.SessionID & "'"
	
	
	rs.open sql, conn, 1

	set tempGetItensGanhos = rs

end function

'salva os subitemPedido na tabela temporaria
function tempGravaSubItensPedido(idPropostaItem, idFornecedor, descricao, qtProduto, qtPedido, vlUnitario)
	dim sql

'	sql = "DELETE FROM tempEmpItemPedido where NumItem='" & NumItem & "' AND sessionID='" & session.SessionID & "'"
'	conn.execute(SQL)
	
	sql = "INSERT into tempEmpItemPedido(idPropostaItem, " & _
			"idFornecedor, " & _	
			"descricao, " & _
			"qtProduto, " & _
			"qtPedido, " & _
			"ValorUnitario, " & _
			"SessionID " & _			
			") Values (" & _
			"'" & replace(idPropostaItem, "'", "''") & "', " & _
			idFornecedor & ", " & _			
			"'" & replace(descricao, "'", "''") & "', " & _
			qtProduto & ", " & _
			qtPedido & ", " & _
			replace(vlUnitario,",",".") & ", " & _
			"'" & Session.SessionID & "')"
			
	conn.execute(SQL)
		
	
end function

'salva os subitemPedido na tabela temporaria
function tempEditaSubItensPedido(ID, idFornecedor, descricao, qtProduto, qtPedido, vlUnitario)
	dim sql

	sql = "UPDATE tempEmpItemPedido" & _
			   " SET descricao = '" & descricao & "'" & _
			   ",qtPedido = " & qtPedido & _
			   ",qtProduto = " & qtProduto & _
			   ",ValorUnitario = " & replace(vlUnitario,",",".") & _
			   ",idFornecedor = " & idFornecedor & _
			   ",SessionID = '" & session.SessionID & "'" & _
			   "WHERE ID=" & ID

	conn.execute(SQL)
		
	
end function


'retorna um recordset contendo todos os subitens ganhos selecionados que já foram enviado para a 
'tabela temporaria atrave da funcao tempGravaSubItensPedido
'se o ID for vazio, então não filtra pelo ID
function tempGetSubItensPedidos(idPropostaItem, ID)
	dim rs
	
	set rs = server.CreateObject("ADODB.recordset")

	SQL = "SELECT * from tempEmpItemPedido where sessionID = '" & session.SessionID & "'"
	if idPropostaItem <> "" then
		sql = sql & " and idPropostaItem = '" & idPropostaItem & "'"
	end if
	if ID <> "" then
		sql = sql & " AND id=" & ID
	end if
	rs.open sql, conn, 1

	set tempGetSubItensPedidos = rs

end function

'deleta um determinado subitem
function tempDeletaSubItemPedido(ID)
	dim sql
	
	sql = "DELETE FROM tempEmpItemPedido where id=" & id
	conn.execute(SQL)
end function


'salva os subitemPedido na tabela temporaria
function tempGravaSubItensEstoque(idPropostaItem, idEstoque, qt)
	dim sql
	dim rs
	
	set rs = conn.execute("sp_AtualizaEstoqueTemp(" & idPropostaItem & "," & idEstoque & "," & qt & "," & session.SessionID & ")")
	
	resp = split(rs("resposta"), ";")
	
	if resp(0) = "Erro" then
		tempGravaSubItensEstoque = resp(1)
	end if
	
	
	rs.close
	set rs = nothing
	
end function

'salva os subitemEstoque na tabela temporaria
function tempEditaSubItensEstoque(ID, idEstoque, qt)

dim sql
dim rs


set rs = conn.execute("sp_AtualizaEstoqueTemp(0," & idEstoque & "," & qt & "," & session.SessionID & "," & ID & ")")

resp = split(rs("resposta"), ";")

if resp(0) = "Erro" then
	tempEditaSubItensEstoque = resp(1)
end if


rs.close
set rs = nothing		
	
end function

'retorna um recordset contendo todos os subitens ganhos selecionados que já foram enviado para a 
'tabela temporaria atrave da funcao tempGravaSubItensEstoque
'se o ID for vazio, então não filtra pelo ID
function tempGetSubItensEstoque(idPropostaItem, ID)
	dim rs
	
	set rs = server.CreateObject("ADODB.recordset")

	SQL = "SELECT * from tempEmpItemEstoque where sessionID = '" & session.SessionID & "'"
	if idPropostaItem <> "" then
		sql = sql & " and idPropostaItem = '" & idPropostaItem & "'"
	end if
	if ID <> "" then
		sql = sql & " AND id=" & ID
	end if
	
	rs.open sql, conn, 1

	set tempGetSubItensEstoque = rs

end function

'deleta um determinado subitem
function tempDeletaSubItemEstoque(ID)
	dim sql
	
	sql = "DELETE FROM tempEmpItemEstoque where id=" & id
	conn.execute(SQL)
end function

'verifica quantidade de itens que ainda FALTA o usuario colocar subItens
function tempFaltaItens(idProcesso)
	dim SQL, rs
	
	if idProcesso = "" then idProcesso = 0
	
	sql = "SELECT idPropostaItem from tempEmpItem " & _
		  "WHERE sessionID=" & session.SessionID & _
			" EXCEPT " & _
				"SELECT idPropostaItem from tempEmpItemPedido " & _
				"WHERE sessionID='" & session.SessionID & "' " & _
			" EXCEPT " & _
				"SELECT idPropostaItem from tempEmpItemEstoque " & _
				"WHERE sessionID='" & session.SessionID & "' "				
					'response.Write(sql)
	set rs = conn.execute(SQL)
	if not rs.eof then
		tempFaltaItens = true
	else
		tempFaltaItens = false
	end if
	rs.close
	set rs = nothing

end function

'verifica se houve alteração no estoque
function EstoqueConferencia
	dim rs, sql, erro
	sql = "select count(*) as tot from tempEmpItemEstoque inner join estoque on estoque.id = tempEmpItemEstoque.idEstoque where (estoque.qt - tempEmpItemEstoque.qt) < 0 and sessionID='" & session.SessionID & "'"
	set rs = conn.execute(SQL)
	if isNull(rs("tot")) then
		erro = "Houve alteração do estoque por outro usuario, estoque de produtos atual é insuficiente."
	else
		if rs("tot") > 0 then
			erro = "Houve alteração do estoque por outro usuario, estoque de produtos atual é insuficiente."
		end if
	end if
	rs.close
	set rs = nothing
	
	EstoqueConferencia = erro
end function

'atualiza a tabela Estoque
sub EstoqueAtualizar
	dim sql, rs
	
	sql = "select * from tempEmpItemEstoque where sessionID='" & session.SessionID & "'"
	set rs = conn.execute(SQL)
	while not rs.eof
		sql = "update Estoque set qt = qt - " & rs("qt") & " where id=" & rs("idEstoque")
		conn.execute(SQL)
				
		rs.movenext
	wend
	
	rs.close
	set rs = nothing
	
end sub

'gera o empenho
sub GerarEmpenho(idProcesso, dataEmpenho, valorTotal, prazo, validade, observacoes)
	
	sql = "INSERT INTO Empenho" & _
           "(idProcesso " & _
		   ",DataEmpenho " & _
           ",ValorTotal " & _
           ",Prazo " & _
           ",Validade " & _
           ",Observacao) " & _
     	"VALUES " & _
           "(" & idProcesso & _
		   ",'" & dataEmpenho & "' " & _
           "," & replace(replace(valorTotal,".",""),",",".") & _
           "," & prazo & _
           "," & validade & _
           ",'" & replace(observacoes, "'", "''") & "')"
		' response.write sql
		' response.End()
	conn.execute(SQL)
	
	'há uma trigger que ao cadastrar um empenho, ela atualiza a proposta como ganhou 
	'automaticamente
	
end sub

'sub item da proposta
function PropostaSubItem(idPropostaItem, quantidade, descricao)
	dim sql, rs
	
	sql = "Insert into PropostaSubItem (idPropostaItem, " & _
						"quantidade, " & _
						"descricao) " & _
						"Values (" & _
						idPropostaItem & ", " & _
						quantidade & ", " & _
						"'" & replace(descricao, "'", "''") & "') "

	conn.execute(SQL)
	sql = "SELECT max(id) as maxID from PropostaSubItem"
	set rs = conn.execute(SQL)
	
	PropostaSubItem = rs("maxID")
	
	rs.close
	set rs = nothing
	
end function
%>