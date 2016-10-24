<%
'FCP = FluxoCaixaPrevisao
sub FCPDeleta(id)
	dim sql
	
	sql = "DELETE FROM FluxoCaixaPrevisao where id=" & id
	
	conn.execute(SQL)
end sub

sub FluxoCaixaLancar(idRelacional, CodPlanoConta, Referente, Valor)
	
	if idRelacional = "" then idRelacional = "NULL"
	
	sql = "INSERT INTO FluxoCaixa" & _
           "(CodPlanoConta" & _
           ",Referente" & _
           ",Valor" & _
           ",DataLancamento" & _
           ",idRelacional)" & _
     "VALUES" & _
           "(" & CodPlanoConta & _
           ",'" & Referente & "'" & _
           "," & Valor & _
           ",getDate()" & _
		   "," & idRelacional & ")"
		 
	conn.execute(SQL)

end sub

sub FluxoCaixaCancelar(id, motivo)
	sql = "UPDATE FluxoCaixa " & _
				"SET Cancelado = 1 " & _
				  ",DataCancelamento = getDate() " & _
				  ",MotivoCancelamento = '" & motivo & "' " & _
			 "WHERE id = " & id
	conn.execute(SQL)
end sub

%>