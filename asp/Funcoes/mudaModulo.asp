<%
sub mudaModulo(pagina)
	if inStr(lcase(pagina),"comercio") > 0 then
		session("modulo_nome") = "Com�rcio"
		session("modulo_id") = 2
		session("modulo_pasta") = "comercio"
	elseif inStr(lcase(pagina),"financeiro") > 0 then
		session("modulo_nome") = "Financeiro"
		session("modulo_id") = 3
		session("modulo_pasta") = "financeiro"
	elseif inStr(lcase(pagina),"propostas") > 0 then
		session("modulo_nome") = "Propostas"
		session("modulo_id") = 1
		session("modulo_pasta") = "Propostas"
	elseif  inStr(lcase(pagina),"usuarios") > 0 then
		session("modulo_nome") = "Usu�rios"
		session("modulo_id") = 4
		session("modulo_pasta") = "usuarios"	
	end if
end sub



%>