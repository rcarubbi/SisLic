<%
dim i, z

m = session("usuario_mnu")

if isArray(m) then
	%>
<link href="../layout/menu.css" rel="stylesheet" type="text/css" />

	<table cellpadding="3" border="0" cellspacing="0" width="90%" align="center" class="menuFonte" >
		<tr>
			<td valign="top" align="left"  width="50%">
				<%	
				response.Write("<ul id=mnu1 class=menuEstilo style='position:relative;z-index:1'>")
				MontaMenu 0, 0
				response.Write("</ul>")
				%>
			</td>
		</tr>
	</table>
	
	<%

end if


'id_superior = ao item que é pai do item atual, caso 0, então não tem pai (root)
'nivel = ao nivel de menu (quantos pais ele tem)
Function MontaMenu(id_superior, nivel)

	dim t, i, z
	
	t = 0
	
	if id_superior <> 0 then
		response.Write("<ul>")
	end if
	
	for i = 0 to ubound(m,2)
	
		if cInt(m(3,i)) = cInt(id_superior) and cInt(m(4,i)) = cInt(session("modulo_id")) then
			if t=4 then
				response.Write("</ul></td><td vAlign=top align=left>")
				response.Write("<ul id=mnu2 class=menuEstilo style='position:relative'>")
						
			end if
			
			%>
			<li  <%if temFilhos(m(0,i)) then response.Write("class='menuparent'")%>>			
				<a  href="<%
				if m(5, i) & "" = "" then
					response.Write("#")
				else
					response.Write("../" & session("modulo_pasta") & "/" & m(5,i))
				end if
							
				%>"><%=m(1,i)%></a>
				<%		
			
				MontaMenu m(0,i), nivel + 1 
				%>
			</li>
			<%

			t = t + 1
		end if	
	next
	response.Write("</ul>")
	
end function

private function temFilhos(pai)
	dim i
	
	for i = 0 to ubound(m,2)
		if cInt(m(3,i)) = cInt(pai) and cInt(m(4,i)) = cInt(session("modulo_id")) then
			temFilhos = true
			exit function
		end if
	next
	temFilhos = false
end function
%>
