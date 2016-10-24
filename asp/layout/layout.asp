<%
select case session("modulo_id")
case 1
	img_modulo = "imagem_modulo_proposta.jpg"
	img_modulo_fundo = "modulo_fundo_proposta.jpg"
	modulo_cor = "#010199"
	modulo = "proposta"
case 2
	img_modulo = "imagem_modulo_comercio.jpg"
	img_modulo_fundo = "modulo_fundo_comercio.jpg"
	modulo_cor = "#cc6600"
	modulo = "comércio"
case 3
	img_modulo = "imagem_modulo_financeiro.jpg"
	img_modulo_fundo = "modulo_fundo_financeiro.jpg"
	modulo_cor = "#006600"
	modulo = "financeiro"
case 4
	img_modulo = "imagem_modulo_usuario.jpg"
	img_modulo_fundo = "modulo_fundo_usuario.jpg"
	modulo_cor = "#990000"
	modulo = "usuários"
end select
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title><%= application("par_1") & " - " & titulo%></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link media="print" rel="alternate" id="LinkPrint" type="text/javascript" href="<%Impressao%>">
<link rel="stylesheet" type="text/css" href="../layout/estilos.css">
<script type="text/javascript">
function IEHoverPseudo(mnu) {
	if(document.getElementById(mnu) != null) {
		var navItems = document.getElementById(mnu).getElementsByTagName("li");
		
		for (var i=0; i<navItems.length; i++) {
			if(navItems[i].className == "menuparent") {
				navItems[i].onmouseover=function() { this.className += " over"; }
				navItems[i].onmouseout=function() { this.className = "menuparent"; }
			}
		}
	}

}

</script>

<style type="text/css">
td img {display: block;}
.mnu {
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-size: 14px;
	color: #333333;
	background-color: #FFFFFF;
	text-align: center;
	vertical-align: middle;
}
.btnModulo:link{
	text-decoration: none;
	color: #333333;
}
.btnModulo:visited{
	text-decoration: none;
	color: #333333;
}
.btnModulo:hover{
	text-decoration: underline;
	color: #333333;
}
.btnModulo:active{
	text-decoration: none;
	color: #333333;
}
.modulo_atual{
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-size: 24px;
	font-weight: bold;
	color: #CCCCCC;
	text-align: center;
	vertical-align: middle;	
}
.mnu_separador{
	border-right-width: 1px;
	border-right-style: dotted;
	border-right-color: #999999;	
}
</style>
</head>
<body bgcolor="#666666" onload="IEHoverPseudo('mnu1'); IEHoverPseudo('mnu2')">
<table border="0" cellpadding="0" cellspacing="0" width="936">
  <tr>
   <td><table align="left" border="0" cellpadding="0" cellspacing="0" width="936">
	  <tr>
	   <td bgcolor="#666666">&nbsp;</td>
	   <td align="right"><img name="layout_r1_c8" src="../layout/images/layout_r1_c8.jpg" width="91" height="32" border="0" id="layout_r1_c8" alt="" /></td>
	  </tr>
	</table></td>
  </tr>
  <tr>
   <td><table align="left" border="0" cellpadding="0" cellspacing="0" width="936">
	  <tr>
	   <td bgcolor="#666666"><img src="../layout/images/spacer.gif" width="12" height="56" border="0" alt="" /></td>
	   <td><img src="../layout/images/borda_superior_esquerda.jpg" width="14" height="56" /></td>
	   <td width="908" bgcolor="#FFFFFF"><table width="100%" border="0">
         <tr>
           <td class="mnu mnu_separador"><a href="../propostas" class="btnModulo">proposta</a></td>
           <td class="mnu mnu_separador"><a href="../comercio" class="btnModulo">com&eacute;rcio</a></td>
           <td class="mnu mnu_separador"><a href="../financeiro" class="btnModulo">financeiro</a></td>
           <td class="mnu"><a href="../usuarios" class="btnModulo">usu&aacute;rio</a></td>
         </tr>
       </table></td>
	   <td><img src="../layout/images/borda_superor_direita.jpg" width="14" height="56" /></td>
	   <td bgcolor="#666666"><img src="../layout/images/spacer.gif" width="16" height="56" border="0" alt="" /></td>
	  </tr>
	</table></td>
  </tr>
  <tr>
   <td><table align="left" border="0" cellpadding="0" cellspacing="0" width="936">
	  <tr>
	   <td bgcolor="<%=modulo_cor%>" width="12">&nbsp;</td>
	   <td  bgcolor="<%=modulo_cor%>"><img src="../layout/images/<%=img_modulo%>" /></td>
	   <td bgcolor="<%=modulo_cor%>" width="242"><!--#include file="menu.asp" --></td>
	   <td bgcolor="<%=modulo_cor%>" background="../layout/images/<%=img_modulo_fundo%>" width="416" class="modulo_atual"> {<%=modulo%>}</td>
	   <td bgcolor="<%=modulo_cor%>" width="12">&nbsp;</td>
	  </tr>
	</table></td>
  </tr>
  <tr>
   <td><table align="left" border="0" cellpadding="0" cellspacing="0" width="936">
	  <tr>
	   <td valign="top"><table align="left" border="0" cellpadding="0" cellspacing="0" width="12">
		  <tr>
		   <td valign="top"><img name="layout_r4_c1" src="../layout/images/layout_r4_c1.jpg" width="12" height="12" border="0" id="layout_r4_c1" alt="" /></td>
		  </tr>
		  <tr>
		   <td bgcolor="#666666"><img src="../layout/images/spacer.gif" width="12" height="408" border="0" alt="" /></td>
		  </tr>
		</table></td>
	   <td valign="top" background="../layout/images/mais_opcoes_fundo.jpg">
	   <table align="left" border="0" cellpadding="0" cellspacing="0" width="250">
		  <tr>
		   <td><img name="mais_opcoes_cabecalho" src="../layout/images/mais_opcoes_cabecalho.jpg" width="250" height="79" border="0" id="mais_opcoes_cabecalho" alt="" /></td>
		  </tr>
		  <tr>
		 	 <td style="padding-left:10px">
		 	   <%auxiliar%>		 	   </td>
		  </tr>
		  
		</table>
		</td>
	   <td valign="top" bgcolor="#FFFFFF"><table align="left" border="0" cellpadding="0" cellspacing="0" width="658">
		  <tr>
		   <td><img name="layout_r4_c4" src="../layout/images/layout_r4_c4.jpg" width="658" height="12" border="0" id="layout_r4_c4" alt="" /></td>
		  </tr>
		  <tr>
		  	<td><div style="margin:5px"><%principal%></div></td>
			</tr>
		  
		</table>
		
		</td>
	   <td valign="top"><table align="left" border="0" cellpadding="0" cellspacing="0" width="16">
		  <tr>
		   <td valign="top"><img name="layout_r4_c9" src="../layout/images/layout_r4_c9.jpg" width="16" height="12" border="0" id="layout_r4_c9" alt="" /></td>
		  </tr>
		  <tr>
		   <td bgcolor="#666666"><img src="../layout/images/spacer.gif" width="16" height="408" border="0" alt="" /></td>
		  </tr>
		</table></td>
	  </tr>
	</table></td>
  </tr>
  <tr>
   <td><img name="layout_r8_c1" src="../layout/images/layout_r8_c1.jpg" width="936" height="11" border="0" id="layout_r8_c1" alt="" /></td>
  </tr>
  <tr>
   <td><table align="left" border="0" cellpadding="0" cellspacing="0" width="936">
	  <tr>
	   <td bgcolor="#666666"><img src="../layout/images/spacer.gif" width="12" height="15" border="0" alt="" /></td>
	   <td width="936"><img src="../layout/images/borda_rodape.jpg" width="908" height="15" /></td>
	   <td bgcolor="#666666"><img src="../layout/images/spacer.gif" width="16" height="15" border="0" alt="" /></td>
	  </tr>
	</table></td>
  </tr>
</table>
</body>
</html>
