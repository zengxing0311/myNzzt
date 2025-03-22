<!--#include file="nzcms_nznet.asp"-->
<% If Request.ServerVariables("SERVER_NAME")<>""&nzcms_bq_bz30&"" and Request.ServerVariables("SERVER_NAME")<>""&nzcms_bq_bz31&""  and left(Request.ServerVariables("SERVER_NAME"),7)<>""&nzcms_bq_bz32&"" Then %><% If nzcms_a_sqok="" or nzcms_a_sqok=1 Then 
response.Redirect(""&nzcms_bq_ts3&"")
response.end
end if
%><% End If %>
<% conn_a.close:set conn_a=nothing %><% conn_by.close:set conn_by=nothing %><% conn_dna.close:set conn_dna=nothing %>
