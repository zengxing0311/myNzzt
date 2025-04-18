
<% 
set conn_a = Server.CreateObject("adodb.connection")
connstr="Provider=Microsoft.Jet.OLEDB.4.0;Jet OLEDB:dataBase Password=**********;data Source=" & Server.MapPath("nzcms_a.asp")
conn_a.open connstr
search=Request.QueryString("search")
set rs_web=server.CreateObject("adodb.recordset")
sql="select * from nzcms order by id desc"

rs_web.open sql,conn_a,1,1
if rs_web.recordcount<>0 then
nzcms_a_web=rs_web("web")
nzcms_a_www=rs_web("www")
nzcms_a_time=rs_web("dataandtime")
nzcms_a_userid=rs_web("userid")
nzcms_a_cdid=rs_web("cdid")
nzcms_a_sqok=rs_web("sqok")

nzcms_a_he=rs_web.recordcount

nzcms_a_he8=3
nzcms_a_he888=3
nzcms_a_he8888=4
nzcms_a_he88888=5
nzcms_a_he888888=6
nzcms_a_he8888888=7
nzcms_a_he88888888=8
nzcms_a_he888888888=9

end if
rs_web.close
set rs_web=nothing
set conn_by = Server.CreateObject("adodb.connection")
connstr="Provider=Microsoft.Jet.OLEDB.4.0;Jet OLEDB:dataBase Password=**********;data Source=" & Server.MapPath("nzcms_bq.asp")
conn_by.open connstr
search=Request.QueryString("search")
set rs_web=server.CreateObject("adodb.recordset")
sql="select * from nzcms order by id desc"

rs_web.open sql,conn_by,1,1

nzcms_bq_www=rs_web("www")
nzcms_bq_ts=rs_web("ts")
nzcms_bq_ts2=rs_web("ts2")
nzcms_bq_ts3=rs_web("ts3")
nzcms_bq_ts4=rs_web("ts4")
nzcms_bq_ts5=rs_web("ts5")
nzcms_bq_ts6=rs_web("ts6")
nzcms_bq_ts7=rs_web("ts7")
nzcms_bq_ts8=rs_web("ts8")
nzcms_bq_ts9=rs_web("ts9")
nzcms_bq_ts10=rs_web("ts10")
nzcms_bq_userid=rs_web("user-id")
nzcms_bq_userqq=rs_web("user-qq")



nzcms_bq_tj=rs_web("tj")
nzcms_bq_51tj=rs_web("51tj")
nzcms_bq_nzcms=rs_web("nzcms")
nzcms_bq_by_images=rs_web("by_images")
nzcms_bq_nzcms_name=rs_web("nzcms_name")
nzcms_bq_bq=rs_web("bq")
nzcms_bq_year=rs_web("year")
nzcms_bq_month=rs_web("month")
nzcms_bq_by=rs_web("by")
nzcms_bq_v=rs_web("v")
nzcms_bq_sodna=rs_web("sodna")
nzcms_bq_sov=rs_web("sov")
nzcms_a_he88=rs_web("heshu")
nzcms_bq_bz1=rs_web("bz1")
nzcms_bq_bz2=rs_web("bz2")
nzcms_bq_bz3=rs_web("bz3")
nzcms_bq_bz4=rs_web("bz4")
nzcms_bq_bz5=rs_web("bz5")
nzcms_bq_bz6=rs_web("bz6")
nzcms_bq_bz7=rs_web("bz7")
nzcms_bq_bz8=rs_web("bz8")
nzcms_bq_bz9=rs_web("bz9")
nzcms_bq_bz10=rs_web("bz10")
nzcms_bq_bz11=rs_web("bz11")
nzcms_bq_bz12=rs_web("bz12")
nzcms_bq_bz13=rs_web("bz13")
nzcms_bq_bz14=rs_web("bz14")
nzcms_bq_bz15=rs_web("bz15")
nzcms_bq_bz16=rs_web("bz16")
nzcms_bq_bz17=rs_web("bz17")
nzcms_bq_bz18=rs_web("bz18")
nzcms_bq_bz19=rs_web("bz19")
nzcms_bq_bz20=rs_web("bz20")
nzcms_bq_bz21=rs_web("bz21")
nzcms_bq_bz22=rs_web("bz22")
nzcms_bq_bz23=rs_web("bz23")
nzcms_bq_bz24=rs_web("bz24")
nzcms_bq_bz25=rs_web("bz25")
nzcms_bq_bz26=rs_web("bz26")
nzcms_bq_bz27=rs_web("bz27")
nzcms_bq_bz28=rs_web("bz28")
nzcms_bq_bz29=rs_web("bz29")
nzcms_bq_bz30=rs_web("bz30")
nzcms_bq_bz31=rs_web("bz31")
nzcms_bq_bz32=rs_web("bz32")
nzcms_bq_bz33=rs_web("bz33")

rs_web.close
set rs_web=nothing


set conn_dna = Server.CreateObject("adodb.connection")
connstr="Provider=Microsoft.Jet.OLEDB.4.0;Jet OLEDB:dataBase Password=!@#$%^&*;data Source=" & Server.MapPath("nzcms_dna.asp")
conn_dna.open connstr
search=Request.QueryString("search")
set rs_web=server.CreateObject("adodb.recordset")
sql="select * from nzcms order by id desc"

rs_web.open sql,conn_dna,1,1
nzcms_dna_user_id=rs_web("user_id")
nzcms_dna_user_zs=rs_web("user_zs")
rs_web.close
set rs_web=nothing

%>

