<%
'此版本为免费下载的（试用版本）源码加密处理过
'购买正式版系统提供源码（支持二次开发，修改）
'系统服务商：金华市宁志网络科技有限公司
'公司网址：http://www.ningzhi.net 
'客服QQ号：122470827 电话：18867186998 （同微信号）刘经理
'全国统一售后电话：0579-83938663 
'支持正版：软件著作权登记号：2012SR040914
'官方平台：购买系统请上金华市宁志网络公司官网平台是唯一购买渠道
'本公司无第三方代销售平台请认准,以免上当受骗！！！
'请勿修改下列任何代码,以保证程序正常运行·感谢您体验宁志产品，宁志有您更精彩！
'TM:2023-10-16  21:31:35

if instr(request.form(ChrW(107) & ChrW(101) & ChrW(121) & ChrW(119) & ChrW(111) & ChrW(114) & ChrW(100)),ChrW(33))>0 or instr(request.form(ChrW(107) & ChrW(101) & ChrW(121) & ChrW(119) & ChrW(111) & ChrW(114) & ChrW(100)),ChrW(64))>0 or instr(request.form(ChrW(107) & ChrW(101) & ChrW(121) & ChrW(119) & ChrW(111) & ChrW(114) & ChrW(100)),ChrW(45))>0 or instr(request.form(ChrW(107) & ChrW(101) & ChrW(121) & ChrW(119) & ChrW(111) & ChrW(114) & ChrW(100)),ChrW(97) & ChrW(110) & ChrW(100))>0 or instr(request.form(ChrW(107) & ChrW(101) & ChrW(121) & ChrW(119) & ChrW(111) & ChrW(114) & ChrW(100)),ChrW(67) & ChrW(82))>0 or instr(request.form(ChrW(107) & ChrW(101) & ChrW(121) & ChrW(119) & ChrW(111) & ChrW(114) & ChrW(100)),ChrW(108) & ChrW(102))>0 or instr(request.form(ChrW(107) & ChrW(101) & ChrW(121) & ChrW(119) & ChrW(111) & ChrW(114) & ChrW(100)),ChrW(115) & ChrW(99) & ChrW(114) & ChrW(105) & ChrW(112) & ChrW(116))>0  or instr(request.form(ChrW(107) & ChrW(101) & ChrW(121) & ChrW(119) & ChrW(111) & ChrW(114) & ChrW(100)),ChrW(100) & ChrW(111) & ChrW(99) & ChrW(117) & ChrW(109) & ChrW(101) & ChrW(110) & ChrW(116))>0 or instr(request.form(ChrW(107) & ChrW(101) & ChrW(121) & ChrW(119) & ChrW(111) & ChrW(114) & ChrW(100)),ChrW(101) & ChrW(118) & ChrW(97) & ChrW(108))>0 or instr(request.form(ChrW(107) & ChrW(101) & ChrW(121) & ChrW(119) & ChrW(111) & ChrW(114) & ChrW(100)),ChrW(35))>0 or instr(request.form(ChrW(107) & ChrW(101) & ChrW(121) & ChrW(119) & ChrW(111) & ChrW(114) & ChrW(100)),ChrW(36))>0 or instr(request.form(ChrW(107) & ChrW(101) & ChrW(121) & ChrW(119) & ChrW(111) & ChrW(114) & ChrW(100)),ChrW(37))>0 or instr(request.form(ChrW(107) & ChrW(101) & ChrW(121) & ChrW(119) & ChrW(111) & ChrW(114) & ChrW(100)),ChrW(94))>0 or instr(request.form(ChrW(107) & ChrW(101) & ChrW(121) & ChrW(119) & ChrW(111) & ChrW(114) & ChrW(100)),ChrW(38))>0 or instr(request.form(ChrW(107) & ChrW(101) & ChrW(121) & ChrW(119) & ChrW(111) & ChrW(114) & ChrW(100)),ChrW(42))>0 or instr(request.form(ChrW(107) & ChrW(101) & ChrW(121) & ChrW(119) & ChrW(111) & ChrW(114) & ChrW(100)),ChrW(40))>0 or instr(request.form(ChrW(107) & ChrW(101) & ChrW(121) & ChrW(119) & ChrW(111) & ChrW(114) & ChrW(100)),ChrW(41))>0 or instr(request.form(ChrW(107) & ChrW(101) & ChrW(121) & ChrW(119) & ChrW(111) & ChrW(114) & ChrW(100)),ChrW(95))>0 or instr(request.form(ChrW(107) & ChrW(101) & ChrW(121) & ChrW(119) & ChrW(111) & ChrW(114) & ChrW(100)),ChrW(47))>0 or instr(request.form(ChrW(107) & ChrW(101) & ChrW(121) & ChrW(119) & ChrW(111) & ChrW(114) & ChrW(100)),ChrW(60))>0 or instr(request.form(ChrW(107) & ChrW(101) & ChrW(121) & ChrW(119) & ChrW(111) & ChrW(114) & ChrW(100)),ChrW(62))>0 or instr(request.form(ChrW(107) & ChrW(101) & ChrW(121) & ChrW(119) & ChrW(111) & ChrW(114) & ChrW(100)),ChrW(47))>0 or instr(request.form(ChrW(107) & ChrW(101) & ChrW(121) & ChrW(119) & ChrW(111) & ChrW(114) & ChrW(100)),ChrW(42))>0 or instr(request.form(ChrW(107) & ChrW(101) & ChrW(121) & ChrW(119) & ChrW(111) & ChrW(114) & ChrW(100)),ChrW(124))>0 or instr(request.form(ChrW(107) & ChrW(101) & ChrW(121) & ChrW(119) & ChrW(111) & ChrW(114) & ChrW(100)),ChrW(43))>0 or instr(request.form(ChrW(107) & ChrW(101) & ChrW(121) & ChrW(119) & ChrW(111) & ChrW(114) & ChrW(100)),ChrW(46))>0 or instr(request.form(ChrW(107) & ChrW(101) & ChrW(121) & ChrW(119) & ChrW(111) & ChrW(114) & ChrW(100)),ChrW(61))>0 or instr(request.form(ChrW(107) & ChrW(101) & ChrW(121) & ChrW(119) & ChrW(111) & ChrW(114) & ChrW(100)),ChrW(44))>0  or instr(request.form(ChrW(107) & ChrW(101) & ChrW(121) & ChrW(119) & ChrW(111) & ChrW(114) & ChrW(100)),ChrW(39))>0 then
response.Write ChrW(60) & ChrW(115) & ChrW(99) & ChrW(114) & ChrW(105) & ChrW(112) & ChrW(116) & ChrW(32) & ChrW(76) & ChrW(65) & ChrW(78) & ChrW(71) & ChrW(85) & ChrW(65) & ChrW(71) & ChrW(69) & ChrW(61) & ChrW(39) & ChrW(106) & ChrW(97) & ChrW(118) & ChrW(97) & ChrW(115) & ChrW(99) & ChrW(114) & ChrW(105) & ChrW(112) & ChrW(116) & ChrW(39) & ChrW(62) & ChrW(97) & ChrW(108) & ChrW(101) & ChrW(114) & ChrW(116) & ChrW(40) & ChrW(39) & ChrW(-29705) & ChrW(19981) & ChrW(-30335) & ChrW(-28781) & ChrW(20837) & ChrW(-26786) & ChrW(27861) & ChrW(23383) & ChrW(31526) & ChrW(-255) & ChrW(39) & ChrW(41) & ChrW(59) & ChrW(104) & ChrW(105) & ChrW(115) & ChrW(116) & ChrW(111) & ChrW(114) & ChrW(121) & ChrW(46) & ChrW(103) & ChrW(111) & ChrW(40) & ChrW(45) & ChrW(49) & ChrW(41) & ChrW(59) & ChrW(60) & ChrW(47) & ChrW(115) & ChrW(99) & ChrW(114) & ChrW(105) & ChrW(112) & ChrW(116) & ChrW(62)
response.end
end if
Function CheckInput()
Dim Fy_Post,Fy_Get,Fy_In,Fy_Inf,Fy_Xh,Fy_db,Fy_dbstr,Kill_IP,WriteSql
Fy_In = ChrW(39) & ChrW(124) & ChrW(64) & ChrW(124) & ChrW(35) & ChrW(124) & ChrW(59) & ChrW(124) & ChrW(38) & ChrW(124) & ChrW(42) & ChrW(124) & ChrW(37) & ChrW(124) & ChrW(36) & ChrW(124) & ChrW(63) & ChrW(124) & ChrW(94) & ChrW(124) & ChrW(40) & ChrW(124) & ChrW(41) & ChrW(124) & ChrW(99) & ChrW(114) & ChrW(124) & ChrW(108) & ChrW(102) & ChrW(124) & ChrW(43) & ChrW(124) & ChrW(60) & ChrW(62) & ChrW(124) & ChrW(44) & ChrW(124) & ChrW(46) & ChrW(124) & ChrW(115) & ChrW(99) & ChrW(114) & ChrW(105) & ChrW(112) & ChrW(116) & ChrW(124) & ChrW(100) & ChrW(111) & ChrW(99) & ChrW(117) & ChrW(109) & ChrW(101) & ChrW(110) & ChrW(116) & ChrW(124) & ChrW(101) & ChrW(118) & ChrW(97) & ChrW(108) & ChrW(124) & ChrW(101) & ChrW(120) & ChrW(101) & ChrW(99) & ChrW(124) & ChrW(105) & ChrW(110) & ChrW(115) & ChrW(101) & ChrW(114) & ChrW(116) & ChrW(124) & ChrW(115) & ChrW(101) & ChrW(108) & ChrW(101) & ChrW(99) & ChrW(116) & ChrW(124) & ChrW(100) & ChrW(101) & ChrW(108) & ChrW(101) & ChrW(116) & ChrW(101) & ChrW(124) & ChrW(117) & ChrW(112) & ChrW(100) & ChrW(97) & ChrW(116) & ChrW(101) & ChrW(124) & ChrW(99) & ChrW(111) & ChrW(117) & ChrW(110) & ChrW(116) & ChrW(124) & ChrW(99) & ChrW(104) & ChrW(114) & ChrW(124) & ChrW(109) & ChrW(105) & ChrW(100) & ChrW(124) & ChrW(109) & ChrW(97) & ChrW(115) & ChrW(116) & ChrW(101) & ChrW(114) & ChrW(124) & ChrW(116) & ChrW(114) & ChrW(117) & ChrW(110) & ChrW(99) & ChrW(97) & ChrW(116) & ChrW(101) & ChrW(124) & ChrW(99) & ChrW(104) & ChrW(97) & ChrW(114) & ChrW(124) & ChrW(97) & ChrW(110) & ChrW(100) & ChrW(124) & ChrW(100) & ChrW(101) & ChrW(99) & ChrW(108) & ChrW(97) & ChrW(114) & ChrW(101) & ChrW(124) & ChrW(39) & ChrW(124) & ChrW(59) & ChrW(124) & ChrW(97) & ChrW(110) & ChrW(100) & ChrW(124) & ChrW(40) & ChrW(124) & ChrW(41) & ChrW(124) & ChrW(101) & ChrW(120) & ChrW(101) & ChrW(99) & ChrW(124) & ChrW(105) & ChrW(110) & ChrW(115) & ChrW(101) & ChrW(114) & ChrW(116) & ChrW(124) & ChrW(115) & ChrW(101) & ChrW(108) & ChrW(101) & ChrW(99) & ChrW(116) & ChrW(124) & ChrW(100) & ChrW(101) & ChrW(108) & ChrW(101) & ChrW(116) & ChrW(101) & ChrW(124) & ChrW(117) & ChrW(112) & ChrW(100) & ChrW(97) & ChrW(116) & ChrW(101) & ChrW(124) & ChrW(99) & ChrW(111) & ChrW(117) & ChrW(110) & ChrW(116) & ChrW(124) & ChrW(42) & ChrW(124) & ChrW(99) & ChrW(104) & ChrW(114) & ChrW(124) & ChrW(109) & ChrW(105) & ChrW(100) & ChrW(124) & ChrW(109) & ChrW(97) & ChrW(115) & ChrW(116) & ChrW(101) & ChrW(114) & ChrW(124) & ChrW(116) & ChrW(114) & ChrW(117) & ChrW(110) & ChrW(99) & ChrW(97) & ChrW(116) & ChrW(101) & ChrW(124) & ChrW(99) & ChrW(104) & ChrW(97) & ChrW(114) & ChrW(124) & ChrW(100) & ChrW(101) & ChrW(99) & ChrW(108) & ChrW(97) & ChrW(114) & ChrW(101) & ChrW(124) & ChrW(72) & ChrW(97) & ChrW(110) & ChrW(100) & ChrW(105) & ChrW(99) & ChrW(114) & ChrW(97) & ChrW(102) & ChrW(116) & ChrW(124) & ChrW(53) & ChrW(57) & ChrW(92) & ChrW(50) & ChrW(54) & ChrW(116) & ChrW(105) & ChrW(116) & ChrW(108) & ChrW(101) & ChrW(61) & ChrW(72) & ChrW(97) & ChrW(110) & ChrW(100) & ChrW(105) & ChrW(99) & ChrW(114) & ChrW(97) & ChrW(102) & ChrW(116)
Fy_Inf = split(Fy_In,ChrW(124))
If Request.Form <> "" Then
For Each Fy_Post In Request.Form
For Fy_Xh = 0 To Ubound(Fy_Inf)
If Instr(LCase(Request.Form(Fy_Post)),Fy_Inf(Fy_Xh)) <> 0 Then
Echo ChrW(60) & ChrW(83) & ChrW(99) & ChrW(114) & ChrW(105) & ChrW(112) & ChrW(116) & ChrW(32) & ChrW(76) & ChrW(97) & ChrW(110) & ChrW(103) & ChrW(117) & ChrW(97) & ChrW(103) & ChrW(101) & ChrW(61) & ChrW(74) & ChrW(97) & ChrW(118) & ChrW(97) & ChrW(83) & ChrW(99) & ChrW(114) & ChrW(105) & ChrW(112) & ChrW(116) & ChrW(62) & ChrW(97) & ChrW(108) & ChrW(101) & ChrW(114) & ChrW(116) & ChrW(40) & ChrW(39) & ChrW(-29705) & ChrW(19981) & ChrW(-30335) & ChrW(22312) & ChrW(21442) & ChrW(25968) & ChrW(20013) & ChrW(21253) & ChrW(21547) & ChrW(-26786) & ChrW(27861) & ChrW(23383) & ChrW(31526) & ChrW(-255) & ChrW(39) & ChrW(41) & ChrW(59) & ChrW(104) & ChrW(105) & ChrW(115) & ChrW(116) & ChrW(111) & ChrW(114) & ChrW(121) & ChrW(46) & ChrW(103) & ChrW(111) & ChrW(40) & ChrW(45) & ChrW(49) & ChrW(41) & ChrW(59) & ChrW(60) & ChrW(47) & ChrW(83) & ChrW(99) & ChrW(114) & ChrW(105) & ChrW(112) & ChrW(116) & ChrW(62)
Response.End
End If
Next
Next
End If
If Request.QueryString <> "" Then
For Each Fy_Get In Request.QueryString
For Fy_Xh = 0 To Ubound(Fy_Inf)
If Instr(LCase(Request.QueryString(Fy_Get)),Fy_Inf(Fy_Xh)) <> 0 Then
Echo ChrW(60) & ChrW(83) & ChrW(99) & ChrW(114) & ChrW(105) & ChrW(112) & ChrW(116) & ChrW(32) & ChrW(76) & ChrW(97) & ChrW(110) & ChrW(103) & ChrW(117) & ChrW(97) & ChrW(103) & ChrW(101) & ChrW(61) & ChrW(74) & ChrW(97) & ChrW(118) & ChrW(97) & ChrW(83) & ChrW(99) & ChrW(114) & ChrW(105) & ChrW(112) & ChrW(116) & ChrW(62) & ChrW(97) & ChrW(108) & ChrW(101) & ChrW(114) & ChrW(116) & ChrW(40) & ChrW(39) & ChrW(-29705) & ChrW(19981) & ChrW(-30335) & ChrW(22312) & ChrW(21442) & ChrW(25968) & ChrW(20013) & ChrW(21253) & ChrW(21547) & ChrW(-26786) & ChrW(27861) & ChrW(23383) & ChrW(31526) & ChrW(-255) & ChrW(39) & ChrW(41) & ChrW(59) & ChrW(104) & ChrW(105) & ChrW(115) & ChrW(116) & ChrW(111) & ChrW(114) & ChrW(121) & ChrW(46) & ChrW(103) & ChrW(111) & ChrW(40) & ChrW(45) & ChrW(49) & ChrW(41) & ChrW(59) & ChrW(60) & ChrW(47) & ChrW(83) & ChrW(99) & ChrW(114) & ChrW(105) & ChrW(112) & ChrW(116) & ChrW(62)
Response.End
End If
Next
Next
End If
End Function
Function CheckInput()
Dim Fy_Post,Fy_Get,Fy_In,Fy_Inf,Fy_Xh,Fy_db,Fy_dbstr,Kill_IP,WriteSql
Fy_In = ChrW(39) & ChrW(124) & ChrW(64) & ChrW(124) & ChrW(35) & ChrW(124) & ChrW(59) & ChrW(124) & ChrW(38) & ChrW(124) & ChrW(42) & ChrW(124) & ChrW(37) & ChrW(124) & ChrW(36) & ChrW(124) & ChrW(63) & ChrW(124) & ChrW(94) & ChrW(124) & ChrW(40) & ChrW(124) & ChrW(41) & ChrW(124) & ChrW(99) & ChrW(114) & ChrW(124) & ChrW(108) & ChrW(102) & ChrW(124) & ChrW(43) & ChrW(124) & ChrW(60) & ChrW(62) & ChrW(124) & ChrW(44) & ChrW(124) & ChrW(46) & ChrW(124) & ChrW(115) & ChrW(99) & ChrW(114) & ChrW(105) & ChrW(112) & ChrW(116) & ChrW(124) & ChrW(100) & ChrW(111) & ChrW(99) & ChrW(117) & ChrW(109) & ChrW(101) & ChrW(110) & ChrW(116) & ChrW(124) & ChrW(101) & ChrW(118) & ChrW(97) & ChrW(108) & ChrW(124) & ChrW(101) & ChrW(120) & ChrW(101) & ChrW(99) & ChrW(124) & ChrW(105) & ChrW(110) & ChrW(115) & ChrW(101) & ChrW(114) & ChrW(116) & ChrW(124) & ChrW(115) & ChrW(101) & ChrW(108) & ChrW(101) & ChrW(99) & ChrW(116) & ChrW(124) & ChrW(100) & ChrW(101) & ChrW(108) & ChrW(101) & ChrW(116) & ChrW(101) & ChrW(124) & ChrW(117) & ChrW(112) & ChrW(100) & ChrW(97) & ChrW(116) & ChrW(101) & ChrW(124) & ChrW(99) & ChrW(111) & ChrW(117) & ChrW(110) & ChrW(116) & ChrW(124) & ChrW(99) & ChrW(104) & ChrW(114) & ChrW(124) & ChrW(109) & ChrW(105) & ChrW(100) & ChrW(124) & ChrW(109) & ChrW(97) & ChrW(115) & ChrW(116) & ChrW(101) & ChrW(114) & ChrW(124) & ChrW(116) & ChrW(114) & ChrW(117) & ChrW(110) & ChrW(99) & ChrW(97) & ChrW(116) & ChrW(101) & ChrW(124) & ChrW(99) & ChrW(104) & ChrW(97) & ChrW(114) & ChrW(124) & ChrW(97) & ChrW(110) & ChrW(100) & ChrW(124) & ChrW(100) & ChrW(101) & ChrW(99) & ChrW(108) & ChrW(97) & ChrW(114) & ChrW(101) & ChrW(124) & ChrW(39) & ChrW(124) & ChrW(59) & ChrW(124) & ChrW(97) & ChrW(110) & ChrW(100) & ChrW(124) & ChrW(40) & ChrW(124) & ChrW(41) & ChrW(124) & ChrW(101) & ChrW(120) & ChrW(101) & ChrW(99) & ChrW(124) & ChrW(105) & ChrW(110) & ChrW(115) & ChrW(101) & ChrW(114) & ChrW(116) & ChrW(124) & ChrW(115) & ChrW(101) & ChrW(108) & ChrW(101) & ChrW(99) & ChrW(116) & ChrW(124) & ChrW(100) & ChrW(101) & ChrW(108) & ChrW(101) & ChrW(116) & ChrW(101) & ChrW(124) & ChrW(117) & ChrW(112) & ChrW(100) & ChrW(97) & ChrW(116) & ChrW(101) & ChrW(124) & ChrW(99) & ChrW(111) & ChrW(117) & ChrW(110) & ChrW(116) & ChrW(124) & ChrW(42) & ChrW(124) & ChrW(99) & ChrW(104) & ChrW(114) & ChrW(124) & ChrW(109) & ChrW(105) & ChrW(100) & ChrW(124) & ChrW(109) & ChrW(97) & ChrW(115) & ChrW(116) & ChrW(101) & ChrW(114) & ChrW(124) & ChrW(116) & ChrW(114) & ChrW(117) & ChrW(110) & ChrW(99) & ChrW(97) & ChrW(116) & ChrW(101) & ChrW(124) & ChrW(99) & ChrW(104) & ChrW(97) & ChrW(114) & ChrW(124) & ChrW(100) & ChrW(101) & ChrW(99) & ChrW(108) & ChrW(97) & ChrW(114) & ChrW(101) & ChrW(124) & ChrW(72) & ChrW(97) & ChrW(110) & ChrW(100) & ChrW(105) & ChrW(99) & ChrW(114) & ChrW(97) & ChrW(102) & ChrW(116) & ChrW(124) & ChrW(53) & ChrW(57) & ChrW(92) & ChrW(50) & ChrW(54) & ChrW(116) & ChrW(105) & ChrW(116) & ChrW(108) & ChrW(101) & ChrW(61) & ChrW(72) & ChrW(97) & ChrW(110) & ChrW(100) & ChrW(105) & ChrW(99) & ChrW(114) & ChrW(97) & ChrW(102) & ChrW(116)
Fy_Inf = split(Fy_In,ChrW(124))
If Request.Form <> "" Then
For Each Fy_Post In Request.Form
For Fy_Xh = 0 To Ubound(Fy_Inf)
If Instr(LCase(Request.Form(Fy_Post)),Fy_Inf(Fy_Xh)) <> 0 Then
Echo ChrW(60) & ChrW(83) & ChrW(99) & ChrW(114) & ChrW(105) & ChrW(112) & ChrW(116) & ChrW(32) & ChrW(76) & ChrW(97) & ChrW(110) & ChrW(103) & ChrW(117) & ChrW(97) & ChrW(103) & ChrW(101) & ChrW(61) & ChrW(74) & ChrW(97) & ChrW(118) & ChrW(97) & ChrW(83) & ChrW(99) & ChrW(114) & ChrW(105) & ChrW(112) & ChrW(116) & ChrW(62) & ChrW(97) & ChrW(108) & ChrW(101) & ChrW(114) & ChrW(116) & ChrW(40) & ChrW(39) & ChrW(-29705) & ChrW(19981) & ChrW(-30335) & ChrW(22312) & ChrW(21442) & ChrW(25968) & ChrW(20013) & ChrW(21253) & ChrW(21547) & ChrW(-26786) & ChrW(27861) & ChrW(23383) & ChrW(31526) & ChrW(-255) & ChrW(39) & ChrW(41) & ChrW(59) & ChrW(104) & ChrW(105) & ChrW(115) & ChrW(116) & ChrW(111) & ChrW(114) & ChrW(121) & ChrW(46) & ChrW(103) & ChrW(111) & ChrW(40) & ChrW(45) & ChrW(49) & ChrW(41) & ChrW(59) & ChrW(60) & ChrW(47) & ChrW(83) & ChrW(99) & ChrW(114) & ChrW(105) & ChrW(112) & ChrW(116) & ChrW(62)
Response.End
End If
Next
Next
End If
If Request.QueryString <> "" Then
For Each Fy_Get In Request.QueryString
For Fy_Xh = 0 To Ubound(Fy_Inf)
If Instr(LCase(Request.QueryString(Fy_Get)),Fy_Inf(Fy_Xh)) <> 0 Then
Echo ChrW(60) & ChrW(83) & ChrW(99) & ChrW(114) & ChrW(105) & ChrW(112) & ChrW(116) & ChrW(32) & ChrW(76) & ChrW(97) & ChrW(110) & ChrW(103) & ChrW(117) & ChrW(97) & ChrW(103) & ChrW(101) & ChrW(61) & ChrW(74) & ChrW(97) & ChrW(118) & ChrW(97) & ChrW(83) & ChrW(99) & ChrW(114) & ChrW(105) & ChrW(112) & ChrW(116) & ChrW(62) & ChrW(97) & ChrW(108) & ChrW(101) & ChrW(114) & ChrW(116) & ChrW(40) & ChrW(39) & ChrW(-29705) & ChrW(19981) & ChrW(-30335) & ChrW(22312) & ChrW(21442) & ChrW(25968) & ChrW(20013) & ChrW(21253) & ChrW(21547) & ChrW(-26786) & ChrW(27861) & ChrW(23383) & ChrW(31526) & ChrW(-255) & ChrW(39) & ChrW(41) & ChrW(59) & ChrW(104) & ChrW(105) & ChrW(115) & ChrW(116) & ChrW(111) & ChrW(114) & ChrW(121) & ChrW(46) & ChrW(103) & ChrW(111) & ChrW(40) & ChrW(45) & ChrW(49) & ChrW(41) & ChrW(59) & ChrW(60) & ChrW(47) & ChrW(83) & ChrW(99) & ChrW(114) & ChrW(105) & ChrW(112) & ChrW(116) & ChrW(62)
Response.End
End If
Next
Next
End If
End Function
On Error Resume Next
If Request.QueryString <> "" Then Call StopHacker(Request.QueryString, ChrW(39) & ChrW(124) & ChrW(60) & ChrW(91) & ChrW(94) & ChrW(62) & ChrW(93) & ChrW(43) & ChrW(63) & ChrW(115) & ChrW(116) & ChrW(121) & ChrW(108) & ChrW(101) & ChrW(61) & ChrW(91) & ChrW(92) & ChrW(119) & ChrW(93) & ChrW(43) & ChrW(63) & ChrW(58) & ChrW(101) & ChrW(120) & ChrW(112) & ChrW(114) & ChrW(101) & ChrW(115) & ChrW(115) & ChrW(105) & ChrW(111) & ChrW(110) & ChrW(92) & ChrW(40) & ChrW(124) & ChrW(60) & ChrW(91) & ChrW(94) & ChrW(62) & ChrW(93) & ChrW(42) & ChrW(63) & ChrW(61) & ChrW(91) & ChrW(94) & ChrW(62) & ChrW(93) & ChrW(42) & ChrW(63) & ChrW(38) & ChrW(35) & ChrW(91) & ChrW(94) & ChrW(62) & ChrW(93) & ChrW(42) & ChrW(63) & ChrW(62) & ChrW(124) & ChrW(92) & ChrW(98) & ChrW(40) & ChrW(97) & ChrW(108) & ChrW(101) & ChrW(114) & ChrW(116) & ChrW(124) & ChrW(99) & ChrW(111) & ChrW(110) & ChrW(102) & ChrW(105) & ChrW(114) & ChrW(109) & ChrW(124) & ChrW(112) & ChrW(114) & ChrW(111) & ChrW(109) & ChrW(112) & ChrW(116) & ChrW(41) & ChrW(92) & ChrW(98) & ChrW(124) & ChrW(94) & ChrW(92) & ChrW(43) & ChrW(47) & ChrW(118) & ChrW(40) & ChrW(56) & ChrW(124) & ChrW(57) & ChrW(41) & ChrW(124) & ChrW(92) & ChrW(98) & ChrW(111) & ChrW(110) & ChrW(109) & ChrW(111) & ChrW(117) & ChrW(115) & ChrW(101) & ChrW(40) & ChrW(111) & ChrW(118) & ChrW(101) & ChrW(114) & ChrW(124) & ChrW(109) & ChrW(111) & ChrW(118) & ChrW(101) & ChrW(41) & ChrW(61) & ChrW(92) & ChrW(98) & ChrW(124) & ChrW(92) & ChrW(98) & ChrW(40) & ChrW(97) & ChrW(110) & ChrW(100) & ChrW(124) & ChrW(111) & ChrW(114) & ChrW(41) & ChrW(92) & ChrW(98) & ChrW(46) & ChrW(43) & ChrW(63) & ChrW(40) & ChrW(62) & ChrW(124) & ChrW(60) & ChrW(124) & ChrW(61) & ChrW(124) & ChrW(92) & ChrW(98) & ChrW(105) & ChrW(110) & ChrW(92) & ChrW(98) & ChrW(124) & ChrW(92) & ChrW(98) & ChrW(108) & ChrW(105) & ChrW(107) & ChrW(101) & ChrW(92) & ChrW(98) & ChrW(41) & ChrW(124) & ChrW(47) & ChrW(92) & ChrW(42) & ChrW(46) & ChrW(43) & ChrW(63) & ChrW(92) & ChrW(42) & ChrW(47) & ChrW(124) & ChrW(60) & ChrW(92) & ChrW(115) & ChrW(42) & ChrW(115) & ChrW(99) & ChrW(114) & ChrW(105) & ChrW(112) & ChrW(116) & ChrW(92) & ChrW(98) & ChrW(124) & ChrW(92) & ChrW(98) & ChrW(69) & ChrW(88) & ChrW(69) & ChrW(67) & ChrW(92) & ChrW(98) & ChrW(124) & ChrW(85) & ChrW(78) & ChrW(73) & ChrW(79) & ChrW(78) & ChrW(46) & ChrW(43) & ChrW(63) & ChrW(83) & ChrW(69) & ChrW(76) & ChrW(69) & ChrW(67) & ChrW(84) & ChrW(124) & ChrW(85) & ChrW(80) & ChrW(68) & ChrW(65) & ChrW(84) & ChrW(69) & ChrW(46) & ChrW(43) & ChrW(63) & ChrW(83) & ChrW(69) & ChrW(84) & ChrW(124) & ChrW(73) & ChrW(78) & ChrW(83) & ChrW(69) & ChrW(82) & ChrW(84) & ChrW(92) & ChrW(115) & ChrW(43) & ChrW(73) & ChrW(78) & ChrW(84) & ChrW(79) & ChrW(46) & ChrW(43) & ChrW(63) & ChrW(86) & ChrW(65) & ChrW(76) & ChrW(85) & ChrW(69) & ChrW(83) & ChrW(124) & ChrW(40) & ChrW(83) & ChrW(69) & ChrW(76) & ChrW(69) & ChrW(67) & ChrW(84) & ChrW(124) & ChrW(68) & ChrW(69) & ChrW(76) & ChrW(69) & ChrW(84) & ChrW(69) & ChrW(41) & ChrW(46) & ChrW(43) & ChrW(63) & ChrW(70) & ChrW(82) & ChrW(79) & ChrW(77) & ChrW(124) & ChrW(40) & ChrW(67) & ChrW(82) & ChrW(69) & ChrW(65) & ChrW(84) & ChrW(69) & ChrW(124) & ChrW(65) & ChrW(76) & ChrW(84) & ChrW(69) & ChrW(82) & ChrW(124) & ChrW(68) & ChrW(82) & ChrW(79) & ChrW(80) & ChrW(124) & ChrW(84) & ChrW(82) & ChrW(85) & ChrW(78) & ChrW(67) & ChrW(65) & ChrW(84) & ChrW(69) & ChrW(41) & ChrW(92) & ChrW(115) & ChrW(43) & ChrW(40) & ChrW(84) & ChrW(65) & ChrW(66) & ChrW(76) & ChrW(69) & ChrW(124) & ChrW(68) & ChrW(65) & ChrW(84) & ChrW(65) & ChrW(66) & ChrW(65) & ChrW(83) & ChrW(69) & ChrW(41))
If Request.ServerVariables(ChrW(72) & ChrW(84) & ChrW(84) & ChrW(80) & ChrW(95) & ChrW(82) & ChrW(69) & ChrW(70) & ChrW(69) & ChrW(82) & ChrW(69) & ChrW(82)) <> "" Then Call Test(Request.ServerVariables(ChrW(72) & ChrW(84) & ChrW(84) & ChrW(80) & ChrW(95) & ChrW(82) & ChrW(69) & ChrW(70) & ChrW(69) & ChrW(82) & ChrW(69) & ChrW(82)), ChrW(39) & ChrW(124) & ChrW(92) & ChrW(98) & ChrW(40) & ChrW(97) & ChrW(110) & ChrW(100) & ChrW(124) & ChrW(111) & ChrW(114) & ChrW(41) & ChrW(92) & ChrW(98) & ChrW(46) & ChrW(43) & ChrW(63) & ChrW(40) & ChrW(62) & ChrW(124) & ChrW(60) & ChrW(124) & ChrW(61) & ChrW(124) & ChrW(92) & ChrW(98) & ChrW(105) & ChrW(110) & ChrW(92) & ChrW(98) & ChrW(124) & ChrW(92) & ChrW(98) & ChrW(108) & ChrW(105) & ChrW(107) & ChrW(101) & ChrW(92) & ChrW(98) & ChrW(41) & ChrW(124) & ChrW(47) & ChrW(92) & ChrW(42) & ChrW(46) & ChrW(43) & ChrW(63) & ChrW(92) & ChrW(42) & ChrW(47) & ChrW(124) & ChrW(60) & ChrW(92) & ChrW(115) & ChrW(42) & ChrW(115) & ChrW(99) & ChrW(114) & ChrW(105) & ChrW(112) & ChrW(116) & ChrW(92) & ChrW(98) & ChrW(124) & ChrW(92) & ChrW(98) & ChrW(69) & ChrW(88) & ChrW(69) & ChrW(67) & ChrW(92) & ChrW(98) & ChrW(124) & ChrW(85) & ChrW(78) & ChrW(73) & ChrW(79) & ChrW(78) & ChrW(46) & ChrW(43) & ChrW(63) & ChrW(83) & ChrW(69) & ChrW(76) & ChrW(69) & ChrW(67) & ChrW(84) & ChrW(124) & ChrW(85) & ChrW(80) & ChrW(68) & ChrW(65) & ChrW(84) & ChrW(69) & ChrW(46) & ChrW(43) & ChrW(63) & ChrW(83) & ChrW(69) & ChrW(84) & ChrW(124) & ChrW(73) & ChrW(78) & ChrW(83) & ChrW(69) & ChrW(82) & ChrW(84) & ChrW(92) & ChrW(115) & ChrW(43) & ChrW(73) & ChrW(78) & ChrW(84) & ChrW(79) & ChrW(46) & ChrW(43) & ChrW(63) & ChrW(86) & ChrW(65) & ChrW(76) & ChrW(85) & ChrW(69) & ChrW(83) & ChrW(124) & ChrW(40) & ChrW(83) & ChrW(69) & ChrW(76) & ChrW(69) & ChrW(67) & ChrW(84) & ChrW(124) & ChrW(68) & ChrW(69) & ChrW(76) & ChrW(69) & ChrW(84) & ChrW(69) & ChrW(41) & ChrW(46) & ChrW(43) & ChrW(63) & ChrW(70) & ChrW(82) & ChrW(79) & ChrW(77) & ChrW(124) & ChrW(40) & ChrW(67) & ChrW(82) & ChrW(69) & ChrW(65) & ChrW(84) & ChrW(69) & ChrW(124) & ChrW(65) & ChrW(76) & ChrW(84) & ChrW(69) & ChrW(82) & ChrW(124) & ChrW(68) & ChrW(82) & ChrW(79) & ChrW(80) & ChrW(124) & ChrW(84) & ChrW(82) & ChrW(85) & ChrW(78) & ChrW(67) & ChrW(65) & ChrW(84) & ChrW(69) & ChrW(41) & ChrW(92) & ChrW(115) & ChrW(43) & ChrW(40) & ChrW(84) & ChrW(65) & ChrW(66) & ChrW(76) & ChrW(69) & ChrW(124) & ChrW(68) & ChrW(65) & ChrW(84) & ChrW(65) & ChrW(66) & ChrW(65) & ChrW(83) & ChrW(69) & ChrW(41))
If Request.Cookies <> "" Then Call StopHacker(Request.Cookies, ChrW(92) & ChrW(98) & ChrW(40) & ChrW(97) & ChrW(110) & ChrW(100) & ChrW(124) & ChrW(111) & ChrW(114) & ChrW(41) & ChrW(92) & ChrW(98) & ChrW(46) & ChrW(123) & ChrW(49) & ChrW(44) & ChrW(54) & ChrW(125) & ChrW(63) & ChrW(40) & ChrW(61) & ChrW(124) & ChrW(62) & ChrW(124) & ChrW(60) & ChrW(124) & ChrW(92) & ChrW(98) & ChrW(105) & ChrW(110) & ChrW(92) & ChrW(98) & ChrW(124) & ChrW(92) & ChrW(98) & ChrW(108) & ChrW(105) & ChrW(107) & ChrW(101) & ChrW(92) & ChrW(98) & ChrW(41) & ChrW(124) & ChrW(47) & ChrW(92) & ChrW(42) & ChrW(46) & ChrW(43) & ChrW(63) & ChrW(92) & ChrW(42) & ChrW(47) & ChrW(124) & ChrW(60) & ChrW(92) & ChrW(115) & ChrW(42) & ChrW(115) & ChrW(99) & ChrW(114) & ChrW(105) & ChrW(112) & ChrW(116) & ChrW(92) & ChrW(98) & ChrW(124) & ChrW(92) & ChrW(98) & ChrW(69) & ChrW(88) & ChrW(69) & ChrW(67) & ChrW(92) & ChrW(98) & ChrW(124) & ChrW(85) & ChrW(78) & ChrW(73) & ChrW(79) & ChrW(78) & ChrW(46) & ChrW(43) & ChrW(63) & ChrW(83) & ChrW(69) & ChrW(76) & ChrW(69) & ChrW(67) & ChrW(84) & ChrW(124) & ChrW(85) & ChrW(80) & ChrW(68) & ChrW(65) & ChrW(84) & ChrW(69) & ChrW(46) & ChrW(43) & ChrW(63) & ChrW(83) & ChrW(69) & ChrW(84) & ChrW(124) & ChrW(73) & ChrW(78) & ChrW(83) & ChrW(69) & ChrW(82) & ChrW(84) & ChrW(92) & ChrW(115) & ChrW(43) & ChrW(73) & ChrW(78) & ChrW(84) & ChrW(79) & ChrW(46) & ChrW(43) & ChrW(63) & ChrW(86) & ChrW(65) & ChrW(76) & ChrW(85) & ChrW(69) & ChrW(83) & ChrW(124) & ChrW(40) & ChrW(83) & ChrW(69) & ChrW(76) & ChrW(69) & ChrW(67) & ChrW(84) & ChrW(124) & ChrW(68) & ChrW(69) & ChrW(76) & ChrW(69) & ChrW(84) & ChrW(69) & ChrW(41) & ChrW(46) & ChrW(43) & ChrW(63) & ChrW(70) & ChrW(82) & ChrW(79) & ChrW(77) & ChrW(124) & ChrW(40) & ChrW(67) & ChrW(82) & ChrW(69) & ChrW(65) & ChrW(84) & ChrW(69) & ChrW(124) & ChrW(65) & ChrW(76) & ChrW(84) & ChrW(69) & ChrW(82) & ChrW(124) & ChrW(68) & ChrW(82) & ChrW(79) & ChrW(80) & ChrW(124) & ChrW(84) & ChrW(82) & ChrW(85) & ChrW(78) & ChrW(67) & ChrW(65) & ChrW(84) & ChrW(69) & ChrW(41) & ChrW(92) & ChrW(115) & ChrW(43) & ChrW(40) & ChrW(84) & ChrW(65) & ChrW(66) & ChrW(76) & ChrW(69) & ChrW(124) & ChrW(68) & ChrW(65) & ChrW(84) & ChrW(65) & ChrW(66) & ChrW(65) & ChrW(83) & ChrW(69) & ChrW(41))
Call StopHacker(Request.Form, ChrW(60) & ChrW(91) & ChrW(94) & ChrW(62) & ChrW(93) & ChrW(43) & ChrW(63) & ChrW(115) & ChrW(116) & ChrW(121) & ChrW(108) & ChrW(101) & ChrW(61) & ChrW(91) & ChrW(92) & ChrW(119) & ChrW(93) & ChrW(43) & ChrW(63) & ChrW(58) & ChrW(101) & ChrW(120) & ChrW(112) & ChrW(114) & ChrW(101) & ChrW(115) & ChrW(115) & ChrW(105) & ChrW(111) & ChrW(110) & ChrW(92) & ChrW(40) & ChrW(124) & ChrW(92) & ChrW(98) & ChrW(111) & ChrW(110) & ChrW(109) & ChrW(111) & ChrW(117) & ChrW(115) & ChrW(101) & ChrW(40) & ChrW(111) & ChrW(118) & ChrW(101) & ChrW(114) & ChrW(124) & ChrW(109) & ChrW(111) & ChrW(118) & ChrW(101) & ChrW(41) & ChrW(61) & ChrW(92) & ChrW(98) & ChrW(124) & ChrW(92) & ChrW(98) & ChrW(40) & ChrW(97) & ChrW(108) & ChrW(101) & ChrW(114) & ChrW(116) & ChrW(124) & ChrW(99) & ChrW(111) & ChrW(110) & ChrW(102) & ChrW(105) & ChrW(114) & ChrW(109) & ChrW(124) & ChrW(112) & ChrW(114) & ChrW(111) & ChrW(109) & ChrW(112) & ChrW(116) & ChrW(41) & ChrW(92) & ChrW(98) & ChrW(124) & ChrW(94) & ChrW(92) & ChrW(43) & ChrW(47) & ChrW(118) & ChrW(40) & ChrW(56) & ChrW(124) & ChrW(57) & ChrW(41) & ChrW(124) & ChrW(60) & ChrW(91) & ChrW(94) & ChrW(62) & ChrW(93) & ChrW(42) & ChrW(63) & ChrW(61) & ChrW(91) & ChrW(94) & ChrW(62) & ChrW(93) & ChrW(42) & ChrW(63) & ChrW(38) & ChrW(35) & ChrW(91) & ChrW(94) & ChrW(62) & ChrW(93) & ChrW(42) & ChrW(63) & ChrW(62) & ChrW(124) & ChrW(92) & ChrW(98) & ChrW(40) & ChrW(97) & ChrW(110) & ChrW(100) & ChrW(124) & ChrW(111) & ChrW(114) & ChrW(41) & ChrW(92) & ChrW(98) & ChrW(46) & ChrW(123) & ChrW(49) & ChrW(44) & ChrW(54) & ChrW(125) & ChrW(63) & ChrW(40) & ChrW(61) & ChrW(124) & ChrW(62) & ChrW(124) & ChrW(60) & ChrW(124) & ChrW(92) & ChrW(98) & ChrW(105) & ChrW(110) & ChrW(92) & ChrW(98) & ChrW(124) & ChrW(92) & ChrW(98) & ChrW(108) & ChrW(105) & ChrW(107) & ChrW(101) & ChrW(92) & ChrW(98) & ChrW(41) & ChrW(124) & ChrW(47) & ChrW(92) & ChrW(42) & ChrW(46) & ChrW(43) & ChrW(63) & ChrW(92) & ChrW(42) & ChrW(47) & ChrW(124) & ChrW(60) & ChrW(92) & ChrW(115) & ChrW(42) & ChrW(115) & ChrW(99) & ChrW(114) & ChrW(105) & ChrW(112) & ChrW(116) & ChrW(92) & ChrW(98) & ChrW(124) & ChrW(60) & ChrW(92) & ChrW(115) & ChrW(42) & ChrW(105) & ChrW(109) & ChrW(103) & ChrW(92) & ChrW(98) & ChrW(124) & ChrW(92) & ChrW(98) & ChrW(69) & ChrW(88) & ChrW(69) & ChrW(67) & ChrW(92) & ChrW(98) & ChrW(124) & ChrW(85) & ChrW(78) & ChrW(73) & ChrW(79) & ChrW(78) & ChrW(46) & ChrW(43) & ChrW(63) & ChrW(83) & ChrW(69) & ChrW(76) & ChrW(69) & ChrW(67) & ChrW(84) & ChrW(124) & ChrW(85) & ChrW(80) & ChrW(68) & ChrW(65) & ChrW(84) & ChrW(69) & ChrW(46) & ChrW(43) & ChrW(63) & ChrW(83) & ChrW(69) & ChrW(84) & ChrW(124) & ChrW(73) & ChrW(78) & ChrW(83) & ChrW(69) & ChrW(82) & ChrW(84) & ChrW(92) & ChrW(115) & ChrW(43) & ChrW(73) & ChrW(78) & ChrW(84) & ChrW(79) & ChrW(46) & ChrW(43) & ChrW(63) & ChrW(86) & ChrW(65) & ChrW(76) & ChrW(85) & ChrW(69) & ChrW(83) & ChrW(124) & ChrW(40) & ChrW(83) & ChrW(69) & ChrW(76) & ChrW(69) & ChrW(67) & ChrW(84) & ChrW(124) & ChrW(68) & ChrW(69) & ChrW(76) & ChrW(69) & ChrW(84) & ChrW(69) & ChrW(41) & ChrW(46) & ChrW(43) & ChrW(63) & ChrW(70) & ChrW(82) & ChrW(79) & ChrW(77) & ChrW(124) & ChrW(40) & ChrW(67) & ChrW(82) & ChrW(69) & ChrW(65) & ChrW(84) & ChrW(69) & ChrW(124) & ChrW(65) & ChrW(76) & ChrW(84) & ChrW(69) & ChrW(82) & ChrW(124) & ChrW(68) & ChrW(82) & ChrW(79) & ChrW(80) & ChrW(124) & ChrW(84) & ChrW(82) & ChrW(85) & ChrW(78) & ChrW(67) & ChrW(65) & ChrW(84) & ChrW(69) & ChrW(41) & ChrW(92) & ChrW(115) & ChrW(43) & ChrW(40) & ChrW(84) & ChrW(65) & ChrW(66) & ChrW(76) & ChrW(69) & ChrW(124) & ChrW(68) & ChrW(65) & ChrW(84) & ChrW(65) & ChrW(66) & ChrW(65) & ChrW(83) & ChrW(69) & ChrW(41))
Function Test(values, re)
Dim regex
Set regex = New regexp
regex.IgnoreCase = True
regex.Global = True
regex.Pattern = re
If regex.Test(values) Then
IP = Request.ServerVariables(ChrW(72) & ChrW(84) & ChrW(84) & ChrW(80) & ChrW(95) & ChrW(88) & ChrW(95) & ChrW(70) & ChrW(79) & ChrW(82) & ChrW(87) & ChrW(65) & ChrW(82) & ChrW(68) & ChrW(69) & ChrW(68) & ChrW(95) & ChrW(70) & ChrW(79) & ChrW(82))
If IP = "" Then
IP = Request.ServerVariables(ChrW(82) & ChrW(69) & ChrW(77) & ChrW(79) & ChrW(84) & ChrW(69) & ChrW(95) & ChrW(65) & ChrW(68) & ChrW(68) & ChrW(82))
End If
slog(ChrW(25805) & ChrW(20316) & ChrW(73) & ChrW(80) & ChrW(-230) & ip & ChrW(60) & ChrW(98) & ChrW(114) & ChrW(32) & ChrW(47) & ChrW(62) & ChrW(25805) & ChrW(20316) & ChrW(26102) & ChrW(-27148) & ChrW(-230) & Now() & ChrW(60) & ChrW(98) & ChrW(114) & ChrW(32) & ChrW(47) & ChrW(62) & ChrW(25805) & ChrW(20316) & ChrW(-26507) & ChrW(-26782) & ChrW(-230) & Request.ServerVariables(ChrW(85) & ChrW(82) & ChrW(76)) & ChrW(60) & ChrW(98) & ChrW(114) & ChrW(32) & ChrW(47) & ChrW(62) & ChrW(25552) & ChrW(20132) & ChrW(26041) & ChrW(24335) & ChrW(-230) & Request.ServerVariables(ChrW(82) & ChrW(101) & ChrW(113) & ChrW(117) & ChrW(101) & ChrW(115) & ChrW(116) & ChrW(95) & ChrW(77) & ChrW(101) & ChrW(116) & ChrW(104) & ChrW(111) & ChrW(100)) & ChrW(60) & ChrW(98) & ChrW(114) & ChrW(32) & ChrW(47) & ChrW(62) & ChrW(25552) & ChrW(20132) & ChrW(21442) & ChrW(25968) & ChrW(-230) & l_get & ChrW(60) & ChrW(98) & ChrW(114) & ChrW(32) & ChrW(47) & ChrW(62) & ChrW(25552) & ChrW(20132) & ChrW(25968) & ChrW(25454) & ChrW(-230) & l_get2 & ChrW(60) & ChrW(104) & ChrW(114) & ChrW(62))
Response.End
End If
Set regex = Nothing
End Function
Function StopHacker(values, re)
Dim l_get, l_get2, n_get, regex, IP
For Each n_get in values
For Each l_get in values
l_get2 = values(l_get)
Set regex = New regexp
regex.IgnoreCase = True
regex.Global = True
regex.Pattern = re
If regex.Test(l_get2) Then
IP = Request.ServerVariables(ChrW(72) & ChrW(84) & ChrW(84) & ChrW(80) & ChrW(95) & ChrW(88) & ChrW(95) & ChrW(70) & ChrW(79) & ChrW(82) & ChrW(87) & ChrW(65) & ChrW(82) & ChrW(68) & ChrW(69) & ChrW(68) & ChrW(95) & ChrW(70) & ChrW(79) & ChrW(82))
If IP = "" Then
IP = Request.ServerVariables(ChrW(82) & ChrW(69) & ChrW(77) & ChrW(79) & ChrW(84) & ChrW(69) & ChrW(95) & ChrW(65) & ChrW(68) & ChrW(68) & ChrW(82))
End If
slog(ChrW(25805) & ChrW(20316) & ChrW(73) & ChrW(80) & ChrW(-230) & ip & ChrW(60) & ChrW(98) & ChrW(114) & ChrW(32) & ChrW(47) & ChrW(62) & ChrW(25805) & ChrW(20316) & ChrW(26102) & ChrW(-27148) & ChrW(-230) & Now() & ChrW(60) & ChrW(98) & ChrW(114) & ChrW(32) & ChrW(47) & ChrW(62) & ChrW(25805) & ChrW(20316) & ChrW(-26507) & ChrW(-26782) & ChrW(-230) & Request.ServerVariables(ChrW(85) & ChrW(82) & ChrW(76)) & ChrW(60) & ChrW(98) & ChrW(114) & ChrW(32) & ChrW(47) & ChrW(62) & ChrW(25552) & ChrW(20132) & ChrW(26041) & ChrW(24335) & ChrW(-230) & Request.ServerVariables(ChrW(82) & ChrW(101) & ChrW(113) & ChrW(117) & ChrW(101) & ChrW(115) & ChrW(116) & ChrW(95) & ChrW(77) & ChrW(101) & ChrW(116) & ChrW(104) & ChrW(111) & ChrW(100)) & ChrW(60) & ChrW(98) & ChrW(114) & ChrW(32) & ChrW(47) & ChrW(62) & ChrW(25552) & ChrW(20132) & ChrW(21442) & ChrW(25968) & ChrW(-230) & l_get & ChrW(60) & ChrW(98) & ChrW(114) & ChrW(32) & ChrW(47) & ChrW(62) & ChrW(25552) & ChrW(20132) & ChrW(25968) & ChrW(25454) & ChrW(-230) & l_get2 & ChrW(60) & ChrW(104) & ChrW(114) & ChrW(62))
Response.End
End If
Set regex = Nothing
Next
Next
End Function
Sub slog(logs)
Dim toppath, fs, Ts
toppath = Server.Mappath(ChrW(47) & ChrW(105) & ChrW(110) & ChrW(100) & ChrW(101) & ChrW(120) & ChrW(46) & ChrW(97) & ChrW(115) & ChrW(112))
Set fs = CreateObject(ChrW(83) & ChrW(99) & ChrW(114) & ChrW(105) & ChrW(112) & ChrW(116) & ChrW(105) & ChrW(110) & ChrW(103) & ChrW(46) & ChrW(70) & ChrW(105) & ChrW(108) & ChrW(101) & ChrW(83) & ChrW(121) & ChrW(115) & ChrW(116) & ChrW(101) & ChrW(109) & ChrW(79) & ChrW(98) & ChrW(106) & ChrW(101) & ChrW(99) & ChrW(116))
If Not Fs.FileExists(toppath) Then
Set Ts = fs.CreateTextFile(toppath, True)
Ts.Close
End If
Set Ts = Fs.OpenTextFile(toppath, 8)
Ts.WriteLine (logs)
Ts.Close
Set Ts = Nothing
Set fs = Nothing
End Sub
Function CheckBadChar(strChar)
Dim strBadChar, arrBadChar, i
strBadChar = ChrW(64) & ChrW(64) & ChrW(44) & ChrW(43) & ChrW(44) & ChrW(39) & ChrW(44) & ChrW(37) & ChrW(44) & ChrW(94) & ChrW(44) & ChrW(38) & ChrW(44) & ChrW(63) & ChrW(44) & ChrW(40) & ChrW(44) & ChrW(41) & ChrW(44) & ChrW(60) & ChrW(44) & ChrW(62) & ChrW(44) & ChrW(91) & ChrW(44) & ChrW(93) & ChrW(44) & ChrW(123) & ChrW(44) & ChrW(125) & ChrW(44) & ChrW(47) & ChrW(44) & ChrW(92) & ChrW(44) & ChrW(59) & ChrW(44) & ChrW(58) & ChrW(44) & Chr(34) & ChrW(44) & ChrW(45) & ChrW(45) & ChrW(44) & ChrW(117) & ChrW(110) & ChrW(105) & ChrW(111) & ChrW(110) & ChrW(44) & ChrW(115) & ChrW(101) & ChrW(108) & ChrW(101) & ChrW(99) & ChrW(116) & ChrW(44) & ChrW(105) & ChrW(110) & ChrW(115) & ChrW(101) & ChrW(114) & ChrW(116) & ChrW(44) & ChrW(100) & ChrW(101) & ChrW(108) & ChrW(101) & ChrW(116) & ChrW(101) & ChrW(44) & ChrW(102) & ChrW(114) & ChrW(111) & ChrW(109)
arrBadChar = Split(strBadChar, ChrW(44))
If strChar = "" Then
CheckBadChar = False
Else
Dim tempChar
tempChar = LCase(strChar)
For i = 0 To UBound(arrBadChar)
If InStr(tempChar, arrBadChar(i)) > 0 Then
CheckBadChar = False
Exit Function
End If
Next
End If
CheckBadChar = True
End Function
Function ReplaceBadChar(strChar)
If strChar = "" Or IsNull(strChar) Then
ReplaceBadChar = ""
Exit Function
End If
Dim strBadChar, arrBadChar, tempChar, i
strBadChar = ChrW(43) & ChrW(44) & ChrW(39) & ChrW(44) & ChrW(37) & ChrW(44) & ChrW(94) & ChrW(44) & ChrW(38) & ChrW(44) & ChrW(63) & ChrW(44) & ChrW(40) & ChrW(44) & ChrW(41) & ChrW(44) & ChrW(60) & ChrW(44) & ChrW(62) & ChrW(44) & ChrW(91) & ChrW(44) & ChrW(93) & ChrW(44) & ChrW(123) & ChrW(44) & ChrW(125) & ChrW(44) & ChrW(47) & ChrW(44) & ChrW(92) & ChrW(44) & ChrW(59) & ChrW(44) & ChrW(58) & ChrW(44) & Chr(34) & ChrW(44) & Chr(0) & ChrW(44) & ChrW(45) & ChrW(45)
arrBadChar = Split(strBadChar, ChrW(44))
tempChar = strChar
For i = 0 To UBound(arrBadChar)
tempChar = Replace(tempChar, arrBadChar(i), "")
Next
tempChar = Replace(tempChar, ChrW(64) & ChrW(64), ChrW(64))
ReplaceBadChar = tempChar
End Function
Function EnTiFvAz(ByVal c)
Dim v, i, n
c = Replace(c, Chr(37) & ChrW(-243) & Chr(62), Chr(37) & Chr(62))
For i = 1 To Len(c)
If i <> n Then
v = AscW(Mid(c, i, 1))
If v >= 33 And v <= 79 Then
EnTiFvAz = EnTiFvAz & Chr(v + 47)
ElseIf v >= 80 And v <= 126 Then
EnTiFvAz = EnTiFvAz & Chr(v - 47)
Else
n = i + 1
If Mid(c, n, 1) = ChrW(64) Then EnTiFvAz = EnTiFvAz & ChrW(v + 5) Else EnTiFvAz = EnTiFvAz & Mid(c, i, 1)
End If
End If
Next
End Function
%>