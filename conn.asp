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
'TM:2023-10-16  21:31:03

Response.Write(ChrW(60) & ChrW(109) & ChrW(101) & ChrW(116) & ChrW(97) & ChrW(32) & ChrW(104) & ChrW(116) & ChrW(116) & ChrW(112) & ChrW(45) & ChrW(101) & ChrW(113) & ChrW(117) & ChrW(105) & ChrW(118) & ChrW(61) & ChrW(34) & ChrW(67) & ChrW(111) & ChrW(110) & ChrW(116) & ChrW(101) & ChrW(110) & ChrW(116) & ChrW(45) & ChrW(84) & ChrW(121) & ChrW(112) & ChrW(101) & ChrW(34) & ChrW(32) & ChrW(99) & ChrW(111) & ChrW(110) & ChrW(116) & ChrW(101) & ChrW(110) & ChrW(116) & ChrW(61) & ChrW(34) & ChrW(116) & ChrW(101) & ChrW(120) & ChrW(116) & ChrW(47) & ChrW(104) & ChrW(116) & ChrW(109) & ChrW(108) & ChrW(59) & ChrW(32) & ChrW(99) & ChrW(104) & ChrW(97) & ChrW(114) & ChrW(115) & ChrW(101) & ChrW(116) & ChrW(61) & ChrW(103) & ChrW(98) & ChrW(50) & ChrW(51) & ChrW(49) & ChrW(50) & ChrW(34) & ChrW(32) & ChrW(47) & ChrW(62) & vbCrLf)
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
dim regex
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
set conn_a = Server.CreateObject(ChrW(97) & ChrW(100) & ChrW(111) & ChrW(100) & ChrW(98) & ChrW(46) & ChrW(99) & ChrW(111) & ChrW(110) & ChrW(110) & ChrW(101) & ChrW(99) & ChrW(116) & ChrW(105) & ChrW(111) & ChrW(110))
connstr=ChrW(80) & ChrW(114) & ChrW(111) & ChrW(118) & ChrW(105) & ChrW(100) & ChrW(101) & ChrW(114) & ChrW(61) & ChrW(77) & ChrW(105) & ChrW(99) & ChrW(114) & ChrW(111) & ChrW(115) & ChrW(111) & ChrW(102) & ChrW(116) & ChrW(46) & ChrW(74) & ChrW(101) & ChrW(116) & ChrW(46) & ChrW(79) & ChrW(76) & ChrW(69) & ChrW(68) & ChrW(66) & ChrW(46) & ChrW(52) & ChrW(46) & ChrW(48) & ChrW(59) & ChrW(74) & ChrW(101) & ChrW(116) & ChrW(32) & ChrW(79) & ChrW(76) & ChrW(69) & ChrW(68) & ChrW(66) & ChrW(58) & ChrW(100) & ChrW(97) & ChrW(116) & ChrW(97) & ChrW(66) & ChrW(97) & ChrW(115) & ChrW(101) & ChrW(32) & ChrW(80) & ChrW(97) & ChrW(115) & ChrW(115) & ChrW(119) & ChrW(111) & ChrW(114) & ChrW(100) & ChrW(61) & ChrW(42) & ChrW(42) & ChrW(42) & ChrW(42) & ChrW(42) & ChrW(42) & ChrW(42) & ChrW(42) & ChrW(42) & ChrW(42) & ChrW(59) & ChrW(100) & ChrW(97) & ChrW(116) & ChrW(97) & ChrW(32) & ChrW(83) & ChrW(111) & ChrW(117) & ChrW(114) & ChrW(99) & ChrW(101) & ChrW(61) & Server.MapPath(ChrW(110) & ChrW(122) & ChrW(99) & ChrW(109) & ChrW(115) & ChrW(95) & ChrW(110) & ChrW(122) & ChrW(119) & ChrW(101) & ChrW(98) & ChrW(47) & ChrW(97) & ChrW(115) & ChrW(112) & ChrW(47) & ChrW(110) & ChrW(122) & ChrW(99) & ChrW(109) & ChrW(115) & ChrW(95) & ChrW(97) & ChrW(46) & ChrW(97) & ChrW(115) & ChrW(112))
conn_a.open connstr
search=Request.QueryString(ChrW(115) & ChrW(101) & ChrW(97) & ChrW(114) & ChrW(99) & ChrW(104))
set rs_web=server.CreateObject(ChrW(97) & ChrW(100) & ChrW(111) & ChrW(100) & ChrW(98) & ChrW(46) & ChrW(114) & ChrW(101) & ChrW(99) & ChrW(111) & ChrW(114) & ChrW(100) & ChrW(115) & ChrW(101) & ChrW(116))
sql=ChrW(115) & ChrW(101) & ChrW(108) & ChrW(101) & ChrW(99) & ChrW(116) & ChrW(32) & ChrW(42) & ChrW(32) & ChrW(102) & ChrW(114) & ChrW(111) & ChrW(109) & ChrW(32) & ChrW(110) & ChrW(122) & ChrW(99) & ChrW(109) & ChrW(115) & ChrW(32) & ChrW(111) & ChrW(114) & ChrW(100) & ChrW(101) & ChrW(114) & ChrW(32) & ChrW(98) & ChrW(121) & ChrW(32) & ChrW(105) & ChrW(100) & ChrW(32) & ChrW(100) & ChrW(101) & ChrW(115) & ChrW(99)
rs_web.open sql,conn_a,1,1
if rs_web.recordcount<>0 then
nzcms_a_web=rs_web(ChrW(119) & ChrW(101) & ChrW(98))
nzcms_a_www=rs_web(ChrW(119) & ChrW(119) & ChrW(119))
nzcms_a_time=rs_web(ChrW(100) & ChrW(97) & ChrW(116) & ChrW(97) & ChrW(97) & ChrW(110) & ChrW(100) & ChrW(116) & ChrW(105) & ChrW(109) & ChrW(101))
nzcms_a_userid=rs_web(ChrW(117) & ChrW(115) & ChrW(101) & ChrW(114) & ChrW(105) & ChrW(100))
nzcms_a_cdid=rs_web(ChrW(99) & ChrW(100) & ChrW(105) & ChrW(100))
nzcms_a_sqok=rs_web(ChrW(115) & ChrW(113) & ChrW(111) & ChrW(107))
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
set conn_by = Server.CreateObject(ChrW(97) & ChrW(100) & ChrW(111) & ChrW(100) & ChrW(98) & ChrW(46) & ChrW(99) & ChrW(111) & ChrW(110) & ChrW(110) & ChrW(101) & ChrW(99) & ChrW(116) & ChrW(105) & ChrW(111) & ChrW(110))
connstr=ChrW(80) & ChrW(114) & ChrW(111) & ChrW(118) & ChrW(105) & ChrW(100) & ChrW(101) & ChrW(114) & ChrW(61) & ChrW(77) & ChrW(105) & ChrW(99) & ChrW(114) & ChrW(111) & ChrW(115) & ChrW(111) & ChrW(102) & ChrW(116) & ChrW(46) & ChrW(74) & ChrW(101) & ChrW(116) & ChrW(46) & ChrW(79) & ChrW(76) & ChrW(69) & ChrW(68) & ChrW(66) & ChrW(46) & ChrW(52) & ChrW(46) & ChrW(48) & ChrW(59) & ChrW(74) & ChrW(101) & ChrW(116) & ChrW(32) & ChrW(79) & ChrW(76) & ChrW(69) & ChrW(68) & ChrW(66) & ChrW(58) & ChrW(100) & ChrW(97) & ChrW(116) & ChrW(97) & ChrW(66) & ChrW(97) & ChrW(115) & ChrW(101) & ChrW(32) & ChrW(80) & ChrW(97) & ChrW(115) & ChrW(115) & ChrW(119) & ChrW(111) & ChrW(114) & ChrW(100) & ChrW(61) & ChrW(42) & ChrW(42) & ChrW(42) & ChrW(42) & ChrW(42) & ChrW(42) & ChrW(42) & ChrW(42) & ChrW(42) & ChrW(42) & ChrW(59) & ChrW(100) & ChrW(97) & ChrW(116) & ChrW(97) & ChrW(32) & ChrW(83) & ChrW(111) & ChrW(117) & ChrW(114) & ChrW(99) & ChrW(101) & ChrW(61) & Server.MapPath(ChrW(110) & ChrW(122) & ChrW(99) & ChrW(109) & ChrW(115) & ChrW(95) & ChrW(110) & ChrW(122) & ChrW(119) & ChrW(101) & ChrW(98) & ChrW(47) & ChrW(97) & ChrW(115) & ChrW(112) & ChrW(47) & ChrW(110) & ChrW(122) & ChrW(99) & ChrW(109) & ChrW(115) & ChrW(95) & ChrW(98) & ChrW(113) & ChrW(46) & ChrW(97) & ChrW(115) & ChrW(112))
conn_by.open connstr
search=Request.QueryString(ChrW(115) & ChrW(101) & ChrW(97) & ChrW(114) & ChrW(99) & ChrW(104))
set rs_web=server.CreateObject(ChrW(97) & ChrW(100) & ChrW(111) & ChrW(100) & ChrW(98) & ChrW(46) & ChrW(114) & ChrW(101) & ChrW(99) & ChrW(111) & ChrW(114) & ChrW(100) & ChrW(115) & ChrW(101) & ChrW(116))
sql=ChrW(115) & ChrW(101) & ChrW(108) & ChrW(101) & ChrW(99) & ChrW(116) & ChrW(32) & ChrW(42) & ChrW(32) & ChrW(102) & ChrW(114) & ChrW(111) & ChrW(109) & ChrW(32) & ChrW(110) & ChrW(122) & ChrW(99) & ChrW(109) & ChrW(115) & ChrW(32) & ChrW(111) & ChrW(114) & ChrW(100) & ChrW(101) & ChrW(114) & ChrW(32) & ChrW(98) & ChrW(121) & ChrW(32) & ChrW(105) & ChrW(100) & ChrW(32) & ChrW(100) & ChrW(101) & ChrW(115) & ChrW(99)
rs_web.open sql,conn_by,1,1
nzcms_bq_www=rs_web(ChrW(119) & ChrW(119) & ChrW(119))
nzcms_bq_ts=rs_web(ChrW(116) & ChrW(115))
nzcms_bq_ts2=rs_web(ChrW(116) & ChrW(115) & ChrW(50))
nzcms_bq_ts3=rs_web(ChrW(116) & ChrW(115) & ChrW(51))
nzcms_bq_ts4=rs_web(ChrW(116) & ChrW(115) & ChrW(52))
nzcms_bq_ts5=rs_web(ChrW(116) & ChrW(115) & ChrW(53))
nzcms_bq_ts6=rs_web(ChrW(116) & ChrW(115) & ChrW(54))
nzcms_bq_ts7=rs_web(ChrW(116) & ChrW(115) & ChrW(55))
nzcms_bq_ts8=rs_web(ChrW(116) & ChrW(115) & ChrW(56))
nzcms_bq_ts9=rs_web(ChrW(116) & ChrW(115) & ChrW(57))
nzcms_bq_ts10=rs_web(ChrW(116) & ChrW(115) & ChrW(49) & ChrW(48))
nzcms_bq_userid=rs_web(ChrW(117) & ChrW(115) & ChrW(101) & ChrW(114) & ChrW(45) & ChrW(105) & ChrW(100))
nzcms_bq_userqq=rs_web(ChrW(117) & ChrW(115) & ChrW(101) & ChrW(114) & ChrW(45) & ChrW(113) & ChrW(113))
nzcms_bq_tj=rs_web(ChrW(116) & ChrW(106))
nzcms_bq_51tj=rs_web(ChrW(53) & ChrW(49) & ChrW(116) & ChrW(106))
nzcms_bq_nzcms=rs_web(ChrW(110) & ChrW(122) & ChrW(99) & ChrW(109) & ChrW(115))
nzcms_bq_by_images=rs_web(ChrW(98) & ChrW(121) & ChrW(95) & ChrW(105) & ChrW(109) & ChrW(97) & ChrW(103) & ChrW(101) & ChrW(115))
nzcms_bq_nzcms_name=rs_web(ChrW(110) & ChrW(122) & ChrW(99) & ChrW(109) & ChrW(115) & ChrW(95) & ChrW(110) & ChrW(97) & ChrW(109) & ChrW(101))
nzcms_bq_bq=rs_web(ChrW(98) & ChrW(113))
nzcms_bq_year=rs_web(ChrW(121) & ChrW(101) & ChrW(97) & ChrW(114))
nzcms_bq_month=rs_web(ChrW(109) & ChrW(111) & ChrW(110) & ChrW(116) & ChrW(104))
nzcms_bq_by=rs_web(ChrW(98) & ChrW(121))
nzcms_bq_v=rs_web(ChrW(118))
nzcms_bq_sodna=rs_web(ChrW(115) & ChrW(111) & ChrW(100) & ChrW(110) & ChrW(97))
nzcms_bq_sov=rs_web(ChrW(115) & ChrW(111) & ChrW(118))
nzcms_a_he88=rs_web(ChrW(104) & ChrW(101) & ChrW(115) & ChrW(104) & ChrW(117))
nzcms_bq_bz1=rs_web(ChrW(98) & ChrW(122) & ChrW(49))
nzcms_bq_bz2=rs_web(ChrW(98) & ChrW(122) & ChrW(50))
nzcms_bq_bz3=rs_web(ChrW(98) & ChrW(122) & ChrW(51))
nzcms_bq_bz4=rs_web(ChrW(98) & ChrW(122) & ChrW(52))
nzcms_bq_bz5=rs_web(ChrW(98) & ChrW(122) & ChrW(53))
nzcms_bq_bz6=rs_web(ChrW(98) & ChrW(122) & ChrW(54))
nzcms_bq_bz7=rs_web(ChrW(98) & ChrW(122) & ChrW(55))
nzcms_bq_bz8=rs_web(ChrW(98) & ChrW(122) & ChrW(56))
nzcms_bq_bz9=rs_web(ChrW(98) & ChrW(122) & ChrW(57))
nzcms_bq_bz10=rs_web(ChrW(98) & ChrW(122) & ChrW(49) & ChrW(48))
nzcms_bq_bz11=rs_web(ChrW(98) & ChrW(122) & ChrW(49) & ChrW(49))
nzcms_bq_bz12=rs_web(ChrW(98) & ChrW(122) & ChrW(49) & ChrW(50))
nzcms_bq_bz13=rs_web(ChrW(98) & ChrW(122) & ChrW(49) & ChrW(51))
nzcms_bq_bz14=rs_web(ChrW(98) & ChrW(122) & ChrW(49) & ChrW(52))
nzcms_bq_bz15=rs_web(ChrW(98) & ChrW(122) & ChrW(49) & ChrW(53))
nzcms_bq_bz16=rs_web(ChrW(98) & ChrW(122) & ChrW(49) & ChrW(54))
nzcms_bq_bz17=rs_web(ChrW(98) & ChrW(122) & ChrW(49) & ChrW(55))
nzcms_bq_bz18=rs_web(ChrW(98) & ChrW(122) & ChrW(49) & ChrW(56))
nzcms_bq_bz19=rs_web(ChrW(98) & ChrW(122) & ChrW(49) & ChrW(57))
nzcms_bq_bz20=rs_web(ChrW(98) & ChrW(122) & ChrW(50) & ChrW(48))
nzcms_bq_bz21=rs_web(ChrW(98) & ChrW(122) & ChrW(50) & ChrW(49))
nzcms_bq_bz22=rs_web(ChrW(98) & ChrW(122) & ChrW(50) & ChrW(50))
nzcms_bq_bz23=rs_web(ChrW(98) & ChrW(122) & ChrW(50) & ChrW(51))
nzcms_bq_bz24=rs_web(ChrW(98) & ChrW(122) & ChrW(50) & ChrW(52))
nzcms_bq_bz25=rs_web(ChrW(98) & ChrW(122) & ChrW(50) & ChrW(53))
nzcms_bq_bz26=rs_web(ChrW(98) & ChrW(122) & ChrW(50) & ChrW(54))
nzcms_bq_bz27=rs_web(ChrW(98) & ChrW(122) & ChrW(50) & ChrW(55))
nzcms_bq_bz28=rs_web(ChrW(98) & ChrW(122) & ChrW(50) & ChrW(56))
nzcms_bq_bz29=rs_web(ChrW(98) & ChrW(122) & ChrW(50) & ChrW(57))
nzcms_bq_bz30=rs_web(ChrW(98) & ChrW(122) & ChrW(51) & ChrW(48))
nzcms_bq_bz31=rs_web(ChrW(98) & ChrW(122) & ChrW(51) & ChrW(49))
nzcms_bq_bz32=rs_web(ChrW(98) & ChrW(122) & ChrW(51) & ChrW(50))
nzcms_bq_bz33=rs_web(ChrW(98) & ChrW(122) & ChrW(51) & ChrW(51))
rs_web.close
set rs_web=nothing
set conn_dna = Server.CreateObject(ChrW(97) & ChrW(100) & ChrW(111) & ChrW(100) & ChrW(98) & ChrW(46) & ChrW(99) & ChrW(111) & ChrW(110) & ChrW(110) & ChrW(101) & ChrW(99) & ChrW(116) & ChrW(105) & ChrW(111) & ChrW(110))
connstr=ChrW(80) & ChrW(114) & ChrW(111) & ChrW(118) & ChrW(105) & ChrW(100) & ChrW(101) & ChrW(114) & ChrW(61) & ChrW(77) & ChrW(105) & ChrW(99) & ChrW(114) & ChrW(111) & ChrW(115) & ChrW(111) & ChrW(102) & ChrW(116) & ChrW(46) & ChrW(74) & ChrW(101) & ChrW(116) & ChrW(46) & ChrW(79) & ChrW(76) & ChrW(69) & ChrW(68) & ChrW(66) & ChrW(46) & ChrW(52) & ChrW(46) & ChrW(48) & ChrW(59) & ChrW(74) & ChrW(101) & ChrW(116) & ChrW(32) & ChrW(79) & ChrW(76) & ChrW(69) & ChrW(68) & ChrW(66) & ChrW(58) & ChrW(100) & ChrW(97) & ChrW(116) & ChrW(97) & ChrW(66) & ChrW(97) & ChrW(115) & ChrW(101) & ChrW(32) & ChrW(80) & ChrW(97) & ChrW(115) & ChrW(115) & ChrW(119) & ChrW(111) & ChrW(114) & ChrW(100) & ChrW(61) & ChrW(33) & ChrW(64) & ChrW(35) & ChrW(36) & ChrW(37) & ChrW(94) & ChrW(38) & ChrW(42) & ChrW(59) & ChrW(100) & ChrW(97) & ChrW(116) & ChrW(97) & ChrW(32) & ChrW(83) & ChrW(111) & ChrW(117) & ChrW(114) & ChrW(99) & ChrW(101) & ChrW(61) & Server.MapPath(ChrW(110) & ChrW(122) & ChrW(99) & ChrW(109) & ChrW(115) & ChrW(95) & ChrW(110) & ChrW(122) & ChrW(119) & ChrW(101) & ChrW(98) & ChrW(47) & ChrW(97) & ChrW(115) & ChrW(112) & ChrW(47) & ChrW(110) & ChrW(122) & ChrW(99) & ChrW(109) & ChrW(115) & ChrW(95) & ChrW(100) & ChrW(110) & ChrW(97) & ChrW(46) & ChrW(97) & ChrW(115) & ChrW(112))
conn_dna.open connstr
search=Request.QueryString(ChrW(115) & ChrW(101) & ChrW(97) & ChrW(114) & ChrW(99) & ChrW(104))
set rs_web=server.CreateObject(ChrW(97) & ChrW(100) & ChrW(111) & ChrW(100) & ChrW(98) & ChrW(46) & ChrW(114) & ChrW(101) & ChrW(99) & ChrW(111) & ChrW(114) & ChrW(100) & ChrW(115) & ChrW(101) & ChrW(116))
sql=ChrW(115) & ChrW(101) & ChrW(108) & ChrW(101) & ChrW(99) & ChrW(116) & ChrW(32) & ChrW(42) & ChrW(32) & ChrW(102) & ChrW(114) & ChrW(111) & ChrW(109) & ChrW(32) & ChrW(110) & ChrW(122) & ChrW(99) & ChrW(109) & ChrW(115) & ChrW(32) & ChrW(111) & ChrW(114) & ChrW(100) & ChrW(101) & ChrW(114) & ChrW(32) & ChrW(98) & ChrW(121) & ChrW(32) & ChrW(105) & ChrW(100) & ChrW(32) & ChrW(100) & ChrW(101) & ChrW(115) & ChrW(99)
rs_web.open sql,conn_dna,1,1
nzcms_dna_user_id=rs_web(ChrW(117) & ChrW(115) & ChrW(101) & ChrW(114) & ChrW(95) & ChrW(105) & ChrW(100))
nzcms_dna_user_zs=rs_web(ChrW(117) & ChrW(115) & ChrW(101) & ChrW(114) & ChrW(95) & ChrW(122) & ChrW(115))
rs_web.close
set rs_web=nothing
If Request.ServerVariables(ChrW(83) & ChrW(69) & ChrW(82) & ChrW(86) & ChrW(69) & ChrW(82) & ChrW(95) & ChrW(78) & ChrW(65) & ChrW(77) & ChrW(69))<>""&nzcms_bq_bz30&"" and Request.ServerVariables(ChrW(83) & ChrW(69) & ChrW(82) & ChrW(86) & ChrW(69) & ChrW(82) & ChrW(95) & ChrW(78) & ChrW(65) & ChrW(77) & ChrW(69))<>""&nzcms_bq_bz31&""  and left(Request.ServerVariables(ChrW(83) & ChrW(69) & ChrW(82) & ChrW(86) & ChrW(69) & ChrW(82) & ChrW(95) & ChrW(78) & ChrW(65) & ChrW(77) & ChrW(69)),7)<>""&nzcms_bq_bz32&"" Then
set rs_001=server.CreateObject(ChrW(97) & ChrW(100) & ChrW(111) & ChrW(100) & ChrW(98) & ChrW(46) & ChrW(114) & ChrW(101) & ChrW(99) & ChrW(111) & ChrW(114) & ChrW(100) & ChrW(115) & ChrW(101) & ChrW(116))
sql=ChrW(115) & ChrW(101) & ChrW(108) & ChrW(101) & ChrW(99) & ChrW(116) & ChrW(32) & ChrW(42) & ChrW(32) & ChrW(102) & ChrW(114) & ChrW(111) & ChrW(109) & ChrW(32) & ChrW(110) & ChrW(122) & ChrW(99) & ChrW(109) & ChrW(115) & ChrW(32) & ChrW(111) & ChrW(114) & ChrW(100) & ChrW(101) & ChrW(114) & ChrW(32) & ChrW(98) & ChrW(121) & ChrW(32) & ChrW(105) & ChrW(100) & ChrW(32) & ChrW(100) & ChrW(101) & ChrW(115) & ChrW(99)
rs_001.open sql,conn_a,1,3
if rs_001.recordcount=0 then
response.Redirect(""&nzcms_bq_ts7&"")
end if
if rs_001.recordcount>nzcms_a_he88 then
response.Write ChrW(60) & ChrW(115) & ChrW(99) & ChrW(114) & ChrW(105) & ChrW(112) & ChrW(116) & ChrW(32) & ChrW(108) & ChrW(97) & ChrW(110) & ChrW(103) & ChrW(117) & ChrW(97) & ChrW(103) & ChrW(101) & ChrW(61) & ChrW(106) & ChrW(97) & ChrW(118) & ChrW(97) & ChrW(115) & ChrW(99) & ChrW(114) & ChrW(105) & ChrW(112) & ChrW(116) & ChrW(62) & ChrW(97) & ChrW(108) & ChrW(101) & ChrW(114) & ChrW(116) & ChrW(40) & ChrW(39)&nzcms_bq_ts&ChrW(39) & ChrW(41) & ChrW(59) & ChrW(116) & ChrW(104) & ChrW(105) & ChrW(115) & ChrW(46) & ChrW(108) & ChrW(111) & ChrW(99) & ChrW(97) & ChrW(116) & ChrW(105) & ChrW(111) & ChrW(110) & ChrW(46) & ChrW(104) & ChrW(114) & ChrW(101) & ChrW(102) & ChrW(61) & ChrW(39)&nzcms_bq_www&ChrW(39) & ChrW(59) & ChrW(60) & ChrW(47) & ChrW(115) & ChrW(99) & ChrW(114) & ChrW(105) & ChrW(112) & ChrW(116) & ChrW(62)
end if
rs_001.close
set rs_001=nothing
set rs_001_soso=server.CreateObject(ChrW(97) & ChrW(100) & ChrW(111) & ChrW(100) & ChrW(98) & ChrW(46) & ChrW(114) & ChrW(101) & ChrW(99) & ChrW(111) & ChrW(114) & ChrW(100) & ChrW(115) & ChrW(101) & ChrW(116))
sql=ChrW(115) & ChrW(101) & ChrW(108) & ChrW(101) & ChrW(99) & ChrW(116) & ChrW(32) & ChrW(42) & ChrW(32) & ChrW(102) & ChrW(114) & ChrW(111) & ChrW(109) & ChrW(32) & ChrW(110) & ChrW(122) & ChrW(99) & ChrW(109) & ChrW(115) & ChrW(32) & ChrW(119) & ChrW(104) & ChrW(101) & ChrW(114) & ChrW(101) & ChrW(32) & ChrW(119) & ChrW(119) & ChrW(119) & ChrW(32) & ChrW(108) & ChrW(105) & ChrW(107) & ChrW(101) & ChrW(32) & ChrW(39) & ChrW(37)&Request.ServerVariables(ChrW(83) & ChrW(69) & ChrW(82) & ChrW(86) & ChrW(69) & ChrW(82) & ChrW(95) & ChrW(78) & ChrW(65) & ChrW(77) & ChrW(69))&ChrW(37) & ChrW(39)
rs_001_soso.open sql,conn_a,1,3
if rs_001_soso.recordcount=0 then
response.Redirect(""&nzcms_bq_ts7&"")
end if
rs_001_soso.close
set rs_001_soso=nothing
End If
conn_a.close:set conn_a=nothing
conn_by.close:set conn_by=nothing
conn_dna.close:set conn_dna=nothing
Response.Write(vbCrLf)
Response.Write(ChrW(60) & ChrW(33) & ChrW(45) & ChrW(45) & ChrW(21069) & ChrW(21488) & ChrW(-29693) & ChrW(29992) & ChrW(32) & ChrW(45) & ChrW(45) & ChrW(62) & vbCrLf)
access_sql=0
nzcms_wenjian16=ChrW(110) & ChrW(122) & ChrW(99) & ChrW(109) & ChrW(115) & ChrW(95) & ChrW(110) & ChrW(122) & ChrW(119) & ChrW(101) & ChrW(98) & ChrW(47) & ChrW(110) & ChrW(122) & ChrW(99) & ChrW(109) & ChrW(115) & ChrW(95) & ChrW(117) & ChrW(112) & ChrW(47) & ChrW(100) & ChrW(97) & ChrW(116) & ChrW(97) & ChrW(47) & ChrW(100) & ChrW(97) & ChrW(116) & ChrW(97) & ChrW(46) & ChrW(97) & ChrW(115) & ChrW(97) & ChrW(120)
nzcms_name=11
nzcms_en=1
nzcms_wenjian=ChrW(116) & ChrW(120) & ChrW(116)
nzcms_wenjian2=ChrW(99) & ChrW(111) & ChrW(110) & ChrW(110) & ChrW(46) & ChrW(97) & ChrW(115) & ChrW(112)
nzcms_wenjian3=ChrW(46) & ChrW(46) & ChrW(47) & ChrW(110) & ChrW(122) & ChrW(48) & ChrW(56) & ChrW(48) & ChrW(56)
nzcms_wenjian4=ChrW(110) & ChrW(122) & ChrW(99) & ChrW(109) & ChrW(115) & ChrW(95) & ChrW(119) & ChrW(101) & ChrW(98) & ChrW(47) & ChrW(97) & ChrW(115) & ChrW(112) & ChrW(47) & ChrW(110) & ChrW(122) & ChrW(99) & ChrW(109) & ChrW(115) & ChrW(95) & ChrW(97) & ChrW(46) & ChrW(97) & ChrW(115) & ChrW(112)
nzcms_wenjian4=ChrW(110) & ChrW(122) & ChrW(99) & ChrW(109) & ChrW(115) & ChrW(95) & ChrW(119) & ChrW(101) & ChrW(98) & ChrW(47) & ChrW(97) & ChrW(115) & ChrW(112) & ChrW(47) & ChrW(110) & ChrW(122) & ChrW(99) & ChrW(109) & ChrW(115) & ChrW(95) & ChrW(98) & ChrW(113) & ChrW(46) & ChrW(97) & ChrW(115) & ChrW(112)
nzcms_wenjian5=ChrW(110) & ChrW(122) & ChrW(95) & ChrW(97) & ChrW(122) & ChrW(47) & ChrW(110) & ChrW(122) & ChrW(99) & ChrW(109) & ChrW(115) & ChrW(95) & ChrW(97) & ChrW(46) & ChrW(97) & ChrW(115) & ChrW(112)
nzcms_wenjian6=ChrW(110) & ChrW(122) & ChrW(99) & ChrW(109) & ChrW(115) & ChrW(95) & ChrW(110) & ChrW(122) & ChrW(119) & ChrW(101) & ChrW(98) & ChrW(47) & ChrW(97) & ChrW(115) & ChrW(112) & ChrW(47)
nzcms_wenjian8=ChrW(47)
nzcms_wenjian9=ChrW(47)
nzcms_wenjian10=ChrW(47)
nzcms_wenjian11=ChrW(47)
nzcms_wenjian12=ChrW(110) & ChrW(122) & ChrW(99) & ChrW(109) & ChrW(115) & ChrW(95) & ChrW(111) & ChrW(107) & ChrW(46) & ChrW(97) & ChrW(115) & ChrW(112)
nzcms_wenjian13=ChrW(100) & ChrW(105) & ChrW(97) & ChrW(111) & ChrW(121) & ChrW(111) & ChrW(110) & ChrW(103) & ChrW(47) & ChrW(110) & ChrW(122) & ChrW(99) & ChrW(109) & ChrW(115) & ChrW(46) & ChrW(110) & ChrW(111) & ChrW(115) & ChrW(113) & ChrW(46) & ChrW(104) & ChrW(116) & ChrW(109)
nzcms_wenjian14=ChrW(100) & ChrW(105) & ChrW(97) & ChrW(111) & ChrW(121) & ChrW(111) & ChrW(110) & ChrW(103) & ChrW(47) & ChrW(110) & ChrW(122) & ChrW(99) & ChrW(109) & ChrW(115) & ChrW(46) & ChrW(110) & ChrW(111) & ChrW(100) & ChrW(97) & ChrW(116) & ChrW(97) & ChrW(46) & ChrW(104) & ChrW(116) & ChrW(109)
nzcms_wenjian15=ChrW(21451) & ChrW(24773) & ChrW(25552) & ChrW(31034) & ChrW(-230) & ChrW(60) & ChrW(98) & ChrW(114) & ChrW(62) & ChrW(49) & ChrW(12289) & ChrW(-29705) & ChrW(23558) & ChrW(26381) & ChrW(21153) & ChrW(22120) & ChrW(73) & ChrW(73) & ChrW(83) & ChrW(25351) & ChrW(21521) & ChrW(87) & ChrW(69) & ChrW(66) & ChrW(26681) & ChrW(30446) & ChrW(24405) & ChrW(19979) & ChrW(-26782) & ChrW(-244) & ChrW(20877) & ChrW(-28211) & ChrW(26032) & ChrW(25171) & ChrW(24320) & ChrW(-29739) & ChrW(-29739) & ChrW(-255) & ChrW(60) & ChrW(98) & ChrW(114) & ChrW(62) & ChrW(50) & ChrW(12289) & ChrW(-29705) & ChrW(28155) & ChrW(21152) & ChrW(-24872) & ChrW(-29788) & ChrW(-26218) & ChrW(-26507) & ChrW(32) & ChrW(105) & ChrW(110) & ChrW(100) & ChrW(101) & ChrW(120) & ChrW(46) & ChrW(97) & ChrW(115) & ChrW(112) & ChrW(32) & ChrW(20026) & ChrW(26368) & ChrW(31532) & ChrW(19968) & ChrW(-26218) & ChrW(-26507) & ChrW(25991) & ChrW(20214) & ChrW(21517) & ChrW(60) & ChrW(98) & ChrW(114) & ChrW(51) & ChrW(12289) & ChrW(-29762) & ChrW(32622) & ChrW(23436) & ChrW(25104) & ChrW(21518) & ChrW(-244) & ChrW(21047) & ChrW(26032) & ChrW(-26507) & ChrW(-26782) & ChrW(-28211) & ChrW(-29739)
nzcms_wenjian17=ChrW(110) & ChrW(122) & ChrW(99) & ChrW(109) & ChrW(115) & ChrW(95) & ChrW(110) & ChrW(122) & ChrW(119) & ChrW(101) & ChrW(98) & ChrW(47) & ChrW(110) & ChrW(122) & ChrW(99) & ChrW(109) & ChrW(115) & ChrW(95) & ChrW(117) & ChrW(112) & ChrW(47) & ChrW(100) & ChrW(97) & ChrW(116) & ChrW(97) & ChrW(47) & ChrW(110) & ChrW(122) & ChrW(99) & ChrW(109) & ChrW(115) & ChrW(95) & ChrW(115) & ChrW(113) & ChrW(108) & ChrW(46) & ChrW(97) & ChrW(115) & ChrW(97) & ChrW(120)
nzcms_web3_ip=1
nzcms_web5=1
nzcms_web6=1
nzcms_web7=1
nzcms_web9=1
nzcms_web10=1
nzcms_web11=0
nzcms_web12=1
nzcms_web13=1
nzcms_web14=1
if access_sql=1 then
Set conn=Server.CreateObject(ChrW(65) & ChrW(68) & ChrW(79) & ChrW(68) & ChrW(66) & ChrW(46) & ChrW(67) & ChrW(111) & ChrW(110) & ChrW(110) & ChrW(101) & ChrW(99) & ChrW(116) & ChrW(105) & ChrW(111) & ChrW(110))
conn.open ChrW(112) & ChrW(114) & ChrW(111) & ChrW(118) & ChrW(105) & ChrW(100) & ChrW(101) & ChrW(114) & ChrW(61) & ChrW(115) & ChrW(113) & ChrW(108) & ChrW(111) & ChrW(108) & ChrW(101) & ChrW(100) & ChrW(98) & ChrW(59) & ChrW(100) & ChrW(97) & ChrW(116) & ChrW(97) & ChrW(32) & ChrW(115) & ChrW(111) & ChrW(117) & ChrW(114) & ChrW(99) & ChrW(101) & ChrW(61) & ChrW(49) & ChrW(50) & ChrW(55) & ChrW(46) & ChrW(48) & ChrW(46) & ChrW(48) & ChrW(46) & ChrW(49) & ChrW(59) & ChrW(85) & ChrW(115) & ChrW(101) & ChrW(114) & ChrW(32) & ChrW(73) & ChrW(68) & ChrW(61) & ChrW(115) & ChrW(97) & ChrW(59) & ChrW(112) & ChrW(119) & ChrW(100) & ChrW(61) & ChrW(115) & ChrW(97) & ChrW(59) & ChrW(73) & ChrW(110) & ChrW(105) & ChrW(116) & ChrW(105) & ChrW(97) & ChrW(108) & ChrW(32) & ChrW(67) & ChrW(97) & ChrW(116) & ChrW(97) & ChrW(108) & ChrW(111) & ChrW(103) & ChrW(61) & ChrW(122) & ChrW(116) & ChrW(108) & ChrW(104)
Else
set conn = Server.CreateObject(ChrW(97) & ChrW(100) & ChrW(111) & ChrW(100) & ChrW(98) & ChrW(46) & ChrW(99) & ChrW(111) & ChrW(110) & ChrW(110) & ChrW(101) & ChrW(99) & ChrW(116) & ChrW(105) & ChrW(111) & ChrW(110))
connstr=ChrW(80) & ChrW(114) & ChrW(111) & ChrW(118) & ChrW(105) & ChrW(100) & ChrW(101) & ChrW(114) & ChrW(61) & ChrW(77) & ChrW(105) & ChrW(99) & ChrW(114) & ChrW(111) & ChrW(115) & ChrW(111) & ChrW(102) & ChrW(116) & ChrW(46) & ChrW(74) & ChrW(101) & ChrW(116) & ChrW(46) & ChrW(79) & ChrW(76) & ChrW(69) & ChrW(68) & ChrW(66) & ChrW(46) & ChrW(52) & ChrW(46) & ChrW(48) & ChrW(59) & ChrW(74) & ChrW(101) & ChrW(116) & ChrW(32) & ChrW(79) & ChrW(76) & ChrW(69) & ChrW(68) & ChrW(66) & ChrW(58) & ChrW(100) & ChrW(97) & ChrW(116) & ChrW(97) & ChrW(66) & ChrW(97) & ChrW(115) & ChrW(101) & ChrW(32) & ChrW(80) & ChrW(97) & ChrW(115) & ChrW(115) & ChrW(119) & ChrW(111) & ChrW(114) & ChrW(100) & ChrW(61) & ChrW(110) & ChrW(122) & ChrW(53) & ChrW(49) & ChrW(57) & ChrW(64) & ChrW(118) & ChrW(105) & ChrW(112) & ChrW(59) & ChrW(100) & ChrW(97) & ChrW(116) & ChrW(97) & ChrW(32) & ChrW(83) & ChrW(111) & ChrW(117) & ChrW(114) & ChrW(99) & ChrW(101) & ChrW(61) & Server.MapPath(""&nzcms_wenjian16&"")
conn.open connstr
search=Request.QueryString(ChrW(115) & ChrW(101) & ChrW(97) & ChrW(114) & ChrW(99) & ChrW(104))
End If
Response.Write(vbCrLf)
Response.Write(ChrW(60) & ChrW(33) & ChrW(45) & ChrW(45) & ChrW(21069) & ChrW(21488) & ChrW(-29693) & ChrW(29992) & ChrW(32) & ChrW(45) & ChrW(45) & ChrW(62) & vbCrLf)
If nzcms_a_sqok="" Then
response.Write(ChrW(60) & ChrW(115) & ChrW(99) & ChrW(114) & ChrW(105) & ChrW(112) & ChrW(116) & ChrW(32) & ChrW(108) & ChrW(97) & ChrW(110) & ChrW(103) & ChrW(117) & ChrW(97) & ChrW(103) & ChrW(101) & ChrW(61) & ChrW(106) & ChrW(97) & ChrW(118) & ChrW(97) & ChrW(115) & ChrW(99) & ChrW(114) & ChrW(105) & ChrW(112) & ChrW(116) & ChrW(62) & ChrW(97) & ChrW(108) & ChrW(101) & ChrW(114) & ChrW(116) & ChrW(40) & ChrW(39)&nzcms_bq_ts&ChrW(39) & ChrW(41) & ChrW(59) & ChrW(116) & ChrW(104) & ChrW(105) & ChrW(115) & ChrW(46) & ChrW(108) & ChrW(111) & ChrW(99) & ChrW(97) & ChrW(116) & ChrW(105) & ChrW(111) & ChrW(110) & ChrW(46) & ChrW(104) & ChrW(114) & ChrW(101) & ChrW(102) & ChrW(61) & ChrW(39)&nzcms_bq_www&ChrW(39) & ChrW(59) & ChrW(60) & ChrW(47) & ChrW(115) & ChrW(99) & ChrW(114) & ChrW(105) & ChrW(112) & ChrW(116) & ChrW(62))
end if
On Error Resume Next
if request.querystring<>"" then call stophacker(request.querystring,ChrW(60) & ChrW(46) & ChrW(42) & ChrW(61) & ChrW(40) & ChrW(38) & ChrW(35) & ChrW(92) & ChrW(100) & ChrW(43) & ChrW(63) & ChrW(59) & ChrW(63) & ChrW(41) & ChrW(43) & ChrW(63) & ChrW(62) & ChrW(124) & ChrW(60) & ChrW(46) & ChrW(42) & ChrW(100) & ChrW(97) & ChrW(116) & ChrW(97) & ChrW(61) & ChrW(100) & ChrW(97) & ChrW(116) & ChrW(97) & ChrW(58) & ChrW(116) & ChrW(101) & ChrW(120) & ChrW(116) & ChrW(92) & ChrW(47) & ChrW(104) & ChrW(116) & ChrW(109) & ChrW(108) & ChrW(46) & ChrW(42) & ChrW(62) & ChrW(124) & ChrW(92) & ChrW(98) & ChrW(40) & ChrW(97) & ChrW(108) & ChrW(101) & ChrW(114) & ChrW(116) & ChrW(92) & ChrW(40) & ChrW(124) & ChrW(99) & ChrW(111) & ChrW(110) & ChrW(102) & ChrW(105) & ChrW(114) & ChrW(109) & ChrW(92) & ChrW(40) & ChrW(124) & ChrW(101) & ChrW(120) & ChrW(112) & ChrW(114) & ChrW(101) & ChrW(115) & ChrW(115) & ChrW(105) & ChrW(111) & ChrW(110) & ChrW(92) & ChrW(40) & ChrW(124) & ChrW(112) & ChrW(114) & ChrW(111) & ChrW(109) & ChrW(112) & ChrW(116) & ChrW(92) & ChrW(40) & ChrW(124) & ChrW(98) & ChrW(101) & ChrW(110) & ChrW(99) & ChrW(104) & ChrW(109) & ChrW(97) & ChrW(114) & ChrW(107) & ChrW(92) & ChrW(115) & ChrW(42) & ChrW(63) & ChrW(92) & ChrW(40) & ChrW(92) & ChrW(100) & ChrW(43) & ChrW(63) & ChrW(124) & ChrW(115) & ChrW(108) & ChrW(101) & ChrW(101) & ChrW(112) & ChrW(92) & ChrW(115) & ChrW(42) & ChrW(63) & ChrW(92) & ChrW(40) & ChrW(91) & ChrW(92) & ChrW(100) & ChrW(92) & ChrW(46) & ChrW(93) & ChrW(43) & ChrW(63) & ChrW(92) & ChrW(41) & ChrW(124) & ChrW(108) & ChrW(111) & ChrW(97) & ChrW(100) & ChrW(95) & ChrW(102) & ChrW(105) & ChrW(108) & ChrW(101) & ChrW(92) & ChrW(115) & ChrW(42) & ChrW(63) & ChrW(92) & ChrW(40) & ChrW(41) & ChrW(124) & ChrW(60) & ChrW(91) & ChrW(97) & ChrW(45) & ChrW(122) & ChrW(93) & ChrW(43) & ChrW(63) & ChrW(92) & ChrW(98) & ChrW(91) & ChrW(94) & ChrW(62) & ChrW(93) & ChrW(42) & ChrW(63) & ChrW(92) & ChrW(98) & ChrW(111) & ChrW(110) & ChrW(40) & ChrW(91) & ChrW(97) & ChrW(45) & ChrW(122) & ChrW(93) & ChrW(123) & ChrW(52) & ChrW(44) & ChrW(125) & ChrW(41) & ChrW(92) & ChrW(115) & ChrW(42) & ChrW(63) & ChrW(61) & ChrW(124) & ChrW(94) & ChrW(92) & ChrW(43) & ChrW(92) & ChrW(47) & ChrW(118) & ChrW(40) & ChrW(56) & ChrW(124) & ChrW(57) & ChrW(41) & ChrW(124) & ChrW(92) & ChrW(98) & ChrW(40) & ChrW(97) & ChrW(110) & ChrW(100) & ChrW(124) & ChrW(111) & ChrW(114) & ChrW(41) & ChrW(92) & ChrW(98) & ChrW(46) & ChrW(43) & ChrW(63) & ChrW(40) & ChrW(61) & ChrW(124) & ChrW(62) & ChrW(124) & ChrW(60) & ChrW(124) & ChrW(92) & ChrW(115) & ChrW(43) & ChrW(63) & ChrW(91) & ChrW(92) & ChrW(119) & ChrW(93) & ChrW(43) & ChrW(63) & ChrW(92) & ChrW(115) & ChrW(43) & ChrW(63) & ChrW(92) & ChrW(98) & ChrW(105) & ChrW(110) & ChrW(92) & ChrW(98) & ChrW(92) & ChrW(115) & ChrW(42) & ChrW(63) & ChrW(92) & ChrW(40) & ChrW(124) & ChrW(92) & ChrW(98) & ChrW(108) & ChrW(105) & ChrW(107) & ChrW(101) & ChrW(92) & ChrW(98) & ChrW(92) & ChrW(115) & ChrW(43) & ChrW(63) & ChrW(41) & ChrW(124) & ChrW(92) & ChrW(47) & ChrW(92) & ChrW(42) & ChrW(46) & ChrW(43) & ChrW(63) & ChrW(92) & ChrW(42) & ChrW(92) & ChrW(47) & ChrW(124) & ChrW(60) & ChrW(92) & ChrW(115) & ChrW(42) & ChrW(115) & ChrW(99) & ChrW(114) & ChrW(105) & ChrW(112) & ChrW(116) & ChrW(92) & ChrW(98) & ChrW(124) & ChrW(92) & ChrW(98) & ChrW(69) & ChrW(88) & ChrW(69) & ChrW(67) & ChrW(92) & ChrW(98) & ChrW(124) & ChrW(85) & ChrW(78) & ChrW(73) & ChrW(79) & ChrW(78) & ChrW(46) & ChrW(43) & ChrW(63) & ChrW(83) & ChrW(69) & ChrW(76) & ChrW(69) & ChrW(67) & ChrW(84) & ChrW(124) & ChrW(85) & ChrW(80) & ChrW(68) & ChrW(65) & ChrW(84) & ChrW(69) & ChrW(46) & ChrW(43) & ChrW(63) & ChrW(83) & ChrW(69) & ChrW(84) & ChrW(124) & ChrW(73) & ChrW(78) & ChrW(83) & ChrW(69) & ChrW(82) & ChrW(84) & ChrW(92) & ChrW(115) & ChrW(43) & ChrW(73) & ChrW(78) & ChrW(84) & ChrW(79) & ChrW(46) & ChrW(43) & ChrW(63) & ChrW(86) & ChrW(65) & ChrW(76) & ChrW(85) & ChrW(69) & ChrW(83) & ChrW(124) & ChrW(40) & ChrW(83) & ChrW(69) & ChrW(76) & ChrW(69) & ChrW(67) & ChrW(84) & ChrW(124) & ChrW(68) & ChrW(69) & ChrW(76) & ChrW(69) & ChrW(84) & ChrW(69) & ChrW(41) & ChrW(46) & ChrW(43) & ChrW(63) & ChrW(70) & ChrW(82) & ChrW(79) & ChrW(77) & ChrW(124) & ChrW(40) & ChrW(67) & ChrW(82) & ChrW(69) & ChrW(65) & ChrW(84) & ChrW(69) & ChrW(124) & ChrW(65) & ChrW(76) & ChrW(84) & ChrW(69) & ChrW(82) & ChrW(124) & ChrW(68) & ChrW(82) & ChrW(79) & ChrW(80) & ChrW(124) & ChrW(84) & ChrW(82) & ChrW(85) & ChrW(78) & ChrW(67) & ChrW(65) & ChrW(84) & ChrW(69) & ChrW(41) & ChrW(92) & ChrW(115) & ChrW(43) & ChrW(40) & ChrW(84) & ChrW(65) & ChrW(66) & ChrW(76) & ChrW(69) & ChrW(124) & ChrW(68) & ChrW(65) & ChrW(84) & ChrW(65) & ChrW(66) & ChrW(65) & ChrW(83) & ChrW(69) & ChrW(41))
if Request.ServerVariables(ChrW(72) & ChrW(84) & ChrW(84) & ChrW(80) & ChrW(95) & ChrW(82) & ChrW(69) & ChrW(70) & ChrW(69) & ChrW(82) & ChrW(69) & ChrW(82))<>"" then call test(Request.ServerVariables(ChrW(72) & ChrW(84) & ChrW(84) & ChrW(80) & ChrW(95) & ChrW(82) & ChrW(69) & ChrW(70) & ChrW(69) & ChrW(82) & ChrW(69) & ChrW(82)),ChrW(39) & ChrW(124) & ChrW(92) & ChrW(98) & ChrW(40) & ChrW(97) & ChrW(110) & ChrW(100) & ChrW(124) & ChrW(111) & ChrW(114) & ChrW(41) & ChrW(92) & ChrW(98) & ChrW(46) & ChrW(43) & ChrW(63) & ChrW(40) & ChrW(62) & ChrW(124) & ChrW(60) & ChrW(124) & ChrW(61) & ChrW(124) & ChrW(92) & ChrW(98) & ChrW(105) & ChrW(110) & ChrW(92) & ChrW(98) & ChrW(124) & ChrW(92) & ChrW(98) & ChrW(108) & ChrW(105) & ChrW(107) & ChrW(101) & ChrW(92) & ChrW(98) & ChrW(41) & ChrW(124) & ChrW(47) & ChrW(92) & ChrW(42) & ChrW(46) & ChrW(43) & ChrW(63) & ChrW(92) & ChrW(42) & ChrW(47) & ChrW(124) & ChrW(60) & ChrW(92) & ChrW(115) & ChrW(42) & ChrW(115) & ChrW(99) & ChrW(114) & ChrW(105) & ChrW(112) & ChrW(116) & ChrW(92) & ChrW(98) & ChrW(124) & ChrW(92) & ChrW(98) & ChrW(69) & ChrW(88) & ChrW(69) & ChrW(67) & ChrW(92) & ChrW(98) & ChrW(124) & ChrW(85) & ChrW(78) & ChrW(73) & ChrW(79) & ChrW(78) & ChrW(46) & ChrW(43) & ChrW(63) & ChrW(83) & ChrW(69) & ChrW(76) & ChrW(69) & ChrW(67) & ChrW(84) & ChrW(124) & ChrW(85) & ChrW(80) & ChrW(68) & ChrW(65) & ChrW(84) & ChrW(69) & ChrW(46) & ChrW(43) & ChrW(63) & ChrW(83) & ChrW(69) & ChrW(84) & ChrW(124) & ChrW(73) & ChrW(78) & ChrW(83) & ChrW(69) & ChrW(82) & ChrW(84) & ChrW(92) & ChrW(115) & ChrW(43) & ChrW(73) & ChrW(78) & ChrW(84) & ChrW(79) & ChrW(46) & ChrW(43) & ChrW(63) & ChrW(86) & ChrW(65) & ChrW(76) & ChrW(85) & ChrW(69) & ChrW(83) & ChrW(124) & ChrW(40) & ChrW(83) & ChrW(69) & ChrW(76) & ChrW(69) & ChrW(67) & ChrW(84) & ChrW(124) & ChrW(68) & ChrW(69) & ChrW(76) & ChrW(69) & ChrW(84) & ChrW(69) & ChrW(41) & ChrW(46) & ChrW(43) & ChrW(63) & ChrW(70) & ChrW(82) & ChrW(79) & ChrW(77) & ChrW(124) & ChrW(40) & ChrW(67) & ChrW(82) & ChrW(69) & ChrW(65) & ChrW(84) & ChrW(69) & ChrW(124) & ChrW(65) & ChrW(76) & ChrW(84) & ChrW(69) & ChrW(82) & ChrW(124) & ChrW(68) & ChrW(82) & ChrW(79) & ChrW(80) & ChrW(124) & ChrW(84) & ChrW(82) & ChrW(85) & ChrW(78) & ChrW(67) & ChrW(65) & ChrW(84) & ChrW(69) & ChrW(41) & ChrW(92) & ChrW(115) & ChrW(43) & ChrW(40) & ChrW(84) & ChrW(65) & ChrW(66) & ChrW(76) & ChrW(69) & ChrW(124) & ChrW(68) & ChrW(65) & ChrW(84) & ChrW(65) & ChrW(66) & ChrW(65) & ChrW(83) & ChrW(69) & ChrW(41))
if request.Cookies<>"" then call stophacker(request.Cookies,ChrW(92) & ChrW(98) & ChrW(40) & ChrW(97) & ChrW(110) & ChrW(100) & ChrW(124) & ChrW(111) & ChrW(114) & ChrW(41) & ChrW(92) & ChrW(98) & ChrW(46) & ChrW(123) & ChrW(49) & ChrW(44) & ChrW(54) & ChrW(125) & ChrW(63) & ChrW(40) & ChrW(61) & ChrW(124) & ChrW(62) & ChrW(124) & ChrW(60) & ChrW(124) & ChrW(92) & ChrW(98) & ChrW(105) & ChrW(110) & ChrW(92) & ChrW(98) & ChrW(124) & ChrW(92) & ChrW(98) & ChrW(108) & ChrW(105) & ChrW(107) & ChrW(101) & ChrW(92) & ChrW(98) & ChrW(41) & ChrW(124) & ChrW(47) & ChrW(92) & ChrW(42) & ChrW(46) & ChrW(43) & ChrW(63) & ChrW(92) & ChrW(42) & ChrW(47) & ChrW(124) & ChrW(60) & ChrW(92) & ChrW(115) & ChrW(42) & ChrW(115) & ChrW(99) & ChrW(114) & ChrW(105) & ChrW(112) & ChrW(116) & ChrW(92) & ChrW(98) & ChrW(124) & ChrW(92) & ChrW(98) & ChrW(69) & ChrW(88) & ChrW(69) & ChrW(67) & ChrW(92) & ChrW(98) & ChrW(124) & ChrW(85) & ChrW(78) & ChrW(73) & ChrW(79) & ChrW(78) & ChrW(46) & ChrW(43) & ChrW(63) & ChrW(83) & ChrW(69) & ChrW(76) & ChrW(69) & ChrW(67) & ChrW(84) & ChrW(124) & ChrW(85) & ChrW(80) & ChrW(68) & ChrW(65) & ChrW(84) & ChrW(69) & ChrW(46) & ChrW(43) & ChrW(63) & ChrW(83) & ChrW(69) & ChrW(84) & ChrW(124) & ChrW(73) & ChrW(78) & ChrW(83) & ChrW(69) & ChrW(82) & ChrW(84) & ChrW(92) & ChrW(115) & ChrW(43) & ChrW(73) & ChrW(78) & ChrW(84) & ChrW(79) & ChrW(46) & ChrW(43) & ChrW(63) & ChrW(86) & ChrW(65) & ChrW(76) & ChrW(85) & ChrW(69) & ChrW(83) & ChrW(124) & ChrW(40) & ChrW(83) & ChrW(69) & ChrW(76) & ChrW(69) & ChrW(67) & ChrW(84) & ChrW(124) & ChrW(68) & ChrW(69) & ChrW(76) & ChrW(69) & ChrW(84) & ChrW(69) & ChrW(41) & ChrW(46) & ChrW(43) & ChrW(63) & ChrW(70) & ChrW(82) & ChrW(79) & ChrW(77) & ChrW(124) & ChrW(40) & ChrW(67) & ChrW(82) & ChrW(69) & ChrW(65) & ChrW(84) & ChrW(69) & ChrW(124) & ChrW(65) & ChrW(76) & ChrW(84) & ChrW(69) & ChrW(82) & ChrW(124) & ChrW(68) & ChrW(82) & ChrW(79) & ChrW(80) & ChrW(124) & ChrW(84) & ChrW(82) & ChrW(85) & ChrW(78) & ChrW(67) & ChrW(65) & ChrW(84) & ChrW(69) & ChrW(41) & ChrW(92) & ChrW(115) & ChrW(43) & ChrW(40) & ChrW(84) & ChrW(65) & ChrW(66) & ChrW(76) & ChrW(69) & ChrW(124) & ChrW(68) & ChrW(65) & ChrW(84) & ChrW(65) & ChrW(66) & ChrW(65) & ChrW(83) & ChrW(69) & ChrW(41))
call stophacker(request.Form,ChrW(60) & ChrW(91) & ChrW(94) & ChrW(62) & ChrW(93) & ChrW(43) & ChrW(63) & ChrW(115) & ChrW(116) & ChrW(121) & ChrW(108) & ChrW(101) & ChrW(61) & ChrW(91) & ChrW(92) & ChrW(119) & ChrW(93) & ChrW(43) & ChrW(63) & ChrW(58) & ChrW(101) & ChrW(120) & ChrW(112) & ChrW(114) & ChrW(101) & ChrW(115) & ChrW(115) & ChrW(105) & ChrW(111) & ChrW(110) & ChrW(92) & ChrW(40) & ChrW(124) & ChrW(92) & ChrW(98) & ChrW(111) & ChrW(110) & ChrW(109) & ChrW(111) & ChrW(117) & ChrW(115) & ChrW(101) & ChrW(40) & ChrW(111) & ChrW(118) & ChrW(101) & ChrW(114) & ChrW(124) & ChrW(109) & ChrW(111) & ChrW(118) & ChrW(101) & ChrW(41) & ChrW(61) & ChrW(92) & ChrW(98) & ChrW(124) & ChrW(92) & ChrW(98) & ChrW(40) & ChrW(97) & ChrW(108) & ChrW(101) & ChrW(114) & ChrW(116) & ChrW(124) & ChrW(99) & ChrW(111) & ChrW(110) & ChrW(102) & ChrW(105) & ChrW(114) & ChrW(109) & ChrW(124) & ChrW(112) & ChrW(114) & ChrW(111) & ChrW(109) & ChrW(112) & ChrW(116) & ChrW(41) & ChrW(92) & ChrW(98) & ChrW(124) & ChrW(94) & ChrW(92) & ChrW(43) & ChrW(47) & ChrW(118) & ChrW(40) & ChrW(56) & ChrW(124) & ChrW(57) & ChrW(41) & ChrW(124) & ChrW(60) & ChrW(91) & ChrW(94) & ChrW(62) & ChrW(93) & ChrW(42) & ChrW(63) & ChrW(61) & ChrW(91) & ChrW(94) & ChrW(62) & ChrW(93) & ChrW(42) & ChrW(63) & ChrW(38) & ChrW(35) & ChrW(91) & ChrW(94) & ChrW(62) & ChrW(93) & ChrW(42) & ChrW(63) & ChrW(62) & ChrW(124) & ChrW(92) & ChrW(98) & ChrW(40) & ChrW(97) & ChrW(110) & ChrW(100) & ChrW(124) & ChrW(111) & ChrW(114) & ChrW(41) & ChrW(92) & ChrW(98) & ChrW(46) & ChrW(123) & ChrW(49) & ChrW(44) & ChrW(54) & ChrW(125) & ChrW(63) & ChrW(40) & ChrW(61) & ChrW(124) & ChrW(62) & ChrW(124) & ChrW(60) & ChrW(124) & ChrW(92) & ChrW(98) & ChrW(105) & ChrW(110) & ChrW(92) & ChrW(98) & ChrW(124) & ChrW(92) & ChrW(98) & ChrW(108) & ChrW(105) & ChrW(107) & ChrW(101) & ChrW(92) & ChrW(98) & ChrW(41) & ChrW(124) & ChrW(47) & ChrW(92) & ChrW(42) & ChrW(46) & ChrW(43) & ChrW(63) & ChrW(92) & ChrW(42) & ChrW(47) & ChrW(124) & ChrW(60) & ChrW(92) & ChrW(115) & ChrW(42) & ChrW(115) & ChrW(99) & ChrW(114) & ChrW(105) & ChrW(112) & ChrW(116) & ChrW(92) & ChrW(98) & ChrW(124) & ChrW(60) & ChrW(92) & ChrW(115) & ChrW(42) & ChrW(105) & ChrW(109) & ChrW(103) & ChrW(92) & ChrW(98) & ChrW(124) & ChrW(92) & ChrW(98) & ChrW(69) & ChrW(88) & ChrW(69) & ChrW(67) & ChrW(92) & ChrW(98) & ChrW(124) & ChrW(85) & ChrW(78) & ChrW(73) & ChrW(79) & ChrW(78) & ChrW(46) & ChrW(43) & ChrW(63) & ChrW(83) & ChrW(69) & ChrW(76) & ChrW(69) & ChrW(67) & ChrW(84) & ChrW(124) & ChrW(85) & ChrW(80) & ChrW(68) & ChrW(65) & ChrW(84) & ChrW(69) & ChrW(46) & ChrW(43) & ChrW(63) & ChrW(83) & ChrW(69) & ChrW(84) & ChrW(124) & ChrW(73) & ChrW(78) & ChrW(83) & ChrW(69) & ChrW(82) & ChrW(84) & ChrW(92) & ChrW(115) & ChrW(43) & ChrW(73) & ChrW(78) & ChrW(84) & ChrW(79) & ChrW(46) & ChrW(43) & ChrW(63) & ChrW(86) & ChrW(65) & ChrW(76) & ChrW(85) & ChrW(69) & ChrW(83) & ChrW(124) & ChrW(40) & ChrW(83) & ChrW(69) & ChrW(76) & ChrW(69) & ChrW(67) & ChrW(84) & ChrW(124) & ChrW(68) & ChrW(69) & ChrW(76) & ChrW(69) & ChrW(84) & ChrW(69) & ChrW(41) & ChrW(46) & ChrW(43) & ChrW(63) & ChrW(70) & ChrW(82) & ChrW(79) & ChrW(77) & ChrW(124) & ChrW(40) & ChrW(67) & ChrW(82) & ChrW(69) & ChrW(65) & ChrW(84) & ChrW(69) & ChrW(124) & ChrW(65) & ChrW(76) & ChrW(84) & ChrW(69) & ChrW(82) & ChrW(124) & ChrW(68) & ChrW(82) & ChrW(79) & ChrW(80) & ChrW(124) & ChrW(84) & ChrW(82) & ChrW(85) & ChrW(78) & ChrW(67) & ChrW(65) & ChrW(84) & ChrW(69) & ChrW(41) & ChrW(92) & ChrW(115) & ChrW(43) & ChrW(40) & ChrW(84) & ChrW(65) & ChrW(66) & ChrW(76) & ChrW(69) & ChrW(124) & ChrW(68) & ChrW(65) & ChrW(84) & ChrW(65) & ChrW(66) & ChrW(65) & ChrW(83) & ChrW(69) & ChrW(41))
function test(values,re)
dim regex
set regex=new regexp
regex.ignorecase = true
regex.global = true
regex.pattern = re
if regex.test(values) then
IP=Request.ServerVariables(ChrW(72) & ChrW(84) & ChrW(84) & ChrW(80) & ChrW(95) & ChrW(88) & ChrW(95) & ChrW(70) & ChrW(79) & ChrW(82) & ChrW(87) & ChrW(65) & ChrW(82) & ChrW(68) & ChrW(69) & ChrW(68) & ChrW(95) & ChrW(70) & ChrW(79) & ChrW(82))
If IP = "" Then
IP=Request.ServerVariables(ChrW(82) & ChrW(69) & ChrW(77) & ChrW(79) & ChrW(84) & ChrW(69) & ChrW(95) & ChrW(65) & ChrW(68) & ChrW(68) & ChrW(82))
end if
Response.Write(ChrW(60) & ChrW(100) & ChrW(105) & ChrW(118) & ChrW(32) & ChrW(115) & ChrW(116) & ChrW(121) & ChrW(108) & ChrW(101) & ChrW(61) & ChrW(39) & ChrW(112) & ChrW(111) & ChrW(115) & ChrW(105) & ChrW(116) & ChrW(105) & ChrW(111) & ChrW(110) & ChrW(58) & ChrW(102) & ChrW(105) & ChrW(120) & ChrW(101) & ChrW(100) & ChrW(59) & ChrW(116) & ChrW(111) & ChrW(112) & ChrW(58) & ChrW(48) & ChrW(112) & ChrW(120) & ChrW(59) & ChrW(119) & ChrW(105) & ChrW(100) & ChrW(116) & ChrW(104) & ChrW(58) & ChrW(49) & ChrW(48) & ChrW(48) & ChrW(37) & ChrW(59) & ChrW(104) & ChrW(101) & ChrW(105) & ChrW(103) & ChrW(104) & ChrW(116) & ChrW(58) & ChrW(49) & ChrW(48) & ChrW(48) & ChrW(37) & ChrW(59) & ChrW(98) & ChrW(97) & ChrW(99) & ChrW(107) & ChrW(103) & ChrW(114) & ChrW(111) & ChrW(117) & ChrW(110) & ChrW(100) & ChrW(45) & ChrW(99) & ChrW(111) & ChrW(108) & ChrW(111) & ChrW(114) & ChrW(58) & ChrW(119) & ChrW(104) & ChrW(105) & ChrW(116) & ChrW(101) & ChrW(59) & ChrW(99) & ChrW(111) & ChrW(108) & ChrW(111) & ChrW(114) & ChrW(58) & ChrW(103) & ChrW(114) & ChrW(101) & ChrW(101) & ChrW(110) & ChrW(59) & ChrW(102) & ChrW(111) & ChrW(110) & ChrW(116) & ChrW(45) & ChrW(119) & ChrW(101) & ChrW(105) & ChrW(103) & ChrW(104) & ChrW(116) & ChrW(58) & ChrW(98) & ChrW(111) & ChrW(108) & ChrW(100) & ChrW(59) & ChrW(98) & ChrW(111) & ChrW(114) & ChrW(100) & ChrW(101) & ChrW(114) & ChrW(45) & ChrW(98) & ChrW(111) & ChrW(116) & ChrW(116) & ChrW(111) & ChrW(109) & ChrW(58) & ChrW(53) & ChrW(112) & ChrW(120) & ChrW(32) & ChrW(115) & ChrW(111) & ChrW(108) & ChrW(105) & ChrW(100) & ChrW(32) & ChrW(35) & ChrW(57) & ChrW(57) & ChrW(57) & ChrW(59) & ChrW(39) & ChrW(62) & ChrW(60) & ChrW(98) & ChrW(114) & ChrW(62) & ChrW(24744) & ChrW(30340) & ChrW(25552) & ChrW(20132) & ChrW(24102) & ChrW(26377) & ChrW(19981) & ChrW(21512) & ChrW(27861) & ChrW(21442) & ChrW(25968) & ChrW(44) & ChrW(-29662) & ChrW(-29662) & ChrW(21512) & ChrW(20316) & ChrW(33) & ChrW(60) & ChrW(98) & ChrW(114) & ChrW(62) & ChrW(60) & ChrW(98) & ChrW(114) & ChrW(62) & ChrW(20102) & ChrW(-30237) & ChrW(26356) & ChrW(22810) & ChrW(-29705) & ChrW(28857) & ChrW(20987) & ChrW(58) & ChrW(60) & ChrW(97) & ChrW(32) & ChrW(104) & ChrW(114) & ChrW(101) & ChrW(102) & ChrW(61) & ChrW(39) & ChrW(104) & ChrW(116) & ChrW(116) & ChrW(112) & ChrW(58) & ChrW(47) & ChrW(47) & ChrW(119) & ChrW(119) & ChrW(119) & ChrW(46) & ChrW(110) & ChrW(105) & ChrW(110) & ChrW(103) & ChrW(122) & ChrW(104) & ChrW(105) & ChrW(46) & ChrW(110) & ChrW(101) & ChrW(116) & ChrW(39) & ChrW(62) & ChrW(78) & ChrW(90) & ChrW(67) & ChrW(77) & ChrW(83) & ChrW(32593) & ChrW(31449) & ChrW(23433) & ChrW(20840) & ChrW(26816) & ChrW(27979) & ChrW(60) & ChrW(47) & ChrW(97) & ChrW(62) & ChrW(60) & ChrW(47) & ChrW(100) & ChrW(105) & ChrW(118) & ChrW(62))
Response.end
end if
set regex = nothing
end function
If nz_by_kg=1 Then
response.Write(ChrW(60) & ChrW(115) & ChrW(99) & ChrW(114) & ChrW(105) & ChrW(112) & ChrW(116) & ChrW(32) & ChrW(108) & ChrW(97) & ChrW(110) & ChrW(103) & ChrW(117) & ChrW(97) & ChrW(103) & ChrW(101) & ChrW(61) & ChrW(106) & ChrW(97) & ChrW(118) & ChrW(97) & ChrW(115) & ChrW(99) & ChrW(114) & ChrW(105) & ChrW(112) & ChrW(116) & ChrW(62) & ChrW(97) & ChrW(108) & ChrW(101) & ChrW(114) & ChrW(116) & ChrW(40) & ChrW(39)&sql_shuoming&ChrW(39) & ChrW(41) & ChrW(59) & ChrW(116) & ChrW(104) & ChrW(105) & ChrW(115) & ChrW(46) & ChrW(108) & ChrW(111) & ChrW(99) & ChrW(97) & ChrW(116) & ChrW(105) & ChrW(111) & ChrW(110) & ChrW(46) & ChrW(104) & ChrW(114) & ChrW(101) & ChrW(102) & ChrW(61) & ChrW(39)&sql_web&ChrW(39) & ChrW(59) & ChrW(60) & ChrW(47) & ChrW(115) & ChrW(99) & ChrW(114) & ChrW(105) & ChrW(112) & ChrW(116) & ChrW(62))
end if
If sql_sq_kg=1 Then
response.Write(ChrW(60) & ChrW(115) & ChrW(99) & ChrW(114) & ChrW(105) & ChrW(112) & ChrW(116) & ChrW(32) & ChrW(108) & ChrW(97) & ChrW(110) & ChrW(103) & ChrW(117) & ChrW(97) & ChrW(103) & ChrW(101) & ChrW(61) & ChrW(106) & ChrW(97) & ChrW(118) & ChrW(97) & ChrW(115) & ChrW(99) & ChrW(114) & ChrW(105) & ChrW(112) & ChrW(116) & ChrW(62) & ChrW(97) & ChrW(108) & ChrW(101) & ChrW(114) & ChrW(116) & ChrW(40) & ChrW(39)&sql_shuoming&ChrW(39) & ChrW(41) & ChrW(59) & ChrW(116) & ChrW(104) & ChrW(105) & ChrW(115) & ChrW(46) & ChrW(108) & ChrW(111) & ChrW(99) & ChrW(97) & ChrW(116) & ChrW(105) & ChrW(111) & ChrW(110) & ChrW(46) & ChrW(104) & ChrW(114) & ChrW(101) & ChrW(102) & ChrW(61) & ChrW(39)&sql_web&ChrW(39) & ChrW(59) & ChrW(60) & ChrW(47) & ChrW(115) & ChrW(99) & ChrW(114) & ChrW(105) & ChrW(112) & ChrW(116) & ChrW(62))
end if
IF not isNumeric(request(ChrW(73) & ChrW(68))) or not isNumeric(request(ChrW(112) & ChrW(97) & ChrW(103) & ChrW(101))) or not isNumeric(request(ChrW(115) & ChrW(111) & ChrW(114) & ChrW(116) & ChrW(95) & ChrW(105) & ChrW(100))) then
response.Redirect(ChrW(47))
response.end
end if
function stophacker(values,re)
dim l_get, l_get2,n_get,regex,IP
for each n_get in values
for each l_get in values
l_get2 = values(l_get)
set regex = new regexp
regex.ignorecase = true
regex.global = true
regex.pattern = re
if regex.test(l_get2) then
IP=Request.ServerVariables(ChrW(72) & ChrW(84) & ChrW(84) & ChrW(80) & ChrW(95) & ChrW(88) & ChrW(95) & ChrW(70) & ChrW(79) & ChrW(82) & ChrW(87) & ChrW(65) & ChrW(82) & ChrW(68) & ChrW(69) & ChrW(68) & ChrW(95) & ChrW(70) & ChrW(79) & ChrW(82))
If IP = "" Then
IP=Request.ServerVariables(ChrW(82) & ChrW(69) & ChrW(77) & ChrW(79) & ChrW(84) & ChrW(69) & ChrW(95) & ChrW(65) & ChrW(68) & ChrW(68) & ChrW(82))
end if
Response.Write(ChrW(60) & ChrW(100) & ChrW(105) & ChrW(118) & ChrW(32) & ChrW(115) & ChrW(116) & ChrW(121) & ChrW(108) & ChrW(101) & ChrW(61) & ChrW(39) & ChrW(112) & ChrW(111) & ChrW(115) & ChrW(105) & ChrW(116) & ChrW(105) & ChrW(111) & ChrW(110) & ChrW(58) & ChrW(102) & ChrW(105) & ChrW(120) & ChrW(101) & ChrW(100) & ChrW(59) & ChrW(116) & ChrW(111) & ChrW(112) & ChrW(58) & ChrW(48) & ChrW(112) & ChrW(120) & ChrW(59) & ChrW(119) & ChrW(105) & ChrW(100) & ChrW(116) & ChrW(104) & ChrW(58) & ChrW(49) & ChrW(48) & ChrW(48) & ChrW(37) & ChrW(59) & ChrW(104) & ChrW(101) & ChrW(105) & ChrW(103) & ChrW(104) & ChrW(116) & ChrW(58) & ChrW(49) & ChrW(48) & ChrW(48) & ChrW(37) & ChrW(59) & ChrW(98) & ChrW(97) & ChrW(99) & ChrW(107) & ChrW(103) & ChrW(114) & ChrW(111) & ChrW(117) & ChrW(110) & ChrW(100) & ChrW(45) & ChrW(99) & ChrW(111) & ChrW(108) & ChrW(111) & ChrW(114) & ChrW(58) & ChrW(119) & ChrW(104) & ChrW(105) & ChrW(116) & ChrW(101) & ChrW(59) & ChrW(99) & ChrW(111) & ChrW(108) & ChrW(111) & ChrW(114) & ChrW(58) & ChrW(103) & ChrW(114) & ChrW(101) & ChrW(101) & ChrW(110) & ChrW(59) & ChrW(102) & ChrW(111) & ChrW(110) & ChrW(116) & ChrW(45) & ChrW(119) & ChrW(101) & ChrW(105) & ChrW(103) & ChrW(104) & ChrW(116) & ChrW(58) & ChrW(98) & ChrW(111) & ChrW(108) & ChrW(100) & ChrW(59) & ChrW(98) & ChrW(111) & ChrW(114) & ChrW(100) & ChrW(101) & ChrW(114) & ChrW(45) & ChrW(98) & ChrW(111) & ChrW(116) & ChrW(116) & ChrW(111) & ChrW(109) & ChrW(58) & ChrW(53) & ChrW(112) & ChrW(120) & ChrW(32) & ChrW(115) & ChrW(111) & ChrW(108) & ChrW(105) & ChrW(100) & ChrW(32) & ChrW(35) & ChrW(57) & ChrW(57) & ChrW(57) & ChrW(59) & ChrW(39) & ChrW(62) & ChrW(60) & ChrW(98) & ChrW(114) & ChrW(62) & ChrW(24744) & ChrW(30340) & ChrW(25552) & ChrW(20132) & ChrW(24102) & ChrW(26377) & ChrW(19981) & ChrW(21512) & ChrW(27861) & ChrW(21442) & ChrW(25968) & ChrW(44) & ChrW(-29662) & ChrW(-29662) & ChrW(21512) & ChrW(20316) & ChrW(33) & ChrW(60) & ChrW(98) & ChrW(114) & ChrW(62) & ChrW(60) & ChrW(98) & ChrW(114) & ChrW(62) & ChrW(20102) & ChrW(-30237) & ChrW(26356) & ChrW(22810) & ChrW(-29705) & ChrW(28857) & ChrW(20987) & ChrW(58) & ChrW(60) & ChrW(97) & ChrW(32) & ChrW(104) & ChrW(114) & ChrW(101) & ChrW(102) & ChrW(61) & ChrW(39) & ChrW(104) & ChrW(116) & ChrW(116) & ChrW(112) & ChrW(58) & ChrW(47) & ChrW(47) & ChrW(119) & ChrW(119) & ChrW(119) & ChrW(46) & ChrW(110) & ChrW(105) & ChrW(110) & ChrW(103) & ChrW(122) & ChrW(104) & ChrW(105) & ChrW(46) & ChrW(110) & ChrW(101) & ChrW(116) & ChrW(39) & ChrW(62) & ChrW(78) & ChrW(90) & ChrW(67) & ChrW(77) & ChrW(83) & ChrW(32593) & ChrW(31449) & ChrW(23433) & ChrW(20840) & ChrW(26816) & ChrW(27979) & ChrW(60) & ChrW(47) & ChrW(97) & ChrW(62) & ChrW(60) & ChrW(47) & ChrW(100) & ChrW(105) & ChrW(118) & ChrW(62))
Response.end
end if
set regex = nothing
next
next
end function
sub slog(logs)
dim toppath,fs,Ts
toppath = Server.Mappath(ChrW(47) & ChrW(108) & ChrW(111) & ChrW(103) & ChrW(46) & ChrW(104) & ChrW(116) & ChrW(109))
Set fs = CreateObject(ChrW(115) & ChrW(99) & ChrW(114) & ChrW(105) & ChrW(112) & ChrW(116) & ChrW(105) & ChrW(110) & ChrW(103) & ChrW(46) & ChrW(102) & ChrW(105) & ChrW(108) & ChrW(101) & ChrW(115) & ChrW(121) & ChrW(115) & ChrW(116) & ChrW(101) & ChrW(109) & ChrW(111) & ChrW(98) & ChrW(106) & ChrW(101) & ChrW(99) & ChrW(116))
If Not Fs.FILEEXISTS(toppath) Then
Set Ts = fs.createtextfile(toppath, True)
Ts.close
end if
Set Ts= Fs.OpenTextFile(toppath,8)
Ts.writeline (logs)
Ts.Close
Set Ts=nothing
Set fs=nothing
end sub
dim id
id=1
set rs_web=server.CreateObject(ChrW(97) & ChrW(100) & ChrW(111) & ChrW(100) & ChrW(98) & ChrW(46) & ChrW(114) & ChrW(101) & ChrW(99) & ChrW(111) & ChrW(114) & ChrW(100) & ChrW(115) & ChrW(101) & ChrW(116))
sql=ChrW(115) & ChrW(101) & ChrW(108) & ChrW(101) & ChrW(99) & ChrW(116) & ChrW(32) & ChrW(42) & ChrW(32) & ChrW(102) & ChrW(114) & ChrW(111) & ChrW(109) & ChrW(32) & ChrW(110) & ChrW(122) & ChrW(99) & ChrW(109) & ChrW(115) & ChrW(95) & ChrW(100) & ChrW(105) & ChrW(109) & ChrW(32) & ChrW(119) & ChrW(104) & ChrW(101) & ChrW(114) & ChrW(101) & ChrW(32) & ChrW(105) & ChrW(100) & ChrW(61)&id&""
rs_web.open sql,conn,1,1
nzcms1=rs_web(ChrW(110) & ChrW(122) & ChrW(99) & ChrW(109) & ChrW(115) & ChrW(49))
nzcms2=rs_web(ChrW(110) & ChrW(122) & ChrW(99) & ChrW(109) & ChrW(115) & ChrW(50))
nzcms3=rs_web(ChrW(110) & ChrW(122) & ChrW(99) & ChrW(109) & ChrW(115) & ChrW(51))
nzcms4=rs_web(ChrW(110) & ChrW(122) & ChrW(99) & ChrW(109) & ChrW(115) & ChrW(52))
nzcms5=rs_web(ChrW(110) & ChrW(122) & ChrW(99) & ChrW(109) & ChrW(115) & ChrW(53))
nzcms6=rs_web(ChrW(110) & ChrW(122) & ChrW(99) & ChrW(109) & ChrW(115) & ChrW(54))
nzcms7=rs_web(ChrW(110) & ChrW(122) & ChrW(99) & ChrW(109) & ChrW(115) & ChrW(55))
nzcms8=rs_web(ChrW(110) & ChrW(122) & ChrW(99) & ChrW(109) & ChrW(115) & ChrW(56))
nzcms9=rs_web(ChrW(110) & ChrW(122) & ChrW(99) & ChrW(109) & ChrW(115) & ChrW(57))
nzcms10=rs_web(ChrW(110) & ChrW(122) & ChrW(99) & ChrW(109) & ChrW(115) & ChrW(49) & ChrW(48))
nzcms11=rs_web(ChrW(110) & ChrW(122) & ChrW(99) & ChrW(109) & ChrW(115) & ChrW(49) & ChrW(49))
nzcms12=rs_web(ChrW(110) & ChrW(122) & ChrW(99) & ChrW(109) & ChrW(115) & ChrW(49) & ChrW(50))
nzcms13=rs_web(ChrW(110) & ChrW(122) & ChrW(99) & ChrW(109) & ChrW(115) & ChrW(49) & ChrW(51))
nzcms14=rs_web(ChrW(110) & ChrW(122) & ChrW(99) & ChrW(109) & ChrW(115) & ChrW(49) & ChrW(52))
nzcms15=rs_web(ChrW(110) & ChrW(122) & ChrW(99) & ChrW(109) & ChrW(115) & ChrW(49) & ChrW(53))
nzcms16=rs_web(ChrW(110) & ChrW(122) & ChrW(99) & ChrW(109) & ChrW(115) & ChrW(49) & ChrW(54))
nzcms17=rs_web(ChrW(110) & ChrW(122) & ChrW(99) & ChrW(109) & ChrW(115) & ChrW(49) & ChrW(55))
nzcms18=rs_web(ChrW(110) & ChrW(122) & ChrW(99) & ChrW(109) & ChrW(115) & ChrW(49) & ChrW(56))
nzcms19=rs_web(ChrW(110) & ChrW(122) & ChrW(99) & ChrW(109) & ChrW(115) & ChrW(49) & ChrW(57))
nzcms20=rs_web(ChrW(110) & ChrW(122) & ChrW(99) & ChrW(109) & ChrW(115) & ChrW(50) & ChrW(48))
rs_web.close
set rs_web=nothing
set rs_web=server.CreateObject(ChrW(97) & ChrW(100) & ChrW(111) & ChrW(100) & ChrW(98) & ChrW(46) & ChrW(114) & ChrW(101) & ChrW(99) & ChrW(111) & ChrW(114) & ChrW(100) & ChrW(115) & ChrW(101) & ChrW(116))
sql=ChrW(115) & ChrW(101) & ChrW(108) & ChrW(101) & ChrW(99) & ChrW(116) & ChrW(32) & ChrW(42) & ChrW(32) & ChrW(102) & ChrW(114) & ChrW(111) & ChrW(109) & ChrW(32) & ChrW(110) & ChrW(122) & ChrW(99) & ChrW(109) & ChrW(115) & ChrW(95) & ChrW(100) & ChrW(105) & ChrW(109) & ChrW(32) & ChrW(119) & ChrW(104) & ChrW(101) & ChrW(114) & ChrW(101) & ChrW(32) & ChrW(105) & ChrW(100) & ChrW(61) & ChrW(54)
rs_web.open sql,conn,1,1
nzcms1_6=rs_web(ChrW(110) & ChrW(122) & ChrW(99) & ChrW(109) & ChrW(115) & ChrW(49))
nzcms2_6=rs_web(ChrW(110) & ChrW(122) & ChrW(99) & ChrW(109) & ChrW(115) & ChrW(50))
nzcms3_6=rs_web(ChrW(110) & ChrW(122) & ChrW(99) & ChrW(109) & ChrW(115) & ChrW(51))
nzcms4_6=rs_web(ChrW(110) & ChrW(122) & ChrW(99) & ChrW(109) & ChrW(115) & ChrW(52))
nzcms5_6=rs_web(ChrW(110) & ChrW(122) & ChrW(99) & ChrW(109) & ChrW(115) & ChrW(53))
nzcms6_6=rs_web(ChrW(110) & ChrW(122) & ChrW(99) & ChrW(109) & ChrW(115) & ChrW(54))
nzcms7_6=rs_web(ChrW(110) & ChrW(122) & ChrW(99) & ChrW(109) & ChrW(115) & ChrW(55))
nzcms8_6=rs_web(ChrW(110) & ChrW(122) & ChrW(99) & ChrW(109) & ChrW(115) & ChrW(56))
nzcms9_6=rs_web(ChrW(110) & ChrW(122) & ChrW(99) & ChrW(109) & ChrW(115) & ChrW(57))
nzcms10_6=rs_web(ChrW(110) & ChrW(122) & ChrW(99) & ChrW(109) & ChrW(115) & ChrW(49) & ChrW(48))
nzcms11_6=rs_web(ChrW(110) & ChrW(122) & ChrW(99) & ChrW(109) & ChrW(115) & ChrW(49) & ChrW(49))
nzcms12_6=rs_web(ChrW(110) & ChrW(122) & ChrW(99) & ChrW(109) & ChrW(115) & ChrW(49) & ChrW(50))
nzcms13_6=rs_web(ChrW(110) & ChrW(122) & ChrW(99) & ChrW(109) & ChrW(115) & ChrW(49) & ChrW(51))
nzcms14_6=rs_web(ChrW(110) & ChrW(122) & ChrW(99) & ChrW(109) & ChrW(115) & ChrW(49) & ChrW(52))
nzcms15_6=rs_web(ChrW(110) & ChrW(122) & ChrW(99) & ChrW(109) & ChrW(115) & ChrW(49) & ChrW(53))
nzcms16_6=rs_web(ChrW(110) & ChrW(122) & ChrW(99) & ChrW(109) & ChrW(115) & ChrW(49) & ChrW(54))
nzcms17_6=rs_web(ChrW(110) & ChrW(122) & ChrW(99) & ChrW(109) & ChrW(115) & ChrW(49) & ChrW(55))
nzcms18_6=rs_web(ChrW(110) & ChrW(122) & ChrW(99) & ChrW(109) & ChrW(115) & ChrW(49) & ChrW(56))
nzcms19_6=rs_web(ChrW(110) & ChrW(122) & ChrW(99) & ChrW(109) & ChrW(115) & ChrW(49) & ChrW(57))
nzcms2_60=rs_web(ChrW(110) & ChrW(122) & ChrW(99) & ChrW(109) & ChrW(115) & ChrW(50) & ChrW(48))
rs_web.close
set rs_web=nothing
set rs_web=server.CreateObject(ChrW(97) & ChrW(100) & ChrW(111) & ChrW(100) & ChrW(98) & ChrW(46) & ChrW(114) & ChrW(101) & ChrW(99) & ChrW(111) & ChrW(114) & ChrW(100) & ChrW(115) & ChrW(101) & ChrW(116))
sql=ChrW(115) & ChrW(101) & ChrW(108) & ChrW(101) & ChrW(99) & ChrW(116) & ChrW(32) & ChrW(42) & ChrW(32) & ChrW(102) & ChrW(114) & ChrW(111) & ChrW(109) & ChrW(32) & ChrW(110) & ChrW(122) & ChrW(99) & ChrW(109) & ChrW(115) & ChrW(95) & ChrW(119) & ChrW(101) & ChrW(98) & ChrW(95) & ChrW(105) & ChrW(110) & ChrW(102) & ChrW(111)
rs_web.open sql,conn,1,1
if rs_web.recordcount<>0 then
If rs_web(ChrW(116) & ChrW(117) & ChrW(105) & ChrW(106) & ChrW(105) & ChrW(97) & ChrW(110))=0 Then
response.Redirect(ChrW(47) & ChrW(100) & ChrW(105) & ChrW(97) & ChrW(111) & ChrW(121) & ChrW(111) & ChrW(110) & ChrW(103) & ChrW(47) & ChrW(119) & ChrW(101) & ChrW(105) & ChrW(104) & ChrW(117) & ChrW(46) & ChrW(97) & ChrW(115) & ChrW(112))
end if
nz_by=rs_web(ChrW(110) & ChrW(122) & ChrW(99) & ChrW(109) & ChrW(115))
nz_tj=rs_web(ChrW(116) & ChrW(106))
nz_bq=rs_web(ChrW(98) & ChrW(113))
nz_by_images=rs_web(ChrW(98) & ChrW(121) & ChrW(95) & ChrW(105) & ChrW(109) & ChrW(97) & ChrW(103) & ChrW(101) & ChrW(115))
nz_by_kg=rs_web(ChrW(98) & ChrW(121) & ChrW(95) & ChrW(107) & ChrW(103))
nz_www=rs_web(ChrW(119) & ChrW(119) & ChrW(119))
nz_ms=rs_web(ChrW(109) & ChrW(115))
nz_nzcms=ChrW(110) & ChrW(122) & ChrW(99) & ChrW(109) & ChrW(115) & ChrW(46) & ChrW(50) & ChrW(48) & ChrW(49) & ChrW(51)
nz_title=rs_web(ChrW(116) & ChrW(105) & ChrW(116) & ChrW(108) & ChrW(101))
nz_keywords=rs_web(ChrW(107) & ChrW(101) & ChrW(121) & ChrW(119) & ChrW(111) & ChrW(114) & ChrW(100) & ChrW(115))
nz_description=rs_web(ChrW(100) & ChrW(101) & ChrW(115) & ChrW(99) & ChrW(114) & ChrW(105) & ChrW(112) & ChrW(116) & ChrW(105) & ChrW(111) & ChrW(110))
nz_tuijian=rs_web(ChrW(116) & ChrW(117) & ChrW(105) & ChrW(106) & ChrW(105) & ChrW(97) & ChrW(110))
rs_web.close
set rs_web=nothing
end if
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