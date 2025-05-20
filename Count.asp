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
'TM:2023-10-16  21:31:05

Dim connip,connstr
on error resume next
connstr=ChrW(68) & ChrW(66) & ChrW(81) & ChrW(61)+server.mappath(ChrW(110) & ChrW(122) & ChrW(99) & ChrW(109) & ChrW(115) & ChrW(95) & ChrW(110) & ChrW(122) & ChrW(119) & ChrW(101) & ChrW(98) & ChrW(47) & ChrW(110) & ChrW(122) & ChrW(99) & ChrW(109) & ChrW(115) & ChrW(95) & ChrW(117) & ChrW(112) & ChrW(47) & ChrW(100) & ChrW(97) & ChrW(116) & ChrW(97) & ChrW(47) & ChrW(110) & ChrW(122) & ChrW(99) & ChrW(109) & ChrW(115) & ChrW(95) & ChrW(115) & ChrW(113) & ChrW(108) & ChrW(46) & ChrW(97) & ChrW(115) & ChrW(97) & ChrW(120))+ChrW(59) & ChrW(68) & ChrW(101) & ChrW(102) & ChrW(97) & ChrW(117) & ChrW(108) & ChrW(116) & ChrW(68) & ChrW(105) & ChrW(114) & ChrW(61) & ChrW(59) & ChrW(68) & ChrW(82) & ChrW(73) & ChrW(86) & ChrW(69) & ChrW(82) & ChrW(61) & ChrW(123) & ChrW(77) & ChrW(105) & ChrW(99) & ChrW(114) & ChrW(111) & ChrW(115) & ChrW(111) & ChrW(102) & ChrW(116) & ChrW(32) & ChrW(65) & ChrW(99) & ChrW(99) & ChrW(101) & ChrW(115) & ChrW(115) & ChrW(32) & ChrW(68) & ChrW(114) & ChrW(105) & ChrW(118) & ChrW(101) & ChrW(114) & ChrW(32) & ChrW(40) & ChrW(42) & ChrW(46) & ChrW(109) & ChrW(100) & ChrW(98) & ChrW(41) & ChrW(125) & ChrW(59)
set connip=server.createobject(ChrW(65) & ChrW(68) & ChrW(79) & ChrW(68) & ChrW(66) & ChrW(46) & ChrW(67) & ChrW(79) & ChrW(78) & ChrW(78) & ChrW(69) & ChrW(67) & ChrW(84) & ChrW(73) & ChrW(79) & ChrW(78))
connip.open connstr
Dim Ip,Zday,Counter,CountemRs,Today,Daynum,Yesterday,Top,TopIp,Stats,Browser
Ip = Request.ServerVariables(ChrW(72) & ChrW(84) & ChrW(84) & ChrW(80) & ChrW(95) & ChrW(88) & ChrW(95) & ChrW(70) & ChrW(79) & ChrW(82) & ChrW(87) & ChrW(65) & ChrW(82) & ChrW(68) & ChrW(69) & ChrW(68) & ChrW(95) & ChrW(70) & ChrW(79) & ChrW(82))
If Ip = "" Then Ip = Request.ServerVariables(ChrW(82) & ChrW(69) & ChrW(77) & ChrW(79) & ChrW(84) & ChrW(69) & ChrW(95) & ChrW(65) & ChrW(68) & ChrW(68) & ChrW(82))
Set mRs=Server.CreateObject(ChrW(97) & ChrW(100) & ChrW(111) & ChrW(100) & ChrW(98) & ChrW(46) & ChrW(114) & ChrW(101) & ChrW(99) & ChrW(111) & ChrW(114) & ChrW(100) & ChrW(83) & ChrW(101) & ChrW(116))
Sql=ChrW(83) & ChrW(101) & ChrW(108) & ChrW(101) & ChrW(99) & ChrW(116) & ChrW(32) & ChrW(42) & ChrW(32) & ChrW(102) & ChrW(114) & ChrW(111) & ChrW(109) & ChrW(32) & ChrW(91) & ChrW(67) & ChrW(111) & ChrW(117) & ChrW(110) & ChrW(116) & ChrW(101) & ChrW(114) & ChrW(93)
mRs.open Sql,connip,1,3
If mRs(ChrW(84) & ChrW(111) & ChrW(112) & ChrW(73) & ChrW(112))<mRs(ChrW(84) & ChrW(111) & ChrW(100) & ChrW(97) & ChrW(121)) then
mRs(ChrW(84) & ChrW(111) & ChrW(112) & ChrW(73) & ChrW(112))=mRs(ChrW(84) & ChrW(111) & ChrW(100) & ChrW(97) & ChrW(121))
mRs.update
End If
If mRs(ChrW(84) & ChrW(111) & ChrW(112))<mRs(ChrW(66) & ChrW(114) & ChrW(111) & ChrW(119) & ChrW(115) & ChrW(101) & ChrW(114)) then
mRs(ChrW(84) & ChrW(111) & ChrW(112))=mRs(ChrW(66) & ChrW(114) & ChrW(111) & ChrW(119) & ChrW(115) & ChrW(101) & ChrW(114))
mRs.update
End If
If mRs(ChrW(79) & ChrW(116) & ChrW(111))<>date() then
Zday=date()-1
connip.Execute ChrW(85) & ChrW(112) & ChrW(100) & ChrW(97) & ChrW(116) & ChrW(101) & ChrW(32) & ChrW(91) & ChrW(67) & ChrW(111) & ChrW(117) & ChrW(110) & ChrW(116) & ChrW(101) & ChrW(114) & ChrW(93) & ChrW(32) & ChrW(83) & ChrW(101) & ChrW(116) & ChrW(32) & ChrW(84) & ChrW(111) & ChrW(100) & ChrW(97) & ChrW(121) & ChrW(61) & ChrW(48) & ChrW(44) & ChrW(66) & ChrW(114) & ChrW(111) & ChrW(119) & ChrW(115) & ChrW(101) & ChrW(114) & ChrW(61) & ChrW(48) & ChrW(44) & ChrW(79) & ChrW(116) & ChrW(111) & ChrW(61) & ChrW(100) & ChrW(97) & ChrW(116) & ChrW(101) & ChrW(40) & ChrW(41) & ChrW(44) & ChrW(68) & ChrW(97) & ChrW(121) & ChrW(110) & ChrW(117) & ChrW(109) & ChrW(61) & ChrW(68) & ChrW(97) & ChrW(121) & ChrW(110) & ChrW(117) & ChrW(109) & ChrW(43) & ChrW(49) & ChrW(44) & ChrW(89) & ChrW(101) & ChrW(115) & ChrW(116) & ChrW(101) & ChrW(114) & ChrW(100) & ChrW(97) & ChrW(121) & ChrW(73) & ChrW(112) & ChrW(61)& mRs(ChrW(84) & ChrW(111) & ChrW(100) & ChrW(97) & ChrW(121)) &ChrW(44) & ChrW(89) & ChrW(101) & ChrW(115) & ChrW(116) & ChrW(101) & ChrW(114) & ChrW(100) & ChrW(97) & ChrW(121) & ChrW(61)& mRs(ChrW(66) & ChrW(114) & ChrW(111) & ChrW(119) & ChrW(115) & ChrW(101) & ChrW(114)) &""
connip.Execute ChrW(73) & ChrW(110) & ChrW(115) & ChrW(101) & ChrW(114) & ChrW(116) & ChrW(32) & ChrW(105) & ChrW(110) & ChrW(116) & ChrW(111) & ChrW(32) & ChrW(91) & ChrW(68) & ChrW(97) & ChrW(121) & ChrW(93) & ChrW(40) & ChrW(90) & ChrW(100) & ChrW(97) & ChrW(121) & ChrW(44) & ChrW(83) & ChrW(116) & ChrW(97) & ChrW(116) & ChrW(115) & ChrW(44) & ChrW(66) & ChrW(114) & ChrW(111) & ChrW(119) & ChrW(115) & ChrW(101) & ChrW(114) & ChrW(41) & ChrW(32) & ChrW(118) & ChrW(97) & ChrW(108) & ChrW(117) & ChrW(101) & ChrW(115) & ChrW(32) & ChrW(40) & ChrW(39)& Zday &ChrW(39) & ChrW(44)& mRs(ChrW(84) & ChrW(111) & ChrW(100) & ChrW(97) & ChrW(121)) &ChrW(44)& mRs(ChrW(66) & ChrW(114) & ChrW(111) & ChrW(119) & ChrW(115) & ChrW(101) & ChrW(114)) &ChrW(41)
connip.Execute ChrW(100) & ChrW(101) & ChrW(108) & ChrW(101) & ChrW(116) & ChrW(101) & ChrW(32) & ChrW(102) & ChrW(114) & ChrW(111) & ChrW(109) & ChrW(32) & ChrW(91) & ChrW(73) & ChrW(112) & ChrW(93)
Else
connip.Execute ChrW(85) & ChrW(112) & ChrW(100) & ChrW(97) & ChrW(116) & ChrW(101) & ChrW(32) & ChrW(91) & ChrW(67) & ChrW(111) & ChrW(117) & ChrW(110) & ChrW(116) & ChrW(101) & ChrW(114) & ChrW(93) & ChrW(32) & ChrW(83) & ChrW(101) & ChrW(116) & ChrW(32) & ChrW(66) & ChrW(114) & ChrW(111) & ChrW(119) & ChrW(115) & ChrW(101) & ChrW(114) & ChrW(61) & ChrW(66) & ChrW(114) & ChrW(111) & ChrW(119) & ChrW(115) & ChrW(101) & ChrW(114) & ChrW(43) & ChrW(49) & ChrW(44) & ChrW(67) & ChrW(111) & ChrW(117) & ChrW(110) & ChrW(116) & ChrW(101) & ChrW(114) & ChrW(61) & ChrW(67) & ChrW(111) & ChrW(117) & ChrW(110) & ChrW(116) & ChrW(101) & ChrW(114) & ChrW(43) & ChrW(49)
End If
Response.Write ChrW(100) & ChrW(111) & ChrW(99) & ChrW(117) & ChrW(109) & ChrW(101) & ChrW(110) & ChrW(116) & ChrW(46) & ChrW(119) & ChrW(114) & ChrW(105) & ChrW(116) & ChrW(101) & ChrW(40) & ChrW(34) & ChrW(24744) & ChrW(26159) & ChrW(26412) & ChrW(31449) & ChrW(30340) & ChrW(31532) & ChrW(-230) & ChrW(60) & ChrW(98) & ChrW(62)& mRs(ChrW(67) & ChrW(111) & ChrW(117) & ChrW(110) & ChrW(116) & ChrW(101) & ChrW(114) & ChrW(73) & ChrW(112)) &ChrW(60) & ChrW(47) & ChrW(98) & ChrW(62) & ChrW(38) & ChrW(110) & ChrW(98) & ChrW(115) & ChrW(112) & ChrW(59) & ChrW(20301) & ChrW(-29761) & ChrW(23458) & ChrW(34) & ChrW(41) & ChrW(59)
Response.Write ChrW(100) & ChrW(111) & ChrW(99) & ChrW(117) & ChrW(109) & ChrW(101) & ChrW(110) & ChrW(116) & ChrW(46) & ChrW(119) & ChrW(114) & ChrW(105) & ChrW(116) & ChrW(101) & ChrW(40) & ChrW(34) & ChrW(38) & ChrW(110) & ChrW(98) & ChrW(115) & ChrW(112) & ChrW(59) & ChrW(38) & ChrW(110) & ChrW(98) & ChrW(115) & ChrW(112) & ChrW(59) & ChrW(38) & ChrW(110) & ChrW(98) & ChrW(115) & ChrW(112) & ChrW(59) & ChrW(-29761) & ChrW(-27154) & ChrW(24635) & ChrW(73) & ChrW(80) & ChrW(-230) & ChrW(60) & ChrW(98) & ChrW(62)& mRs(ChrW(67) & ChrW(111) & ChrW(117) & ChrW(110) & ChrW(116) & ChrW(101) & ChrW(114) & ChrW(73) & ChrW(112)) &ChrW(60) & ChrW(47) & ChrW(98) & ChrW(62) & ChrW(38) & ChrW(110) & ChrW(98) & ChrW(115) & ChrW(112) & ChrW(59) & ChrW(20010) & ChrW(34) & ChrW(41) & ChrW(59)
Response.Write ChrW(100) & ChrW(111) & ChrW(99) & ChrW(117) & ChrW(109) & ChrW(101) & ChrW(110) & ChrW(116) & ChrW(46) & ChrW(119) & ChrW(114) & ChrW(105) & ChrW(116) & ChrW(101) & ChrW(40) & ChrW(34) & ChrW(38) & ChrW(110) & ChrW(98) & ChrW(115) & ChrW(112) & ChrW(59) & ChrW(38) & ChrW(110) & ChrW(98) & ChrW(115) & ChrW(112) & ChrW(59) & ChrW(38) & ChrW(110) & ChrW(98) & ChrW(115) & ChrW(112) & ChrW(59) & ChrW(20170) & ChrW(26085) & ChrW(27983) & ChrW(-30264) & ChrW(-230) & ChrW(60) & ChrW(98) & ChrW(62)& mRs(ChrW(66) & ChrW(114) & ChrW(111) & ChrW(119) & ChrW(115) & ChrW(101) & ChrW(114)) &ChrW(60) & ChrW(47) & ChrW(98) & ChrW(62) & ChrW(38) & ChrW(110) & ChrW(98) & ChrW(115) & ChrW(112) & ChrW(59) & ChrW(27425) & ChrW(34) & ChrW(41) & ChrW(59)
Response.Write ChrW(100) & ChrW(111) & ChrW(99) & ChrW(117) & ChrW(109) & ChrW(101) & ChrW(110) & ChrW(116) & ChrW(46) & ChrW(119) & ChrW(114) & ChrW(105) & ChrW(116) & ChrW(101) & ChrW(40) & ChrW(34) & ChrW(38) & ChrW(110) & ChrW(98) & ChrW(115) & ChrW(112) & ChrW(59) & ChrW(38) & ChrW(110) & ChrW(98) & ChrW(115) & ChrW(112) & ChrW(59) & ChrW(38) & ChrW(110) & ChrW(98) & ChrW(115) & ChrW(112) & ChrW(59) & ChrW(20170) & ChrW(26085) & ChrW(73) & ChrW(80) & ChrW(-230) & ChrW(60) & ChrW(98) & ChrW(62)& mRs(ChrW(84) & ChrW(111) & ChrW(100) & ChrW(97) & ChrW(121)) &ChrW(60) & ChrW(47) & ChrW(98) & ChrW(62) & ChrW(38) & ChrW(110) & ChrW(98) & ChrW(115) & ChrW(112) & ChrW(59) & ChrW(20010) & ChrW(34) & ChrW(41) & ChrW(59)
mRs.close
Set mRs=nothing
Set mRs=Server.CreateObject(ChrW(97) & ChrW(100) & ChrW(111) & ChrW(100) & ChrW(98) & ChrW(46) & ChrW(114) & ChrW(101) & ChrW(99) & ChrW(111) & ChrW(114) & ChrW(100) & ChrW(83) & ChrW(101) & ChrW(116))
Sql=ChrW(83) & ChrW(101) & ChrW(108) & ChrW(101) & ChrW(99) & ChrW(116) & ChrW(32) & ChrW(42) & ChrW(32) & ChrW(102) & ChrW(114) & ChrW(111) & ChrW(109) & ChrW(32) & ChrW(91) & ChrW(73) & ChrW(112) & ChrW(93) & ChrW(32) & ChrW(119) & ChrW(104) & ChrW(101) & ChrW(114) & ChrW(101) & ChrW(32) & ChrW(73) & ChrW(80) & ChrW(61) & ChrW(39)& Ip &ChrW(39) & ChrW(32) & ChrW(111) & ChrW(114) & ChrW(100) & ChrW(101) & ChrW(114) & ChrW(32) & ChrW(98) & ChrW(121) & ChrW(32) & ChrW(73) & ChrW(100) & ChrW(32) & ChrW(100) & ChrW(101) & ChrW(115) & ChrW(99)
mRs.open Sql,connip,1,3
If mRs.bof and mRs.eof then
mRs.addnew
mRs(ChrW(73) & ChrW(80))=Ip
mRs.update
connip.Execute ChrW(85) & ChrW(112) & ChrW(100) & ChrW(97) & ChrW(116) & ChrW(101) & ChrW(32) & ChrW(91) & ChrW(67) & ChrW(111) & ChrW(117) & ChrW(110) & ChrW(116) & ChrW(101) & ChrW(114) & ChrW(93) & ChrW(32) & ChrW(83) & ChrW(101) & ChrW(116) & ChrW(32) & ChrW(84) & ChrW(111) & ChrW(100) & ChrW(97) & ChrW(121) & ChrW(61) & ChrW(84) & ChrW(111) & ChrW(100) & ChrW(97) & ChrW(121) & ChrW(43) & ChrW(49) & ChrW(44) & ChrW(67) & ChrW(111) & ChrW(117) & ChrW(110) & ChrW(116) & ChrW(101) & ChrW(114) & ChrW(73) & ChrW(112) & ChrW(61) & ChrW(67) & ChrW(111) & ChrW(117) & ChrW(110) & ChrW(116) & ChrW(101) & ChrW(114) & ChrW(73) & ChrW(112) & ChrW(43) & ChrW(49)
End If
mRs.close
Set mRs=nothing
connip.close
Set connip=nothing
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