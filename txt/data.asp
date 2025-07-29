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
'TM:2023-10-16  21:34:45

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