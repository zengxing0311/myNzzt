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

Set conn=Server.CreateObject(ChrW(65) & ChrW(68) & ChrW(79) & ChrW(68) & ChrW(66) & ChrW(46) & ChrW(67) & ChrW(111) & ChrW(110) & ChrW(110) & ChrW(101) & ChrW(99) & ChrW(116) & ChrW(105) & ChrW(111) & ChrW(110))
conn.open ChrW(112) & ChrW(114) & ChrW(111) & ChrW(118) & ChrW(105) & ChrW(100) & ChrW(101) & ChrW(114) & ChrW(61) & ChrW(115) & ChrW(113) & ChrW(108) & ChrW(111) & ChrW(108) & ChrW(101) & ChrW(100) & ChrW(98) & ChrW(59) & ChrW(100) & ChrW(97) & ChrW(116) & ChrW(97) & ChrW(32) & ChrW(115) & ChrW(111) & ChrW(117) & ChrW(114) & ChrW(99) & ChrW(101) & ChrW(61) & ChrW(49) & ChrW(50) & ChrW(55) & ChrW(46) & ChrW(48) & ChrW(46) & ChrW(48) & ChrW(46) & ChrW(49) & ChrW(59) & ChrW(85) & ChrW(115) & ChrW(101) & ChrW(114) & ChrW(32) & ChrW(73) & ChrW(68) & ChrW(61) & ChrW(115) & ChrW(97) & ChrW(59) & ChrW(112) & ChrW(119) & ChrW(100) & ChrW(61) & ChrW(115) & ChrW(97) & ChrW(59) & ChrW(73) & ChrW(110) & ChrW(105) & ChrW(116) & ChrW(105) & ChrW(97) & ChrW(108) & ChrW(32) & ChrW(67) & ChrW(97) & ChrW(116) & ChrW(97) & ChrW(108) & ChrW(111) & ChrW(103) & ChrW(61) & ChrW(122) & ChrW(116) & ChrW(108) & ChrW(104)
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