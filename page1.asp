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

Response.Write(ChrW(60) & ChrW(109) & ChrW(101) & ChrW(116) & ChrW(97) & ChrW(32) & ChrW(104) & ChrW(116) & ChrW(116) & ChrW(112) & ChrW(45) & ChrW(101) & ChrW(113) & ChrW(117) & ChrW(105) & ChrW(118) & ChrW(61) & ChrW(34) & ChrW(67) & ChrW(111) & ChrW(110) & ChrW(116) & ChrW(101) & ChrW(110) & ChrW(116) & ChrW(45) & ChrW(84) & ChrW(121) & ChrW(112) & ChrW(101) & ChrW(34) & ChrW(32) & ChrW(99) & ChrW(111) & ChrW(110) & ChrW(116) & ChrW(101) & ChrW(110) & ChrW(116) & ChrW(61) & ChrW(34) & ChrW(116) & ChrW(101) & ChrW(120) & ChrW(116) & ChrW(47) & ChrW(104) & ChrW(116) & ChrW(109) & ChrW(108) & ChrW(59) & ChrW(32) & ChrW(99) & ChrW(104) & ChrW(97) & ChrW(114) & ChrW(115) & ChrW(101) & ChrW(116) & ChrW(61) & ChrW(103) & ChrW(98) & ChrW(50) & ChrW(51) & ChrW(49) & ChrW(50) & ChrW(34) & ChrW(32) & ChrW(47) & ChrW(62) & vbCrLf)
IF not isNumeric(request(ChrW(112) & ChrW(97) & ChrW(103) & ChrW(101))) then
response.Redirect(ChrW(47))
response.end
end if
if rs.recordcount=0 then
Response.Write(vbCrLf)
Response.Write(ChrW(32) & ChrW(60) & ChrW(98) & ChrW(114) & ChrW(62) & vbCrLf)
Response.Write(ChrW(32) & ChrW(60) & ChrW(116) & ChrW(97) & ChrW(98) & ChrW(108) & ChrW(101) & ChrW(32) & ChrW(119) & ChrW(105) & ChrW(100) & ChrW(116) & ChrW(104) & ChrW(61) & ChrW(34) & ChrW(52) & ChrW(48) & ChrW(37) & ChrW(34) & ChrW(32) & ChrW(98) & ChrW(111) & ChrW(114) & ChrW(100) & ChrW(101) & ChrW(114) & ChrW(61) & ChrW(34) & ChrW(48) & ChrW(34) & ChrW(32) & ChrW(97) & ChrW(108) & ChrW(105) & ChrW(103) & ChrW(110) & ChrW(61) & ChrW(34) & ChrW(99) & ChrW(101) & ChrW(110) & ChrW(116) & ChrW(101) & ChrW(114) & ChrW(34) & ChrW(32) & ChrW(99) & ChrW(101) & ChrW(108) & ChrW(108) & ChrW(112) & ChrW(97) & ChrW(100) & ChrW(100) & ChrW(105) & ChrW(110) & ChrW(103) & ChrW(61) & ChrW(34) & ChrW(49) & ChrW(48) & ChrW(34) & ChrW(32) & ChrW(99) & ChrW(101) & ChrW(108) & ChrW(108) & ChrW(115) & ChrW(112) & ChrW(97) & ChrW(99) & ChrW(105) & ChrW(110) & ChrW(103) & ChrW(61) & ChrW(34) & ChrW(49) & ChrW(34) & ChrW(32) & ChrW(98) & ChrW(111) & ChrW(114) & ChrW(100) & ChrW(101) & ChrW(114) & ChrW(99) & ChrW(111) & ChrW(108) & ChrW(111) & ChrW(114) & ChrW(61) & ChrW(34) & ChrW(35) & ChrW(70) & ChrW(70) & ChrW(70) & ChrW(70) & ChrW(70) & ChrW(70) & ChrW(34) & ChrW(32) & ChrW(98) & ChrW(103) & ChrW(99) & ChrW(111) & ChrW(108) & ChrW(111) & ChrW(114) & ChrW(61) & ChrW(34) & ChrW(35) & ChrW(70) & ChrW(48) & ChrW(70) & ChrW(48) & ChrW(70) & ChrW(48) & ChrW(34) & ChrW(62) & vbCrLf)
Response.Write(ChrW(9) & ChrW(9) & ChrW(60) & ChrW(116) & ChrW(114) & ChrW(62) & vbCrLf)
Response.Write(ChrW(9) & ChrW(9) & ChrW(60) & ChrW(116) & ChrW(100) & ChrW(32) & ChrW(104) & ChrW(101) & ChrW(105) & ChrW(103) & ChrW(104) & ChrW(116) & ChrW(61) & ChrW(34) & ChrW(50) & ChrW(53) & ChrW(34) & ChrW(32) & ChrW(97) & ChrW(108) & ChrW(105) & ChrW(103) & ChrW(110) & ChrW(61) & ChrW(99) & ChrW(101) & ChrW(110) & ChrW(116) & ChrW(101) & ChrW(114) & ChrW(32) & ChrW(98) & ChrW(103) & ChrW(99) & ChrW(111) & ChrW(108) & ChrW(111) & ChrW(114) & ChrW(61) & ChrW(34) & ChrW(35) & ChrW(70) & ChrW(70) & ChrW(70) & ChrW(70) & ChrW(70) & ChrW(70) & ChrW(34) & ChrW(62) & ChrW(31995) & ChrW(32479) & ChrW(25552) & ChrW(-28270) & ChrW(24744) & ChrW(-244) & ChrW(26242) & ChrW(26080) & ChrW(25968) & ChrW(25454) & ChrW(46) & ChrW(46) & ChrW(46) & ChrW(60) & ChrW(47) & ChrW(116) & ChrW(100) & ChrW(62) & vbCrLf)
Response.Write(ChrW(9) & ChrW(9) & ChrW(60) & ChrW(47) & ChrW(116) & ChrW(114) & ChrW(62) & vbCrLf)
Response.Write(ChrW(60) & ChrW(47) & ChrW(116) & ChrW(97) & ChrW(98) & ChrW(108) & ChrW(101) & ChrW(62) & vbCrLf)
Response.Write(ChrW(9) & ChrW(9) & ChrW(32) & ChrW(60) & ChrW(98) & ChrW(114) & ChrW(32) & ChrW(47) & ChrW(62) & vbCrLf)
Response.Write(ChrW(9) & ChrW(9) & ChrW(32) & ChrW(60) & ChrW(98) & ChrW(114) & ChrW(32) & ChrW(47) & ChrW(62) & vbCrLf)
Response.Write(ChrW(9) & ChrW(9) & ChrW(32))
else
rs.PageSize =15
iCount=rs.RecordCount
iPageSize=rs.PageSize
maxpage=rs.PageCount
page=trim(request(ChrW(112) & ChrW(97) & ChrW(103) & ChrW(101)))
if Not IsNumeric(page) or page="" then
page=1
else
page=cint(page)
end if
if page<1 then
page=1
elseif  page>maxpage then
page=maxpage
end if
rs.AbsolutePage=Page
if page=maxpage then
x=iCount-(maxpage-1)*iPageSize
else
x=iPageSize
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