<%
'�˰汾Ϊ������صģ����ð汾��Դ����ܴ����
'������ʽ��ϵͳ�ṩԴ�루֧�ֶ��ο������޸ģ�
'ϵͳ�����̣�������־����Ƽ����޹�˾
'��˾��ַ��http://www.ningzhi.net 
'�ͷ�QQ�ţ�122470827 �绰��18867186998 ��ͬ΢�źţ�������
'ȫ��ͳһ�ۺ�绰��0579-83938663 
'֧�����棺�������Ȩ�ǼǺţ�2012SR040914
'�ٷ�ƽ̨������ϵͳ���Ͻ�����־���繫˾����ƽ̨��Ψһ��������
'����˾�޵�����������ƽ̨����׼,�����ϵ���ƭ������
'�����޸������κδ���,�Ա�֤�����������С���л��������־��Ʒ����־���������ʣ�
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