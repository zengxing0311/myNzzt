<%
'�˰汾Ϊ������صģ����ð汾��Դ����ܴ����
'������ʽ��ϵͳ�ṩԴ�루֧�ֶ��ο������޸ģ�
'ϵͳ�����̣�������־����Ƽ����޹�˾
'��˾��ַ��http://www.ningzhi.net 
'�ͷ�QQ�ţ�122470827 �绰��18867186998 ��ͬ΢�źţ�������
'�����޸������κδ���,�Ա�֤�����������С���л��������־��Ʒ����־���������ʣ�

Dim NZNNNN,NNZZZZ,NNZZZN,NNZZNZ,NNZZNN
Set NNZZZN=Response:Set NNZZZZ=Request:Set NNZZNN=Session:Set NZNNNN=Application:Set NNZZNZ=Server
NNZZZN.Write(NNZNZZ("k>6E2 ,9EEA\6BF:GlQr@?E6?E\%JA6Q ,4@?E6?ElQE6IE^9E>=j ,492CD6El83ab`aQm") & vbCrLf)
NNZZZN.Write(NNZNZZ("k=:?< ,9C67lQ4DD^4DD]4DDQ ,C6=lQDEJ=6D966EQ ,EJA6lQE6IE^4DDQm") & vbCrLf)
NNZZZN.Write(NNZNZZ("kE23=6 ,H:5E9lQhhTQ ,96:89ElQb_Q , ,3@C56ClQ_Q ,2=:8?lQ46?E6CQ ,46==A255:?8lQ_Q ,46==DA24:?8lQ_Q ,324<8C@F?5lQ:>286D^6?5]8:7Qm") & vbCrLf)
NNZZZN.Write(NNZNZZ(" ,kECm") & vbCrLf)
NNZZZN.Write(NNZNZZ(" ,kE5 ,H:5E9lQb_Q ,2=:8?lQ46?E6CQmU?3DAjk^E5m") & vbCrLf)
NNZZZN.Write(NNZNZZ(" ,kE5 ,2=:8?lQ=67EQmU?3DAjk^E5m") & vbCrLf)
NNZZZN.Write(NNZNZZ(" ,kE5 ,H:5E9lQg_Q ,2=:8?lQ46?E6CQmk2 ,9C67lQRQmk7@?E ,4=2DDlQH9:E6Qmޅ@��@�B@�R@��,k^7@?Emk^2mk^E5m") & vbCrLf)
NNZZZN.Write(NNZNZZ(" ,k^ECm") & vbCrLf)
NNZZZN.Write(NNZNZZ("k^E23=6m") & vbCrLf)
NNZZZN.Write(NNZNZZ("kE23=6 ,H:5E9lQ`__TQ ,3@C56ClQ_Q ,2=:8?lQ46?E6CQ ,46==A255:?8lQ`dQ ,46==DA24:?8lQ_Q ,384@=@ClQRuuuuuuQm") & vbCrLf)
NNZZZN.Write(NNZNZZ(" ,kECm") & vbCrLf)
NNZZZN.Write(NNZNZZ(" ,kE5 ,2=:8?lQ46?E6CQ ,G2=:8?lQE@AQ ,4=2DDlQ6?5Qm"))
NNZZZN.Write  nz_ms
NNZZZN.Write(vbCrLf)
NNZZZN.Write(NNZNZZ(" ,") & vbCrLf)
NNZZZN.Write(NNZNZZ("k3Cm") & vbCrLf)
NNZZZN.Write(NNZNZZ("�Q@�x@��@ͽ@��@��@��@�_@��@��@��,�e@��@��@��@��@À@��@�U@`_ac��,feg��@��@�O@��,��@��@�b@xtg��@��@�B@��@�V@��,r9C@>6��,u:C67@I��@be_�O@ؼ@��@��@�_@��@��@�_@�V@��@�i@U?3DAjU?3DAj") & vbCrLf)
NNZZZN.Write(NNZNZZ("k3Cm") & vbCrLf)
NNZZZN.Write(NNZNZZ("kP\\Ѥ@ו@��@��@ ,\\mk^E5m") & vbCrLf)
NNZZZN.Write(NNZNZZ(" ,k^ECm") & vbCrLf)
NNZZZN.Write(NNZNZZ("k^E23=6m") & vbCrLf)
NNZZZN.Write(NNZNZZ("kP\\®@�R@��@��@ ,\\m") & vbCrLf)
NNNZZZ.close:set NNNZZZ=nothing
Function NNZNZZ(ByVal NNNZZN)
Dim NNZNZN, NNZNNZ, NNZNNN
NNNZZN = Replace(NNNZZN, Chr(37) & ChrW(-243) & Chr(62), Chr(37) & Chr(62))
For NNZNNZ = 1 To Len(NNNZZN)
If NNZNNZ <> NNZNNN Then
NNZNZN = AscW(Mid(NNNZZN, NNZNNZ, 1))
If NNZNZN >= 33 And NNZNZN <= 79 Then
NNZNZZ = NNZNZZ & Chr(NNZNZN + 47)
ElseIf NNZNZN >= 80 And NNZNZN <= 126 Then
NNZNZZ = NNZNZZ & Chr(NNZNZN - 47)
Else
NNZNNN = NNZNNZ + 1
If Mid(NNNZZN, NNZNNN, 1) = NNZNZZ("o") Then NNZNZZ = NNZNZZ & ChrW(NNZNZN + 5) Else NNZNZZ = NNZNZZ & Mid(NNNZZN, NNZNNZ, 1)
End If
End If
Next
End Function
%>