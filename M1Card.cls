VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "M1Card"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim akey(6) As Byte
Dim bkey(6) As Byte
Dim hexkey As String * 12
Dim displaydata(8) As Byte
Dim eepromdata As String * 32
Dim hardver As String * 20
Dim libver As String * 16
Dim ser As String * 16
Dim Snr As Long
Dim Rsnr As Long
Dim data16 As String * 16
Dim data32 As String * 32
Dim databuff32 As String * 32
Dim databuff16 As String * 16
Dim rvalue As Long
Dim wvalue As Long
Dim mlrvalue As Integer
Dim mlwvalue As Integer
Dim cardmode As Integer
Dim loadmode As Integer
Dim sector As Integer
Dim address As Integer
Dim ptrdest As String * 16
Dim ptrsource As String * 16
Dim time As String * 14
Dim timebuff As String * 14
Dim tagtype As Integer
Dim Size As Integer
Const port As Integer = 0 '�˿�
Const baud As Long = 9600 '115200 'ͨѶ������9600��115200
Const dataGroup As String = "000000000000000000000000000" '00000"




Public Function PutCard(ByVal strCardNo As String) As Integer '����
    Dim strCardRead As String
    
    If Len(strCardNo) <> 5 Then
        PutCard = 20 '������Ϊ5λ
        Exit Function
    End If
    icdev = rf_init(port, baud)
    If icdev < 0 Then
        PutCard = 1 '�豸��ʼ���˿�ʧ��
        Exit Function
    End If
        
    'Ѱ��
    st = rf_reset(icdev, 10) '��Ƶ��дģ�鸴λ
    
    st = rf_request(icdev, 1, tagtype)
    If st <> 0 Then
    PutCard = 11 'Ѱ������ʧ��
    Call quit
    Exit Function
    End If

    st = rf_anticoll(icdev, 0, Snr)
    If st <> 0 Then
        PutCard = 12 '������ͻ�����ؿ������к�
        Call quit
        Exit Function
    End If
            
    st = rf_select(icdev, Snr, Size)
    If st <> 0 Then
        PutCard = 13 '�Ӷ������ѡȡһ���������кŵĿ�
        Call quit
        Exit Function
    End If
    
        'װ��B����
    bkey(0) = &HB0
    bkey(1) = &H11
    bkey(2) = &H4C
    bkey(3) = &HE9
    bkey(4) = &HA3
    bkey(5) = &HD7
    
    
    loadmode = 4
    sector = 1
        
   st = rf_load_key(ByVal icdev, loadmode, sector, bkey(0))
    If st <> 0 Then
        PutCard = 3 'װ������B����
        Call quit
        Exit Function
    End If

    '��֤����B
    
    st = rf_authentication(ByVal icdev, loadmode, sector)
    If st <> 0 Then
        PutCard = 5 '�ÿ�Ƭ�����ڱ�ϵͳ��ʹ��
        Call quit
        Exit Function
    End If
    
    
    'װ������A
    akey(0) = &HA3 '0xa3
    akey(1) = &HD4 '0xd4
    akey(2) = &H56 '0x56
    akey(3) = &H8C '0x8c
    akey(4) = &HB9 '0xb9
    akey(5) = &HE5 '0xe5
    
    loadmode = 0
    sector = 1
       
    st = rf_load_key(ByVal icdev, loadmode, sector, akey(0))
    If st <> 0 Then
        PutCard = 2 'װ������A����
        Call quit
        Exit Function
     End If

    '��֤����A
    st = rf_authentication(ByVal icdev, loadmode, sector)
    If st <> 0 Then
        PutCard = 6 'A������֤����
        Call quit
        Exit Function
    End If
    


    '��ȡ�����
    'cardmode = 1
    
    'st = rf_card(ByVal icdev, cardmode, Snr)
    'If st <> 0 Then
    '    PutCard = 4 '��ȡ��Ա����ų���
    '    Call quit
    '    Exit Function
    'End If

    'д���ݣ�����
    address = 4
    'databuff32 = Mid(dataGroup, 1, 31 - Len(strCardNo)) & "9" & strCardNo
    data32 = dataGroup & strCardNo
    st = rf_write_hex(ByVal icdev, address, ByVal data32) 'д��1������0����,���Ŀ�
    If st <> 0 Then
        PutCard = 7 '������ʼ��ʧ��"
        Call quit
        Exit Function
    End If

    '��֤д����
    st = rf_check_writehex(ByVal icdev, Snr, loadmode, address, data32)
    If st <> 0 Then
        PutCard = 8 '������ʼ��ʧ��"
        Call quit
        Exit Function
    End If

    '�����ݣ���֤
    address = 4
    st = rf_read_hex(ByVal icdev, address, databuff32)
    If st <> 0 Then
        PutCard = 9 '������ʧ��
        Call quit
        Exit Function
    End If

    'If InStr(data32, databuff32) <> 0 Then
    If data32 <> databuff32 Then
        PutCard = 10 '��������ʧ�ܣ��ÿ��ݲ�Ҫʹ��"
        Call quit
        Exit Function
    End If
    
    st = rf_beep(icdev, 5)

    PutCard = 0
    'ȡ���豸
    Call quit
        
       
End Function

Public Function ReadCard(strCardNo As String) As Integer '����
     '��ʼ���˿�
    icdev = rf_init(port, baud)

    If icdev < 0 Then
        ReadCard = 1 '�豸��ʼ���˿�ʧ��
        Exit Function
    End If

     'Ѱ��
    st = rf_reset(icdev, 10) '��Ƶ��дģ�鸴λ
    
    st = rf_request(icdev, 1, tagtype)
    If st <> 0 Then
        ReadCard = 11 'Ѱ������ʧ��
        Call quit
        Exit Function
    End If

    st = rf_anticoll(icdev, 0, Snr)
    If st <> 0 Then
        ReadCard = 12 '������ͻ�����ؿ������к�
        Call quit
        Exit Function
    End If
            
    st = rf_select(icdev, Snr, Size)
    If st <> 0 Then
        ReadCard = 13 '�Ӷ������ѡȡһ���������кŵĿ�
        Call quit
        Exit Function
    End If
    
    'װ��B����
    bkey(0) = &HB0
    bkey(1) = &H11
    bkey(2) = &H4C
    bkey(3) = &HE9
    bkey(4) = &HA3
    bkey(5) = &HD7
    loadmode = 4
    sector = 1
    st = rf_load_key(ByVal icdev, loadmode, sector, bkey(0))
    If st <> 0 Then
          ReadCard = 2 'װ���������_B
          Call quit
          Exit Function
    End If
    
    '��֤����B
    st = rf_authentication(ByVal icdev, loadmode, sector)
    If st <> 0 Then
          ReadCard = 4 '�ÿ�Ƭ�����ڱ�ϵͳ��ʹ��
          Call quit
          Exit Function
    End If

    '��ȡ�����
    'cardmode = 1
    'st = rf_card(ByVal icdev, cardmode, Snr)
    'If st <> 0 Then
    '      ReadCard = 3 '��ȡ��Ա����ų���
    '      Call quit
    '      Exit Function
    'End If

    

    'װ������A
    akey(0) = &HA3
    akey(1) = &HD4
    akey(2) = &H56
    akey(3) = &H8C
    akey(4) = &HB9
    akey(5) = &HE5
    loadmode = 0
    sector = 1
    st = rf_load_key(ByVal icdev, loadmode, sector, akey(0))
    If st <> 0 Then
          ReadCard = 5 'װ������A����
          Call quit
          Exit Function
    End If

    '��֤����A
    st = rf_authentication(ByVal icdev, loadmode, sector)
    If st <> 0 Then
          ReadCard = 6 '����A����ȷ
          Call quit
          Exit Function
    End If

    address = 4
    st = rf_read_hex(ByVal icdev, address, databuff32)
    If st <> 0 Then
         ReadCard = 7 '��ȡ���Ŵ���
         Call quit
         Exit Function
    End If
    'strCardNo = Trim(Str$(Val(databuff32)))
    'strCardNo = Mid(strCardNo, 2, Len(strCardNo))
    strCardNo = Mid(databuff32, 28)
    If strCardNo = "" Then
        ReadCard = 8 '����������Ϊ��
        Call quit
        Exit Function
    End If

   

    st = rf_beep(icdev, 5)
    
    ReadCard = 0
    'ȡ���豸
    Call quit
    
End Function



