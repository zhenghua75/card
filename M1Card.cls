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
Const port As Integer = 0 '端口
Const baud As Long = 9600 '115200 '通讯波特率9600～115200
Const dataGroup As String = "000000000000000000000000000" '00000"




Public Function PutCard(ByVal strCardNo As String) As Integer '发卡
    Dim strCardRead As String

    If Len(strCardNo) <> 5 Then
        PutCard = 20 '卡号需为5位
        Exit Function
    End If
    icdev = rf_init(port, baud)
    If icdev < 0 Then
        PutCard = 1 '设备初始化端口失败
        Exit Function
    End If

    '寻卡
    st = rf_reset(icdev, 10) '射频读写模块复位

    st = rf_request(icdev, 1, tagtype)
    If st <> 0 Then
    PutCard = 11 '寻卡请求失败
    Call quit
    Exit Function
    End If

    st = rf_anticoll(icdev, 0, Snr)
    If st <> 0 Then
        PutCard = 12 '卡防冲突，返回卡的序列号
        Call quit
        Exit Function
    End If

    st = rf_select(icdev, Snr, Size)
    If st <> 0 Then
        PutCard = 13 '从多个卡中选取一个给定序列号的卡
        Call quit
        Exit Function
    End If

        '装载B密码
    bkey(0) = &H01
    bkey(1) = &H02
    bkey(2) = &H03
    bkey(3) = &H04
    bkey(4) = &H05
    bkey(5) = &H06


    loadmode = 4
    sector = 1

   st = rf_load_key(ByVal icdev, loadmode, sector, bkey(0))
    If st <> 0 Then
        PutCard = 3 '装载密码B出错
        Call quit
        Exit Function
    End If

    '验证密码B

    st = rf_authentication(ByVal icdev, loadmode, sector)
    If st <> 0 Then
        PutCard = 5 '该卡片不属于本系统所使用
        Call quit
        Exit Function
    End If


    '装载密码A
    akey(0) = &H01
    akey(1) = &H02
    akey(2) = &H03
    akey(3) = &H04
    akey(4) = &H05
    akey(5) = &H06

    loadmode = 0
    sector = 1

    st = rf_load_key(ByVal icdev, loadmode, sector, akey(0))
    If st <> 0 Then
        PutCard = 2 '装载密码A出错
        Call quit
        Exit Function
     End If

    '验证密码A
    st = rf_authentication(ByVal icdev, loadmode, sector)
    If st <> 0 Then
        PutCard = 6 'A密码验证错误
        Call quit
        Exit Function
    End If



    '获取卡序号
    'cardmode = 1

    'st = rf_card(ByVal icdev, cardmode, Snr)
    'If st <> 0 Then
    '    PutCard = 4 '获取会员卡序号出错
    '    Call quit
    '    Exit Function
    'End If

    '写数据，卡号
    address = 4
    'databuff32 = Mid(dataGroup, 1, 31 - Len(strCardNo)) & "9" & strCardNo
    data32 = dataGroup & strCardNo
    st = rf_write_hex(ByVal icdev, address, ByVal data32) '写在1扇区的0块上,第四块
    If st <> 0 Then
        PutCard = 7 '发卡初始化失败"
        Call quit
        Exit Function
    End If

    '验证写数据
    st = rf_check_writehex(ByVal icdev, Snr, loadmode, address, data32)
    If st <> 0 Then
        PutCard = 8 '发卡初始化失败"
        Call quit
        Exit Function
    End If

    '读数据，验证
    address = 4
    st = rf_read_hex(ByVal icdev, address, databuff32)
    If st <> 0 Then
        PutCard = 9 '读数据失败
        Call quit
        Exit Function
    End If

    'If InStr(data32, databuff32) <> 0 Then
    If data32 <> databuff32 Then
        PutCard = 10 '发卡操作失败，该卡暂不要使用"
        Call quit
        Exit Function
    End If

    st = rf_beep(icdev, 5)

    PutCard = 0
    '取消设备
    Call quit


End Function

Public Function ReadCard(strCardNo As String) As Integer '读卡
     '初始化端口
    icdev = rf_init(port, baud)

    If icdev < 0 Then
        ReadCard = 1 '设备初始化端口失败
        Exit Function
    End If

     '寻卡
    st = rf_reset(icdev, 10) '射频读写模块复位

    st = rf_request(icdev, 1, tagtype)
    If st <> 0 Then
        ReadCard = 11 '寻卡请求失败
        Call quit
        Exit Function
    End If

    st = rf_anticoll(icdev, 0, Snr)
    If st <> 0 Then
        ReadCard = 12 '卡防冲突，返回卡的序列号
        Call quit
        Exit Function
    End If

    st = rf_select(icdev, Snr, Size)
    If st <> 0 Then
        ReadCard = 13 '从多个卡中选取一个给定序列号的卡
        Call quit
        Exit Function
    End If

    '装载B密码
    bkey(0) = &H01
    bkey(1) = &H02
    bkey(2) = &H03
    bkey(3) = &H04
    bkey(4) = &H05
    bkey(5) = &H06
    loadmode = 4
    sector = 1
    st = rf_load_key(ByVal icdev, loadmode, sector, bkey(0))
    If st <> 0 Then
          ReadCard = 2 '装载密码出错_B
          Call quit
          Exit Function
    End If

    '验证密码B
    st = rf_authentication(ByVal icdev, loadmode, sector)
    If st <> 0 Then
          ReadCard = 4 '该卡片不属于本系统所使用
          Call quit
          Exit Function
    End If

    '获取卡序号
    'cardmode = 1
    'st = rf_card(ByVal icdev, cardmode, Snr)
    'If st <> 0 Then
    '      ReadCard = 3 '获取会员卡序号出错
    '      Call quit
    '      Exit Function
    'End If



    '装载密码A
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
          ReadCard = 5 '装载密码A出错
          Call quit
          Exit Function
    End If

    '验证密码A
    st = rf_authentication(ByVal icdev, loadmode, sector)
    If st <> 0 Then
          ReadCard = 6 '密码A不正确
          Call quit
          Exit Function
    End If

    address = 4
    st = rf_read_hex(ByVal icdev, address, databuff32)
    If st <> 0 Then
         ReadCard = 7 '读取卡号错误
         Call quit
         Exit Function
    End If
    'strCardNo = Trim(Str$(Val(databuff32)))
    'strCardNo = Mid(strCardNo, 2, Len(strCardNo))
    strCardNo = Mid(databuff32, 28)
    If strCardNo = "" Then
        ReadCard = 8 '所读卡卡号为空
        Call quit
        Exit Function
    End If



    st = rf_beep(icdev, 5)

    ReadCard = 0
    '取消设备
    Call quit

End Function




