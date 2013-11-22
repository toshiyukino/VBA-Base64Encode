Public Function Base64Encode(b() As Byte) As String
  Dim bytBase64() As Byte '変換テーブル
  Dim bytSTR() As Byte    'エンコード後の文字列を格納する変数
  Dim lngSize As Long     '元のサイズを入れておく変数
  Dim i As Integer, j As Integer

  If Not IsArray(b) Then Exit Function

  '変換テーブル
  bytBase64 = StrConv("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/", vbFromUnicode)

  '引数のサイズを調べておく
  lngSize = UBound(b)

  'とりあえず4倍は確保して多く
  ReDim bytSTR((UBound(b) + 1) * 4 + 1)

  '3バイトの倍数になるように足りないバイト数を ReDim Preserve する(ゼロ埋めしなくていいのかな？)
  If (lngSize + 1 + 3) Mod 3 = 1 Then
    ReDim Preserve b(lngSize + 2) '2バイト増やす
  End If
  If (lngSize + 1 + 3) Mod 3 = 2 Then
    ReDim Preserve b(lngSize + 1) '1バイト増やす
  End If

  '3バイトずつエンコード
  j = 0
  For i = 0 To UBound(b) Step 3
    '10進数 = 16進数 = 2進数 : 3 = &H03 = 00000011, 15 = &H0F = 00001111, 63 = &H3F = 00111111
    bytSTR(j) = bytBase64(Int(b(i) / (2 ^ 2)))                                              '右に2ビットシフト
    bytSTR(j + 1) = bytBase64(Int((b(i) And &H3) * (2 ^ 4)) + Int(b(i + 1) / (2 ^ 4)))      '上位6ビットを0にしてこれを4ビット左シフトして、次の要素を右に4ビットシフトして足す
    bytSTR(j + 2) = bytBase64(Int((b(i + 1) And &HF) * (2 ^ 2)) + Int(b(i + 2) / (2 ^ 6)))  '上位4ビットを0にしてこれを2ビット左シフトして、次の要素を右に6ビットシフトして足す
    bytSTR(j + 3) = bytBase64(Int(b(i + 2) And &H3F))                                       '上位2ビットを0にする
    j = j + 4 'エンコードは4バイトずつ進む
  Next

  'pad処理(足りない部分は'='で埋める)
  If (lngSize + 1) Mod 3 = 1 Then
    bytSTR(j - 2) = AscB("=")
    bytSTR(j - 1) = AscB("=")
  End If
  If (lngSize + 1) Mod 3 = 2 Then
    bytSTR(j - 1) = AscB("=")
  End If

  'いらない部分は削除
  ReDim Preserve bytSTR(j - 1)

  Base64Encode = StrConv(bytSTR, vbUnicode)
End Function
