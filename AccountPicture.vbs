'Copyright 2023 xinminn
'Licensed under the Apache License, Version 2.0 (the "License");
'you may not use this file except in compliance with the License.
'You may obtain a copy of the License at
'    http://www.apache.org/licenses/LICENSE-2.0
'Unless required by applicable law or agreed to in writing, software
'distributed under the License is distributed on an "AS IS" BASIS,
'WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
'See the License for the specific language governing permissions and
'limitations under the License.
'
'=====================================
'USAGE:
'    converter jpg to accountpicture-ms format.
'    AccountPicture.vbs <path of 96x96.jpg> <path of 448x448.jpg> <output>
'
'EXAMPLE:
'    AccountPicture.vbs c:\p1.jpg c:\p2.jpg c:\out.accountpicture-ms
'=====================================
dim p1Path, p2Path
if wscript.arguments.length >2 then
  p1Path = wscript.arguments(0)
  p2Path = wscript.arguments(1)
  outPath = wscript.arguments(2)
  p1Buffer = ReadFile(p1Path)
  p2Buffer = ReadFile(p2Path)
  p1Len = UBound(p1Buffer)
  p2Len = UBound(p2Buffer)
  outputLen=77+4+p1Len+52+4+p2len+9
  dim outputBuffer()
  ReDim outputBuffer(outputLen)
  offset=0

  '写文件头
  WriteInt32 outputBuffer, offset, outputLen-4
  WriteInt32 outputBuffer, offset, outputLen-8
  WriteBytes outputBuffer, offset, HexToBuffer("3153505318B08B0B2527444B92BA7933AEB2DDE7")
  WriteInt32 outputBuffer, offset, p1Len+56
  WriteBytes outputBuffer, offset, HexToBuffer("0400000000420000001E000000700072006F007000340032003900340039003600370032003900350000000000")
  '写p1长度
  WriteInt32 outputBuffer, offset, p1Len
  '写p1内容
  WriteBytes outputBuffer, offset, p1Buffer
  '写中间头
  WriteBytes outputBuffer, offset, HexToBuffer("000000")
  WriteInt32 outputBuffer, offset, p2Len+54
  WriteBytes outputBuffer, offset, HexToBuffer("0300000000420000001E000000700072006F007000340032003900340039003600370032003900350000000000")
  '写p2长度
  WriteInt32 outputBuffer, offset, p2Len
  '写p2内容
  WriteBytes outputBuffer, offset, p2Buffer
  '写续尾
  WriteBytes outputBuffer, offset, HexToBuffer("000000000000000000")

  '输出文件
  WriteFile outPath, outputBuffer
end if

Function ReadFile(filePath)
  Dim Buf(), I
  With CreateObject("ADODB.Stream")
    .Mode = 3: .Type = 1: .Open: .LoadFromFile filePath
    ReDim Buf(.Size)
    For I = 0 To .Size - 1: Buf(I) = AscB(.Read(1)): Next
    .Close
  End With
  ReadFile = Buf
End Function

Function HexToBuffer(hex)
  bufferLen = len(hex)
  Dim Buf()
  ReDim Buf(bufferLen/2)
  off=0
  for i = 0 to bufferLen-1 Step 2
    hexStr = Mid(hex, i+1, 2)
    WriteByte Buf, off, cint("&h"+hexStr)
  next
  HexToBuffer = Buf
End Function


'指定位数
Function WriteBytes(buffer, ByRef offset, val)
  for i = 0 to UBound(val) -1 Step 1
    buffer(offset + i) = val(i)
  next
  offset = offset + UBound(val)
End Function

'1位
Function WriteByte(buffer, ByRef offset, val)
  buffer(offset) = val
  offset = offset + 1
End Function

'2位
Function WriteInt16(buffer, ByRef offset, val)
  buffer(offset + 1) = val \ 255
  buffer(offset) = val mod 255
  offset = offset + 2
End Function

'4位
Function WriteInt32(buffer, ByRef offset, val)
  buffer(offset + 3) = val \ ( 2 ^ 24)
  buffer(offset + 2) = (val and &HFF0000)  \ ( 2 ^ 16)
  buffer(offset + 1) = (val and &HFF00)  \ ( 2 ^ 8)
  buffer(offset + 0) = (val and &HFF)
  offset = offset + 4
End Function


Sub WriteFile(filePath, Buf)
 Dim BufferData
 BufferData=""
 Size = UBound(Buf)
  For I = 0 To Size - 1 Step 1
     BufferData =BufferData & right("00" & Hex(Buf(I)),2)
  Next

 Dim Stream, ObjXML, MyNode

 Set ObjXML = CreateObject("Microsoft.XMLDOM")
 Set MyNode = ObjXML.CreateElement("binary")
 Set Stream = CreateObject("ADODB.Stream")

 MyNode.DataType = "bin.hex"
 MyNode.Text = BufferData

 Stream.Type = 1
 Stream.Open
 Stream.Write MyNode.NodeTypedValue
 Stream.SaveToFile filePath, 2
 Stream.Close

 Set stream = Nothing
 Set MyNode = Nothing
 Set ObjXML = Nothing
End Sub