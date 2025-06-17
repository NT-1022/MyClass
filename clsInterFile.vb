Imports System.IO

Public Class clsInterFile
    ' ヘッダー情報
    Public Property Offset As Long
    Public Property BigEndianFlag As Boolean
    Public Property MatrixX As Integer
    Public Property MatrixY As Integer
    Public Property MatrixZ As Integer
    Public Property SizeX As Single
    Public Property SizeY As Single
    Public Property SizeZ As Single
    Public Property RescaleMax As Single
    Public Property DirectionRL As Boolean
    Public Property DirectionAP As Boolean
    Public Property DirectionSI As Boolean
    Public Property DataFormat As DataFormatType
    Public Property BytesPerPixel As Integer

    ' 画像データ
    Public Property Pixel As Double(,,)

    Public Enum DataFormatType
        IntegerFormat = 1
        UnsignedIntegerFormat = 2
        FloatFormat = 3
    End Enum

    Sub New()
        Offset = 0
        BigEndianFlag = False
        MatrixX = 0
        MatrixY = 0
        MatrixZ = 0
        SizeX = 0
        SizeY = 0
        SizeZ = 0
        RescaleMax = 32700
        DirectionRL = True
        DirectionAP = True
        DirectionSI = True
        DataFormat = DataFormatType.IntegerFormat
        UpdateBytesPerPixel()
    End Sub

    Sub CloneHeader(Source As clsInterFile)
        Offset = Source.Offset
        BigEndianFlag = Source.BigEndianFlag
        MatrixX = Source.MatrixX
        MatrixY = Source.MatrixY
        MatrixZ = Source.MatrixZ
        SizeX = Source.SizeX
        SizeY = Source.SizeY
        SizeZ = Source.SizeZ
        RescaleMax = Source.RescaleMax
        DirectionRL = Source.DirectionRL
        DirectionAP = Source.DirectionAP
        DirectionSI = Source.DirectionSI
        DataFormat = Source.DataFormat
        UpdateBytesPerPixel()
    End Sub

    Private Sub UpdateBytesPerPixel()
        BytesPerPixel = If(DataFormat = DataFormatType.FloatFormat, 4, 2)
    End Sub

    Public Sub Read(SourceFilePath As String)

        Dim hdrFilePath As String
        Dim imgFilePath As String

        Select Case Path.GetExtension(SourceFilePath)
            Case ".hdr"
                hdrFilePath = SourceFilePath
                imgFilePath = Path.ChangeExtension(SourceFilePath, "img")
            Case ".img"
                hdrFilePath = Path.ChangeExtension(SourceFilePath, "hdr")
                imgFilePath = SourceFilePath
            Case Else
                Throw New FileNotFoundException("Header/Image file not found.", SourceFilePath)
        End Select

        If Not File.Exists(hdrFilePath) Then
            Throw New FileNotFoundException("Header file not found.", hdrFilePath)
        End If

        If Not File.Exists(imgFilePath) Then
            Throw New FileNotFoundException("Image file not found.", imgFilePath)
        End If

        'hdrファイル読み込み
        Dim HeaderLines As String() = File.ReadAllLines(hdrFilePath)
        For Each HeaderEntry As String In HeaderLines
            HeaderEntry = HeaderEntry.ToLower
            Dim separatorIndex As Integer = HeaderEntry.LastIndexOf(":=")
            If separatorIndex >= 0 Then
                Dim Value As String = HeaderEntry.Substring(separatorIndex + 2).Trim()
                Try
                    Select Case True
                        Case HeaderEntry.Contains("!data offset in bytes")
                            Offset = Convert.ToInt64(Value)
                        Case HeaderEntry.Contains("!imagedata byte order")
                            BigEndianFlag = (Value = "bigendian")
                        Case HeaderEntry.Contains("!matrix size [1]")
                            MatrixX = Convert.ToInt32(Value)
                        Case HeaderEntry.Contains("!matrix size [2]")
                            MatrixY = Convert.ToInt32(Value)
                        Case HeaderEntry.Contains("!number of slices")
                            MatrixZ = Convert.ToInt32(Value)
                        Case HeaderEntry.Contains("!data format")
                            DataFormat = CType(Convert.ToInt32(Value), DataFormatType)
                        Case HeaderEntry.Contains("scaling factor (mm/pixel) [1]")
                            SizeX = Convert.ToSingle(Value)
                        Case HeaderEntry.Contains("scaling factor (mm/pixel) [2]")
                            SizeY = Convert.ToSingle(Value)
                        Case HeaderEntry.Contains("!slice thickness (mm/pixel)")
                            SizeZ = Convert.ToSingle(Value)
                        Case HeaderEntry.Contains("!pixel scaling value")
                            RescaleMax = Convert.ToSingle(Value)
                        Case HeaderEntry.Contains("!the right brain on the left")
                            DirectionRL = (Value = "1")
                        Case HeaderEntry.Contains("!the anterior to the posterior")
                            DirectionAP = (Value = "1")
                        Case HeaderEntry.Contains("!the superior to the inferior")
                            DirectionSI = (Value = "1")
                    End Select
                Catch ex As Exception
                    Console.WriteLine("Error processing header entry: " & ex.Message)
                End Try
            End If
        Next

        UpdateBytesPerPixel()

        Pixel = New Double(MatrixX - 1, MatrixY - 1, MatrixZ - 1) {}

        'imgファイル読み込み

        Using reader As New BinaryReader(File.OpenRead(imgFilePath))
            Dim AllPixelBuff() As Byte = reader.ReadBytes(Convert.ToInt64(MatrixX) * Convert.ToInt64(MatrixY) * Convert.ToInt64(MatrixZ) * BytesPerPixel)
            Dim imgOffset As Long = 0

            For z As Integer = 0 To MatrixZ - 1
                For y As Integer = 0 To MatrixY - 1
                    For x As Integer = 0 To MatrixX - 1
                        Select Case DataFormat
                            Case 1
                                Pixel(x, y, z) = Convert.ToDouble(ReadValue(Of Int16)(AllPixelBuff, imgOffset, BigEndianFlag))
                            Case 2
                                Pixel(x, y, z) = Convert.ToDouble(ReadValue(Of UInt16)(AllPixelBuff, imgOffset, BigEndianFlag))
                            Case 3
                                Pixel(x, y, z) = Convert.ToDouble(ReadValue(Of Single)(AllPixelBuff, imgOffset, BigEndianFlag))
                            Case Else
                                Throw New NotSupportedException("Type " & DataFormat.ToString() & " is not supported")
                        End Select
                        Pixel(x, y, z) = Math.Round(Pixel(x, y, z) / 32700 * RescaleMax, 6, MidpointRounding.AwayFromZero)
                        imgOffset += BytesPerPixel
                    Next
                Next
            Next
        End Using

        '画素値再現後、次元反転処理を追加
        If Not DirectionRL Then
            FlipDimension(0) ' X次元を反転
            DirectionRL = True  ' 処理済み
        End If

        If Not DirectionAP Then
            FlipDimension(1) ' Y次元を反転
            DirectionAP = True  ' 処理済み
        End If

        If Not DirectionSI Then
            FlipDimension(2) ' Z次元を反転
            DirectionSI = True  ' 処理済み
        End If

    End Sub

    Private Function ReadValue(Of T)(Buffer() As Byte, Offset As Long, isBigEndian As Boolean) As T

        ' サイズを型ごとに指定
        Dim BuffSize As Integer = Runtime.InteropServices.Marshal.SizeOf(GetType(T))

        ' 指定したサイズ分のバイトを取り出す
        Dim TempBuff(BuffSize - 1) As Byte
        Array.Copy(Buffer, Offset, TempBuff, 0, BuffSize)

        ' BigEndian対応
        If isBigEndian Then
            Array.Reverse(TempBuff)
        End If

        ' 型に応じたBitConverterのメソッドで変換
        Dim result As Object = Nothing
        Select Case GetType(T)
            Case GetType(Short)
                result = BitConverter.ToInt16(TempBuff, 0)
            Case GetType(Integer)
                result = BitConverter.ToInt32(TempBuff, 0)
            Case GetType(Long)
                result = BitConverter.ToInt64(TempBuff, 0)
            Case GetType(Single)
                result = BitConverter.ToSingle(TempBuff, 0)
            Case GetType(Double)
                result = BitConverter.ToDouble(TempBuff, 0)
            Case GetType(Byte)
                result = TempBuff(0)
            Case GetType(UShort)
                result = BitConverter.ToUInt16(TempBuff, 0)
            Case GetType(ULong)
                result = BitConverter.ToUInt64(TempBuff, 0)
        End Select

        Return CType(result, T)
    End Function

    Sub Write(ByVal DestFilePath As String, OverWriteFlag As Boolean)
        Dim imgFilePath As String
        Dim hdrFilePath As String

        Select Case Path.GetExtension(DestFilePath)
            Case "hdr"
                hdrFilePath = DestFilePath
                imgFilePath = Path.ChangeExtension(DestFilePath, "img")
            Case "img"
                hdrFilePath = Path.ChangeExtension(DestFilePath, "hdr")
                imgFilePath = DestFilePath
            Case Else
                hdrFilePath = Path.ChangeExtension(DestFilePath, "hdr")
                imgFilePath = Path.ChangeExtension(DestFilePath, "img")
        End Select

        If OverWriteFlag = False And (File.Exists(hdrFilePath) OrElse File.Exists(imgFilePath)) Then
            Console.WriteLine("The file already exists. Do you want to overwrite it? (Y/N)")
            Dim response As String = Console.ReadLine()
            If response.ToUpper() <> "Y" Then Exit Sub
        End If

        Dim formatMap As New Dictionary(Of Integer, String) From {
            {1, "signed integer"},
            {2, "unsigned integer"},
            {3, "floating point"}
        }

        'float、LittleEndian 固定で書き込む
        DataFormat = 3
        BigEndianFlag = False

        'ヘッダ構築
        Dim HeaderBuff As New System.Text.StringBuilder()

        HeaderBuff.AppendLine("!data offset in bytes :=" & Offset)
        HeaderBuff.AppendLine("!imagedata byte order :=" & If(BigEndianFlag, "BIGENDIAN", "LITTLEENDIAN"))
        HeaderBuff.AppendLine("!matrix size [1] :=" & MatrixX)
        HeaderBuff.AppendLine("!matrix size [2] :=" & MatrixY)

        If formatMap.ContainsKey(DataFormat) Then
            HeaderBuff.AppendLine("!data format :=" & DataFormat)
            HeaderBuff.AppendLine("!number format :=" & formatMap(DataFormat))
            HeaderBuff.AppendLine("!number of bytes per pixel :=" & If(DataFormat = 3, 4, 2))
        End If

        HeaderBuff.AppendLine("scaling factor (mm/pixel) [1] :=" & SizeX)
        HeaderBuff.AppendLine("scaling factor (mm/pixel) [2] :=" & SizeY)
        HeaderBuff.AppendLine("!pixel scaling value :=32700.000000")
        HeaderBuff.AppendLine("!number of slices :=" & MatrixZ)
        HeaderBuff.AppendLine("!slice thickness (mm/pixel) :=" & SizeZ)

        HeaderBuff.AppendLine("!the right brain on the left   :=" & If(DirectionRL, "1", "0"))
        HeaderBuff.AppendLine("!the anterior to the posterior :=" & If(DirectionAP, "1", "0"))
        HeaderBuff.AppendLine("!the superior to the inferior  :=" & If(DirectionSI, "1", "0"))

        HeaderBuff.AppendLine("!END OF HEADER:=")

        Using writer As New StreamWriter(File.OpenWrite(hdrFilePath))
            writer.Write(HeaderBuff.ToString())
        End Using

        '画素値をバッファリング
        '小数点以下6桁まで担保
        Dim BytesPerPixel As Integer = 4
        Dim DestBuff(Convert.ToInt64(MatrixX) * Convert.ToInt64(MatrixY) * Convert.ToInt64(MatrixZ) * BytesPerPixel - 1) As Byte
        For z As Integer = 0 To MatrixZ - 1
            For y As Integer = 0 To MatrixY - 1
                For x As Integer = 0 To MatrixX - 1
                    Dim TempValue As Single = Convert.ToSingle(Math.Round(Pixel(x, y, z), 6, MidpointRounding.AwayFromZero))
                    Buffer.BlockCopy(BitConverter.GetBytes(TempValue), 0, DestBuff, ((MatrixX * MatrixY * z) + (MatrixX * y) + x) * BytesPerPixel, BytesPerPixel)
                Next
            Next
        Next

        'ピクセル書き込み
        Using writer As New BinaryWriter(File.OpenWrite(imgFilePath))
            writer.Write(DestBuff)
        End Using

    End Sub

    Public Function GetPixels() As Double(,,)
        Return Pixel.Clone()
    End Function

    Public Function GetPixelValue(x As Long, y As Long, z As Long) As Double
        Return Pixel(x, y, z)
    End Function

    Public Sub SetPixels(Source As Double(,,))

        Dim SourceX As Integer = Source.GetLength(0) ' X軸のサイズ
        Dim SourceY As Integer = Source.GetLength(1) ' Y軸のサイズ
        Dim SourceZ As Integer = Source.GetLength(2) ' Z軸のサイズ

        ' サイズ検証
        If SourceX <> MatrixX OrElse SourceY <> MatrixY OrElse SourceZ <> MatrixZ Then
            Throw New ArgumentException($"入力配列のサイズが不正です。" & vbCrLf &
                                    $"期待されるサイズ: ({MatrixX}, {MatrixY}, {MatrixZ})" & vbCrLf &
                                    $"実際のサイズ: ({SourceX}, {SourceY}, {SourceZ})")
        End If

        Pixel = Source.Clone()
    End Sub

    Public Sub SetPixelValue(x As Long, y As Long, z As Long, Value As Double)
        Pixel(x, y, z) = Math.Round(Value, 6, MidpointRounding.AwayFromZero)
    End Sub

    Sub CreatePixelBuff()
        Pixel = New Double(MatrixX - 1, MatrixY - 1, MatrixZ - 1) {}
    End Sub

    Public Sub FlipDimension(axis As Integer)
        ' 指定された軸 (0: X, 1: Y, 2: Z) に沿ってPixel配列を反転
        Select Case axis
            Case 0 ' X軸 (横方向)
                Dim tempVal As Double
                For z As Integer = 0 To MatrixZ - 1
                    For y As Integer = 0 To MatrixY - 1
                        For x1 As Integer = 0 To (MatrixX \ 2) - 1
                            Dim x2 As Integer = MatrixX - 1 - x1
                            tempVal = Pixel(x1, y, z)
                            Pixel(x1, y, z) = Pixel(x2, y, z)
                            Pixel(x2, y, z) = tempVal
                        Next
                    Next
                Next

            Case 1 ' Y軸 (縦方向)
                Dim tempRow(MatrixX - 1) As Double
                For z As Integer = 0 To MatrixZ - 1
                    For y1 As Integer = 0 To (MatrixY \ 2) - 1
                        Dim y2 As Integer = MatrixY - 1 - y1
                        For x As Integer = 0 To MatrixX - 1
                            tempRow(x) = Pixel(x, y1, z)
                            Pixel(x, y1, z) = Pixel(x, y2, z)
                            Pixel(x, y2, z) = tempRow(x)
                        Next
                    Next
                Next

            Case 2 ' Z軸 (スライス方向)
                Dim tempSlice(MatrixY - 1, MatrixX - 1) As Double
                For z1 As Integer = 0 To (MatrixZ \ 2) - 1
                    Dim z2 As Integer = MatrixZ - 1 - z1
                    For y As Integer = 0 To MatrixY - 1
                        For x As Integer = 0 To MatrixX - 1
                            tempSlice(y, x) = Pixel(x, y, z1)
                            Pixel(x, y, z1) = Pixel(x, y, z2)
                            Pixel(x, y, z2) = tempSlice(y, x)
                        Next
                    Next
                Next
        End Select
    End Sub

End Class

