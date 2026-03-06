Imports System.IO

Public Class clsAnalyzeFile
    Public Property BigEndianFlag As Boolean    'バイトオーダー（TrueでBigEndian）
    Public Property Dimension0 As Short
    Public Property MatrixX As Short             '横方向のマトリクスサイズ
    Public Property MatrixY As Short            '縦方向のマトリクスサイズ
    Public Property MatrixZ As Short         'スライス枚数
    Public Property Dimension4 As Short
    Public Property Dimension5 As Short
    Public Property Dimension6 As Short
    Public Property Dimension7 As Short
    Public Property DataType As Short           'データタイプ（1:Binary 2:Unsigned char 4:signed short　8:signed long 16:float 64:double 512:unsigned short)
    Public Property BitsPerPixel As Short       'BitsPerPixel
    Public Property SizeDim0 As Single
    Public Property SizeX As Single             '横方向のピクセルサイズ
    Public Property SizeY As Single             '縦方向のピクセルサイズ
    Public Property SizeZ As Single             'スライス厚
    Public Property SizeDim4 As Single
    Public Property SizeDim5 As Single
    Public Property SizeDim6 As Single
    Public Property SizeDim7 As Single
    Public Property VoxelOffset As Single       'Imageのオフセット(通常は0）
    Public Property RescaleSlope As Single      'リスケールスロープ
    Public Property RescaleIntercept As Single  'リスケール切片

    Public Property Pixel As Double(,,)     ' 3次元の画素値配列（z,y,x)

    Private Const OFFSET_DIMENSION As Integer = 40
    Private Const OFFSET_MATRIX_X As Integer = 42
    Private Const OFFSET_MATRIX_Y As Integer = 44
    Private Const OFFSET_MATRIX_Z As Integer = 46
    Private Const OFFSET_DATATYPE As Integer = 70
    Private Const OFFSET_BITPERPIXEL As Integer = 72
    Private Const OFFSET_PIXDIM As Integer = 76
    Private Const OFFSET_SIZE_X As Integer = 80
    Private Const OFFSET_SIZE_Y As Integer = 84
    Private Const OFFSET_SIZE_Z As Integer = 88
    Private Const OFFSET_VOXEL_OFFSET As Integer = 108
    Private Const OFFSET_RESCALE_SLOPE As Integer = 112
    Private Const OFFSET_RESCALE_INTERCEPT As Integer = 116
    Private Const OFFSET_QFORM_CODE As Integer = 252
    Private Const OFFSET_SFORM_CODE As Integer = 254
    Private Const OFFSET_QPARAM_B As Integer = 256
    Private Const OFFSET_QPARAM_C As Integer = 260
    Private Const OFFSET_QPARAM_D As Integer = 264
    Private Const OFFSET_QOFFSET_X As Integer = 268
    Private Const OFFSET_QOFFSET_Y As Integer = 272
    Private Const OFFSET_QOFFSET_Z As Integer = 276
    Private Const OFFSET_SROW_X As Integer = 280
    Private Const OFFSET_SROW_Y As Integer = 296
    Private Const OFFSET_SROW_Z As Integer = 312
    Private Const OFFSET_MAGIC_WORD As Integer = 344
    Private Const MAGIC_WORD As String = "n+1" & Chr(0)

    Public Enum DataFormatType As Short
        DT_NONE = 0
        DT_BINARY = 1
        DT_UNSIGNED_CHAR = 2
        DT_SIGNED_SHORT = 4
        DT_SIGNED_INT = 8
        DT_FLOAT = 16
        DT_COMPLEX = 32
        DT_DOUBLE = 64
        DT_RGB = 128
        DT_ALL = 255
        DT_UINT16 = 512
        DT_UINT32 = 768
    End Enum

    Public Sub New()
        BigEndianFlag = False
        MatrixX = 0
        MatrixY = 0
        MatrixZ = 0
        SizeX = 0
        SizeY = 0
        SizeZ = 0
    End Sub

    Public Sub Read(SourceFilePath As String)
        Dim hdrFilePath As String
        Dim imgFilePath As String
        Dim BytesPerPixel As Integer

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

        Using reader As New BinaryReader(File.OpenRead(hdrFilePath))
            Dim HeaderBuff() As Byte = reader.ReadBytes(352)

            'Endian判定
            If ReadValue(Of Int32)(HeaderBuff, 0, 0) = 348 Then
                BigEndianFlag = False
            ElseIf ReadValue(Of Int32)(HeaderBuff, 0, 1) = 348 Then
                BigEndianFlag = True
            Else
                Throw New InvalidDataException($"File is not NIfTI: {SourceFilePath}")
            End If

            '次元数
            Dimension0 = ReadValue(Of Int16)(HeaderBuff, OFFSET_DIMENSION, BigEndianFlag)

            'Matrixサイズ
            MatrixX = ReadValue(Of Int16)(HeaderBuff, OFFSET_MATRIX_X, BigEndianFlag)
            MatrixY = ReadValue(Of Int16)(HeaderBuff, OFFSET_MATRIX_Y, BigEndianFlag)
            MatrixZ = ReadValue(Of Int16)(HeaderBuff, OFFSET_MATRIX_Z, BigEndianFlag)

            '未使用次元
            Dimension4 = ReadValue(Of Int16)(HeaderBuff, OFFSET_MATRIX_Z + 2, BigEndianFlag)
            Dimension5 = ReadValue(Of Int16)(HeaderBuff, OFFSET_MATRIX_Z + 4, BigEndianFlag)
            Dimension6 = ReadValue(Of Int16)(HeaderBuff, OFFSET_MATRIX_Z + 6, BigEndianFlag)
            Dimension7 = ReadValue(Of Int16)(HeaderBuff, OFFSET_MATRIX_Z + 8, BigEndianFlag)

            'Pixelバッファ
            Pixel = New Double(MatrixX - 1, MatrixY - 1, MatrixZ - 1) {}

            'データタイプ
            DataType = ReadValue(Of Int16)(HeaderBuff, OFFSET_DATATYPE, BigEndianFlag)

            'BitsPerPixel
            BitsPerPixel = ReadValue(Of Int16)(HeaderBuff, OFFSET_BITPERPIXEL, BigEndianFlag)
            BytesPerPixel = BitsPerPixel / 8

            'ピクセルサイズ
            SizeDim0 = ReadValue(Of Single)(HeaderBuff, OFFSET_PIXDIM, BigEndianFlag)
            SizeX = ReadValue(Of Single)(HeaderBuff, OFFSET_SIZE_X, BigEndianFlag)
            SizeY = ReadValue(Of Single)(HeaderBuff, OFFSET_SIZE_Y, BigEndianFlag)
            SizeZ = ReadValue(Of Single)(HeaderBuff, OFFSET_SIZE_Z, BigEndianFlag)
            '未使用サイズ
            SizeDim4 = ReadValue(Of Single)(HeaderBuff, OFFSET_SIZE_Z + 4, BigEndianFlag)
            SizeDim5 = ReadValue(Of Single)(HeaderBuff, OFFSET_SIZE_Z + 8, BigEndianFlag)
            SizeDim6 = ReadValue(Of Single)(HeaderBuff, OFFSET_SIZE_Z + 12, BigEndianFlag)
            SizeDim7 = ReadValue(Of Single)(HeaderBuff, OFFSET_SIZE_Z + 16, BigEndianFlag)

            '画素データのオフセット
            VoxelOffset = ReadValue(Of Single)(HeaderBuff, OFFSET_VOXEL_OFFSET, BigEndianFlag)

            'スケーリングファクタ
            RescaleSlope = ReadValue(Of Single)(HeaderBuff, OFFSET_RESCALE_SLOPE, BigEndianFlag)
            RescaleIntercept = ReadValue(Of Single)(HeaderBuff, OFFSET_RESCALE_INTERCEPT, BigEndianFlag)

        End Using

        Pixel = New Double(MatrixX - 1, MatrixY - 1, MatrixZ - 1) {}

        'imgファイル読み込み
        Using reader As New BinaryReader(File.OpenRead(imgFilePath))
            Dim AllPixelBuff() As Byte = reader.ReadBytes(Convert.ToInt64(MatrixX) * Convert.ToInt64(MatrixY) * Convert.ToInt64(MatrixZ) * BytesPerPixel)

            Dim Offset As Long = 0

            '画素値再現(小数点以下6桁まで担保）
            For z As Integer = 0 To MatrixZ - 1
                For y As Integer = 0 To MatrixY - 1
                    For x As Integer = 0 To MatrixX - 1
                        Select Case DataType
                            Case 1
                                If ReadValue(Of Byte)(AllPixelBuff, Offset, BigEndianFlag) = 0 Then
                                    Pixel(x, y, z) = 0
                                Else
                                    Pixel(x, y, z) = 1
                                End If
                            Case 2
                                Pixel(x, y, z) = Convert.ToDouble(ReadValue(Of Byte)(AllPixelBuff, Offset, BigEndianFlag))
                            Case 4
                                Pixel(x, y, z) = Convert.ToDouble(ReadValue(Of Int16)(AllPixelBuff, Offset, BigEndianFlag))
                            Case 8
                                Pixel(x, y, z) = Convert.ToDouble(ReadValue(Of Int32)(AllPixelBuff, Offset, BigEndianFlag))
                            Case 16
                                Pixel(x, y, z) = Convert.ToDouble(ReadValue(Of Single)(AllPixelBuff, Offset, BigEndianFlag))
                            Case 64
                                Pixel(x, y, z) = Convert.ToDouble(ReadValue(Of Double)(AllPixelBuff, Offset, BigEndianFlag))
                            Case 512
                                Pixel(x, y, z) = Convert.ToDouble(ReadValue(Of UInt16)(AllPixelBuff, Offset, BigEndianFlag))
                            Case Else
                                Throw New NotSupportedException("Type " & DataType.ToString() & " is not supported")
                        End Select

                        If DataType <> 1 AndAlso (RescaleSlope <> 1 Or RescaleIntercept <> 0) Then
                            Pixel(x, y, z) = Math.Round(Pixel(x, y, z) * RescaleSlope + RescaleIntercept, 6, MidpointRounding.AwayFromZero)
                        End If

                        Offset += BytesPerPixel
                    Next
                Next
            Next

        End Using
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

        Return result
    End Function

    Public Function GetPixels() As Double(,,)
        Return CType(Pixel.Clone(), Double(,,))
    End Function

    Public Function GetPixelValue(x As Long, y As Long, z As Long) As Double
        Return CDbl(Pixel(x, y, z))
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

        Pixel = CType(Source.Clone(), Double(,,))
    End Sub

    Public Sub SetPixelValue(x As Long, y As Long, z As Long, Value As Double)
        Pixel(x, y, z) = CDbl(Value)
    End Sub

    Sub CloneHeader(Source As clsAnalyzeFile)
        BigEndianFlag = Source.BigEndianFlag
        MatrixX = Source.MatrixX
        MatrixY = Source.MatrixY
        MatrixZ = Source.MatrixZ
        DataType = Source.DataType
        BitsPerPixel = Source.BitsPerPixel
        SizeX = Source.SizeX
        SizeY = Source.SizeY
        SizeZ = Source.SizeZ
        VoxelOffset = Source.VoxelOffset
        RescaleSlope = Source.RescaleSlope
        RescaleIntercept = Source.RescaleIntercept
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

        Dim HeaderBuff(347) As Byte

        'Endian判定
        HeaderBuff(0) = &H5C
        HeaderBuff(1) = &H1
        HeaderBuff(2) = &H0
        HeaderBuff(3) = &H0

        '次元数（三次元固定）
        HeaderBuff(OFFSET_DIMENSION) = &H3
        HeaderBuff(OFFSET_DIMENSION + 1) = &H0

        'Matrixサイズ,スライス枚数
        Buffer.BlockCopy(BitConverter.GetBytes(MatrixX), 0, HeaderBuff, OFFSET_MATRIX_X, 2)
        Buffer.BlockCopy(BitConverter.GetBytes(MatrixY), 0, HeaderBuff, OFFSET_MATRIX_Y, 2)
        Buffer.BlockCopy(BitConverter.GetBytes(MatrixZ), 0, HeaderBuff, OFFSET_MATRIX_Z, 2)
        Buffer.BlockCopy(BitConverter.GetBytes(Dimension4), 0, HeaderBuff, OFFSET_MATRIX_Z + 4, 2)
        Buffer.BlockCopy(BitConverter.GetBytes(Dimension5), 0, HeaderBuff, OFFSET_MATRIX_Z + 8, 2)
        Buffer.BlockCopy(BitConverter.GetBytes(Dimension6), 0, HeaderBuff, OFFSET_MATRIX_Z + 12, 2)
        Buffer.BlockCopy(BitConverter.GetBytes(Dimension7), 0, HeaderBuff, OFFSET_MATRIX_Z + 16, 2)

        'データタイプ（float型16固定）
        Buffer.BlockCopy(BitConverter.GetBytes(CShort(16)), 0, HeaderBuff, OFFSET_DATATYPE, 2)
        'HeaderBuff(OFFSET_DATATYPE) = &H10

        'bit/pixel（float型=32固定）
        Buffer.BlockCopy(BitConverter.GetBytes(CShort(32)), 0, HeaderBuff, OFFSET_BITPERPIXEL, 2)
        'HeaderBuff(OFFSET_BITPERPIXEL) = &H20

        'ピクセルサイズ
        Buffer.BlockCopy(BitConverter.GetBytes(SizeDim0), 0, HeaderBuff, OFFSET_PIXDIM, 4)
        Buffer.BlockCopy(BitConverter.GetBytes(SizeX), 0, HeaderBuff, OFFSET_SIZE_X, 4)
        Buffer.BlockCopy(BitConverter.GetBytes(SizeY), 0, HeaderBuff, OFFSET_SIZE_Y, 4)
        Buffer.BlockCopy(BitConverter.GetBytes(SizeZ), 0, HeaderBuff, OFFSET_SIZE_Z, 4)
        Buffer.BlockCopy(BitConverter.GetBytes(SizeDim4), 0, HeaderBuff, OFFSET_SIZE_Z + 4, 4)
        Buffer.BlockCopy(BitConverter.GetBytes(SizeDim5), 0, HeaderBuff, OFFSET_SIZE_Z + 8, 4)
        Buffer.BlockCopy(BitConverter.GetBytes(SizeDim6), 0, HeaderBuff, OFFSET_SIZE_Z + 12, 4)
        Buffer.BlockCopy(BitConverter.GetBytes(SizeDim7), 0, HeaderBuff, OFFSET_SIZE_Z + 16, 4)

        '画素データのオフセット
        Buffer.BlockCopy(BitConverter.GetBytes(VoxelOffset), 0, HeaderBuff, OFFSET_VOXEL_OFFSET, 4)

        'スケーリングファクタ(スケーリングなし固定）
        Buffer.BlockCopy(BitConverter.GetBytes(CSng(1)), 0, HeaderBuff, OFFSET_RESCALE_SLOPE, 4)
        Buffer.BlockCopy(BitConverter.GetBytes(CSng(0)), 0, HeaderBuff, OFFSET_RESCALE_INTERCEPT, 4)


        '画素値をバッファリング
        '小数点以下6桁まで担保
        Dim BytesPerPixel As Integer = 4
        Dim DestBuff(Convert.ToInt64(MatrixX) * Convert.ToInt64(MatrixY) * Convert.ToInt64(MatrixZ) * BytesPerPixel - 1) As Byte
        Dim DestOffset As Long = 0

        For z As Integer = 0 To MatrixZ - 1
            For y As Integer = 0 To MatrixY - 1
                For x As Integer = 0 To MatrixX - 1
                    Buffer.BlockCopy(BitConverter.GetBytes(Convert.ToSingle(Math.Round(Pixel(x, y, z), 6, MidpointRounding.AwayFromZero))), 0, DestBuff, DestOffset, BytesPerPixel)
                    DestOffset += BytesPerPixel
                Next
            Next
        Next

        'ファイル書き込み
        Using writer As New BinaryWriter(File.OpenWrite(hdrFilePath))
            'ヘッダ書き込み
            writer.Write(HeaderBuff)
        End Using

        Using writer As New BinaryWriter(File.OpenWrite(imgFilePath))
            'ピクセル書き込み
            writer.Write(DestBuff)
        End Using
    End Sub

End Class