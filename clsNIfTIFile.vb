Imports System.IO
Public Class clsNIfTIFile
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
    Public Property VoxelOffset As Single       'Imageのオフセット(通常は352）
    Public Property RescaleSlope As Single      'リスケールスロープ
    Public Property RescaleIntercept As Single  'リスケール切片
    Public Property QFormCode As Short          ' qform座標系コード
    Public Property SFormCode As Short          ' sform座標系コード
    Public Property QParamB As Single           ' 四元数パラメータB
    Public Property QParamC As Single           ' 四元数パラメータC
    Public Property QParamD As Single           ' 四元数パラメータD
    Public Property QOffsetX As Single          ' 四元数座標オフセットX
    Public Property QOffsetY As Single          ' 四元数座標オフセットY
    Public Property QOffsetZ As Single          ' 四元数座標オフセットZ
    Public Property SrowX As Single()           ' sform行列の第1行
    Public Property SrowY As Single()           ' sform行列の第2行
    Public Property SrowZ As Single()           ' sform行列の第3行
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

    Sub New()
        BigEndianFlag = False
        Dimension0 = 3
        MatrixX = 0
        MatrixY = 0
        MatrixZ = 0
        Dimension4 = 1
        Dimension5 = 1
        Dimension6 = 1
        Dimension7 = 1
        DataType = 0
        BitsPerPixel = 0
        SizeDim0 = 0
        SizeX = 0
        SizeY = 0
        SizeZ = 0
        SizeDim4 = 0
        SizeDim5 = 0
        SizeDim6 = 0
        SizeDim7 = 0
        VoxelOffset = 352
        RescaleSlope = 1
        RescaleIntercept = 0
        QFormCode = 1
        SFormCode = 1
        QParamB = 0
        QParamC = 0
        QParamD = 0
        QOffsetX = 0
        QOffsetY = 0
        QOffsetZ = 0
        SrowX = New Single(3) {}
        SrowY = New Single(3) {}
        SrowZ = New Single(3) {}
    End Sub

    Sub Read(ByVal FilePath As String)

        If Not File.Exists(FilePath) Then
            Throw New FileNotFoundException($"File not found: {FilePath}")
        End If

        ' streamから読み込むためのBinaryReaderを作成
        Using reader As New BinaryReader(File.OpenRead(FilePath))
            Dim HeaderBuff() As Byte = reader.ReadBytes(352)

            'Endian判定
            If ReadValue(Of Int32)(HeaderBuff, 0, 0) = 348 Then
                BigEndianFlag = False
            ElseIf ReadValue(Of Int32)(HeaderBuff, 0, 1) = 348 Then
                BigEndianFlag = True
            Else
                Throw New InvalidDataException($"File is not NIfTI: {FilePath}")
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
            Dim BytesPerPixel As Integer = BitsPerPixel / 8

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

            'ORIENTATION AND LOCATION
            QFormCode = ReadValue(Of Short)(HeaderBuff, OFFSET_QFORM_CODE, BigEndianFlag)
            SFormCode = ReadValue(Of Short)(HeaderBuff, OFFSET_SFORM_CODE, BigEndianFlag)
            QParamB = ReadValue(Of Single)(HeaderBuff, OFFSET_QPARAM_B, BigEndianFlag)
            QParamC = ReadValue(Of Single)(HeaderBuff, OFFSET_QPARAM_C, BigEndianFlag)
            QParamD = ReadValue(Of Single)(HeaderBuff, OFFSET_QPARAM_D, BigEndianFlag)
            QOffsetX = ReadValue(Of Single)(HeaderBuff, OFFSET_QOFFSET_X, BigEndianFlag)
            QOffsetY = ReadValue(Of Single)(HeaderBuff, OFFSET_QOFFSET_Y, BigEndianFlag)
            QOffsetZ = ReadValue(Of Single)(HeaderBuff, OFFSET_QOFFSET_Z, BigEndianFlag)

            For col As Integer = 0 To 3
                SrowX(col) = ReadValue(Of Single)(HeaderBuff, OFFSET_SROW_X + col * 4, BigEndianFlag)
                SrowY(col) = ReadValue(Of Single)(HeaderBuff, OFFSET_SROW_Y + col * 4, BigEndianFlag)
                SrowZ(col) = ReadValue(Of Single)(HeaderBuff, OFFSET_SROW_Z + col * 4, BigEndianFlag)
            Next

            If SFormCode > 0 Then
                AffineToQuaternion()
                QFormCode = 1
            ElseIf SFormCode = 0 And QFormCode > 0 Then
                QuaternionToAffine()
                SFormCode = 1
            End If

            '画素データを取り込み
            Dim AllPixelBuff() As Byte = reader.ReadBytes(Convert.ToInt64(MatrixX) * Convert.ToInt64(MatrixY) * Convert.ToInt64(MatrixZ) * BytesPerPixel)

            '画素値再現(小数点以下6桁まで担保）
            Dim Offset As Long = 0
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

        '反転処理 
        If SrowX(0) < 0 Then
            FlipDimension(0)
            SrowX(0) *= -1
            SrowX(3) *= -1
        End If

        If SrowY(1) < 0 Then
            FlipDimension(1)
            SrowY(1) *= -1
            SrowY(3) *= -1
        End If

        If SrowZ(2) < 0 Then
            FlipDimension(2)
            SrowZ(2) *= -1
            SrowZ(3) *= -1
        End If

    End Sub

    Sub Write(ByVal FilePath As String, OverWriteFlag As Boolean)

        If File.Exists(FilePath) And OverWriteFlag = False Then
            Console.WriteLine("The file already exists. Do you want to overwrite it? (Y/N)")
            Dim response As String = Console.ReadLine()
            If response.ToUpper() <> "Y" Then Exit Sub
        End If

        'クォータニオンとAffine行列の整合性確保
        AffineToQuaternion()

        Dim HeaderBuff(351) As Byte

        'Endian判定
        'Array.Copy(BitConverter.GetBytes(CLng(348)), 0, HeaderBuff, 0, 4)
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

        'ORIENTATION AND LOCATION
        Buffer.BlockCopy(BitConverter.GetBytes(QFormCode), 0, HeaderBuff, OFFSET_QFORM_CODE, 2)
        Buffer.BlockCopy(BitConverter.GetBytes(SFormCode), 0, HeaderBuff, OFFSET_SFORM_CODE, 2)

        Buffer.BlockCopy(BitConverter.GetBytes(QParamB), 0, HeaderBuff, OFFSET_QPARAM_B, 4)
        Buffer.BlockCopy(BitConverter.GetBytes(QParamC), 0, HeaderBuff, OFFSET_QPARAM_C, 4)
        Buffer.BlockCopy(BitConverter.GetBytes(QParamD), 0, HeaderBuff, OFFSET_QPARAM_D, 4)
        Buffer.BlockCopy(BitConverter.GetBytes(QOffsetX), 0, HeaderBuff, OFFSET_QOFFSET_X, 4)
        Buffer.BlockCopy(BitConverter.GetBytes(QOffsetY), 0, HeaderBuff, OFFSET_QOFFSET_Y, 4)
        Buffer.BlockCopy(BitConverter.GetBytes(QOffsetZ), 0, HeaderBuff, OFFSET_QOFFSET_Z, 4)

        For col As Integer = 0 To 3
            Buffer.BlockCopy(BitConverter.GetBytes(SrowX(col)), 0, HeaderBuff, OFFSET_SROW_X + col * 4, 4)
            Buffer.BlockCopy(BitConverter.GetBytes(SrowY(col)), 0, HeaderBuff, OFFSET_SROW_Y + col * 4, 4)
            Buffer.BlockCopy(BitConverter.GetBytes(SrowZ(col)), 0, HeaderBuff, OFFSET_SROW_Z + col * 4, 4)
        Next

        'Magic Word("n+1" 固定)
        Buffer.BlockCopy(Text.Encoding.ASCII.GetBytes(MAGIC_WORD), 0, HeaderBuff, OFFSET_MAGIC_WORD, MAGIC_WORD.Length)


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
        Using writer As New BinaryWriter(File.OpenWrite(FilePath))
            'ヘッダ書き込み
            writer.Write(HeaderBuff)

            'ピクセル書き込み
            writer.Write(DestBuff)
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
                result = Math.Round(BitConverter.ToSingle(TempBuff, 0), 6, MidpointRounding.AwayFromZero)
            Case GetType(Double)
                result = Math.Round(BitConverter.ToDouble(TempBuff, 0), 6, MidpointRounding.AwayFromZero)
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

    Sub CloneHeader(Source As clsNIfTIFile)
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
        QFormCode = Source.QFormCode
        SFormCode = Source.SFormCode
        QParamB = Source.QParamB
        QParamC = Source.QParamC
        QParamD = Source.QParamD
        QOffsetX = Source.QOffsetX
        QOffsetY = Source.QOffsetY
        QOffsetZ = Source.QOffsetZ
        SrowX = Source.SrowX
        SrowY = Source.SrowY
        SrowZ = Source.SrowZ
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

    Private Sub AffineToQuaternion()

        ' Extract the 3x3 rotation part of the affine matrix
        Dim R As Double(,) = {{SrowX(0), SrowX(1), SrowX(2)},
                          {SrowY(0), SrowY(1), SrowY(2)},
                          {SrowZ(0), SrowZ(1), SrowZ(2)}}

        ' スケール係数を取り除く（列ベクトルの長さ）
        Dim dx As Double = Math.Sqrt(R(0, 0) ^ 2 + R(1, 0) ^ 2 + R(2, 0) ^ 2)
        Dim dy As Double = Math.Sqrt(R(0, 1) ^ 2 + R(1, 1) ^ 2 + R(2, 1) ^ 2)
        Dim dz As Double = Math.Sqrt(R(0, 2) ^ 2 + R(1, 2) ^ 2 + R(2, 2) ^ 2)

        ' qfac判定（右手系 or 左手系）
        Dim det = R(0, 0) * (R(1, 1) * R(2, 2) - R(1, 2) * R(2, 1)) - R(0, 1) * (R(1, 0) * R(2, 2) - R(1, 2) * R(2, 0)) + R(0, 2) * (R(1, 0) * R(2, 1) - R(1, 1) * R(2, 0))
        Dim qfac As Single = If(det > 0, 1.0F, -1.0F)
        dz *= qfac
        SizeDim0 = qfac

        ' 正規化回転行列
        For i = 0 To 2
            R(i, 0) /= dx
            R(i, 1) /= dy
            R(i, 2) /= dz
        Next

        ' Initialize quaternion components
        Dim qw, qx, qy, qz As Double
        Dim trace As Double = R(0, 0) + R(1, 1) + R(2, 2)

        ' Calculate the quaternion based on the trace of the matrix
        If trace > 0 Then
            Dim s As Double = Math.Sqrt(trace + 1.0) * 2
            qw = 0.25 * s
            qx = (R(2, 1) - R(1, 2)) / s
            qy = (R(0, 2) - R(2, 0)) / s
            qz = (R(1, 0) - R(0, 1)) / s
        ElseIf (R(0, 0) > R(1, 1)) AndAlso (R(0, 0) > R(2, 2)) Then
            Dim s As Double = Math.Sqrt(1.0 + R(0, 0) - R(1, 1) - R(2, 2)) * 2
            qw = (R(2, 1) - R(1, 2)) / s
            qx = 0.25 * s
            qy = (R(0, 1) + R(1, 0)) / s
            qz = (R(0, 2) + R(2, 0)) / s
        ElseIf R(1, 1) > R(2, 2) Then
            Dim s As Double = Math.Sqrt(1.0 + R(1, 1) - R(0, 0) - R(2, 2)) * 2
            qw = (R(0, 2) - R(2, 0)) / s
            qx = (R(0, 1) + R(1, 0)) / s
            qy = 0.25 * s
            qz = (R(1, 2) + R(2, 1)) / s
        Else
            Dim s As Double = Math.Sqrt(1.0 + R(2, 2) - R(0, 0) - R(1, 1)) * 2
            qw = (R(1, 0) - R(0, 1)) / s
            qx = (R(0, 2) + R(2, 0)) / s
            qy = (R(1, 2) + R(2, 1)) / s
            qz = 0.25 * s
        End If

        ' Normalize the quaternion
        Dim magnitude As Double = Math.Sqrt(qw * qw + qx * qx + qy * qy + qz * qz)
        qw /= magnitude
        qx /= magnitude
        qy /= magnitude
        qz /= magnitude

        '小数点以下6桁まで担保
        QParamB = Math.Round(qx, 6, MidpointRounding.AwayFromZero)
        QParamC = Math.Round(qy, 6, MidpointRounding.AwayFromZero)
        QParamD = Math.Round(qz, 6, MidpointRounding.AwayFromZero)
        QOffsetX = SrowX(3)
        QOffsetY = SrowY(3)
        QOffsetZ = SrowZ(3)

        QFormCode = 1

    End Sub

    Private Sub QuaternionToAffine()

        Dim b As Double = QParamB
        Dim c As Double = QParamC
        Dim d As Double = QParamD
        Dim a As Double = Math.Sqrt(1.0 - (b * b + c * c + d * d))

        ' 正規化されていない場合の補正
        If (a < 0.0000001) Then
            a = 0.0
            Dim norm As Double = Math.Sqrt(b * b + c * c + d * d)
            If norm > 0.0 Then
                b /= norm
                c /= norm
                d /= norm
            End If
        End If

        '' pixdim: pixdim(1)=dx, pixdim(2)=dy, pixdim(3)=dz
        Dim dx As Double = If(SizeX = 0, 1.0, SizeX)
        Dim dy As Double = If(SizeY = 0, 1.0, SizeY)
        Dim dz As Double = If(SizeZ = 0, 1.0, SizeZ)
        Dim qfac As Double = If(SizeDim0 = 0, 1.0, SizeDim0) ' pixdim(0) = qfac

        dz *= qfac

        ' 回転行列の各列ベクトルを求める
        Dim R(2, 2) As Double

        R(0, 0) = a * a + b * b - c * c - d * d
        R(0, 1) = 2 * (b * c - a * d)
        R(0, 2) = 2 * (b * d + a * c)

        R(1, 0) = 2 * (b * c + a * d)
        R(1, 1) = a * a + c * c - b * b - d * d
        R(1, 2) = 2 * (c * d - a * b)

        R(2, 0) = 2 * (b * d - a * c)
        R(2, 1) = 2 * (c * d + a * b)
        R(2, 2) = a * a + d * d - c * c - b * b

        ' スケーリング適用
        For i As Integer = 0 To 2
            R(i, 0) *= dx
            R(i, 1) *= dy
            R(i, 2) *= dz
        Next

        ' sformに格納(小数点6桁まで担保）
        For i As Integer = 0 To 2
            SrowX(i) = Math.Round(R(0, i), 6, MidpointRounding.AwayFromZero)
            SrowY(i) = Math.Round(R(1, i), 6, MidpointRounding.AwayFromZero)
            SrowZ(i) = Math.Round(R(2, i), 6, MidpointRounding.AwayFromZero)
        Next
        SrowX(3) = QOffsetX
        SrowY(3) = QOffsetY
        SrowZ(3) = QOffsetZ
        SFormCode = 1
    End Sub

    Public Function GetPixelMax() As Double
        Dim MaxValue As Double = Double.MinValue

        For x As Integer = 0 To MatrixX - 1
            For y As Integer = 0 To MatrixY - 1
                For z As Integer = 0 To MatrixZ - 1
                    If MaxValue < Pixel(x, y, z) Then
                        MaxValue = Pixel(x, y, z)
                    End If
                Next
            Next
        Next
        Return Math.Round(MaxValue, 6, MidpointRounding.AwayFromZero)

    End Function

    Public Function GetPixelMin() As Double
        Dim MinValue As Double = Double.MaxValue

        For x As Integer = 0 To MatrixX - 1
            For y As Integer = 0 To MatrixY - 1
                For z As Integer = 0 To MatrixZ - 1
                    If MinValue > Pixel(x, y, z) Then
                        MinValue = Pixel(x, y, z)
                    End If
                Next
            Next
        Next
        Return Math.Round(MinValue, 6, MidpointRounding.AwayFromZero)
    End Function

    Public Function GetPixelAverage() As Double
        Dim SumValue As Double = 0
        Dim PixelCount As Long = 0

        For x As Integer = 0 To MatrixX - 1
            For y As Integer = 0 To MatrixY - 1
                For z As Integer = 0 To MatrixZ - 1
                    SumValue += Pixel(x, y, z)
                    PixelCount += 1
                Next
            Next
        Next

        Return Math.Round(SumValue / PixelCount, 6, MidpointRounding.AwayFromZero)

    End Function

End Class

