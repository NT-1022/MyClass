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
    Public Property DataType As Short           'データタイプ（1:Binary 2:Unsigned char 4:signed short　8:signed int (int32) 16:float 64:double 512:unsigned short)
    Public Property BitsPerPixel As Short       'BitsPerPixel
    Public Property SizeDim0 As Single
    Public Property SizeX As Single             '横方向のピクセルサイズ
    Public Property SizeY As Single             '縦方向のピクセルサイズ
    Public Property SizeZ As Single             'スライス厚
    Public Property SizeDim4 As Single
    Public Property SizeDim5 As Single
    Public Property SizeDim6 As Single
    Public Property SizeDim7 As Single
    Public Property VoxelOffset As Single       'Imageのオフセット(ni1ならゼロ、n+1なら通常352）
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
    Public Property _Pixel As Double(,,)         ' 3次元の画素値配列（x,y,z)

    Public Property Pixel As Double(,,)
        Get
            Return _Pixel
        End Get
        Private Set(value As Double(,,))
            _Pixel = value
        End Set
    End Property

    Private Const OFFSET_DIMENSION As Integer = 40
    Private Const OFFSET_MATRIX_X As Integer = 42
    Private Const OFFSET_MATRIX_Y As Integer = 44
    Private Const OFFSET_MATRIX_Z As Integer = 46
    Private Const OFFSET_DIM4 As Integer = 48
    Private Const OFFSET_DIM5 As Integer = 50
    Private Const OFFSET_DIM6 As Integer = 52
    Private Const OFFSET_DIM7 As Integer = 54
    Private Const OFFSET_DATATYPE As Integer = 70
    Private Const OFFSET_BITPERPIXEL As Integer = 72
    Private Const OFFSET_PIXDIM As Integer = 76
    Private Const OFFSET_SIZE_X As Integer = 80
    Private Const OFFSET_SIZE_Y As Integer = 84
    Private Const OFFSET_SIZE_Z As Integer = 88
    Private Const OFFSET_SIZE4 As Integer = 92
    Private Const OFFSET_SIZE5 As Integer = 96
    Private Const OFFSET_SIZE6 As Integer = 100
    Private Const OFFSET_SIZE7 As Integer = 104
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
    Private Const RAW_HDR_SIZE As Integer = 348

    Private Enum NiftiStorageKind
        SingleFile   ' .nii
        PairFile     ' .hdr + .img
    End Enum

    Public Enum NiftiDataType As Short
        DT_BINARY = 1
        DT_UINT8 = 2
        DT_INT16 = 4
        DT_INT32 = 8
        DT_FLOAT32 = 16
        DT_FLOAT64 = 64
        DT_UINT16 = 512
    End Enum
    Public ReadOnly Property DataTypeEnum As NiftiDataType
        Get
            Return CType(DataType, NiftiDataType)
        End Get
    End Property

    Private Structure NiftiPaths
        Public Kind As NiftiStorageKind
        Public HeaderPath As String
        Public ImagePath As String
    End Structure

    Private Function ResolveNiftiPaths(inputPath As String) As NiftiPaths
        Dim ext As String = Path.GetExtension(inputPath).ToLowerInvariant()
        Dim base As String = Path.Combine(Path.GetDirectoryName(inputPath), Path.GetFileNameWithoutExtension(inputPath))

        Select Case ext
            Case ".nii"
                Return New NiftiPaths With {.Kind = NiftiStorageKind.SingleFile, .HeaderPath = inputPath, .ImagePath = inputPath}
            Case ".hdr"
                Dim img As String = base & ".img"
                Return New NiftiPaths With {.Kind = NiftiStorageKind.PairFile, .HeaderPath = inputPath, .ImagePath = img}
            Case ".img"
                Dim hdr As String = base & ".hdr"
                Return New NiftiPaths With {.Kind = NiftiStorageKind.PairFile, .HeaderPath = hdr, .ImagePath = inputPath}
            Case Else
                ' 拡張子未指定は .nii とみなす
                Return New NiftiPaths With {.Kind = NiftiStorageKind.SingleFile, .HeaderPath = inputPath, .ImagePath = inputPath}
        End Select
    End Function

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
        Pixel = New Double(-1, -1, -1) {}
    End Sub

    Sub Read(ByVal FilePath As String)

        If Not File.Exists(FilePath) Then
            Throw New FileNotFoundException($"File not found: {FilePath}")
        End If

        Dim Paths As NiftiPaths = ResolveNiftiPaths(FilePath)
        If Paths.Kind = NiftiStorageKind.PairFile Then
            Select Case Path.GetExtension(FilePath).ToLowerInvariant()
                Case ".hdr"
                    If Not File.Exists(Paths.ImagePath) Then Throw New FileNotFoundException("対応する .img が見つかりません: " & Paths.ImagePath)
                Case ".img"
                    If Not File.Exists(Paths.HeaderPath) Then Throw New FileNotFoundException("対応する .hdr が見つかりません: " & Paths.HeaderPath)
            End Select
        End If


        'ヘッダを読み込み
        Dim HeaderBuff() As Byte
        Using hdrReader As New BinaryReader(File.OpenRead(Paths.HeaderPath))
            HeaderBuff = hdrReader.ReadBytes(RAW_HDR_SIZE)
        End Using

        If HeaderBuff.Length < RAW_HDR_SIZE Then
            Throw New InvalidDataException(
        $"ヘッダが短すぎます ({HeaderBuff.Length} bytes)。有効な NIfTI ファイルではありません: {FilePath}")
        End If

        'Endian判定
        If ReadValue(Of Int32)(HeaderBuff, 0, False) = 348 Then
            BigEndianFlag = False
        ElseIf ReadValue(Of Int32)(HeaderBuff, 0, True) = 348 Then
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
        Dimension4 = ReadValue(Of Int16)(HeaderBuff, OFFSET_DIM4, BigEndianFlag)
        Dimension5 = ReadValue(Of Int16)(HeaderBuff, OFFSET_DIM5, BigEndianFlag)
        Dimension6 = ReadValue(Of Int16)(HeaderBuff, OFFSET_DIM6, BigEndianFlag)
        Dimension7 = ReadValue(Of Int16)(HeaderBuff, OFFSET_DIM7, BigEndianFlag)

        'Pixelバッファ
        Pixel = New Double(MatrixX - 1, MatrixY - 1, MatrixZ - 1) {}

        'データタイプ
        DataType = ReadValue(Of Int16)(HeaderBuff, OFFSET_DATATYPE, BigEndianFlag)

        'BitsPerPixel
        BitsPerPixel = ReadValue(Of Int16)(HeaderBuff, OFFSET_BITPERPIXEL, BigEndianFlag)
        Dim BytesPerPixel As Integer = BitsPerPixel \ 8

        'ピクセルサイズ（小数点6桁まで担保）
        SizeDim0 = Math.Round(ReadValue(Of Single)(HeaderBuff, OFFSET_PIXDIM, BigEndianFlag), 6, MidpointRounding.AwayFromZero)
        SizeX = Math.Round(ReadValue(Of Single)(HeaderBuff, OFFSET_SIZE_X, BigEndianFlag), 6, MidpointRounding.AwayFromZero)
        SizeY = Math.Round(ReadValue(Of Single)(HeaderBuff, OFFSET_SIZE_Y, BigEndianFlag), 6, MidpointRounding.AwayFromZero)
        SizeZ = Math.Round(ReadValue(Of Single)(HeaderBuff, OFFSET_SIZE_Z, BigEndianFlag), 6, MidpointRounding.AwayFromZero)

        '未使用サイズ
        SizeDim4 = Math.Round(ReadValue(Of Single)(HeaderBuff, OFFSET_SIZE4, BigEndianFlag), 6, MidpointRounding.AwayFromZero)
        SizeDim5 = Math.Round(ReadValue(Of Single)(HeaderBuff, OFFSET_SIZE5, BigEndianFlag), 6, MidpointRounding.AwayFromZero)
        SizeDim6 = Math.Round(ReadValue(Of Single)(HeaderBuff, OFFSET_SIZE6, BigEndianFlag), 6, MidpointRounding.AwayFromZero)
        SizeDim7 = Math.Round(ReadValue(Of Single)(HeaderBuff, OFFSET_SIZE7, BigEndianFlag), 6, MidpointRounding.AwayFromZero)

        '画素データのオフセット
        VoxelOffset = ReadValue(Of Single)(HeaderBuff, OFFSET_VOXEL_OFFSET, BigEndianFlag)

        'スケーリングファクタ（小数点6桁まで担保）
        RescaleSlope = Math.Round(ReadValue(Of Single)(HeaderBuff, OFFSET_RESCALE_SLOPE, BigEndianFlag), 6, MidpointRounding.AwayFromZero)
        RescaleIntercept = Math.Round(ReadValue(Of Single)(HeaderBuff, OFFSET_RESCALE_INTERCEPT, BigEndianFlag), 6, MidpointRounding.AwayFromZero)

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

        Dim VoxelCount As Long = CLng(MatrixX) * CLng(MatrixY) * CLng(MatrixZ)
        Dim ByteCount As Long
        If DataType = 1 Then
            ByteCount = (VoxelCount + 7) \ 8   ' ceiling division
        Else
            ByteCount = VoxelCount * BytesPerPixel
        End If

        If ByteCount > Integer.MaxValue Then
            Throw New NotSupportedException(
        $"画像データが大きすぎます ({ByteCount:N0} bytes)。対応上限は {Integer.MaxValue:N0} bytes です。")
        End If

        Dim AllPixelBuff(ByteCount - 1) As Byte

        If Paths.Kind = NiftiStorageKind.SingleFile Then
            ' 単一ファイル：画像先頭は vox_offset（拡張がなければ 352）
            Using fs As New FileStream(Paths.ImagePath, FileMode.Open, FileAccess.Read, FileShare.Read)
                fs.Seek(CLng(Math.Truncate(VoxelOffset)), SeekOrigin.Begin)
                Dim actuallyRead As Integer = fs.Read(AllPixelBuff, 0, CInt(ByteCount))
                If actuallyRead <> ByteCount Then
                    Throw New EndOfStreamException("画像データが不足しています。")
                End If
            End Using
        Else
            ' 分割ペア：.img 側の先頭から画像データ（vox_offset は通常 0 と解釈）
            Using fs As New FileStream(Paths.ImagePath, FileMode.Open, FileAccess.Read, FileShare.Read)
                Dim actuallyRead As Integer = fs.Read(AllPixelBuff, 0, CInt(ByteCount))
                If actuallyRead <> ByteCount Then
                    Throw New EndOfStreamException(".img の画像データが不足しています。")
                End If
            End Using
        End If

        ' Endian が Big のときは型幅に応じて一括バイトスワップ
        If BigEndianFlag Then
            Select Case DataType
                Case 1  ' binary: ビットパック配列。バイト内のビット順はEndianに依存しないためスワップ不要
            ' nothing
                Case 2  ' uint8: 1バイト単位のためスワップ不要
            ' nothing
                Case 4, 512, 8, 16, 64
                    Select Case BytesPerPixel
                        Case 2 : SwapBytes2(AllPixelBuff)
                        Case 4 : SwapBytes4(AllPixelBuff)
                        Case 8 : SwapBytes8(AllPixelBuff)
                    End Select
                Case Else
                    Throw New NotSupportedException("Type " & DataType.ToString() & " is not supported in endian swap.")
            End Select
        End If

        ' 中間の一次元バッファ（型配列）に一括コピー → その後に Double(,,) へ詰め替え
        ' 必要に応じて slope/intercept をベクトル適用
        Select Case DataType
            Case 1 ' binary: 1ビット/ボクセル、MSBファーストのビットパック
                ' ビットインデックスを直接カウントし、対応するバイトとビット位置を求める。
                ' 例: ボクセル0 → byte[0] の bit7(MSB)
                '     ボクセル7 → byte[0] の bit0(LSB)
                '     ボクセル8 → byte[1] の bit7(MSB)
                Dim BitIndex As Long = 0
                For z As Integer = 0 To MatrixZ - 1
                    For y As Integer = 0 To MatrixY - 1
                        For x As Integer = 0 To MatrixX - 1
                            Dim BytePos As Long = BitIndex >> 3              ' BitIndex \ 8
                            Dim BitPos As Integer = 7 - CInt(BitIndex And 7) ' MSBファースト: bit7→6→...→0
                            Pixel(x, y, z) = If((AllPixelBuff(BytePos) And (1 << BitPos)) <> 0, 1.0R, 0.0R)
                            BitIndex += 1
                        Next
                    Next
                Next

            Case 2 ' uint8
                Dim TempBuff(VoxelCount - 1) As Byte
                Buffer.BlockCopy(AllPixelBuff, 0, TempBuff, 0, AllPixelBuff.Length)

                Dim PixelIndex As Long = 0
                If RescaleSlope <> 1.0F OrElse RescaleIntercept <> 0.0F Then
                    Dim s As Double = RescaleSlope
                    Dim c As Double = RescaleIntercept
                    For z As Integer = 0 To MatrixZ - 1
                        For y As Integer = 0 To MatrixY - 1
                            For x As Integer = 0 To MatrixX - 1
                                Pixel(x, y, z) = TempBuff(PixelIndex) * s + c
                                PixelIndex += 1
                            Next
                        Next
                    Next
                Else
                    For z As Integer = 0 To MatrixZ - 1
                        For y As Integer = 0 To MatrixY - 1
                            For x As Integer = 0 To MatrixX - 1
                                Pixel(x, y, z) = TempBuff(PixelIndex)
                                PixelIndex += 1
                            Next
                        Next
                    Next
                End If

            Case 512 ' uint16
                Dim TempBuff(VoxelCount - 1) As UShort
                Buffer.BlockCopy(AllPixelBuff, 0, TempBuff, 0, AllPixelBuff.Length)

                Dim PixelIndex As Long = 0
                If RescaleSlope <> 1.0F OrElse RescaleIntercept <> 0.0F Then
                    Dim s As Double = RescaleSlope
                    Dim c As Double = RescaleIntercept
                    For z As Integer = 0 To MatrixZ - 1
                        For y As Integer = 0 To MatrixY - 1
                            For x As Integer = 0 To MatrixX - 1
                                Pixel(x, y, z) = TempBuff(PixelIndex) * s + c
                                PixelIndex += 1
                            Next
                        Next
                    Next
                Else
                    For z As Integer = 0 To MatrixZ - 1
                        For y As Integer = 0 To MatrixY - 1
                            For x As Integer = 0 To MatrixX - 1
                                Pixel(x, y, z) = TempBuff(PixelIndex)
                                PixelIndex += 1
                            Next
                        Next
                    Next
                End If

            Case 4 ' int16
                Dim TempBuff(VoxelCount - 1) As Short
                Buffer.BlockCopy(AllPixelBuff, 0, TempBuff, 0, AllPixelBuff.Length)

                Dim PixelIndex As Long = 0
                If RescaleSlope <> 1.0F OrElse RescaleIntercept <> 0.0F Then
                    Dim s As Double = RescaleSlope
                    Dim c As Double = RescaleIntercept
                    For z As Integer = 0 To MatrixZ - 1
                        For y As Integer = 0 To MatrixY - 1
                            For x As Integer = 0 To MatrixX - 1
                                Pixel(x, y, z) = CLng(TempBuff(PixelIndex)) * s + c
                                PixelIndex += 1
                            Next
                        Next
                    Next
                Else
                    For z As Integer = 0 To MatrixZ - 1
                        For y As Integer = 0 To MatrixY - 1
                            For x As Integer = 0 To MatrixX - 1
                                Pixel(x, y, z) = TempBuff(PixelIndex)
                                PixelIndex += 1
                            Next
                        Next
                    Next
                End If

            Case 8 ' int32
                Dim TempBuff(VoxelCount - 1) As Integer
                Buffer.BlockCopy(AllPixelBuff, 0, TempBuff, 0, AllPixelBuff.Length)

                Dim PixelIndex As Long = 0
                If RescaleSlope <> 1.0F OrElse RescaleIntercept <> 0.0F Then
                    Dim s As Double = RescaleSlope
                    Dim c As Double = RescaleIntercept
                    For z As Integer = 0 To MatrixZ - 1
                        For y As Integer = 0 To MatrixY - 1
                            For x As Integer = 0 To MatrixX - 1
                                Pixel(x, y, z) = CLng(TempBuff(PixelIndex)) * s + c
                                PixelIndex += 1
                            Next
                        Next
                    Next
                Else
                    For z As Integer = 0 To MatrixZ - 1
                        For y As Integer = 0 To MatrixY - 1
                            For x As Integer = 0 To MatrixX - 1
                                Pixel(x, y, z) = TempBuff(PixelIndex)
                                PixelIndex += 1
                            Next
                        Next
                    Next
                End If

            Case 16 ' float32
                Dim TempBuff(VoxelCount - 1) As Single
                Buffer.BlockCopy(AllPixelBuff, 0, TempBuff, 0, AllPixelBuff.Length)

                If RescaleSlope <> 1.0F OrElse RescaleIntercept <> 0.0F Then
                    Dim s As Double = RescaleSlope
                    Dim c As Double = RescaleIntercept
                    For i As Long = 0 To VoxelCount - 1
                        TempBuff(i) = TempBuff(i) * s + c
                    Next
                End If

                ' Single() → Double(,,) へ詰め替え
                '小数点6桁まで担保
                Dim PixelIndex As Long = 0
                For z As Integer = 0 To MatrixZ - 1
                    For y As Integer = 0 To MatrixY - 1
                        For x As Integer = 0 To MatrixX - 1
                            Pixel(x, y, z) = Math.Round(TempBuff(PixelIndex), 6, MidpointRounding.AwayFromZero)
                            PixelIndex += 1
                        Next
                    Next
                Next

            Case 64 ' float64
                Dim TempBuff(VoxelCount - 1) As Double
                Buffer.BlockCopy(AllPixelBuff, 0, TempBuff, 0, AllPixelBuff.Length)

                If RescaleSlope <> 1.0F OrElse RescaleIntercept <> 0.0F Then
                    Dim s As Double = RescaleSlope
                    Dim c As Double = RescaleIntercept
                    For i As Long = 0 To VoxelCount - 1
                        TempBuff(i) = TempBuff(i) * s + c
                    Next
                End If

                '小数点6桁まで担保
                Dim PixelIndex As Long = 0
                For z As Integer = 0 To MatrixZ - 1
                    For y As Integer = 0 To MatrixY - 1
                        For x As Integer = 0 To MatrixX - 1
                            Pixel(x, y, z) = Math.Round(TempBuff(PixelIndex), 6, MidpointRounding.AwayFromZero)
                            PixelIndex += 1
                        Next
                    Next
                Next

            Case Else
                Throw New NotSupportedException("Type " & DataType.ToString() & " is not supported")
        End Select

        'LPS並びを強制する 
        EnforceLPS()

    End Sub

    Sub Write(ByVal FilePath As String, ForceOverWriteFlag As Boolean)

        Dim Paths As NiftiPaths = ResolveNiftiPaths(FilePath)

        If Paths.Kind = NiftiStorageKind.SingleFile Then
            If File.Exists(Paths.ImagePath) AndAlso Not ForceOverWriteFlag Then
                Throw New IOException("出力先 .nii が既に存在します: " & Paths.ImagePath)
            End If
        Else
            If (File.Exists(Paths.HeaderPath) OrElse File.Exists(Paths.ImagePath)) AndAlso Not ForceOverWriteFlag Then
                Throw New IOException("出力先 .hdr/.img のいずれかが既に存在します。")
            End If
        End If

        If Pixel Is Nothing OrElse Pixel.Length = 0 Then
            Throw New InvalidOperationException(
        "Pixel バッファが初期化されていません。書き込み前に Read() または CreatePixelBuff() を呼び出してください。")
        End If

        'クォータニオンとAffine行列の整合性確保
        AffineToQuaternion()

        Dim HeaderBuff(RAW_HDR_SIZE - 1) As Byte

        'Endian判定（LittleEndian固定）
        Buffer.BlockCopy(BitConverter.GetBytes(348), 0, HeaderBuff, 0, 4)
        'BigEndianFlag = False

        '次元数（三次元固定）
        'Dimension0 = 3
        Dim WriteDimension0 As Short = 3

        'Matrixサイズ,スライス枚数
        Buffer.BlockCopy(BitConverter.GetBytes(WriteDimension0), 0, HeaderBuff, OFFSET_DIMENSION, 2)
        Buffer.BlockCopy(BitConverter.GetBytes(MatrixX), 0, HeaderBuff, OFFSET_MATRIX_X, 2)
        Buffer.BlockCopy(BitConverter.GetBytes(MatrixY), 0, HeaderBuff, OFFSET_MATRIX_Y, 2)
        Buffer.BlockCopy(BitConverter.GetBytes(MatrixZ), 0, HeaderBuff, OFFSET_MATRIX_Z, 2)
        Buffer.BlockCopy(BitConverter.GetBytes(Dimension4), 0, HeaderBuff, OFFSET_DIM4, 2)
        Buffer.BlockCopy(BitConverter.GetBytes(Dimension5), 0, HeaderBuff, OFFSET_DIM5, 2)
        Buffer.BlockCopy(BitConverter.GetBytes(Dimension6), 0, HeaderBuff, OFFSET_DIM6, 2)
        Buffer.BlockCopy(BitConverter.GetBytes(Dimension7), 0, HeaderBuff, OFFSET_DIM7, 2)

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
        Buffer.BlockCopy(BitConverter.GetBytes(SizeDim4), 0, HeaderBuff, OFFSET_SIZE4, 4)
        Buffer.BlockCopy(BitConverter.GetBytes(SizeDim5), 0, HeaderBuff, OFFSET_SIZE5, 4)
        Buffer.BlockCopy(BitConverter.GetBytes(SizeDim6), 0, HeaderBuff, OFFSET_SIZE6, 4)
        Buffer.BlockCopy(BitConverter.GetBytes(SizeDim7), 0, HeaderBuff, OFFSET_SIZE7, 4)

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

        '画素データのオフセットとMAGIC_WORD
        Dim MAGIC_WORD As Byte()
        Dim WriteVoxelOffset As Single
        If Paths.Kind = NiftiStorageKind.SingleFile Then
            WriteVoxelOffset = 352.0F
            MAGIC_WORD = Text.Encoding.ASCII.GetBytes("n+1" & ChrW(0))
        Else
            WriteVoxelOffset = 0.0F
            MAGIC_WORD = Text.Encoding.ASCII.GetBytes("ni1" & ChrW(0))
        End If
        Buffer.BlockCopy(BitConverter.GetBytes(WriteVoxelOffset), 0, HeaderBuff, OFFSET_VOXEL_OFFSET, 4)
        Buffer.BlockCopy(MAGIC_WORD, 0, HeaderBuff, OFFSET_MAGIC_WORD, MAGIC_WORD.Length)

        '画素値のバッファリング
        Dim VoxelCount As Long = CLng(MatrixX) * CLng(MatrixY) * CLng(MatrixZ)

        ' Double(,,) → Single() 一括変換（丸めは必要時のみ）
        Dim TempBuff(VoxelCount - 1) As Single ' Single の一次元バッファ
        Dim PixelIndex As Long = 0
        For z As Integer = 0 To MatrixZ - 1
            For y As Integer = 0 To MatrixY - 1
                For x As Integer = 0 To MatrixX - 1
                    TempBuff(PixelIndex) = CSng(Math.Round(Pixel(x, y, z), 6, MidpointRounding.AwayFromZero))
                    PixelIndex += 1
                Next
            Next
        Next

        ' Single() → byte[] 一括コピー
        Dim DestBuff As Byte() = New Byte(VoxelCount * 4 - 1) {}
        Buffer.BlockCopy(TempBuff, 0, DestBuff, 0, DestBuff.Length)

        'ファイル書き込み
        If Paths.Kind = NiftiStorageKind.SingleFile Then
            Using fs As New FileStream(Paths.ImagePath, FileMode.Create, FileAccess.Write, FileShare.None)
                fs.Write(HeaderBuff, 0, HeaderBuff.Length)                 ' 348B
                fs.Write(New Byte() {0, 0, 0, 0}, 0, 4)                    ' extension flag (0)
                fs.Write(DestBuff, 0, DestBuff.Length)                     ' pixels
            End Using
        Else
            Using fsHdr As New FileStream(Paths.HeaderPath, FileMode.Create, FileAccess.Write, FileShare.None)
                fsHdr.Write(HeaderBuff, 0, HeaderBuff.Length)
            End Using
            Using fsImg As New FileStream(Paths.ImagePath, FileMode.Create, FileAccess.Write, FileShare.None)
                fsImg.Write(DestBuff, 0, DestBuff.Length)
            End Using
        End If

    End Sub

    Private Function ReadValue(Of T)(Buffer() As Byte, Offset As Long, isBigEndian As Boolean) As T

        Dim BuffSize As Integer = Runtime.InteropServices.Marshal.SizeOf(GetType(T))
        Dim TempBuff(BuffSize - 1) As Byte
        Array.Copy(Buffer, Offset, TempBuff, 0, BuffSize)

        If isBigEndian Then
            Array.Reverse(TempBuff)
        End If

        Dim result As Object
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
            Case Else
                Throw New NotSupportedException(
                $"ReadValue: 未対応の型です: {GetType(T).Name}")
        End Select

        Return DirectCast(result, T)
    End Function

    Public Function GetPixels() As Double(,,)
        Return DirectCast(Pixel.Clone(), Double(,,))
    End Function

    Public Function GetPixelValue(x As Integer, y As Integer, z As Integer) As Double
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

    Public Sub SetPixelValue(x As Integer, y As Integer, z As Integer, Value As Double)
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
        SrowX = DirectCast(Source.SrowX.Clone(), Single())
        SrowY = DirectCast(Source.SrowY.Clone(), Single())
        SrowZ = DirectCast(Source.SrowZ.Clone(), Single())
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

            Case Else
                Throw New ArgumentOutOfRangeException(NameOf(axis), "axis は 0（X）、1（Y）、2（Z）のいずれかを指定してください。")
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

        If dx < 0.0000000001 OrElse dy < 0.0000000001 OrElse dz < 0.0000000001 Then
            Throw New InvalidOperationException("sform 行列の列ノルムがゼロです。有効なアフィン行列が設定されていません。")
        End If

        'スケーリング済み行列の行列式で符号を判定
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
        If magnitude < 0.0000000001 Then
            Throw New InvalidOperationException("四元数の正規化に失敗しました。回転行列が不正です。")
        End If
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
        If a < 0.0000001 Then
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
        If Pixel Is Nothing OrElse Pixel.Length = 0 Then
            Throw New InvalidOperationException(
            "Pixel バッファが初期化されていません。Read() または CreatePixelBuff() を先に呼び出してください。")
        End If

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
        If Pixel Is Nothing OrElse Pixel.Length = 0 Then
            Throw New InvalidOperationException(
            "Pixel バッファが初期化されていません。Read() または CreatePixelBuff() を先に呼び出してください。")
        End If

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
        If Pixel Is Nothing OrElse Pixel.Length = 0 Then
            Throw New InvalidOperationException(
            "Pixel バッファが初期化されていません。Read() または CreatePixelBuff() を先に呼び出してください。")
        End If

        Dim PixelCount As Long = CLng(MatrixX) * CLng(MatrixY) * CLng(MatrixZ)
        Dim SumValue As Double = 0
        For x As Integer = 0 To MatrixX - 1
            For y As Integer = 0 To MatrixY - 1
                For z As Integer = 0 To MatrixZ - 1
                    SumValue += Pixel(x, y, z)
                Next
            Next
        Next
        Return Math.Round(SumValue / PixelCount, 6, MidpointRounding.AwayFromZero)
    End Function

    ' --- 2,4,8バイト幅の一括スワップ（.NET Framework/Standard対応） ---
    Private Shared Sub SwapBytes2(ByRef a() As Byte)
        For i As Integer = 0 To a.Length - 1 Step 2
            Dim t As Byte = a(i)
            a(i) = a(i + 1)
            a(i + 1) = t
        Next
    End Sub

    Private Shared Sub SwapBytes4(ByRef a() As Byte)
        For i As Integer = 0 To a.Length - 1 Step 4
            Dim t0 As Byte = a(i) : Dim t1 As Byte = a(i + 1)
            a(i) = a(i + 3) : a(i + 1) = a(i + 2)
            a(i + 2) = t1 : a(i + 3) = t0
        Next
    End Sub

    Private Shared Sub SwapBytes8(ByRef a() As Byte)
        For i As Integer = 0 To a.Length - 1 Step 8
            Dim t0 As Byte = a(i) : Dim t1 As Byte = a(i + 1)
            Dim t2 As Byte = a(i + 2) : Dim t3 As Byte = a(i + 3)
            a(i) = a(i + 7) : a(i + 1) = a(i + 6)
            a(i + 2) = a(i + 5) : a(i + 3) = a(i + 4)
            a(i + 4) = t3 : a(i + 5) = t2
            a(i + 6) = t1 : a(i + 7) = t0
        Next
    End Sub

    '''' <summary>1
    '''' 読み込んだ NIfTI の sform/qform に基づき、Pixel を LPS 並び（+i=Left, +j=Posterior, +k=Superior）になるよう
    '''' 必要な軸をフリップする。フリップ後は srow と qform を整合更新する。
    '''' 斜め撮像（オフダイアゴナルが大きい）や軸入替がある場合はフリップだけでは LPS 化できない点に注意。
    '''' </summary>
    'Private Sub EnforceLPS(Optional offdiagTolerance As Double = 0.0001)
    '    ' まず sform を優先。なければ qform から sform を構築
    '    If SFormCode = 0 AndAlso QFormCode > 0 Then
    '        QuaternionToAffine() ' srowX/Y/Z が埋まる
    '        SFormCode = 1
    '    End If

    '    ' 列ベクトル（voxel i/j/k が世界座標に与える寄与）
    '    '   col0 = (SrowX(0), SrowY(0), SrowZ(0))  ← i 方向
    '    '   col1 = (SrowX(1), SrowY(1), SrowZ(1))  ← j 方向
    '    '   col2 = (SrowX(2), SrowY(2), SrowZ(2))  ← k 方向
    '    Dim col0x As Double = SrowX(0), col0y As Double = SrowY(0), col0z As Double = SrowZ(0)
    '    Dim col1x As Double = SrowX(1), col1y As Double = SrowY(1), col1z As Double = SrowZ(1)
    '    Dim col2x As Double = SrowX(2), col2y As Double = SrowY(2), col2z As Double = SrowZ(2)

    '    ' --- 軸整合の簡易チェック（斜めや permutation の検出） ---
    '    ' 世界 X は主に col0.x、世界 Y は主に col1.y、世界 Z は主に col2.z が支配的であることを期待。
    '    ' オフダイアゴナルが対角より大きい場合は、フリップのみでは LPS 化できない可能性がある。
    '    Dim okAxisAligned As Boolean =
    '        (Math.Abs(col0x) >= Math.Abs(col0y) + Math.Abs(col0z) - offdiagTolerance) AndAlso
    '        (Math.Abs(col1y) >= Math.Abs(col1x) + Math.Abs(col1z) - offdiagTolerance) AndAlso
    '        (Math.Abs(col2z) >= Math.Abs(col2x) + Math.Abs(col2y) - offdiagTolerance)

    '    If Not okAxisAligned Then
    '        ' 必要ならここでログや例外を投げる。ここでは注意だけ出して、符号のみ合わせる実装を継続。
    '        Console.WriteLine("Warning: Oblique/permuted orientation detected. Flip-only LPS normalization may be insufficient.")
    '        Return   ' 斜め撮像はフリップのみでは対処不可なため処理を中断
    '    End If

    '    ' --- LPS の符号合わせ：+X(L), +Y(P), +Z(S) が正方向になるように判定 ---
    '    ' 対角成分の符号を見る（軸整合を仮定）。負であればその軸をフリップ。
    '    If SrowX(0) < 0 Then            ' +i が -X に向いている → 反転
    '        FlipAxisAndUpdateAffine(0)
    '    End If

    '    If SrowY(1) < 0 Then            ' +j が -Y に向いている → 反転
    '        FlipAxisAndUpdateAffine(1)
    '    End If

    '    If SrowZ(2) < 0 Then            ' +k が -Z に向いている → 反転
    '        FlipAxisAndUpdateAffine(2)
    '    End If

    '    ' srow を更新したので qform も整合させる
    '    AffineToQuaternion()
    '    QFormCode = 1
    '    SFormCode = 1
    'End Sub

    ''' <summary>
    ''' 指定軸（0:X=i, 1:Y=j, 2:Z=k）で Pixel を反転し、srow の列ベクトルと平行移動を正しく更新する。
    ''' </summary>
    Private Sub FlipAxisAndUpdateAffine(axis As Integer)
        ' まず画素を反転（既存の高速 FlipDimension を利用）
        FlipDimension(axis)

        ' 反転前の列ベクトル c = (SrowX(axis), SrowY(axis), SrowZ(axis))
        Dim cx As Double = SrowX(axis)
        Dim cy As Double = SrowY(axis)
        Dim cz As Double = SrowZ(axis)

        ' 平行移動 t = t + c * (N-1)
        Dim n As Integer = If(axis = 0, MatrixX, If(axis = 1, MatrixY, MatrixZ))
        SrowX(3) = CSng(SrowX(3) + cx * (n - 1))
        SrowY(3) = CSng(SrowY(3) + cy * (n - 1))
        SrowZ(3) = CSng(SrowZ(3) + cz * (n - 1))

        ' 列 c を -c に反転
        SrowX(axis) = CSng(-cx)
        SrowY(axis) = CSng(-cy)
        SrowZ(axis) = CSng(-cz)
    End Sub


    ''' <summary>
    ''' sform/qform に基づき Pixel を LPS 方向に正規化する。
    ''' 軸の並び替えが必要な場合は Permute、符号が逆な場合はフリップを行う。
    ''' 真の斜め撮像（oblique）の場合は最近接軸に近似して処理し、警告を記録する。
    ''' 処理後は srow・qform を整合更新する。
    ''' </summary>
    Private Sub EnforceLPS()

        ' sform 優先。なければ qform から sform を構築
        If SFormCode = 0 AndAlso QFormCode > 0 Then
            QuaternionToAffine()
            SFormCode = 1
        End If

        If SFormCode = 0 AndAlso QFormCode = 0 Then
            ' 座標情報が一切ない場合はスキップ（Method 1 = ピクセルサイズのみ）
            Console.WriteLine("Warning: Both sform_code and qform_code are 0. LPS normalization skipped.")
            Return
        End If

        ' -------------------------------------------------------
        ' Step 1: アフィン行列から各ボクセル軸（i/j/k）が
        '         世界座標の X/Y/Z のどの軸に最も近いかを求める
        '
        ' col(n) = (SrowX(n), SrowY(n), SrowZ(n)) が
        ' ボクセル軸 n の世界座標への寄与ベクトル。
        ' 絶対値が最大の成分がその軸の「支配的な世界軸」。
        ' -------------------------------------------------------
        Dim col(2, 2) As Double   ' col(voxelAxis, worldXYZ)
        For n As Integer = 0 To 2
            col(n, 0) = SrowX(n)
            col(n, 1) = SrowY(n)
            col(n, 2) = SrowZ(n)
        Next

        ' 各ボクセル軸について支配的な世界軸インデックス（0=X,1=Y,2=Z）を求める
        Dim dominantWorld(2) As Integer   ' dominantWorld(voxelAxis) = 世界軸インデックス
        For n As Integer = 0 To 2
            Dim maxAbs As Double = -1
            Dim best As Integer = 0
            For w As Integer = 0 To 2
                If Math.Abs(col(n, w)) > maxAbs Then
                    maxAbs = Math.Abs(col(n, w))
                    best = w
                End If
            Next
            dominantWorld(n) = best
        Next

        ' 支配的な世界軸が重複していないか確認（permutation の一意性チェック）
        Dim worldUsed(2) As Boolean
        Dim isValidPermutation As Boolean = True
        For n As Integer = 0 To 2
            Dim w As Integer = dominantWorld(n)
            If worldUsed(w) Then
                isValidPermutation = False
                Exit For
            End If
            worldUsed(w) = True
        Next

        If Not isValidPermutation Then
            ' 2つ以上のボクセル軸が同じ世界軸を支配 → 真の斜め撮像で近似不可
            Throw New InvalidDataException(
                "アフィン行列が著しく斜め（oblique）のため LPS 正規化を安全に実行できません。" & vbCrLf &
                "このファイルは斜め撮像データです。スキャナ座標系のまま使用してください。")
        End If

        ' -------------------------------------------------------
        ' Step 2: oblique 度チェック（警告のみ、処理は継続）
        '
        ' 支配的な世界軸成分の絶対値 vs オフダイアゴナル成分の合計を比較。
        ' 支配成分が全体の cos(30°) ≒ 0.866 未満なら oblique と判定。
        ' -------------------------------------------------------
        Const OBLIQUE_THRESHOLD As Double = 0.866
        Dim isOblique As Boolean = False
        For n As Integer = 0 To 2
            Dim w As Integer = dominantWorld(n)
            Dim norm As Double = Math.Sqrt(col(n, 0) ^ 2 + col(n, 1) ^ 2 + col(n, 2) ^ 2)
            If norm > 0 AndAlso Math.Abs(col(n, w)) / norm < OBLIQUE_THRESHOLD Then
                isOblique = True
            End If
        Next
        If isOblique Then
            Console.WriteLine("Warning: Oblique orientation detected (tilt > 30 degrees). " &
                              "LPS normalization is approximate.")
        End If

        ' -------------------------------------------------------
        ' Step 3: 軸並び替え（Permute）
        '
        ' LPS の目標は：
        '   ボクセル軸 i → 世界 X（L）
        '   ボクセル軸 j → 世界 Y（P）
        '   ボクセル軸 k → 世界 Z（S）
        ' dominantWorld が {0→0, 1→1, 2→2} でなければ並び替えが必要。
        ' -------------------------------------------------------

        ' dominantWorld の逆写像：世界軸 w を支配するボクセル軸を求める
        ' targetVoxelAxis(w) = 世界軸 w に対応するボクセル軸インデックス
        Dim targetVoxelAxis(2) As Integer
        For n As Integer = 0 To 2
            targetVoxelAxis(dominantWorld(n)) = n
        Next

        ' permOrder(新ボクセル軸) = 旧ボクセル軸
        ' 新ボクセル軸0（i）には世界X（0）を支配していた旧ボクセル軸を割り当てる
        Dim permOrder(2) As Integer
        For newAxis As Integer = 0 To 2
            permOrder(newAxis) = targetVoxelAxis(newAxis)
        Next

        ' 並び替えが必要かどうか確認
        Dim needPermute As Boolean = (permOrder(0) <> 0 OrElse permOrder(1) <> 1 OrElse permOrder(2) <> 2)
        If needPermute Then
            PermuteAxes(permOrder)   ' Pixel 配列と srow を並び替え
        End If

        ' -------------------------------------------------------
        ' Step 4: フリップ
        ' 並び替え後のアフィン行列対角成分の符号を確認し、
        ' 負なら反転して LPS 正方向に揃える。
        ' -------------------------------------------------------
        If SrowX(0) < 0 Then FlipAxisAndUpdateAffine(0)
        If SrowY(1) < 0 Then FlipAxisAndUpdateAffine(1)
        If SrowZ(2) < 0 Then FlipAxisAndUpdateAffine(2)

        ' qform を srow に整合させる
        AffineToQuaternion()
        QFormCode = 1
        SFormCode = 1

    End Sub

    ''' <summary>
    ''' Pixel 配列の軸を permOrder に従って並び替え、srow・pixdim も対応して更新する。
    ''' permOrder(新軸) = 旧軸 という定義。
    '''   例: permOrder = {2, 0, 1} → 新X軸←旧Z軸、新Y軸←旧X軸、新Z軸←旧Y軸
    '''
    ''' oldIdx の計算:
    '''   「新軸 newAxis のボクセルは旧軸 permOrder(newAxis) から来る」ため、
    '''   旧インデックス配列 oldIdx に対して
    '''     oldIdx(permOrder(newAxis)) = newIdx(newAxis)
    '''   と代入することで旧座標に変換する。
    ''' </summary>
    Private Sub PermuteAxes(permOrder() As Integer)

        ' 並び替え後の新しいマトリクスサイズ
        Dim oldMatrix(2) As Integer
        oldMatrix(0) = MatrixX
        oldMatrix(1) = MatrixY
        oldMatrix(2) = MatrixZ

        Dim newNX As Integer = oldMatrix(permOrder(0))
        Dim newNY As Integer = oldMatrix(permOrder(1))
        Dim newNZ As Integer = oldMatrix(permOrder(2))

        ' 新しい Pixel バッファを確保してコピー
        Dim newPixel(newNX - 1, newNY - 1, newNZ - 1) As Double

        ' ループ外で宣言することで、全ボクセル分のヒープ確保を回避する
        'Dim newIdx(2) As Integer
        Dim oldIdx(2) As Integer

        For newX As Integer = 0 To newNX - 1
            For newY As Integer = 0 To newNY - 1
                For newZ As Integer = 0 To newNZ - 1

                    'newIdx(0) = newX
                    'newIdx(1) = newY
                    'newIdx(2) = newZ

                    '' 新軸インデックスから旧軸インデックスに変換
                    '' permOrder(newAxis) = oldAxis なので：
                    ''   oldIdx(oldAxis) = newIdx(newAxis)
                    ''   すなわち oldIdx(permOrder(newAxis)) = newIdx(newAxis)
                    'For newAxis As Integer = 0 To 2
                    '    oldIdx(permOrder(newAxis)) = newIdx(newAxis)
                    'Next
                    oldIdx(permOrder(0)) = newX
                    oldIdx(permOrder(1)) = newY
                    oldIdx(permOrder(2)) = newZ
                    newPixel(newX, newY, newZ) = Pixel(oldIdx(0), oldIdx(1), oldIdx(2))
                Next
            Next
        Next

        Pixel = newPixel
        MatrixX = CShort(newNX)
        MatrixY = CShort(newNY)
        MatrixZ = CShort(newNZ)

        ' srow の列を並び替える（平行移動列 col=3 はそのまま）
        ' 並び替え前の列を退避
        Dim oldSrowX(3) As Single
        Dim oldSrowY(3) As Single
        Dim oldSrowZ(3) As Single
        Array.Copy(SrowX, oldSrowX, 4)
        Array.Copy(SrowY, oldSrowY, 4)
        Array.Copy(SrowZ, oldSrowZ, 4)

        ' 新しい列 newAxis には旧列 permOrder(newAxis) を割り当てる
        For newAxis As Integer = 0 To 2
            Dim oldAxis As Integer = permOrder(newAxis)
            SrowX(newAxis) = oldSrowX(oldAxis)
            SrowY(newAxis) = oldSrowY(oldAxis)
            SrowZ(newAxis) = oldSrowZ(oldAxis)
        Next
        ' 平行移動はそのまま
        SrowX(3) = oldSrowX(3)
        SrowY(3) = oldSrowY(3)
        SrowZ(3) = oldSrowZ(3)

        ' pixdim（ボクセルサイズ）も並び替える
        Dim oldSizes(2) As Single
        oldSizes(0) = SizeX
        oldSizes(1) = SizeY
        oldSizes(2) = SizeZ

        SizeX = oldSizes(permOrder(0))
        SizeY = oldSizes(permOrder(1))
        SizeZ = oldSizes(permOrder(2))

    End Sub
End Class

