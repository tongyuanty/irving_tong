Public Class Form1
    Shared sPath = ""
    '读ini API函数
    Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Int32, ByVal lpFileName As String) As Int32
    Const ESL_SUCCESS = &H0 '  Operation succeeded
    Const ESL_CANCEL = &H1 '  Operation is canceled
    Const ESL_ERR_GENERAL = &H80000001 '  General error
    Const ESL_ERR_NOT_INITIALIZED = &H80000002 '  Library is not initialized =Internal error '
    Const ESL_ERR_FILE_MISSING = &H80000003 '  Required file is missing
    Const ESL_ERR_INVALID_PARAM = &H80000004 '  Invalid parameters are set
    Const ESL_ERR_LOW_MEMORY = &H80000005 '  Not enough memory to operate
    Const ESL_ERR_LOW_DISKSPACE = &H80000006 '  Not enough free disk space to operate
    Const ESL_ERR_WRITE_FAIL = &H80000007 '  Failed to write to disk
    Const ESL_ERR_READ_FAIL = &H80000008 '  Failed to read from disk
    Const ESL_ERR_INVALID_KEY = &H80010001 '  License key is invalid
    Const ESL_ERR_NOT_SUPPORTED = &H80020001 '  Specified model is not supported
    Const ESL_ERR_NO_DRIVER = &H80020002 '  Scanner driver for specified model is not installed
    Const ESL_ERR_OPEN_FAIL = &H80020003 '  Failed to open scanner driver
    Const ESL_ERR_SCAN_GENERAL = &H80030001 '  Scanning operation failed
    '读取ini文件内容
    Public Shared Function GetINI(ByVal Section As String, ByVal AppName As String, ByVal lpDefault As String, ByVal FileName As String) As String
        Dim Str As String = LSet(Str, 256)
        GetPrivateProfileString(Section, AppName, lpDefault, Str, Len(Str), FileName)
        Return Microsoft.VisualBasic.Left(Str, InStr(Str, Chr(0)) - 1)
    End Function

    Public Shared Function App_Path() As String
        Return System.AppDomain.CurrentDomain.BaseDirectory()
    End Function

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        ' Scan()

    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Public Sub XOpenScan(ByVal dictPara As KFO.Dictionary)
        Dim dictCfg As Object

        sPath = App_Path() & "ScannerSetup.ini"

        dictCfg = CreateObject("KFO.Dictionary")

        dictCfg("docSource") = GetINI("ScanConfig", "docSource", "-1", sPath)
        dictCfg("docSize") = GetINI("ScanConfig", "docSize", "-1", sPath)
        dictCfg("skewCorrect") = GetINI("ScanConfig", "skewCorrect", "1", sPath)
        dictCfg("docRotate") = GetINI("ScanConfig", "docRotate", "2", sPath)
        dictCfg("imgType") = GetINI("ScanConfig", "imgType", "0", sPath)
        dictCfg("resolution") = GetINI("ScanConfig", "resolution", "200", sPath)
        dictCfg("docEnhance") = GetINI("ScanConfig", "docEnhance", "0", sPath)
        dictCfg("optBlankPageSkip") = GetINI("ScanConfig", "optBlankPageSkip", "0", sPath)

        dictCfg("sizeUnit") = GetINI("ScanConfig", "sizeUnit", "1", sPath)
        dictCfg("sizeRight") = GetINI("ScanConfig", "sizeRight", "0", sPath)
        dictCfg("sizeBottom") = GetINI("ScanConfig", "sizeBottom", "0", sPath)
        dictCfg("imgFilter") = GetINI("ScanConfig", "imgFilter", "0", sPath)
        dictCfg("brightness") = GetINI("ScanConfig", "brightness", "0", sPath)
        dictCfg("contrast") = GetINI("ScanConfig", "contrast", "0", sPath)

        With Me
            .AxEScanControl1.LicenseKey = "SSCC0100_141020140017D6"
            ' .AxEScanControl1.LicenseKey = "TEST_VERSION_ONLY"
            .AxEScanControl1.EParamDeviceInfo.DeviceName = dictPara("DeviceName")
            .AxEScanControl1.EParamScan.DocSource = dictCfg("docSource")
            .AxEScanControl1.EParamScan.DocSize = dictCfg("docSize")
            .AxEScanControl1.EParamScan.SkewCorrect = dictCfg("skewCorrect")
            .AxEScanControl1.EParamScan.DocRotate = dictCfg("docRotate")
            .AxEScanControl1.EParamScan.ImageType = dictCfg("imgType")
            .AxEScanControl1.EParamScan.Resolution = dictCfg("resolution")
            .AxEScanControl1.EParamScan.DocEnhance = dictCfg("docEnhance")
            .AxEScanControl1.EParamScan.BlankPageSkip = dictCfg("optBlankPageSkip")
            .AxEScanControl1.EParamScan.DocSizeUnit = dictCfg("sizeUnit")
            .AxEScanControl1.EParamScan.DocSizeUserLeft = 0
            .AxEScanControl1.EParamScan.DocSizeUserTop = 0
            .AxEScanControl1.EParamScan.DocSizeUserRight = dictCfg("sizeRight")
            .AxEScanControl1.EParamScan.DocSizeUserBottom = dictCfg("sizeBottom")
            .AxEScanControl1.EParamScan.ImageFilter = dictCfg("imgFilter")
            .AxEScanControl1.EParamScan.Brightness = dictCfg("brightness")
            .AxEScanControl1.EParamScan.Contrast = dictCfg("contrast")

            Select Case UCase(dictPara("FileFormat"))
                Case ".BMP"
                    .AxEScanControl1.EParamSave.FileFormat = 0
                Case ".JPG"
                    .AxEScanControl1.EParamSave.FileFormat = 1
                Case Else
                    .AxEScanControl1.EParamSave.FileFormat = 0
            End Select
            .AxEScanControl1.EParamScan.NumScan = dictPara("NumScan")
            .AxEScanControl1.EParamSave.FilePath = dictPara("FilePath")
            .AxEScanControl1.EParamSave.FileNameType = 0 '文件名使用前缀
            .AxEScanControl1.EParamSave.FileNamePrefix = dictPara("FileNamePrefix") '前缀
            .AxEScanControl1.EParamSave.FileNameTimeFormat = 0 '不使用日期时间前缀
            .AxEScanControl1.EParamSave.FileNameNumCounterDigits = 1 '编号位数
            .AxEScanControl1.EParamSave.FileNameStartCount = dictPara("StartCount") '起始编号
            .Hide()
        End With
    End Sub

    Private Shared obj = 1

    Public Function Scan(ByVal dictPara As KFO.Dictionary)
        Dim dictRet As Object
        Dim ret As UInteger
        Dim dictCfg As Object
        dictRet = New KFO.Dictionary


        sPath = App_Path() & "ScannerSetup.ini"

        dictCfg = CreateObject("KFO.Dictionary")

        dictCfg("docSource") = GetINI("ScanConfig", "docSource", "-1", sPath)
        dictCfg("docSize") = GetINI("ScanConfig", "docSize", "-1", sPath)
        dictCfg("skewCorrect") = GetINI("ScanConfig", "skewCorrect", "1", sPath)
        dictCfg("docRotate") = GetINI("ScanConfig", "docRotate", "2", sPath)
        dictCfg("imgType") = GetINI("ScanConfig", "imgType", "0", sPath)
        dictCfg("resolution") = GetINI("ScanConfig", "resolution", "200", sPath)
        dictCfg("docEnhance") = GetINI("ScanConfig", "docEnhance", "0", sPath)
        dictCfg("optBlankPageSkip") = GetINI("ScanConfig", "optBlankPageSkip", "0", sPath)

        dictCfg("sizeUnit") = GetINI("ScanConfig", "sizeUnit", "1", sPath)
        dictCfg("sizeRight") = GetINI("ScanConfig", "sizeRight", "0", sPath)
        dictCfg("sizeBottom") = GetINI("ScanConfig", "sizeBottom", "0", sPath)
        dictCfg("imgFilter") = GetINI("ScanConfig", "imgFilter", "0", sPath)
        dictCfg("brightness") = GetINI("ScanConfig", "brightness", "0", sPath)
        dictCfg("contrast") = GetINI("ScanConfig", "contrast", "0", sPath)

        With Me
            .AxEScanControl1.LicenseKey = "SSCC0100_141020140017D6"
            ' .AxEScanControl1.LicenseKey = "TEST_VERSION_ONLY"
            .AxEScanControl1.EParamDeviceInfo.DeviceName = dictPara("DeviceName")
            .AxEScanControl1.EParamScan.DocSource = dictCfg("docSource")
            .AxEScanControl1.EParamScan.DocSize = dictCfg("docSize")
            .AxEScanControl1.EParamScan.SkewCorrect = dictCfg("skewCorrect")
            .AxEScanControl1.EParamScan.DocRotate = dictCfg("docRotate")
            .AxEScanControl1.EParamScan.ImageType = dictCfg("imgType")
            .AxEScanControl1.EParamScan.Resolution = dictCfg("resolution")
            .AxEScanControl1.EParamScan.DocEnhance = dictCfg("docEnhance")
            .AxEScanControl1.EParamScan.BlankPageSkip = dictCfg("optBlankPageSkip")
            .AxEScanControl1.EParamScan.DocSizeUnit = dictCfg("sizeUnit")
            .AxEScanControl1.EParamScan.DocSizeUserLeft = 0
            .AxEScanControl1.EParamScan.DocSizeUserTop = 0
            .AxEScanControl1.EParamScan.DocSizeUserRight = dictCfg("sizeRight")
            .AxEScanControl1.EParamScan.DocSizeUserBottom = dictCfg("sizeBottom")
            .AxEScanControl1.EParamScan.ImageFilter = dictCfg("imgFilter")
            .AxEScanControl1.EParamScan.Brightness = dictCfg("brightness")
            .AxEScanControl1.EParamScan.Contrast = dictCfg("contrast")

            Select Case UCase(dictPara("FileFormat"))
                Case ".BMP"
                    .AxEScanControl1.EParamSave.FileFormat = 0
                Case ".JPG"
                    .AxEScanControl1.EParamSave.FileFormat = 1
                Case Else
                    .AxEScanControl1.EParamSave.FileFormat = 0
            End Select
            .AxEScanControl1.EParamScan.NumScan = dictPara("NumScan")
            .AxEScanControl1.EParamSave.FilePath = dictPara("FilePath")
            .AxEScanControl1.EParamSave.FileNameType = 0 '文件名使用前缀
            .AxEScanControl1.EParamSave.FileNamePrefix = dictPara("FileNamePrefix") '前缀
            .AxEScanControl1.EParamSave.FileNameTimeFormat = 0 '不使用日期时间前缀
            .AxEScanControl1.EParamSave.FileNameNumCounterDigits = 1 '编号位数
            .AxEScanControl1.EParamSave.FileNameStartCount = dictPara("StartCount") '起始编号
            .Hide()
        End With

        ' On Error GoTo ErrHandle

       
        AxEScanControl1.OpenScanner()

        ret = AxEScanControl1.Execute(0)
        Dim i = 0
        While (obj = 0 And i < 60)
            Threading.Thread.Sleep(100)
            i = i + 1
        End While
        dictRet("obj") = obj
        obj = 0

        If ret = &H0 Then
            dictRet("isSuccess") = 1
            'Dim fcount As Long
            'fcount = System.IO.Directory.GetFiles(dictPara("FilePath"), "*", System.IO.SearchOption.TopDirectoryOnly).Length()
            'dictRet("NumScan") = fcount
        Else
            dictRet("isSuccess") = 0
            dictRet("NumScan") = 0

            dictRet("ErrDesc") = GetErrDesc(ret)
        End If

        
        Return dictRet
    End Function

    Private Function GetErrDesc(ByVal ErrNo As Integer) As String
        Dim ErrDesc As String
        '' '' ''厂商给的API文档中描述的返回值与实际测试不符,
        '' '' ''例如不放纸扫描失败时不会返回ret,而是直接中断,
        '' '' ''只能通过ErrHandle来处理,并且此时ret依然是0,
        '' '' ''无法区分错误原因,只能归到Case Else
        Select Case ErrNo

            'Case ESL_SUCCESS '=&H0 
            '    ErrDesc = "Operation succeeded"
            'Case ESL_CANCEL '=&H1 
            '    ErrDesc = "Operation is canceled"
            Case ESL_ERR_GENERAL '=&H80000001 
                ErrDesc = "General error"
            Case ESL_ERR_NOT_INITIALIZED '=&H80000002 
                ErrDesc = "Library is not initialized (Internal error) "
            Case ESL_ERR_FILE_MISSING '=&H80000003 
                ErrDesc = "Required file is missing"
            Case ESL_ERR_INVALID_PARAM '=&H80000004 
                ErrDesc = "Invalid parameters are set"
            Case ESL_ERR_LOW_MEMORY '=&H80000005 
                ErrDesc = "Not enough memory to operate"
            Case ESL_ERR_LOW_DISKSPACE '=&H80000006 
                ErrDesc = "Not enough free disk space to operate"
            Case ESL_ERR_WRITE_FAIL '=&H80000007 
                ErrDesc = "Failed to write to disk"
            Case ESL_ERR_READ_FAIL '=&H80000008 
                ErrDesc = "Failed to read from disk"
            Case ESL_ERR_INVALID_KEY '=&H80010001 
                ErrDesc = "License key is invalid"
            Case ESL_ERR_NOT_SUPPORTED '=&H80020001 
                ErrDesc = "Specified model is not supported"
            Case ESL_ERR_NO_DRIVER '=&H80020002 
                ErrDesc = "Scanner driver for specified model is not installed"
            Case ESL_ERR_OPEN_FAIL '=&H80020003 
                ErrDesc = "Failed to open scanner driver"
            Case ESL_ERR_SCAN_GENERAL '=&H80030001 
                ErrDesc = "Scanning operation failed"
            Case Else
                ErrDesc = "请确认选择了正确的扫描仪型号并且已经在自动文稿送纸器中放入了文稿。"
        End Select

        Return ErrDesc
    End Function

    Private Sub AxEScanControl1_OnCompleted(ByVal sender As System.Object, ByVal e As AxESCANOCX2Lib._IEScanControl2Events_OnCompletedEvent) Handles AxEScanControl1.OnCompleted
        obj = 1
        AxEScanControl1.CloseScanner()
    End Sub

End Class
