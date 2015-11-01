Attribute VB_Name = "RLaserV5"
'***************************************************************************************************
'                                    Reader (Shenzhen) Ltd.
'
'                        (c) Coyright 2009, Shen Zhen Reader Ltd.,
'
'* All rights reserved. Reader's  source code is an unpublished work and the use of a copyright
'* notice does not imply otherwise. This source code contains confidential, trade secret material of
'* Shen Zhen Reader Ltd. Any attempt or participation in deciphering, decoding, reverse engineering
'* or in any way altering the source code is strictly prohibited, unless the prior written consent
'* of Shen Zhen Reader Ltd. is obtained.
'*
'* Filename:       RLaser.bas
'* Programmer:     Chen Zhi Xin
'* Created:        09-11-10
'* Description:    Execute RLaser.exe file with starting coreldraw.exe file and release ai/dxf/plt file
'* Note:
'*
'***************************************************************************************************

Private Type VariableBuffer
   VariableParameter() As Byte
End Type

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function LoadModule Lib "kernel32.dll" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeModule Lib "kernel32.dll" Alias "FreeLibrary" (ByVal hLibModule As Long) As Boolean
Private Declare Function GetModuleEntry Lib "kernel32.dll" Alias "GetProcAddress" (ByVal hModule As Long, ByRef lpProcName As String) As Long
Private Declare Function StrCat Lib "kernel32.dll" Alias "lstrcatW" (ByVal lpstring1 As String, ByVal lpstring2 As String)
Private Declare Function StrCpy Lib "kernel32.dll" Alias "lstrcpy" (ByVal lpstring1 As String, ByVal lpstring2 As String)
Private Declare Function CallWindowProc Lib "User32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDest As Any, lpSource As Any, ByVal cBytes As Long)
Private Declare Function SaveEmbroiderData Lib "drIptBdrFe.dll" (ByVal lpFileName As String) As Boolean
Private Declare Function GetLang Lib "rdloadV5.dll" () As Integer
Private Declare Function RunProgram Lib "rdloadV5.dll" () As Integer


Dim VectorFilterType As Integer
Private m_opIndex As Long
Private m_OpCode() As Byte
Dim ret As Long
Dim VectorFiltType As Integer
Dim ProgramType As Integer
Dim AIPrecision As Double
Dim DxfPrecision As Double
Dim DxfEnOptText As Integer
Dim DelReapt As Integer
Dim VelUnit As Integer
Dim nLanuage As Integer

Private Function GetProgPath() As String
    GetProgPath = Application.SetupPath + "Programs\"
End Function
Sub LaserWorking()
    '================================================================================================
    '判断CorelDraw是否有数据，若无数据则直接打开LaserWork应用程序
    If ActiveDocument Is Nothing Then
         RunProgram
        Exit Sub
    End If

    If ActivePage Is Nothing Then
         RunProgram
        Exit Sub
    End If

    If ActivePage.Shapes.Count <= 0 Then
         RunProgram
        Exit Sub
    End If
   
    ActiveDocument.Unit = cdrMillimeter
    ActiveDocument.ReferencePoint = cdrBottomLeft
    'ActivePage.Shapes.All.UngroupAll
        
    '===================================================================================================
    '获取程序的当前路径
    '检查保存参数的文件夹是否存在，若不存在则创建文件夹
    Dim ProgPath As String
    Dim fileName As String
    
    ProgPath = GetProgPath
    fileName = ProgPath + "LaserWorkV5\temp"
    If 0 = DirExists(fileName) Then
        MkDir fileName
    End If
    
    '===================================================================================================
    Set srAll = ActivePage.Shapes.All
    Set srBmp = ActivePage.Shapes.FindShapes(Type:=cdrBitmapShape)
    Set srText = ActivePage.Shapes.FindShapes(Type:=cdrTextShape)
    
    If srBmp.Count <> 0 Then

        srAll.RemoveRange srBmp

    End If
    
    If srText.Count <> 0 Then
    
        For Each s In srText
            If s.Outline.Type = cdrNoOutline Then
                s.Outline.Width = 0.0762
                
                If s.Fill.Type = cdrUniformFill Then '有颜色填充
                    s.Outline.Color = s.Fill.UniformColor
                End If
                
            End If
            
            s.Fill.ApplyNoFill
            
        Next s
    
        srText.ConvertToCurves
    End If
    
    Dim nBmp As Integer
    Dim PositionX As Double
    Dim PositionY As Double
    Dim Width As Double
    Dim Height As Double
    
    Dim expflt As ExportFilter
    '===================================================================================================
    '判断当前图形中是否包含位图,若不包含则直接输出矢量图形
    nBmp = srBmp.Count
    If nBmp = 0 Then
        fileName = ProgPath + "LaserWorkV5\temp\dest.tmp"
        Open fileName For Binary Access Write As #1
        Put #1, , 0
        Put #1, , 0
        Close #1
        
        ActiveDocument.CreateSelection srAll
        fileName = ProgPath + "LaserWorkV5\temp\dest2.tmp"
        Width = srAll.SizeWidth
        Height = srAll.SizeHeight
        PositionX = srAll.PositionX
        PositionY = srAll.PositionY
        Open fileName For Binary Access Write As #1
        Put #1, , Width
        Put #1, , Height
        Put #1, , PositionX
        Put #1, , PositionY
        Close #1
            
        fileName = ProgPath + "LaserWorkV5\temp\RD.ai"
        Set expflt = ActiveDocument.ExportEx(fileName, cdrAI, cdrSelection)
             With expflt
              .Version = 0
              .TextAsCurves = True
              .Platform = 0
              .ConvertSpotColors = False
              .UseColorProfile = False
              .SimulateOutlines = False
              .SimulateFills = False
              .IncludePlacedImages = False
              .IncludePreview = False
              .Finish
            End With
    
        ActiveDocument.ClearSelection
        RunProgram
        Exit Sub
    End If
    
    fileName = ProgPath + "LaserWorkV5\temp\dest.tmp"
    Open fileName For Binary Access Write As #1
    Put #1, , 0
    Put #1, , nBmp
    
    Dim PixelWidth  As Long
    Dim PixelHeight As Long
    Dim ResolutionX As Long
    Dim ResolutionY As Long
    
    Dim CurShape As Shape
    Dim ShapeType As Integer
    
    nBmp = 0
    For Each CurShape In srBmp
        ActiveDocument.ClearSelection
        ShapeType = CurShape.Type
            
        If cdrBitmapShape = ShapeType Then
        CurShape.Selected = True
                
        CurShape.GetPosition PositionX, PositionY
        CurShape.GetSize Width, Height
                
        fileName = ProgPath + "LaserWorkV5\temp\" + Format(nBmp + 1, y) + ".bmp"
                
        PixelWidth = CurShape.Bitmap.SizeWidth
        PixelHeight = CurShape.Bitmap.SizeHeight
        ResolutionX = CurShape.Bitmap.ResolutionX
        ResolutionY = CurShape.Bitmap.ResolutionY
                
        Dim eFilt As ExportFilter
        Set eFilt = ActiveDocument.ExportBitmap(fileName, cdrBMP, cdrSelection, cdrGrayscaleImage, _
                                                PixelWidth, PixelHeight, ResolutionX, ResolutionY, cdrNomalAntiAliasing, _
                                                False, False, False)
                
        eFilt.Finish
                
        PositionX = PositionX + Width / 2
        PositionY = PositionY + Height / 2
        Put #1, , Width
        Put #1, , Height
        Put #1, , PositionX
        Put #1, , PositionY
                
        nBmp = nBmp + 1
        End If
    Next CurShape
    
    Close #1
    
    '=======================================================再输出矢量图============================================================================
    ActiveDocument.CreateSelection srAll
    
    Dim Xmin As Double
    Dim Xmax As Double
    Dim Ymin As Double
    Dim Ymax As Double
    Xmin = Xmax = Ymin = Ymax = 0
    
    Dim VectCount As Long
    VectCount = 0
    
    '统计矢量图形个数及矢量图形的边框====================
    For Each CurShape In srAll
        If cdrBitmapShape <> CurShape.Type Then
            VectCount = VectCount + 1
            
            CurShape.GetSize Width, Height
            CurShape.GetPosition PositionX, PositionY
            
            If 1 = VectCount Then
                Xmin = PositionX
                Xmax = Xmin + Width
                Ymin = PositionY
                Ymax = Ymin + Height
            Else
                If Xmin > PositionX Then
                    Xmin = PositionX
                End If
                    
                If Xmax < (PositionX + Width) Then
                    Xmax = PositionX + Width
                End If
                
                If Ymin > PositionY Then
                    Ymin = PositionY
                End If
                
                If Ymax < (PositionY + Height) Then
                    Ymax = PositionY + Height
                End If
            End If
            
            bHaveFile = True
        Else
            CurShape.Selected = False
        End If
    Next CurShape
    
    If VectCount > 0 Then
       '输出矢量图形边框数据==========================================
       fileName = ProgPath + "LaserWorkV5\temp\dest2.tmp"
       Width = Abs(Xmax - Xmin)
       Height = Abs(Ymax - Ymin)
       PositionX = (Xmax + Xmin) / 2
       PositionY = (Ymax + Ymin) / 2
       
       Open fileName For Binary Access Write As #1
       Put #1, , Width
       Put #1, , Height
       Put #1, , PositionX
       Put #1, , PositionY
       Close #1
       
       
        fileName = ProgPath + "LaserWorkV5\temp\RD.ai"
        Set expflt = ActiveDocument.ExportEx(fileName, cdrAI, cdrSelection)
             With expflt
              .Version = 0
              .TextAsCurves = True
              .Platform = 0
              .ConvertSpotColors = False
              .UseColorProfile = False
              .SimulateOutlines = False
              .SimulateFills = False
              .IncludePlacedImages = False
              .IncludePreview = False
              .Finish
            End With
    End If
    ActiveDocument.ClearSelection
    RunProgram
End Sub

Sub ImportDstFile()
    Dim strTemp As String
    Dim strCurPath As String
    Dim rtn As Long
    Dim im As ImportFilter
    Dim sio As StructImportOptions
   
    strTemp = Application.UserDataPath + "CorelEmb.ai"
    If ActiveLayer Is Nothing Then
    Else
        rtn = SaveEmbroiderData(strTemp)
        If rtn > 0 Then
            Set im = ActiveLayer.ImportEx(strTemp, cdrAI)
            im.Finish
            
            ActiveDocument.SelectionRange.SetOutlineProperties 0.001, , CreateRGBColor(0, 0, 0)
            
        End If
    End If
End Sub

Sub UserInit()
   Dim i As Long
   Dim bInstallRD As Boolean
   Dim strRunning As String, strImport As String
   
   bInstallRD = False
   
   For i = 1 To CommandBars.Count
        With CommandBars(i)
             If .Name = "RLaserCut5.0" Then
                bInstallRD = True
                Exit For
             End If
        End With
  Next
  
    Dim fileName As String
    Dim ProgPath As String
    ProgPath = GetProgPath
          
    nLanuage = GetLang()
    If 0 = nLanuage Then '简体中文
        strRunning = "激光加工"
        strImport = "导入Dst/Dsb数据"
    ElseIf 1 = nLanuage Then '繁体中文
        strRunning = "激光加工"
        strImport = "導入Dst/Dsb數據"
    Else '英文
        strRunning = "Laser Running"
        strImport = "Import Dst/Dsb Data"
    End If

  If bInstallRD = False Then
  
    Application.FrameWork.CommandBars.Add "RLaserCut5.0"
    Application.FrameWork.CommandBars("RLaserCut5.0").Enabled = True
    Application.FrameWork.CommandBars("RLaserCut5.0").Visible = True
    
    fileName = ProgPath + "LaserWorkV5\icon\runV5.bmp"
    With Application.FrameWork.CommandBars("RLaserCut5.0").Controls.AddCustomButton("2cc24a3e-fe24-4708-9a74-9c75406eebcd", "GlobalMacros.RLaserV5.LaserWorking")
                .SetCustomIcon (fileName)
    End With
    
     
    fileName = ProgPath + "LaserWorkV5\icon\importV5.bmp"
    With Application.FrameWork.CommandBars("RLaserCut5.0").Controls.AddCustomButton("2cc24a3e-fe24-4708-9a74-9c75406eebcd", "GlobalMacros.RLaserV5.ImportDstFile")
                .SetCustomIcon (fileName)
    End With
  
     
  End If
  
  Application.FrameWork.CommandBars("RLaserCut5.0").Controls(1).ToolTipText = strRunning
  Application.FrameWork.CommandBars("RLaserCut5.0").Controls(1).DescriptionText = ""
  
  Application.FrameWork.CommandBars("RLaserCut5.0").Controls(2).ToolTipText = strImport
  Application.FrameWork.CommandBars("RLaserCut5.0").Controls(2).DescriptionText = ""
  
End Sub

Public Function DirExists(ByVal strDirName As String) As Integer
    Const strWILDCARD$ = "*.*"
        
    Dim strDummy     As String

    On Error Resume Next
    If Trim(strDirName) = "" Then
        DirExists = 0
        Exit Function
    End If
    
    'AddDirStep strDirName
    strDummy = Dir$(strDirName & strWILDCARD, vbDirectory)
    DirExists = Not (strDummy = vbNullString)
    
    Err = 0
End Function






