import clr
from typing import List
import System  # type: ignore
import os
from utils import *
clr.AddReference("Core\\Tools\\ReportGenerator\\Report_Generator")  # type: ignore

from Report_Generator import ReportGenerator_Excel  # type: ignore

Report_Generator_Excel_CSharp_Class = ReportGenerator_Excel()


def Report_Export_Single(TemplateReportPath: str, DataPath: str):
    # Parameter_Filepath = Datapath\FolderName.txt
    # target file path = Datapath\FolderName.xlsx

    S = Setting("", "", "", "", "", "")
    S.Template_Filepath = TemplateReportPath
    folder = DataPath
    S.Target_Filepath = os.path.join(folder, os.path.basename(folder) + ".xlsx")
    S.Templage_PageName = "Report_Template"
    S.Target_PageName = os.path.basename(folder)
    S.PictureFolder = os.path.join(folder, "Picture")
    S.Parameter_Filepath = os.path.join(folder, os.path.basename(folder) + ".txt")
    Do_Excel_Report(S)


def Do_Excel_Report(Setting: Setting):
    Setting.CleanPath()
    Header = Array_2D_2_CSharp_Format([[""]])
    # 如果Template_Filepath不存在，顯示並跳過
    if not System.IO.File.Exists(Setting.Template_Filepath):
        print("Template_Filepath not exist")
        return
    try:
        Report_Generator_Excel_CSharp_Class.Do_Report2(
            Setting.Template_Filepath,
            Setting.Templage_PageName,
            Setting.Target_Filepath,
            Setting.Target_PageName,
            Header,
            Setting.Parameter_Filepath,
            Setting.PictureFolder,
        )
    except Exception as e:
        print("Do_Report2 Error")
        print(e)


def Write_2_Excel(Setting: Setting, Data: List[List[str]], Row: int, Col: int):
    if not System.IO.File.Exists(Setting.Template_Filepath):
        print("Template_Filepath not exist")
        return
    # WriteToExcel(SourceFilePath,TargetFilePath,PageName,Params,StartRow,StartCol)
    D = Array_2D_2_CSharp_Format(Data)

    try:
        Report_Generator_Excel_CSharp_Class.WriteToExcel(
            Setting.Template_Filepath,
            Setting.Target_Filepath,
            Setting.Target_PageName,
            D,
            Row,
            Col,
        )
    except Exception as e:
        print("WriteToExcel Error")
        print(e)


#        public string WriteToExcel_Multi(string SourceFilePath, string TargetFilePath, string PageName, string[] Params, uint[] Target_Row, uint[] Target_Column)
def Write_2_Excel_Multi(Setting: Setting, Data: List[str], Row: List[int], Col: List[int]):
    # if not System.IO.File.Exists(Setting.Template_Filepath):
    #     print("Template_Filepath not exist")
    #     return
    # WriteToExcel(SourceFilePath,TargetFilePath,PageName,Params,StartRow,StartCol)
    D = System.Array.CreateInstance(System.String, len(Data))
    for i in range(len(Data)):
        D[i] = str(Data[i])

    R = System.Array.CreateInstance(System.UInt32, len(Row))
    for i in range(len(Row)):
        R[i] = Row[i]

    C = System.Array.CreateInstance(System.UInt32, len(Col))
    for i in range(len(Col)):
        C[i] = Col[i]

    try:
        Report_Generator_Excel_CSharp_Class.WriteToExcel_Multi(
            str(Setting.Template_Filepath),
            str(Setting.Target_Filepath),
            Setting.Target_PageName,
            D,
            R,
            C,
        )
    except Exception as e:
        print("WriteToExcel_Multi Error")
        print(e)


def Array_2D_2_CSharp_Format(Param: List[List[str]]):
    rows = len(Param)
    cols = len(Param[0])
    csharp_array = System.Array.CreateInstance(System.String, rows, cols)
    for i in range(rows):
        for j in range(cols):
            csharp_array[i, j] = Param[i][j]
    return csharp_array


# 還沒執行成功

# reportGeneratorExcel.Do_Report2.argtypes = [ctypes.c_char_p, ctypes.c_char_p, ctypes.c_char_p, ctypes.c_char_p, ctypes.POINTER(ctypes.c_char_p), ctypes.c_char_p, ctypes.c_char_p]
# reportGeneratorExcel.Do_Report2.restype = ctypes.c_char_p


# def Do_Report_Excel(Template_Filepath: str, Templage_PageName: str, Target_Filepath: str, Target_PageName: str, Header: List[List[str]], Parameter_Filepath: str, PictureFolder: str):
# reportGeneratorExcel.Do_Report2(Template_Filepath, Templage_PageName, Target_Filepath, Target_PageName, Header, Parameter_Filepath, PictureFolder)
