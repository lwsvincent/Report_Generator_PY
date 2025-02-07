from .utils import Setting
from .ReportGenerator import Report_Export_Single, Do_Excel_Report, Write_2_Excel,Write_2_Excel_Multi

    ###
    # Setting Class
    # Parameters:
    # Template_Filepath(str): The path of the template file, should be a xlsx or docx file
    # Templage_PageName(str): The name of the page in the template file (xlsx only)
    # Target_Filepath(str): The path of the target file, should be a xlsx or docx file
    # Target_PageName(str): The name of the page in the target file (xlsx only)
    # Parameter_Filepath(str): The path of the parameter file, should be a txt file
    # PictureFolder(str): The folder path of the pictures
    
    # Functions:
    # GetImagePath(str): Get the path of the image file
    # CleanPath(): Clean the path of the files, replace // with /
    
    # ###
    
    
    
    ### Report_Export_Single 
    # This function is used to export a single report and replace the final report with the template file
    # It should have a picture folder, paramter file in the data folder
    #
    # call function:
    # Report_Export_Single("template/template.xlsx", "data")
    
    
    # Parameters:
    # TemplateReportPath(str): The path of the template file, should be a xlsx or docx file
    # DataPath(str): The path of the data folder
    
    # Return:
    # None
    ###
    
    ### Do_Excel_Report 
    # This function is used to export a report (recommanded)
    
    # Parameters:
    # Setting(Setting): The setting of the report
    
    # Return:
    # None
    
    #e.g.
    # S = Setting("", "", "", "", "", "")
    # S.Template_Filepath = TemplateReportPath
    # folder = DataPath
    # S.Target_Filepath = os.path.join(folder, os.path.basename(folder) + ".xlsx")
    # S.Templage_PageName = "Report_Template"
    # S.Target_PageName = os.path.basename(folder)
    # S.PictureFolder = os.path.join(folder, "Picture")
    # S.Parameter_Filepath = os.path.join(folder, os.path.basename(folder) + ".txt")
    # Do_Excel_Report(S)
    ###
    
    ### Write_2_Excel
    # This function is used to write data to the target file
    # you can input a location to write the 2D data
    
    # Parameters:
    # Setting(Setting): The setting of the report, just need to set the Template_Filepath, Target_Filepath, Target_PageName
    # Data(List[List[str]]): The data to write
    # Row(int): The start row
    # Col(int): The start column
    
    # Return:
    # None
    
    #e.g.
    # S = Setting("", "", "", "", "", "")
    # Write_2_Excel(S, [["1","2","3"],["4","5","6"]], 1, 1)
    
    # ###
    
    ### Write_2_Excel_Multi 
    # This function is used to write data to the target file
    # you can input multiple location to write the multiple data
    # one data for one location
    # so the length of the data should be the same as the length of the row and col
    
    # Parameters:
    # Setting(Setting): The setting of the report, just need to set the Template_Filepath, Target_Filepath, Target_PageName
    # Data(List[List[str]]): The data to write
    # Row(List[int]): The start row
    # Col(List[int]): The start column
    
    # Return:
    # None
    # 
    
    #e.g.
    # S = Setting("", "", "", "", "", "")
    # Write_2_Excel_Multi(S, ["1","2","3"], [1,2,3], [1,2,3])
    # ###