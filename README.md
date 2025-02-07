# Report_Generator_PY
這個是python版本的report gen, 但核心的DLL還是用C#完成
report generator for python version, but dll is construct from report gen C# project

## API Documentation

### Setting Class

**Parameters:**
- `Template_Filepath` (str): The path of the template file, should be a xlsx or docx file
- `Templage_PageName` (str): The name of the page in the template file (xlsx only)
- `Target_Filepath` (str): The path of the target file, should be a xlsx or docx file
- `Target_PageName` (str): The name of the page in the target file (xlsx only)
- `Parameter_Filepath` (str): The path of the parameter file, should be a txt file
- `PictureFolder` (str): The folder path of the pictures

**Functions:**
- `GetImagePath(str)`: Get the path of the image file
- `CleanPath()`: Clean the path of the files, replace // with /

### Report_Export_Single

This function is used to export a single report and replace the final report with the template file. It should have a picture folder, parameter file in the data folder.

**Parameters:**
- `TemplateReportPath` (str): The path of the template file, should be a xlsx or docx file
- `DataPath` (str): The path of the data folder

**Return:**
- None

**Example:**
```python
Report_Export_Single("template/template.xlsx", "data")
```

### Do_Excel_Report

This function is used to export a report (recommended).

**Parameters:**
- `Setting` (Setting): The setting of the report

**Return:**
- None

**Example:**
```python
S = Setting("", "", "", "", "", "")
S.Template_Filepath = TemplateReportPath
folder = DataPath
S.Target_Filepath = os.path.join(folder, os.path.basename(folder) + ".xlsx")
S.Templage_PageName = "Report_Template"
S.Target_PageName = os.path.basename(folder)
S.PictureFolder = os.path.join(folder, "Picture")
S.Parameter_Filepath = os.path.join(folder, os.path.basename(folder) + ".txt")
Do_Excel_Report(S)
```

### Write_2_Excel

This function is used to write data to the target file. You can input a location to write the 2D data.

**Parameters:**
- `Setting` (Setting): The setting of the report, just need to set the Template_Filepath, Target_Filepath, Target_PageName
- `Data` (List[List[str]]): The data to write
- `Row` (int): The start row
- `Col` (int): The start column

**Return:**
- None

**Example:**
```python
S = Setting("", "", "", "", "", "")
Write_2_Excel(S, [["1","2","3"],["4","5","6"]], 1, 1)
```

### Write_2_Excel_Multi

This function is used to write data to the target file. You can input multiple locations to write the multiple data. One data for one location, so the length of the data should be the same as the length of the row and col.

**Parameters:**
- `Setting` (Setting): The setting of the report, just need to set the Template_Filepath, Target_Filepath, Target_PageName
- `Data` (List[List[str]]): The data to write
- `Row` (List[int]): The start row
- `Col` (List[int]): The start column

**Return:**
- None

**Example:**
```python
S = Setting("", "", "", "", "", "")
Write_2_Excel_Multi(S, ["1","2","3"], [1,2,3], [1,2,3])
