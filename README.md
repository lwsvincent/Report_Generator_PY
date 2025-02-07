# Report_Generator_PY
這個是python版本的report gen, 但核心的DLL還是用C#完成
report generator for python version, but dll is construct from report gen C# project

# Report_Generator_PY
Report generator for Python version, but DLL is constructed from the report gen C# project.

### Setting Class

**Parameters:**
- `Template_Filepath` (str): The path of the template file, should be a xlsx or docx file.
- `Templage_PageName` (str): The name of the page in the template file (xlsx only).
- `Target_Filepath` (str): The path of the target file, should be a xlsx or docx file.
- `Target_PageName` (str): The name of the page in the target file (xlsx only).
- `Parameter_Filepath` (str): The path of the parameter file, should be a txt file.
- `PictureFolder` (str): The folder path of the pictures.

**Functions:**
- `GetImagePath(str)`: Get the path of the image file.
- `CleanPath()`: Clean the path of the files, replace `//` with `/`.

### Report_Export_Single

This function is used to export a single report and replace the final report with the template file. It should have a picture folder and parameter file in the data folder.

**Example:**