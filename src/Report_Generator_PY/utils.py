
import pathlib


class Setting:
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
    def __init__(
        self,
        Template_Filepath: str = "",
        Templage_PageName: str = "",
        Target_Filepath: str = "",
        Target_PageName: str = "",
        Parameter_Filepath: str = "",
        PictureFolder: str = "",
    ) -> None:
        self.Template_Filepath = Template_Filepath
        self.Templage_PageName = Templage_PageName
        self.Target_Filepath = Target_Filepath
        self.Target_PageName = Target_PageName
        self.Parameter_Filepath = Parameter_Filepath
        self.PictureFolder = PictureFolder

    def GetImagePath(self, ImageName: str) -> pathlib.Path:
        if self.PictureFolder == "":
            print("PictureFolder is empty")
            return pathlib.Path("")
        res = pathlib.Path(self.PictureFolder, ImageName)
        # 如果副檔名不是jpg，則加上jpg
        if res.suffix != ".jpg":
            res = res.with_suffix(".jpg")
        return res

    def CleanPath(self):
        # 將//轉換成/
        self.Template_Filepath = self.Template_Filepath.replace("//", "/")
        self.Target_Filepath = self.Target_Filepath.replace("//", "/")
        self.Parameter_Filepath = self.Parameter_Filepath.replace("//", "/")
        self.PictureFolder = self.PictureFolder.replace("//", "/")