import clr, os, winreg
from itertools import islice
import matplotlib.pyplot as plt
import numpy as np

# This boilerplate requires the 'pythonnet' module.
# The following instructions are for installing the 'pythonnet' module via pip:
#    1. Ensure you are running Python 3.4, 3.5, 3.6, or 3.7. PythonNET does not work with Python 3.8 yet.
#    2. Install 'pythonnet' from pip via a command prompt (type 'cmd' from the start menu or press Windows + R and type 'cmd' then enter)
#
#        python -m pip install pythonnet

class PythonStandaloneApplication(object):
    class LicenseException(Exception):
        pass
    class ConnectionException(Exception):
        pass
    class InitializationException(Exception):
        pass
    class SystemNotPresentException(Exception):
        pass

    def __init__(self, path=None):
        # determine location of ZOSAPI_NetHelper.dll & add as reference
        aKey = winreg.OpenKey(winreg.ConnectRegistry(None, winreg.HKEY_CURRENT_USER), r"Software\Zemax", 0, winreg.KEY_READ)
        zemaxData = winreg.QueryValueEx(aKey, 'ZemaxRoot')
        NetHelper = os.path.join(os.sep, zemaxData[0], r'ZOS-API\Libraries\ZOSAPI_NetHelper.dll')
        winreg.CloseKey(aKey)
        clr.AddReference(NetHelper)
        import ZOSAPI_NetHelper
        
        # Find the installed version of OpticStudio
        #if len(path) == 0:
        if path is None:
            isInitialized = ZOSAPI_NetHelper.ZOSAPI_Initializer.Initialize()
        else:
            # Note -- uncomment the following line to use a custom initialization path
            isInitialized = ZOSAPI_NetHelper.ZOSAPI_Initializer.Initialize(path)
        
        # determine the ZOS root directory
        if isInitialized:
            dir = ZOSAPI_NetHelper.ZOSAPI_Initializer.GetZemaxDirectory()
        else:
            raise PythonStandaloneApplication.InitializationException("Unable to locate Zemax OpticStudio.  Try using a hard-coded path.")

        # add ZOS-API referencecs
        clr.AddReference(os.path.join(os.sep, dir, "ZOSAPI.dll"))
        clr.AddReference(os.path.join(os.sep, dir, "ZOSAPI_Interfaces.dll"))
        import ZOSAPI

        # create a reference to the API namespace
        self.ZOSAPI = ZOSAPI

        # create a reference to the API namespace
        self.ZOSAPI = ZOSAPI

        # Create the initial connection class
        self.TheConnection = ZOSAPI.ZOSAPI_Connection()

        if self.TheConnection is None:
            raise PythonStandaloneApplication.ConnectionException("Unable to initialize .NET connection to ZOSAPI")

        self.TheApplication = self.TheConnection.CreateNewApplication()
        if self.TheApplication is None:
            raise PythonStandaloneApplication.InitializationException("Unable to acquire ZOSAPI application")

        if self.TheApplication.IsValidLicenseForAPI == False:
            raise PythonStandaloneApplication.LicenseException("License is not valid for ZOSAPI use")

        self.TheSystem = self.TheApplication.PrimarySystem
        if self.TheSystem is None:
            raise PythonStandaloneApplication.SystemNotPresentException("Unable to acquire Primary system")

    def __del__(self):
        if self.TheApplication is not None:
            self.TheApplication.CloseApplication()
            self.TheApplication = None
        
        self.TheConnection = None
    
    def OpenFile(self, filepath, saveIfNeeded):
        if self.TheSystem is None:
            raise PythonStandaloneApplication.SystemNotPresentException("Unable to acquire Primary system")
        self.TheSystem.LoadFile(filepath, saveIfNeeded)

    def CloseFile(self, save):
        if self.TheSystem is None:
            raise PythonStandaloneApplication.SystemNotPresentException("Unable to acquire Primary system")
        self.TheSystem.Close(save)

    def SamplesDir(self):
        if self.TheApplication is None:
            raise PythonStandaloneApplication.InitializationException("Unable to acquire ZOSAPI application")

        return self.TheApplication.SamplesDir

    def ExampleConstants(self):
        if self.TheApplication.LicenseStatus == self.ZOSAPI.LicenseStatusType.PremiumEdition:
            return "Premium"
        elif self.TheApplication.LicenseStatus == self.ZOSAPI.LicenseStatusTypeProfessionalEdition:
            return "Professional"
        elif self.TheApplication.LicenseStatus == self.ZOSAPI.LicenseStatusTypeStandardEdition:
            return "Standard"
        else:
            return "Invalid"
    
    def reshape(self, data, x, y, transpose = False):
        """Converts a System.Double[,] to a 2D list for plotting or post processing
        
        Parameters
        ----------
        data      : System.Double[,] data directly from ZOS-API 
        x         : x width of new 2D list [use var.GetLength(0) for dimension]
        y         : y width of new 2D list [use var.GetLength(1) for dimension]
        transpose : transposes data; needed for some multi-dimensional line series data
        
        Returns
        -------
        res       : 2D list; can be directly used with Matplotlib or converted to
                    a numpy array using numpy.asarray(res)
        """
        if type(data) is not list:
            data = list(data)
        var_lst = [y] * x;
        it = iter(data)
        res = [list(islice(it, i)) for i in var_lst]
        if transpose:
            return self.transpose(res);
        return res
    
    def transpose(self, data):
        """Transposes a 2D list (Python3.x or greater).  
        
        Useful for converting mutli-dimensional line series (i.e. FFT PSF)
        
        Parameters
        ----------
        data      : Python native list (if using System.Data[,] object reshape first)    
        
        Returns
        -------
        res       : transposed 2D list
        """
        if type(data) is not list:
            data = list(data)
        return list(map(list, zip(*data)))

if __name__ == '__main__':
    zos = PythonStandaloneApplication()
    
    # load local variables
    ZOSAPI = zos.ZOSAPI
    TheApplication = zos.TheApplication
    TheSystem = zos.TheSystem
    TheLDE = TheSystem.LDE
    TheMFE = TheSystem.MFE
    
    
    # I. Double Gauss filepath (should be installed by default)
    SampleDirectory = TheApplication.SamplesDir
    DoubleGauss = r'\Sequential\Objectives\Double Gauss 28 degree field.zmx'
    Filepath = SampleDirectory + DoubleGauss


    # II. Try opening the Double Gauss file without saving it (False flag)
    try:
        TheSystem.LoadFile(Filepath, False)
        print('File: ' + Filepath + ' opened successfully')
    except:
        del zos
        zos = None
        print('Error: Double Gauss file not found')
    
    
    # For point 1. there are two options (1a. and 1b.) detailed below:
        
    # ========================================================================    
    # 1a. Run a macro as a solve. This is glitchy, and one needs to carefully
    # verify that the macro is doing what it is intended to do.
    
    # Create a definition for the Macro Solve type
    MacroSolveDef = ZOSAPI.Editors.SolveType.ZPLMacro
    
    # Add a dummy surface at the end of the LDE
    DummySurface = TheLDE.AddSurface()
    
    # Create a Macro Solve on the radius of the dummy surface
    MacroSolve = DummySurface.RadiusCell.CreateSolveType(MacroSolveDef)
    
    # Select the macro to be run (this dummy macro writes a text file in the
    # same directory as the lens file and is attached to the repository)
    MacroSolve.Macro = 'Standalone_text_example'
    
    # Apply the macro solve
    DummySurface.RadiusCell.SetSolveData(MacroSolve)
    
    # Update text display
    print('1a. File saved from macro')


    # ========================================================================
    # 1b. Run a POP analysis in ZOS-API, this would be my prefered method as
    # it gives more flexibility, and is probably more robust (you have a clear
    # look at what the code does). Optionally, you can also run the high-
    # sampling POP, which is exclusive to the ZOS-API (see Syntax Help)
    
    # Create a definition for the POP analysis
    POPDef = ZOSAPI.Analysis.AnalysisIDM.PhysicalOpticsPropagation
    
    # Open a new POP analysis
    MyPOP = TheSystem.Analyses.New_Analysis(POPDef)
    
    # Retrieve POP settings
    MyPOPSettings = MyPOP.GetSettings()
    
    # Adjust POP settings (I will just change a couple of settings, the full
    # list is available from the Syntax Help, and I'll add it to the
    # repository)
    
    # Wavelength number
    WavelengthNumber = 1
    MyPOPSettings.Wavelength.SetWavelengthNumber(WavelengthNumber)
    
    # Field number
    FieldNumber = 1
    MyPOPSettings.Field.SetFieldNumber(FieldNumber)
    
    # Start surface
    StartSurfaceNumber = 2
    MyPOPSettings.StartSurface.SetSurfaceNumber(StartSurfaceNumber)
    
    # Surface-to-beam distance [in lens unit I believe]
    SurfaceToBeamDistance = 0.0
    MyPOPSettings.SurfaceToBeam = SurfaceToBeamDistance
    
    # Use polarization
    MyPOPSettings.UsePolarization = False
    
    # Apply new POP settings
    MyPOP.ApplyAndWaitForCompletion()
    
    # Retrieve results
    MyPOPResults = MyPOP.GetResults()
    
    # Save the POP grid (if you need a result from the text window, it might
    # be easier to use the POPD operand, and I have a GIST you can find at
    # https://gist.github.com/Omnistic/cfff35796e7cf9cbbda9b1de90d104e2)
    if MyPOPResults.NumberOfDataGrids:
        POPGrid = MyPOPResults.GetDataGrid(0).Values
        POPGrid = zos.reshape(POPGrid,
                              POPGrid.GetLength(0), POPGrid.GetLength(1))
        
        # Display the POP grid in a figure
        plt.figure()
        plt.title('POP Analysis')
        plt.imshow(POPGrid)
        plt.show()
        
        # Update text display
        print('1b. POP analysis completed')
    else:
        # Update text display
        print('1b. Something went wrong witht the POP analysis')
    
    
    # ========================================================================
    # 2. Here you can do whatever you like with the data from the POP analysis
    
    # For example, I can convert the grid to a numpy array 
    POPGrid = np.array(POPGrid)
    
    # And print the mean intensity value
    POPGridMean = POPGrid.mean()
    print('2. Mean POP intensity = ' + str(POPGridMean))
    
    # Update text display
    print('2. Processing completed')
    
    
    # ========================================================================
    # 3. Upload the calculated values in 2. to a Merit Function Operand. Again
    # I'm doing something dummy, but it's just to show the idea
    
    # Define an operand, here TTHI
    TTHIDef = ZOSAPI.Editors.MFE.MeritOperandType.TTHI
    
    # Insert a new Merit Function Operand at the end of the MFE
    TTHIOperand = TheMFE.AddOperand()
    
    # Change its type to TTHI
    if TTHIOperand.ChangeType(TTHIDef):
        # Change its target to the mean POP intensity value (meanlingless)
        TTHIOperand.Target = POPGridMean
        
        # Update text display
        print('3. New operand added successfully')
    else:
        # Update text display
        print('3. Something went wrong when setting up the new MF operand')
    
    
    # ========================================================================
    # 4. I did not understand the last point
    
    
    # This will clean up the connection to OpticStudio.
    # Note that it closes down the server instance of OpticStudio, so you for maximum performance do not do
    # this until you need to.
    del zos
    zos = None