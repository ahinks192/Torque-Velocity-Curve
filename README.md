ADDITIONAL INSTRUCSTIONS ON USING THIS CODE

Ensure Python 3 is downloaded from: https://www.python.org/downloads/

Next, download Visual Studio Code, which you will use as an interface to run the Python code: https://code.visualstudio.com/download

Use "Open With" to open "CurveFitting_InVivo.py" in Visual Studio Code. In Visual Studi Code, install the Python plugin.

At Line 133 in the code, edit the text that says '/Users/averyhinks/Desktop/excel_output.xlsx' to be the new user's file path information. 
IMPORTANT: After editing Line 133, press Save.

In the Terminal at the bottom of the screen in Visual Studio Code, use the following command to install the Python modules that will be needed to run this code: 
Pip3 install numpy pandas scipy matplotlib sklearn tkinter

If errors come up for any of these, try entering these shortforms for the module instead when running the above command:
For numpy -> np
For pandas -> pd
For matplotlib -> plt
For sklearn -> skl
For tkinter -> Tk

If everything above has been completed successfully, you should be able to run the code by clicking the Play button (looks like a triangle pointing to the right) in the top right of Visual Studio Code. A window should appear prompting you to select the Excel file from which you are extracting data.
