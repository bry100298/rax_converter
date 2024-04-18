<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Installation and Usage Guide</title>
</head>
<body>

<h1>Installation and Usage Guide</h1>

<h2>Installation</h2>

<ol>
    <li>
        <strong>Install Python 3.12.2:</strong><br>
        You can download and install Python 3.12.2 from the <a href="https://www.python.org/downloads/">official Python website</a>.
    </li>
    <li>
        <strong>Install Required Packages:</strong><br>
        Run the following commands to install the necessary packages:
        <pre><code>pip install pandas
pip install openpyxl
pip install beautifulsoup4</code></pre>
pip install pdfplumber
pip install --upgrade xlrd
pip install pyinstaller #not working
pyinstaller your_script.py

pip install cx-Freeze #it does not work
python setup.py build
python setup.py bdist_msi




pip install py2exe #it does not work


pip install nuitka
nuitka sm_grp.py

</code></pre>
    </li>
</ol>

<h2>Required Imports for Python Scripts</h2>

<p>Make sure to include the following imports in your Python scripts:</p>

<pre><code>import os
import shutil
import pandas as pd
from datetime import datetime
from xml.etree import ElementTree as ET
import time
from bs4 import BeautifulSoup</code></pre>

<h2>Usage</h2>

<p>After installing Python and the required packages, you can use the provided Python scripts. Make sure to include the required imports mentioned above in your Python scripts.</p>

<pre><code># Example Python script using the required imports
import os
import shutil
import pandas as pd
from datetime import datetime
from xml.etree import ElementTree as ET
import time
from bs4 import BeautifulSoup
import pdfplumber

# Your code here...
</code></pre>

<p>Feel free to customize this README.md file as needed for your project.</p>

</body>
</html>


It should be install in Rax_converter root folder in order to work for example
PS C:\Users\User\Documents\Project\rax_converter>
it does not work if you run the script in PS C:\Users\User\Documents\Project\rax_converter>\Robinson or anuthing.


create an environment
C:\Users\User\AppData\Local\Programs\Python\Python310\
/path/to/python3.10 -m venv rax3_10
C:\Users\User\AppData\Local\Programs\Python\Python310 -m venv rax3_10

PS C:\Users\User\Documents\Project\rax_converter\SM_Group> .\SM_GrpConverter.exe
pyinstaller --onefile --name SM_GrpConverter --hidden-import pandas sm_grp.py


if python version can view as python --version no need to put python3.10 just use python -m straightforward
python -m venv rax3_10
source rax3_10/Scripts/activate
cd /c/Users/User/Documents/Project/rax_converter/SM_Group

source /c/Users/User/Documents/Project/python_projects/rax3_10/Scripts/activate
python sm_grp.py