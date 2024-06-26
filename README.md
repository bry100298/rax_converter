<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Rax Converter</title>
</head>
<body>

<h1>Rax Converter</h1>

<h2>Introduction</h2>

<p>This tool is designed to perform conversions for Rax. It includes Python scripts to handle various data manipulation tasks.</p>

<h2>Installation</h2>

<h3>Prerequisites</h3>

<ul>
    <li>Python 3.10 or higher</li>
</ul>

<h3>Setup Instructions</h3>

<ol>
    <li>
        <strong>Clone Repository:</strong> Clone this repository to your local machine.
    </li>
    <li>
        <strong>Navigate to Root Folder:</strong> Open your terminal or command prompt and navigate to the root folder of the cloned repository.
    </li>
    <li>
        <strong>Create Virtual Environment:</strong> It's recommended to create a virtual environment to isolate dependencies. Run the following commands to create and activate a virtual environment:
        <pre><code>python -m venv rax_env

####
source rax_env/bin/activate    # On Windows, use "rax_env\Scripts\activate"
####
source <your_virtual_env>/bin/activate  # for Unix/Linux
#### or
<your_virtual_env>\Scripts\activate    # for Windows




deactivate
</code></pre>
    </li>
    <li>
        <strong>Install Dependencies:</strong> Install the required Python packages using pip:
        <pre><code>pip install pandas
pip install openpyxl
pip install beautifulsoup4
pip install pdfplumber
pip install tabula-py
pip install jpype1
pip install xlrd
pip install pyinstaller
pip install cx-Freeze
pip install py2exe
pip install nuitka
pip install auto-py-to-exe</code></pre>
    </li>
    <li>
        <strong>Run the Script:</strong> Once the virtual environment is activated and dependencies are installed, you can run the Python scripts from the root folder of the project. For example:
        <pre><code>python SM_Group/SM_GrpConverter.py</code></pre>
    </li>
</ol>

<h2>Usage</h2>

<p>After following the installation steps, you can utilize the provided Python scripts for your desired conversions. Ensure that you include the necessary imports mentioned below in your Python scripts.</p>

<h3>Required Imports for Python Scripts</h3>

<p>Make sure to include the following imports in your Python scripts:</p>

<pre><code>import os
import shutil
import pandas as pd
from datetime import datetime
from xml.etree import ElementTree as ET
import time
from bs4 import BeautifulSoup
import pdfplumber</code></pre>

<h2>Dependencies</h2>
<ul>
    <li>Python 3.10.1</li>
    <li>pandas</li>
</ul>
<h2>Author</h2>
<p>Created by James Bryant Tin.</p>
<h2>License</h2>
<p>This project is licensed under the [License Name] License - see the [LICENSE.md](link-to-license-file) file for details.</p>

<p>To remove all .gitkeep</p>
<pre><code>
find . -type f -name ".gitkeep" -exec rm {} \;
find . -type f -name ".gitkeep"
find . -type f -name ".git*" -exec rm {} \;
find . -type f -name ".git*"
rm -rf .git
</code></pre>

<p>Use requirement.txt to see the dependencies version</p>

<pre><code>
pip install -r requirements.txt
pip list
python -m pip list
</code></pre>




<h2>Updating README.md for pyinstaller inclusion of tabula jar file

To ensure that the tabula jar file is included when using pyinstaller with `ever.py`, you need to manually add it as a data file. Here's how to do it efficiently:</h2>
<pre><code>
pyinstaller --noconfirm --onefile --console --icon "C:/Users/User/Downloads/benbytree_icon.ico" --add-data "C:/Users/User/AppData/Local/Programs/Python/Python310/Lib/site-packages/tabula/tabula-1.0.5-jar-with-dependencies.jar;tabula" "C:/Users/User/Documents/Project/rax_converter/ever.py"
</code></pre>


<h2>To create custom_jre /rax_converter/custom_jre;./jre | dependencies.jar;./tabula | --add-data for ./jre</h2>
<pre><code>
jlink --module-path "C:\RedhatJava\Java17\java-17-openjdk-17.0.10.0.7-2.win.x86_64\jmods" --add-modules java.base,java.logging,java.xml --output custom_jre
</code></pre>
<pre><code>
pyinstaller --noconfirm --onefile --console --icon "C:/Users/User/Downloads/benbytree_icon.ico" --add-data "C:/Users/User/Documents/Project/rax_converter/custom_jre;./jre" --add-data "C:/Users/User/AppData/Local/Programs/Python/Python310/Lib/site-packages/tabula/tabula-1.0.5-jar-with-dependencies.jar;./tabula" --hidden-import jpype "C:/Users/User/Documents/Project/rax_converter/ever.py"
</code></pre>

<h2>/rax_converter/custom_jre;custom_jre | dependencies.jar;tabula | --add-data for custom_jre</h2>
<pre><code>
pyinstaller --noconfirm --onefile --console --icon "C:/Users/User/Downloads/benbytree_icon.ico" --add-data "C:/Users/User/Documents/Project/rax_converter/custom_jre;custom_jre" --add-data "C:/Users/User/AppData/Local/Programs/Python/Python310/Lib/site-packages/tabula/tabula-1.0.5-jar-with-dependencies.jar;tabula" --hidden-import jpype "C:/Users/User/Documents/Project/rax_converter/ever.py"
</code></pre>

<h2>Starting of using --add-binary for custom_jre</h2>
<pre><code>
pyinstaller --noconfirm --onefile --console --icon "C:/Users/User/Downloads/benbytree_icon.ico" --add-binary "C:/Users/User/Documents/Project/rax_converter/custom_jre;custom_jre" --add-data "C:/Users/User/AppData/Local/Programs/Python/Python310/Lib/site-packages/tabula/tabula-1.0.5-jar-with-dependencies.jar;tabula" --hidden-import jpype "C:/Users/User/Documents/Project/rax_converter/ever.py"
</code></pre>

<h2>Starting of getting direct to C:/RedhatJava/Java17/java-17-openjdk-17.0.10.0.7-2.win.x86_64/bin/server/jvm.dll;jre/bin/server

and removing --hidden-import jpype

--add-binary "C:/Users/User/AppData/Local/Programs/Python/Python310/Lib/site-packages/jpype;lib/jpype"
</h2>
<pre><code>
pyinstaller --noconfirm --onefile --console --icon "C:/Users/User/Downloads/benbytree_icon.ico" --add-binary "C:/RedhatJava/Java17/java-17-openjdk-17.0.10.0.7-2.win.x86_64/bin/server/jvm.dll;jre/bin/server" --add-binary "C:/Users/User/AppData/Local/Programs/Python/Python310/Lib/site-packages/jpype;lib/jpype" --add-data "C:/Users/User/AppData/Local/Programs/Python/Python310/Lib/site-packages/tabula/tabula-1.0.5-jar-with-dependencies.jar;tabula" "C:/Users/User/Documents/Project/rax_converter/ever.py"
</code></pre>


<h2>60mb   --add-data for jpype because /jpype/*.py;jpype</h2>
<pre><code>
pyinstaller --noconfirm --onefile --console --icon "C:/Users/User/Downloads/benbytree_icon.ico" --add-binary "C:/RedhatJava/Java17/java-17-openjdk-17.0.10.0.7-2.win.x86_64/bin/server/jvm.dll;jre/bin/server" --add-data "C:/Users/User/AppData/Local/Programs/Python/Python310/Lib/site-packages/jpype/*.py;jpype" --add-data "C:/Users/User/AppData/Local/Programs/Python/Python310/Lib/site-packages/tabula/tabula-1.0.5-jar-with-dependencies.jar;tabula" "C:/Users/User/Documents/Project/rax_converter/ever.py"
</code></pre>


<h2>60mb    --add-data for jpype because /jpype;jpype</h2>
<pre><code>
pyinstaller --noconfirm --onefile --console --icon "C:/Users/User/Downloads/benbytree_icon.ico" --add-binary "C:/RedhatJava/Java17/java-17-openjdk-17.0.10.0.7-2.win.x86_64/bin/server/jvm.dll;jre/bin/server" --add-data "C:/Users/User/AppData/Local/Programs/Python/Python310/Lib/site-packages/jpype;jpype" --add-data "C:/Users/User/AppData/Local/Programs/Python/Python310/Lib/site-packages/tabula/tabula-1.0.5-jar-with-dependencies.jar;tabula" "C:/Users/User/Documents/Project/rax_converter/ever.py"
</code></pre>

<h2>60mb   --add-data for jpype because /jpype;jpype and -hidden-import jpype</h2>
<pre><code>
pyinstaller --noconfirm --onefile --console --icon "C:/Users/User/Downloads/benbytree_icon.ico" --add-binary "C:/RedhatJava/Java17/java-17-openjdk-17.0.10.0.7-2.win.x86_64/bin/server/jvm.dll;jre/bin/server" --add-data "C:/Users/User/AppData/Local/Programs/Python/Python310/Lib/site-packages/jpype;jpype" --add-data "C:/Users/User/AppData/Local/Programs/Python/Python310/Lib/site-packages/tabula/tabula-1.0.5-jar-with-dependencies.jar;tabula" --hidden-import jpype "C:/Users/User/Documents/Project/rax_converter/ever.py"
</code></pre>



<h1>We dont need jpype or jvm, just installed directly it will work</h1>
<h2>56mb   -Removing --hidden-import jpype and Removing /jpype;jpype and removing jvm.dll</h2>
<pre><code>
pyinstaller --noconfirm --onefile --console --icon "C:/Users/User/Downloads/benbytree_icon.ico" --add-data "C:/Users/User/AppData/Local/Programs/Python/Python310/Lib/site-packages/tabula/tabula-1.0.5-jar-with-dependencies.jar;tabula" "C:/Users/User/Documents/Project/rax_converter/ever.py"
</code></pre>

<p>Feel free to customize this README.md file as needed for your project.</p>

</body>
</html>














pyinstaller --noconfirm --onefile --console --icon "C:/Users/User/Downloads/benbytree_icon.ico" --runtime-hook "hook-jpylyzer.py" --add-binary "C:\Users\User\Documents\Project\rax_converter\custom_jre\bin;bin" --add-data "C:/Users/User/AppData/Local/Programs/Python/Python310/Lib/site-packages/tabula/tabula-1.0.5-jar-with-dependencies.jar;tabula" "C:/Users/User/Documents/Project/rax_converter/ever.py"


os.environ['JAVA_HOME'] = os.path.abspath('./custom_jre')
--79mb but no --runtime-hook "hook-jpylyzer.py"
pyinstaller --noconfirm --onefile --console --icon "C:/Users/User/Downloads/benbytree_icon.ico" --add-binary "C:/RedhatJava/Java17/java-17-openjdk-17.0.10.0.7-2.win.x86_64/bin/server/jvm.dll;jre/bin/server" --add-data "C:/Users/User/Documents/Project/rax_converter/custom_jre;custom_jre" --add-data "C:/Users/User/AppData/Local/Programs/Python/Python310/Lib/site-packages/tabula/tabula-1.0.5-jar-with-dependencies.jar;tabula" --hidden-import jpype "C:/Users/User/Documents/Project/rax_converter/ever.py"













pyinstaller --noconfirm --onefile --console --icon "C:/Users/User/Downloads/benbytree_icon.ico" --add-binary "C:/RedhatJava/Java17/java-17-openjdk-17.0.10.0.7-2.win.x86_64/bin/server/jvm.dll;jre/bin/server" --add-data "C:/Users/User/Documents/Project/rax_converter/custom_jre;jre" --add-data "C:/Users/User/AppData/Local/Programs/Python/Python310/Lib/site-packages/tabula/tabula-1.0.5-jar-with-dependencies.jar;tabula" --hidden-import jpype "C:/Users/User/Documents/Project/rax_converter/ever.py"




#last failed
pyinstaller --noconfirm --onefile --console --icon "C:/Users/User/Downloads/benbytree_icon.ico" --runtime-hook "hook-jpylyzer.py" --add-binary "C:\Users\User\Documents\Project\rax_converter\custom_jre\bin;bin" --add-data "C:/Users/User/AppData/Local/Programs/Python/Python310/Lib/site-packages/tabula/tabula-1.0.5-jar-with-dependencies.jar;tabula" "C:/Users/User/Documents/Project/rax_converter/ever.py"


pyinstaller --noconfirm --onefile --console --icon "C:/Users/User/Downloads/benbytree_icon.ico" --add-binary "C:\Users\User\Documents\Project\rax_converter\custom_jre\bin;bin" --add-data "C:/Users/User/AppData/Local/Programs/Python/Python310/Lib/site-packages/tabula/tabula-1.0.5-jar-with-dependencies.jar;tabula" "C:/Users/User/Documents/Project/rax_converter/ever.py"