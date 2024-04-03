<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>README</title>
    <style>
        body {
            font-family: Arial, sans-serif;
            margin: 20px;
        }
        h2 {
            margin-bottom: 10px;
        }
        pre {
            background-color: #f4f4f4;
            padding: 10px;
            border-radius: 5px;
        }
        img {
            max-width: 100%;
            height: auto;
            display: block;
            margin: 20px auto;
        }
        ul {
            list-style-type: none;
        }
        ul ul {
            margin-left: 20px;
        }
    </style>
</head>
<body>

<h2>Script Flow Diagram</h2>


<ul>
    <li>Start
        <ul>
            <li>Check Inbound Folder for XML files
                <ul>
                    <li>Iterate over XML files
                        <ul>
                            <li>Check if filename starts with "RA" (case sensitive)
                                <ul>
                                    <li>Yes: Continue</li>
                                    <li>No: Move file to Error Folder and Return</li>
                                </ul>
                            </li>
                            <li>Parse XML file</li>
                            <li>Check if company name is "SANFORD MARKETING CORPORATION"
                                <ul>
                                    <li>Yes: Continue</li>
                                    <li>No: Move file to Error Folder and Return</li>
                                </ul>
                            </li>
                            <li>Extract data from XML</li>
                            <li>Create DataFrame</li>
                            <li>Create Excel file path</li>
                            <li>Write DataFrame to Excel</li>
                            <li>Move XML file to Archive Folder</li>
                        </ul>
                    </li>
                </ul>
            </li>
        </ul>
    </li>
    <li>End</li>
</ul>

<h2>Description</h2>

<ul>
    <li>This diagram outlines the main steps performed by the script:
        <ul>
            <li>Checking the Inbound Folder for XML files.</li>
            <li>Iterating over each XML file found.</li>
            <li>Verifying conditions such as filename starting with "RA" and company name being "SANFORD MARKETING CORPORATION".</li>
            <li>Extracting data from the XML files.</li>
            <li>Creating a DataFrame from the extracted data.</li>
            <li>Writing the DataFrame to an Excel file.</li>
            <li>Moving the processed XML file to the Archive Folder.</li>
        </ul>
    </li>
    <li>Additionally, there's a delay of 15 seconds between processing each XML file, implemented using the <code>time.sleep()</code> function.</li>
</ul>


<img src="https://i.imgur.com/0piWgBm.png" alt="Script Flow Diagram">

<p>Feel free to customize this README.md file as needed for your project.</p>

</body>
</html>
