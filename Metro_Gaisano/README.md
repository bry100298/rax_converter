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

<h2>Script Overview</h2>

<ul>
    <li>The Metro Gaisano script converts XML files to Excel files for Metro Retail Stores Group, Inc.</li>
</ul>

<h2>Description</h2>

<ul>
    <li>The script parses XML files containing invoice data and converts them into Excel files.</li>
    <li>It then archives the XML and Excel files and moves the Excel files to the outbound folder.</li>
</ul>

<h2>Functionality</h2>

<ul>
    <li><strong>xml_to_excel(xml_file, parent_dir):</strong></li>
    <ul>
        <li>Parses XML file and extracts data.</li>
        <li>Populates DataFrame with extracted data.</li>
        <li>Writes DataFrame to Excel file.</li>
        <li>Archives XML and Excel files.</li>
        <li>Moves Excel files to the outbound folder.</li>
    </ul>
    <li><strong>main():</strong></li>
    <ul>
        <li>Iterates over XML files in the inbound folder.</li>
        <li>Calls <code>xml_to_excel</code> function for each XML file.</li>
    </ul>
</ul>

<h2>Dependencies</h2>

<ul>
    <li>Python 3.x</li>
    <li>pandas</li>
    <li>xml.etree.ElementTree</li>
</ul>

<h2>Usage</h2>

<ul>
    <li>Ensure Python and required libraries are installed.</li>
    <li>Place XML files containing invoice data in the inbound folder.</li>
    <li>Run the script.</li>
</ul>

<h2>Author</h2>

<p>Created by [Your Name].</p>

<h2>License</h2>

<p>This project is licensed under the [License Name] License - see the [LICENSE.md](link-to-license-file) file for details.</p>

</body>
</html>
