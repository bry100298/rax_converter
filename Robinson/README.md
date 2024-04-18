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
    <li>The Robinson script processes Excel files for Robinsons Supermarket and Robinsons Department Store.</li>
    <li>It merges specific Excel files and generates outbound Excel files for each company.</li>
</ul>

<h2>Description</h2>

<ul>
    <li>The script performs the following actions:</li>
    <li>Reads specific Excel files from the Inbound folder.</li>
    <li>Merges data from these files and saves them as new Excel files.</li>
    <li>Generates outbound Excel files with customized columns for each company.</li>
</ul>

<h2>Functionality</h2>

<ul>
    <li><strong>merge_excel_files_robd(company_code):</strong></li>
    <ul>
        <li>Merges specific Excel files for Robinsons Department Store.</li>
        <li>Saves the merged data to a new Excel file.</li>
        <li>Archives original files.</li>
    </ul>
    <li><strong>merge_excel_files_robs(company_code):</strong></li>
    <ul>
        <li>Merges specific Excel files for Robinsons Supermarket.</li>
        <li>Saves the merged data to a new Excel file.</li>
        <li>Archives original files.</li>
    </ul>
    <li><strong>generate_inbound_outbound_excel(company_folder, company_names):</strong></li>
    <ul>
        <li>Generates outbound Excel files with customized columns for each company.</li>
        <li>Archives original merged files.</li>
    </ul>
</ul>

<h2>Dependencies</h2>

<ul>
    <li>Python 3.x</li>
    <li>pandas</li>
    <li>openpyxl</li>
</ul>

<h2>Usage</h2>

<ul>
    <li>Ensure Python and required libraries are installed.</li>
    <li>Run the script with appropriate Excel files in the Inbound folder.</li>
</ul>

<h2>Author</h2>

<p>Created by [Your Name].</p>

<h2>License</h2>

<p>This project is licensed under the [License Name] License - see the [LICENSE.md](link-to-license-file) file for details.</p>

</body>
</html>
