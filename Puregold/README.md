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
    <li>The Puregold script converts HTML files to Excel files for Puregold Price Club Inc. and Ayagold Retailers, Inc.</li>
</ul>

<h2>Description</h2>

<ul>
    <li>The script reads HTML files containing invoice data and converts them into Excel files.</li>
    <li>It then archives the HTML and Excel files and moves the Excel files to the outbound folder.</li>
</ul>

<h2>Functionality</h2>

<ul>
    <li><strong>html_to_excel(html_file, parent_dir):</strong></li>
    <ul>
        <li>Converts HTML file to Excel format.</li>
        <li>Extracts data from HTML and populates DataFrame.</li>
        <li>Writes DataFrame to Excel file.</li>
        <li>Archives HTML and Excel files.</li>
        <li>Moves Excel files to the outbound folder.</li>
    </ul>
    <li><strong>main():</strong></li>
    <ul>
        <li>Iterates over HTML files in the inbound folder.</li>
        <li>Calls <code>html_to_excel</code> function for each HTML file.</li>
    </ul>
</ul>

<h2>Dependencies</h2>

<ul>
    <li>Python 3.x</li>
    <li>pandas</li>
    <li>Beautiful Soup 4</li>
</ul>

<h2>Usage</h2>

<ul>
    <li>Ensure Python and required libraries are installed.</li>
    <li>Place HTML files containing invoice data in the inbound folder.</li>
    <li>Run the script.</li>
</ul>

<h2>Author</h2>

<p>Created by [Your Name].</p>

<h2>License</h2>

<p>This project is licensed under the [License Name] License - see the [LICENSE.md](link-to-license-file) file for details.</p>

</body>
</html>
