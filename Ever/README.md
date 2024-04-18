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
    <li>The Ever script converts PDF files to Excel files for Ever Plus Superstore, Inc.</li>
</ul>

<h2>Description</h2>

<ul>
    <li>The script first converts PDF files to HTML using the tabula library.</li>
    <li>It then parses the HTML files, extracts relevant data, and converts it into Excel files.</li>
    <li>Finally, it archives the HTML, PDF, and Excel files, and moves the Excel files to the outbound folder.</li>
</ul>

<h2>Functionality</h2>

<ul>
    <li><strong>pdf_to_html(pdf_file, output_folder):</strong></li>
    <ul>
        <li>Extracts tables from PDF files using tabula.</li>
        <li>Concatenates tables into a single DataFrame.</li>
        <li>Converts DataFrame to HTML.</li>
        <li>Writes HTML content to file.</li>
        <li>Returns the path to the generated HTML file.</li>
    </ul>
    <li><strong>html_to_excel(html_file, parent_dir, pdf_file):</strong></li>
    <ul>
        <li>Parses HTML file and extracts data.</li>
        <li>Populates DataFrame with extracted data.</li>
        <li>Writes DataFrame to Excel file.</li>
        <li>Archives HTML, PDF, and Excel files.</li>
        <li>Moves Excel files to the outbound folder.</li>
    </ul>
    <li><strong>main():</strong></li>
    <ul>
        <li>Iterates over PDF files in the inbound folder.</li>
        <li>Converts PDF files to HTML using <code>pdf_to_html</code>.</li>
        <li>Processes HTML files to Excel using <code>html_to_excel</code>.</li>
    </ul>
</ul>

<h2>Dependencies</h2>

<ul>
    <li>Python 3.x</li>
    <li>pandas</li>
    <li>bs4 (BeautifulSoup)</li>
    <li>pdfplumber</li>
    <li>tabula</li>
</ul>

<h2>Usage</h2>

<ul>
    <li>Ensure Python and required libraries are installed.</li>
    <li>Place PDF files containing invoice data in the inbound folder.</li>
    <li>Run the script.</li>
</ul>

<h2>Author</h2>

<p>Created by [Your Name].</p>

<h2>License</h2>

<p>This project is licensed under the [License Name] License - see the [LICENSE.md](link-to-license-file) file for details.</p>

</body>
</html>
