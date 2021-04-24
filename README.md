<!DOCTYPE html>
<html>

<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>README</title>
  <link rel="stylesheet" href="https://stackedit.io/style.css" />
</head>

<body class="stackedit">
  <div class="stackedit__html"><h1 id="amazon-to-excel-sheet">Amazon to Excel Sheet</h1>
<p>Finds electronics deals on Amazon and organizes the data on an excel spreadsheet</p>
<h2 id="goal">Goal</h2>
<p>Add new products to the Hardware Hub website(<a href="http://zed.exioite.com">zed.exioite.com</a>)</p>
<h2 id="technologies">Technologies</h2>
<p>-Selenium: Webscraper for C#</p>
<blockquote>
<p><a href="https://www.selenium.dev/">https://www.selenium.dev/</a></p>
</blockquote>
<p>-NPOI: Read and write excel files in C#.</p>
<blockquote>
<p><a href="https://github.com/nissl-lab/npoi">https://github.com/nissl-lab/npoi</a></p>
</blockquote>
<h2 id="procedure">Procedure</h2>
<ol>
<li>getInfo();</li>
</ol>
<ul>
<li>Goes to different Amazon electronics pages and scrapes the pages for products that are on sale. Grabs the name, picture, old price, sale price, and URL of the product and saves it into an arraylist. Also assigns each product with its own GUID</li>
</ul>
<ol start="2">
<li>removeDuplicates(products);</li>
</ol>
<ul>
<li>Removes any duplicate products from the arraylist since some pages repeat sale items</li>
</ul>
<ol start="3">
<li>writeExcel(products);</li>
</ol>
<ul>
<li>clears the products.xlsx to make way for new products on sale this week. Then populates the sheets with products and their data</li>
</ul>
<h2 id="improvements">Improvements</h2>
<ol>
<li>
<p>Find a faster way to grab data. Runtime can stretch up to 15 minutes</p>
</li>
<li>
<p>Set up an SQL database on my laptop so the data is stored in a database instead of an excel file</p>
</li>
</ol>
</div>
</body>

</html>
