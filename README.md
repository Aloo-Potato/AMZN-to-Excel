<!DOCTYPE html>
<html>

<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>README</title>
  <link rel="stylesheet" href="https://stackedit.io/style.css" />
</head>

<body class="stackedit">
  <div class="stackedit__html"><h1 id="excel-to-hardware-hub">Excel to Hardware Hub</h1>
<p>Automatically manages and cleans the website every times it is run</p>
<h2 id="goal">Goal</h2>
<p>Manage the hardware Hub website automatically</p>
<h2 id="technologies">Technologies</h2>
<p>-Selenium: Webscraper for C#</p>
<blockquote>
<p><a href="https://www.selenium.dev/">https://www.selenium.dev/</a></p>
</blockquote>
<p>-WordpressPCL: Wordpress API for C#</p>
<blockquote>
<p><a href="https://github.com/wp-net/WordPressPCL">https://github.com/wp-net/WordPressPCL</a></p>
</blockquote>
<p>-NPOI: Read and write excel files in C#.</p>
<blockquote>
<p><a href="https://github.com/nissl-lab/npoi">https://github.com/nissl-lab/npoi</a></p>
</blockquote>
<h2 id="procedure">Procedure</h2>
<ol>
<li>cleanMedia();</li>
</ol>
<ul>
<li>Cleans the media library of the wordpress site. Removes any pictures that are not being used</li>
</ul>
<ol start="2">
<li>removeDuplicates(products);</li>
</ol>
<ul>
<li>Removes any duplicate products from the arraylist since some pages repeat sale items</li>
</ul>
<ol start="3">
<li>Wordpress.readWebsite();</li>
</ol>
<ul>
<li>Reads the products on the website and stores them in an arraylist to be compared to later</li>
</ul>
<ol start="4">
<li>Selenium.GetOldPostData();</li>
</ol>
<ul>
<li>Goes to the URL of the amazon page for each product and checks to see if the sale is still going on, if the sale has changed, or if the sale has ended.</li>
</ul>
<ol start="5">
<li>Wordpress.UpdatePosts();</li>
</ol>
<ul>
<li>Uses the data from the above method to update each post with the correct amazon sale information</li>
</ul>
<ol start="6">
<li>excel.ReadProducts();</li>
</ol>
<ul>
<li>Grabs new product data from products.xlsx (populated from Amazon to Excel Program) and stores it in an arraylist</li>
</ul>
<ol start="7">
<li>Formatting.correctCategories();</li>
</ol>
<ul>
<li>The categories assigned from amazon and the categorization on the website is different. This method fixes the category of each new product so it can be added to the website</li>
</ul>
<ol start="8">
<li>Product.removeUpdates();</li>
</ol>
<ul>
<li>some of the new products to be added are already on the website. This method removes those products already on the site from the arraylist</li>
</ul>
<ol start="9">
<li>Wordpress.AddPics();</li>
</ol>
<ul>
<li>Before adding new products to the site, their photos need to be added to the site’s media library</li>
</ul>
<ol start="10">
<li>Wordpress.CreatePost();</li>
</ol>
<ul>
<li>Posts each new product to the website</li>
</ul>
<ol start="11">
<li>Wordpress.removeDuplicates();</li>
</ol>
<ul>
<li>Sometimes some duplicates seep through the Product.removeUpdates() method so this method double checks the site to make sure there are no duplicates</li>
</ul>
<ol start="12">
<li>Wordpress.CleanImagesFolder();</li>
</ol>
<ul>
<li>cross references the media library of the site with the photo library on my hard drive and removes unused photos.</li>
</ul>
<ol start="13">
<li>excel.WriteHHPosts();</li>
</ol>
<ul>
<li>creates a copy of all products on the website and puts their data onto an excel sheet</li>
</ul>
<p>Done Updating!</p>
<h2 id="improvements">Improvements</h2>
<ol>
<li>Find a way to use the WordpressPCL API to add the amazon page link to each product instead of using the selenium browser to do it</li>
</ol>
<blockquote>
<p>Selenium can be buddy at times and browsing isn’t always predictable. This method of adding links also takes forever. Ideally id like to just add it directly with he API</p>
</blockquote>
<ol start="2">
<li>
<p>Find a faster way to check the sale data of existing products in Selenium.GetOldPostData(). The runtime can stretch to 10 minutes</p>
</li>
<li>
<p>Set up an SQL database on my laptop so the data is stored in a database instead of an excel file</p>
</li>
<li>
<p>Find a better way to organize data. Everything is a mess and hard to think through when debugging</p>
</li>
</ol>
</div>
</body>

</html>

