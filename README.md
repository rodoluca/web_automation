# web_automation

Project Web Automation - Price Comparison
Objective: Develop a project where we use web automation with Selenium to search for the information we need.

How it will work:
Imagine that you work in the purchasing department of a company and need to compare suppliers for your inputs/products.

At this point, you will constantly search the websites of these suppliers for available products and prices because each supplier may have promotions at different times and with different prices.

Your objective: If the product prices are below a maximum price limit defined by you, you will discover the cheapest products and update this information in a spreadsheet.

Then, you will send an email with the list of products below your maximum purchase price.

In our case, we will do it with common products on websites like Google Shopping and Buscapé, but the idea is the same for other websites.

What do we have available?
-Product spreadsheet with the product names, maximum price, minimum price (to avoid "wrong" products or "too good to be true" prices), and the terms we want to avoid in our searches.

What should we do:
-Search for each product on Google Shopping and retrieve all results that have prices within the range and are the correct products.
-Do the same for Buscapé.
-Send an email to your email address with the notification and a table containing the items and prices found, along with the purchase link.
