# ProductApp
Simple user interface for quick search in excel document.

Searching is based on matching any of the search terms
with any of the terms found in the description of the products
and includes the capability of searching by alliases.

Results list includes basic information about the matching products,
as well as extended information if a product is selected.

-Excel document requirements
The excel document contains various products along with their respective 
codes, cost, price and other information.
ProductApp searches the excel document based on a product's name or code
and return those information.

The excel document is provided by the user and all the necessary information
should be in a continuous table in a specific sheet. 

Since the excel document is read every time ProductApp is started, if any
changes are made to the excel, a simple restart of ProductApp gives always
updated info.
