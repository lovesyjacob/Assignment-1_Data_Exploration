	1) Sum, Count, Average:			
		• What is the total price of all products in the dataset?		
       Formula =SUM(D2:D35) 
       Ans = $10,100
		• How many products are there in the dataset?		
       Formula =COUNT(D2:D35)
       Ans = 34 products
		• Calculate the average price of the products
       Formula =AVERAGE(D2:D35)
       Ans = $297.06
       
	2) Min and Max:				
		• Determine the minimum price among all products
       Formula =MIN(D2:D35)
       Ans = $30
		• Find the maximum price among all products
       Formula =MAX(D2:D35)
       Ans =$1,000
       													
	3) IF Function:												
		• Using an IF function, create a new column named Price Range to categorize products with a price greater than or equal to $500 as 'High Price' and          others as 'Standard Price'.		
       Created a new column called Price Range in G cell
       Formula =IF(D2>=500,"High Price","Standard Price")
       I used the IF function to check the price of each product.
       If the price is greater than or equal to $500, Excel displays "High Price".
       If the price is less than $500, Excel displays "Standard Price".
       Then draged the formula down to G35.
  							
	4) SUMIF and COUNTIF:						
		• Calculate the total price for products in the 'Electronics' category using the SUMIF function.	
        Formula =SUMIF(F2:F35,"Electronics",D2:D35)
        Ans = $8,050
		• Determine the count of products with a price less than $100 using the COUNTIF function.
        Formula =COUNTIF(D2:D35,"<100")
        Ans = 11 products
        
	5) Text Formatting - LEFT, RIGHT, MID:									
		• Create a new column named Day with the first 2 characters of each 'Product ID' using the LEFT function.	
        Created a new column named Day in H cell
        Formula =LEFT(A2,2)
        The LEFT function extracts characters from the left side of the Product ID.
        For example:
        28-JAN-US
        LEFT(A2,2) → 28
		• Create a new column named Country Code by extracting the last 2 characters from the 'Product ID' column using the RIGHT function.
        Created a new column named Country Code in I cell
        Formula =RIGHT(A2,2)
        The RIGHT function extracts characters from the right side of the Product ID.
        For example:
        28-JAN-US
        RIGHT(A2,2) → US
		• Create a new column named Month by extracting 4th to 6th characters from the 'Product ID' column using the MID function.
        Created a new column named Month in J cell
        Formula =MID(A2,4,3)
        means:
          Start at the 4th character
          Extract 3 characters
        For example:
        28-JAN-US
        Therefore, characters 4–6 are JAN.
