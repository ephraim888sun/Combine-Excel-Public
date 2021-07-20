# Combine_Excel

During Summer of 2021 internship at Scout Energy Patners. One of the side projects were to combine multiple excel sheets together. There is about 42 excel sheets each having 157 (should have 157) columns. The final combined excel sheet had around 20615 rows!

To combine each excel sheet, I was able to use the pandas modules to help me. First, I created an excel sheet with only the column headers. Then, I would loop through all the excel files, allow pandas to read the files (xlsx or csv), and then clean up the data a little bit. Then, I would add it to a dataframe, repeating this process until all the excel sheets were part of the same dataframe (used pd.concat()). 

After, I made sure to delete duplicate rows as well as make sure that we get the expected data amount. Then, I exported the dataframe to an excel sheet, where I was able to give it to my mentor.

Couple of issues: When adding the dataframes together, one of the column names weren't the same. Specifically, one of the column names had a dash and the other didn't.
I was able to separate these two by running the script twice, one with the dash and one without. Also, all of the excel sheets had 157 columns except for two of them. I was able to isolate these two files and manually include them at the end. 

Overall, I was able to save a lot of time by automating the process instead of manually going into each excel file to copy and paste. This also help me with my python skills, as it was one of my first time putting my pandas modules skills to the test.
