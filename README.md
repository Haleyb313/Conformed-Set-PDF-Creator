# Conformed-Set-PDF-Creator
Collect thousands of architectural drawing pdfs, rename each file, sort into folders, then append into single pdfs

Goal:  Create one combined PDF out of all the individual PDFs located in respective folders

       1) Open Excel file to find out 1) folder locations and 2) Combined PDF file names

       2) Go into those folders, get all PDFS, save as one document using the designated name
       
           a) include page labels
           
           b) include page bookmarks (work in progress)
           
       3) Do this for all 3 Volumes!
       

       Prior to running this code, several Macros within the Excel file are run.  These include:

              1) Building a list of all Conformed Drawings (the most recent version of each drawing)

              2) Getting a handle to each PDF, updating the name, and saving it within a designated folder


1)  To start, get a handle to each volume's Master PDF Excel list
       While the file is renamed each week with a new date, we do know it is the only Excel file in the folder
       So we can go in, get a list of all Excel files, then grab the 1st (and only) file


       Within each file, we created a table designating folder names and locations.
              Make a dataframe of that information to reference later


2) Loop through all 3 of those dataframe tables:  For each row in Folder Path column:

       a) Follow that Path into that folder
       
       b) Get a list of all PDFs
       
       c) Combine them
       
       d) Save the single file using the name in the File Name column
