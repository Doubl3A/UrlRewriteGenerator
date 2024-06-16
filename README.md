# UrlRewriteGenerator
Generate basic URL rewrite rules for .NET with a .xlsx files

# Features:
- Create a standardized rewrite xml file with rewrite URLs from a xlsx file

# Usage
1. Change the global variables in ``config.py`` to your needs
   * Change the file path to the file you want to generate rewrites from
   * Enable the cell color filter, and add the hex color codes that should be matched
   * Optional: Change the output file name
2. Run ``main.py`` to generate the rewrite file
3. Check the logs for anny error rows, and see if they have invalid URLs
4. Add the created rewrite xml file to your web application, or insert the automatically created rules into a preexisting rewrite file

# Remarks
- Redirects URL will not have www in them
- The rows that fail the URL parse will be printed out for debugging at the end of the job. The printed numbers is the row number in the xlsx file

# Future work
- Multiple URLs can match to the same URL. Find a way to join similar redirects in to one rewrite rule
