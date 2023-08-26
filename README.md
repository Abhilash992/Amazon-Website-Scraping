# Amazon-Website-Scraping
To scrape at least 20 pages of product listing pages Items to scrape • Product URL • Product Name • Product Price • Rating • Number of reviews. With the Product URL received, hit each URL, and add below items: • Description • ASIN • Product Description • Manufacturer by going into each and every URL.

# Requirements to run the program
1. Install python into your system using the link: https://www.python.org/downloads
2. Use a code editor to run the code. I have used VS code, to download it use the link: https://code.visualstudio.com/
3. I have used MS Edge driver for scraping. So, you should have MS Edge of version:115.0.1901.188 and download the driver from the repository itself. If you have a different version of MS Edge then download the driver based on the version of your browser from the link:https://developer.microsoft.com/en-us/microsoft-edge/tools/webdriver. {Make sure that the driver and the code are in the same folder.
4. The code requires selenium, sys, openpyxl and pandas. To install the mentioned open command prompt and use the below commands
   * pip install sys
   * pip install openpyxl
   * pip install pandas
   * pip install selenium
5. The output is in .csv format, if you want the output to be in Excel don't run line number 309.
6. After running the code, the output will be saved in the same folder.
