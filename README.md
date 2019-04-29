# Keyword-Extractor
Description: This module parses excel documents looking for specified keywords. It then outputs the frequency of the
keywords into a output file.
Requirements: xlrd, nltk,
Input: excel files where the 4th column has the data to be extracted.
Output: txt file with distribution data on keywords
Restrictions: Does not work with non excel files even if in the same csv format. It also will not work with other excel
files that aren't looking for the words specified in the dict. Changing the values in the dict may cause errors and is
not recommended.
