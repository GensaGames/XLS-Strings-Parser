# XLS-Strings-Parser
Simple Python Script for generation String Resources for Android Applications from existing XLS and XLSX files. Will be usefull, when your localization in table view. As practice shows, it's very often. Please, make sure, you have installed <b>Python Modules xlrd and xlwt</b>

- Write sample below for more information.

  `python strings_xls_parser.py --help` 

- Scripts requires input XLS file, index for sheets and column for key value.
	Correct order for script caling below. 
  
  `python strings_xls_parser.py
	--file <path/to/file.xls>	
	--sheet <sheet index>
	--sheet_key <key sheet index>
	--sheet_val <value sheet index>`
