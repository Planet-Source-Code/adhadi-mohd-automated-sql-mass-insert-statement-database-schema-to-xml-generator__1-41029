'--------------------------------
' Automated SQL MASS Insert Statement Generator
' Implementation of Automated SQL Insert Statement Generator (with VB6)
' By Adhadi Mohd, 
' Kuala Lumpur.
' Adhadi@hotpop.com
' 22 Nov 2002
'
'
' The reason of doing is because I need to install my application together
' with database (schema & lookup data) at the client site.
' Of course we can do this using different approach like BACKUP/RESTORE,
' or access upsizing. In my case, I need to script all the schema and
' the data to have a better control to the setup file
' Previously, i wrote the lookup data manually.
' Luckilly, Josh Carderonello's article gave me the basic idea on how to perform the automation.
' MS Project Server 2002 installation also doing the database installation
' using the same approach
'
' Original Idea:
'       From article:
'       How to Write SQL to Dynamically Script Mass INSERT Statement Scripts
'       http://www.sql-server-performance.com/jc_write_sql_script.asp
'       by Josh Calderonello
'
'
' Requirements
' a. VB6X.DLL (if it not included here, you can find this component at planetsourcecode)
' b. SQLDMO components (installed with SQL7, SQL2000 or MSDE)
' c. MSXML 3.0
'
' Issuess
' a. Can't extract IMAGE data (Image)
' b. Identity data field was skipped. (It's my requirement)