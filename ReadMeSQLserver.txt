*********************************************************
****Creating MetaData Database in MS SQLserver****
*********************************************************

Files involved with Meta Data Creation on MS SQLserver.

this readme
sqlmeta.ddl
sqlmeta.bat


FIRST:
Create a new Database on your SQLserver.
Call it whatever you like.

SECOND:
Create an ODBC DSN for this Database on the PC that
you will run the Studio from.

THIRD: 
The included Batch file will create the Studio
metadata Tables in your New DataBase,

*** You Must Make Sure to edit the sqlmeta.bat file as follows:

The command in the Batch file is this:

osql -DTCmeta -Usa -Psqdata -i sqlmeta.ddl

command and switches: (all switches are case sensitive)

osql -> The sqlserver command

-D[TCmeta] -> -D is for database name. here we have called it TCmeta.

-U[sa] -> -U is for user id

-P[sqdata] -> -P is for password

-i [sqdmeta.ddl] -> -i means include or insert this structure file.


Aslong as this command is run from a PC with an ODBC connection
to your Metadata Database, the new tables will be created.

