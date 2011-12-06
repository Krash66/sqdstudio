This folder is for files to build Metadata in 4 different environments.

There are two files for DB2:
--ALLTBLS_DB2_LUW.txt --- creates all SQDmetadata tables in DB2
--DB2_LUW DROP TABLES.txt --- drops all SQDmetadata tables in DB2

Two files for Oracle:
--ORCL_meta_create.ddl --- creates all SQDmetadata tables in Oracle
--ORCL_meta_drop.ddl --- drops all SQDmetadata tables in Oracle

One file for MS Access:
--V3MetaData.mdb --- This is an empty MS Access database with all SQD Metadata tables
*** NOTE *** AS BEST PRACTICE: 
*** Make a copy of this file to use as your working metadata and leave this database file empty

Two files for SQL Server:
--SQLsvrMetaV3_2.sql --- This file contains both create and drop statements for adding and removing SQDmetadata tables on MS SQLserver
--SQDmetaSQLsvr.mdf --- This is an empty MS SQLserver database. It is a database already built and can be attached as an existing database in SQLserver. It is recommended to make a copy of this file to attach to SQLserver so that a fresh, blank Database always exists in case it is needed for future use.

 
Please retain this file as a reference for metadata.