Client for RW OPC, memory RW device reader v0.11.41

description:
-------------------
Programm has four part. First OPC Client for universal RW (ReadWin) OPC server, 
run server and obtain data to log-file, second ODBC Client write data to database, 
third read memory content from compatible devices, fourth initial interface for 
management this (and other in future maybe) processes.

Function:
- read memory from devices;
- create and show server logs, client logs, database access logs;
- retrieve system information (OS,processor,drives,page size,address);
- show current values from server (data retrieves from log-files);
- show all OPC Servers installed on machine;
- show all OPC Items on Hart OPC Server;
- has function to enable/disable cycle read/write;
- creating and modifiyng table on database fully automatically;
- allow select which items write to DB;
- writting period may be select manually;

installation notes:
-------------------
not installed utility yet

tested and compatible devices:
------------------------------
EcoGraph (ELD****)
MemoGraph (GLY****)

version history:
----------------
v0.11 build 41
released;