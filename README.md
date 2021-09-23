# OCLC-renew

Script will retrieve updated OCLC numbers and holdings status. Expected input is xlsx with headers. Column 1 is local ID, column 2 is OCLC#. All xlsx files in downloads folder will be processed. If all records are held by your institution and all OCLC#s match within each record, then no action is taken. If a record is unheld, or if any OCLC#s do not match, then the results are output to csv. Contact OCLC to set up the WorldCat Metadata API for use with bookops-worldcat.
