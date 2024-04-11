# originally created by dmarkle/ConvertFrom-Vcf

This version of this script is setup to automatically create CSV files from any .vcf files within a folder.

The CSV files that this script outputs are compatible for an immediate upload to Outlook on the Web.


This script currently works on the following folder structure

c:\contact-conversion\   -  This is where you want to store this script file

c:\contact-conversion\inbound   -  This is where you will put any .vcf files to convert

c:\contact-conversion\outbound   - This is where you will receive your CSV files.


The CSV files are named the same as the incoming files, so make sure if you are doing a bulk lot of files that they all have names that make sense!

All incoming files, once processed are renamed to name.vcf.bak. This is so that they are not re-processed, but it also means you have a backup copy of the original vcf with no changes made (except the file extension).

