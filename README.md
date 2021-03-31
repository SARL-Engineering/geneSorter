# geneSorter
This program searches a specifically formatted gene spreadsheet and outputs results into a separate sheet

How to use these scripts:

These two scripts take respectively one or multiple .xlsx files placed in the Input folder and convert them into a single .xlsx which ends up in the Output folder. 

The 'geneSorter All' script will provide a series of prompts asking the user to format their input spreadsheet with a list of desired genes and a list of chemicals that was tested on them (as is on the spreadsheet's header line). It will then create a new sheet within that .xlsx containing the data on every gene in the gene list.
Note: make sure that each chemical is spelled and capitalized like the header line, or else the program won't work.

The 'One or Multiple Sheet' script will simply find all unique gene IDs in all of the inputted .xlsx and create a new .xlsx which contains the data on all unique genes. Make sure that the gene ID is in the first column!

----------------
"Publish it. Write it. Sing it. Swing to it. Yodel it. We wrote it, that's all we wanted to do."
-Woodie Guthrie
