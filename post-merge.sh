#!/bin/sh
#
# This script is called by "git merge" (and therefore also "git pull") with no arguments.
# This script determines if there are any changes to the non-binary Excel files, and 
# recreates the binary xlsm file if there were changes
# mtilasha@ford.com
# Copyright: Ford Motor Company Limited

# get the file path to the Excel files as the present working directory
FILEPATH="$(pwd)"
cd "${FILEPATH}"

# create a text file and output the names of all files currently staged
touch changed.txt
git diff --name-only HEAD~ > changed.txt

# initialize a variable for later
DONE=""

# loop through each changed file
shopt -s nullglob
while read i; do
	# only care about changed Excel files
	if [[ "${i}" == *"Excel_Unzipped"* || "${i}" == *"Exported_VBA"* ]]; then
		# find the foldername/path where the non-binary files exist
		LOOP="true"
		FOLDERNAME="${i%/*}"
		while [[ "${FOLDERNAME}" == *"Excel_Unzipped"* || "${FOLDERNAME}" == *"Exported_VBA"* ]]; do
			FOLDERNAME="${FOLDERNAME%/*}"		
		done
		
		# ignore a file if a different non-binary file for that xlsm has already been done
		if [[ "${DONE}" != *"${FOLDERNAME}"* ]]; then
			# rezip Excel file as zip archive and then rename to an xlsm file
			7z a "${FOLDERNAME}.zip" "./${FOLDERNAME}/Excel_Unzipped/*" -uq0
			mv -fu "./${FOLDERNAME}.zip" "./${FOLDERNAME}.xlsm"
			
			# import the previous exported vba code
			cscript "import_vba.vbs" "${FILEPATH}/${FOLDERNAME}.xlsm"
			
			# mark the xlsm as done
			DONE+="${FOLDERNAME}"
		fi
	fi
done < changed.txt

# delete the text file
rm -f changed.txt