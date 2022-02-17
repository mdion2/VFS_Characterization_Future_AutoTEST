#!/bin/sh
#
# This script is called by "git commit" with no arguments.
# This script prepares Excel files for commit in a non-binary format.
# The script removes any staged xlsm files, unzips them, exports the
# vba code, and then stages the unzipped/exported non-binary files to be commited.
# The hook should exit with non-zero status after issuing an appropriate message if it wants to stop the commit.
# mtilasha@ford.com
# Copyright: Ford Motor Company Limited

# get the file path to the Excel files as the present working directory
FILEPATH="$(pwd)"
cd "${FILEPATH}"	# move into the file path

# create a text file and output the names of all files currently staged
touch staged.txt
git diff --name-only --cached > staged.txt

# loop through each staged file
shopt -s nullglob
while read i; do

	# if it is a xlsm file
	if [ "${i##*.}" = "xlsm" ]; then
		# create a backup of the Excel file
		mkdir -p "${FILEPATH}/Backups"
		cp -f "${FILEPATH}/${i}" "${FILEPATH}/Backups/${i##*/}"
		
		# unstage the Excel file
		git restore --staged "${i}"
	
		# set the name of the folder to be created equal to the name of the Excel file
		FOLDERNAME="${i%.*}"
		
		# unzip the Excel file and delete the vbaProject.bin file
		7z x "${FILEPATH}/${i}" -o"${FILEPATH}/${FOLDERNAME}/Excel_Unzipped" -aoa
		find "${FILEPATH}/${FOLDERNAME}/Excel_Unzipped" -name "vbaProject.bin" -delete

		# export the vba code
		cscript "export_vba.vbs" "${FILEPATH}/${i}"

		# stage the files in git
		git add "${FILEPATH}/${FOLDERNAME}"
	fi	
done < staged.txt

# delete the text file
rm -f staged.txt