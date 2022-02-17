#!/bin/sh
#
# This script only needs to run once, immediately after cloning the GitHub repo locally
# This script runs the post-merge.sh script once to create the xlsm files, and then copies
# and renames the post-merge.sh and pre-commit.sh scripts to the .git/hooks directory
# mtilasha@ford.com
# Copyright: Ford Motor Company Limited

# copy scripts to .git/hooks directory
cp post-merge.sh ./.git/hooks
cp pre-commit.sh ./.git/hooks

# rename scripts so they can be run as hooks
cd .git/hooks
mv -f post-merge.sh post-merge
mv -f pre-commit.sh pre-commit

# run the post-merge script
cd -
sh post-merge.sh
