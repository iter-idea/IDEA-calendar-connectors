#!/bin/bash

# other parameters
SRC_FOLDER='./src'
C='\033[4;32m' # color
NC='\033[0m'   # reset (no color)

# set the script to exit in case of errors
set -o errexit

# lint the code in search for errors
echo -e "${C}Linting...${NC}"
npm run lint ${SRC_FOLDER} 1>/dev/null

# compile the project's typescript code
echo -e "${C}Compiling...${NC}"
npm run compile 1>/dev/null

# generate documentation
echo -e "${C}Documenting...${NC}"
npm run docs 1>/dev/null