#!/bin/bash
# The only reason this script exists is for cross platform on Windows
if [ "$1" != "node" ]; then
  # pwd must be project root
  echo "Please run 'npm run build:readme'"
fi

echo '<!-- This file is auto generated, change src/README.md.src instead. -->' | cat - ./src/README.md.src | prettier --parser markdown > ./README.md
