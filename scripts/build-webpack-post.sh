#!/bin/bash
# The only reason this script exists is for cross platform on Windows
if [ "$1" != "build" ]; then
  # pwd must be project root
  echo "Please run 'yarn build'"
fi

cat ./manifest.xml \
  | sed "s/Markdown[ ]Styler[ ]Dogfood/Markdown Styler/g" \
  | sed "s/https[:][/][/]localhost[:]3000/https:\/\/poifuture.github.io\/markdown-styler-for-word/g" \
  | sed "s/[<]Version[>].*[<][/]Version[>]/<Version>$(date '+%y').$(date '+%m').$(date '+%d').1<\/Version>/g" \
  | sed "s/[<]Id[>].*[<][/]Id[>]/<Id>05c2e1c9-3e1d-406e-9a91-e9ac64854143<\/Id>/g" \
  > ./dist/manifest.xml

cp ./src/index.html ./dist/index.html
cp ./src/_config.yml ./dist/_config.yml
cp ./Markdown.dotx ./dist/Markdown.dotx

cat ./src/README.md \
  | sed "s/^.*prettier-ignore.*$/<!-- This file is auto generated, change src\/README.md instead. -->/g" \
  | prettier --parser markdown > ./README.md
