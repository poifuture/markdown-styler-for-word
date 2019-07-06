#!/bin/bash
# The only reason this script exists is for cross platform on Windows
if [ "$1" != "build" ]; then
  # pwd must be project root
  echo "Please run 'yarn build'"
fi

version=$(cat nextversion)

cat ./manifest.xml \
  | sed "s/Markdown[ ]Styler[ ]Dogfood/Markdown Styler/g" \
  | sed "s/https[:][/][/]localhost[:]3000/https:\/\/poifuture.github.io\/markdown-styler-for-word/g" \
  | sed "s/[<]Version[>].*[<][/]Version[>]/<Version>$(date '+%Y').$(date '+%m%d').$version.0<\/Version>/g" \
  > ./dist/manifest.xml

cat ./src/README.md \
  | sed "s/^.*prettier-ignore.*$/<!-- This file is auto generated, change src\/README.md instead. -->/g" \
  | prettier --parser markdown > ./README.md

echo $((version+1)) > nextversion
