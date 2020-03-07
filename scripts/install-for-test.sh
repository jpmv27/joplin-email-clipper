#!/usr/bin/env bash

readonly SOURCE_DIR=joplin-email-clipper@thunderbird.jpmvrealtime.com

if [ ! -d $SOURCE_DIR ]; then
  echo Directory $SOURCE_DIR not found. Running from repo root?
  exit 2
fi

readonly EXTENSIONS_DIR=$(readlink -f ~/.thunderbird/*.default/extensions)
readonly TARGET_DIR=$(readlink -f ./$SOURCE_DIR)
echo $TARGET_DIR > $EXTENSIONS_DIR/$SOURCE_DIR
