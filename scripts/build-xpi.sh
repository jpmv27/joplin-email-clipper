#!/usr/bin/env bash

readonly XPI_NAME=joplin-email-clipper
readonly SOURCE_DIR=joplin-email-clipper@thunderbird.jpmvrealtime.com
readonly INSTALL_RDF=$SOURCE_DIR/install.rdf

if ! zip --help &> /dev/null; then
  echo zip was not found in the PATH
  exit 2
fi

if ! grep --help &> /dev/null; then
  echo grep was not found in the PATH
  exit 2
fi

if [ ! -d $SOURCE_DIR ]; then
  echo Directory $SOURCE_DIR not found. Running from repo root?
  exit 2
fi

if [ ! -r $INSTALL_RDF ]; then
  echo $INSTALL_RDF not found
  exit 2
fi

if [[ "$(grep '<em:version>' $INSTALL_RDF)" =~ \>(.*)\< ]]; then
  readonly version=${BASH_REMATCH[1]}
else
  echo No version found in $INSTALL_RDF
  exit 2
fi

cd $SOURCE_DIR
zip -r ../$XPI_NAME-$version.xpi *
