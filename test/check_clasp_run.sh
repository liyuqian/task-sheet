#!/bin/bash
set -e

clasp run $@ 2> >(tee stderr.log >&2)

if [[ $(<stderr.log) == "- Running function: $@" ]]; then
  exit 0
fi

if [ -s stderr.log ]; then
  exit 1;
fi
