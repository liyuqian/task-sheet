#!/bin/bash
set -e

clasp run $@ 2> >(tee stderr.log >&2)

if [ -s stderr.log ]; then
  exit 1;
fi
