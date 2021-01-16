#!/bin/bash
set -e

COMMANDS=(
  'sudo npm install -g @google/clasp'
  'npm i -S @types/google-apps-script'
  'npx eslint .'
  'clasp push'
  'bash test/check_clasp_run.sh testAll'
)

for ((i = 0; i < ${#COMMANDS[@]}; i += 1))
do
  echo ${COMMANDS[$i]}
  ${COMMANDS[$i]}
  echo ''
done
