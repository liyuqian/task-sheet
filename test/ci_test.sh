#!/bin/bash
set -e

COMMANDS=(
  'sudo npm install -g @google/clasp'
  'npm install'
  'npm run lint'
  'npm run build'
  'npm run deploy'
  'npm run test'
)

for ((i = 0; i < ${#COMMANDS[@]}; i += 1))
do
  echo ${COMMANDS[$i]}
  ${COMMANDS[$i]}
  echo ''
done
