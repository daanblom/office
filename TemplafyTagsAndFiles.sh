#!/usr/bin/env bash

input_file="$1"
delimiter=";"

jq -r '
  .[] |
  "\(.Name),\(.Tags | map(select(test("^[0-9]+$"))) | join(","))"
' "$input_file"
