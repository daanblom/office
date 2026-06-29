#!/usr/bin/env bash

jq -r '
  .[] |
  "\(.Name),\(.Tags | map(select(test("^[0-9]+$"))) | join(","))"
' "$1"
