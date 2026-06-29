#!/usr/bin/env bash

read -r input

result=$(echo "$input" \
  | sed 's/[^[:alnum:]]\+/ /g' \
  | awk '{
      for (i = 1; i <= NF; i++) {
        $i = toupper(substr($i,1,1)) tolower(substr($i,2))
      }
      gsub(/ /, "")
      print
    }')

printf "%s\n" "$result"
