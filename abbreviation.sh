#!/usr/bin/env bash

sentence="$*"

abbr=$(printf '%s\n' "$sentence" \
  | perl -CSDA -pe 's/\p{P}//g' \
  | awk '{
      for (i = 1; i <= NF; i++) {
        printf toupper(substr($i, 1, 1))
      }
    }')

echo "$abbr"
