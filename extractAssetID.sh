#!/bin/bash

# Check that a JSON file was provided
if [ -z "$1" ]; then
    echo "Usage: $0 input.json"
    exit 1
fi

# Read each entry and print Name followed by Id
jq -r '.[] | "\(.Name)\n\(.Id)\n"' "$1"
