#!/bin/bash

# Check if target name is provided
if [ -z "$1" ]; then
  echo "Usage: $0 <target-name>"
  echo "Available targets in .clasp.all.json:"
  grep -o '"name": "[^"]*"' .clasp.all.json 2>/dev/null | cut -d'"' -f4 | sed 's/^/  - /'
  exit 1
fi

TARGET="$1"
FOUND=false

# Save original .clasp.json
mv .clasp.json .clasp.json.original 2>/dev/null

# Read and parse .clasp.all.json
while IFS= read -r line; do
  if [[ $line =~ \"name\":\ \"([^\"]+)\" ]]; then
    CURRENT_NAME="${BASH_REMATCH[1]}"
  elif [[ $line =~ \"scriptId\":\ \"([^\"]+)\" ]]; then
    CURRENT_SCRIPT="${BASH_REMATCH[1]}"
    # Remove any whitespace
    CURRENT_SCRIPT="${CURRENT_SCRIPT// /}"

    if [ "$CURRENT_NAME" = "$TARGET" ]; then
      FOUND=true
      echo "Pushing to $CURRENT_NAME..."

      # Create temporary .clasp.json for this project
      echo "{\"scriptId\":\"$CURRENT_SCRIPT\",\"rootDir\":\"./\"}" > .clasp.json

      # Push code using clasp
      clasp push -f
      break
    fi
  fi
done < .clasp.all.json

# Restore original .clasp.json
mv .clasp.json.original .clasp.json 2>/dev/null

if [ "$FOUND" = true ]; then
  echo "Push to $TARGET completed!"
else
  echo "Error: Target '$TARGET' not found in .clasp.all.json"
  exit 1
fi
