#!/bin/bash

# Save original .clasp.json
mv .clasp.json .clasp.json.original 2>/dev/null

# Read and parse .clasp.all.json
while IFS= read -r line; do
  if [[ $line =~ \"name\":\ \"([^\"]+)\" ]]; then
    name="${BASH_REMATCH[1]}"
    echo "Pushing to $name..."
  elif [[ $line =~ \"scriptId\":\ \"([^\"]+)\" ]]; then
    scriptId="${BASH_REMATCH[1]}"
    # Remove any whitespace that might have snuck in
    scriptId="${scriptId// /}"

    # Create temporary .clasp.json for this project
    echo "{\"scriptId\":\"$scriptId\",\"rootDir\":\"./\"}" > .clasp.json

    # Push code using clasp
    clasp push -f
  fi
done < .clasp.all.json

# Restore original .clasp.json
mv .clasp.json.original .clasp.json 2>/dev/null

echo "All pushes completed!"
