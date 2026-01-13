#!/bin/sh
# Simple mock template processor that logs its arguments for tests
OUTFILE="$(dirname "$0")/mock-processor.output"
printf "%s %s\n" "$(date --iso-8601=seconds)" "$*" >> "$OUTFILE"
exit 0
