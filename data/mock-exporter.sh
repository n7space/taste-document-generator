#!/bin/sh
BASEDIR="$(dirname "$0")"
LOGFILE="$BASEDIR/mock-exporter.output"
printf "%s %s\n" "$(date --iso-8601=seconds)" "$*" >> "$LOGFILE"

OUTPUT=""
TYPE=""

while [ $# -gt 0 ]; do
	case "$1" in
		--output)
			OUTPUT="$2"
			shift 2
			;;
		--system-object-type)
			TYPE="$2"
			shift 2
			;;
		*)
			shift
			;;
	esac
done

if [ -n "$OUTPUT" ]; then
	mkdir -p "$(dirname "$OUTPUT")"
	{
		printf "name,value\n"
		printf "%s,%s\n" "${TYPE:-system-object}" "${TYPE:-value}"
	} > "$OUTPUT"
fi

exit 0