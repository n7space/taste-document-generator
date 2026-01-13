#!/bin/sh
# Simple mock template processor that logs its arguments for tests
BASEDIR="$(dirname "$0")"
OUTFILE="$BASEDIR/mock-processors.output"
printf "%s %s\n" "$(date --iso-8601=seconds)" "$*" >> "$OUTFILE"

# Extract -o and -t parameters
OUTDIR=""
TFILE=""
while [ $# -gt 0 ]; do
	case "$1" in
		-o) OUTDIR="$2"; shift 2;;
		-t) TFILE="$2"; shift 2;;
		--) shift; break;;
		*) shift;;
	esac
done

if [ -n "$OUTDIR" ] && [ -n "$TFILE" ]; then
	BASENAME=$(basename "$TFILE")
	DESTPATH="$OUTDIR/$BASENAME"
	SRC="$BASEDIR/test_in_tmplt.docx"
	if [ -f "$SRC" ]; then
		mkdir -p "$OUTDIR"
		cp "$SRC" "$DESTPATH"
		echo "Copied $SRC to $DESTPATH" >> "$OUTFILE"
		exit 0
	else
		echo "Source template $SRC not found" >> "$OUTFILE"
		exit 2
	fi
fi

exit 0
