#!/bin/sh
# Builds the DOCXClass samples: docxwriter, docxtemplate, docxreader, docxchart
#   ./build.sh              with the default C compiler of hbmk2
#   ./build.sh -comp=clang  other hbmk2 options are passed on
# HBMK2  hbmk2 to use (default: hbmk2 from the PATH)
cd "$(dirname "$0")" || exit 1
HBMK2=${HBMK2:-hbmk2}
FAILED=
for p in docxwriter docxtemplate docxreader docxchart; do
    echo
    echo "=== $p"
    "$HBMK2" "$p.hbp" "$@" || FAILED="$FAILED $p"
done
echo
if [ -n "$FAILED" ]; then
    echo "Not built:$FAILED"
    exit 1
fi
echo "Built: docxwriter docxtemplate docxreader docxchart"
