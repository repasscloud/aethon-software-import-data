#!/usr/bin/env bash
set -euo pipefail

# ---------------------------------------------------------------------------
# Configuration — edit these or pass as env vars before running
# ---------------------------------------------------------------------------
API_HOST="${API_HOST:-}"
IMPORT_API_KEY="${IMPORT_API_KEY:-}"
INPUT_FILE="${1:-}"
BATCH_SIZE=500
# ---------------------------------------------------------------------------

if [[ ! -f "$INPUT_FILE" ]]; then
  echo "ERROR: File not found: $INPUT_FILE" >&2
  exit 1
fi

TOTAL=$(jq 'length' "$INPUT_FILE")
echo "Found $TOTAL records in $INPUT_FILE"

if [[ "$TOTAL" -eq 0 ]]; then
  echo "Nothing to import."
  exit 0
fi

ENDPOINT="${API_HOST}/api/v1/import/jobs/bulk"
BATCH_NUM=0
OFFSET=0

while [[ "$OFFSET" -lt "$TOTAL" ]]; do
  BATCH_NUM=$((BATCH_NUM + 1))
  END=$((OFFSET + BATCH_SIZE))
  ACTUAL_END=$(( END < TOTAL ? END : TOTAL ))
  COUNT=$((ACTUAL_END - OFFSET))

  echo ""
  echo "--- Batch $BATCH_NUM: records $OFFSET-$((ACTUAL_END - 1)) ($COUNT jobs) ---"

  HTTP_STATUS=$(
    jq ".[$OFFSET:$ACTUAL_END]" "$INPUT_FILE" \
    | curl -s -o /tmp/import_response.json -w "%{http_code}" \
        -X POST "$ENDPOINT" \
        -H "Content-Type: application/json" \
        -H "X-Import-Api-Key: $IMPORT_API_KEY" \
        -d @-
  )

  if [[ "$HTTP_STATUS" -eq 200 ]]; then
    IMPORTED=$(jq '[.[] | select(.wasDuplicate == false)] | length' /tmp/import_response.json)
    DUPES=$(jq '[.[] | select(.wasDuplicate == true)] | length' /tmp/import_response.json)
    echo "  OK — imported: $IMPORTED, duplicates skipped: $DUPES"
  else
    echo "  ERROR — HTTP $HTTP_STATUS"
    cat /tmp/import_response.json
    exit 1
  fi

  OFFSET=$((OFFSET + BATCH_SIZE))
done

echo ""
echo "Done. $BATCH_NUM batch(es) sent."
