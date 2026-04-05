#!/usr/bin/env bash
set -euo pipefail

# ---------------------------------------------------------------------------
# Configuration — pass as env vars or edit defaults below
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

# Validate the file is a non-empty JSON array before doing anything
FILE_TYPE=$(jq -r 'type' "$INPUT_FILE" 2>/dev/null || echo "invalid")
if [[ "$FILE_TYPE" != "array" ]]; then
  echo "ERROR: $INPUT_FILE is not a JSON array (got: $FILE_TYPE)" >&2
  exit 1
fi

TOTAL=$(jq 'length' "$INPUT_FILE")

echo "========================================"
echo "  Importing: $INPUT_FILE"
echo "  Records:   $TOTAL"
echo "========================================"

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
    # Response is always BulkImportResponseDto — no type-checking needed
    IMPORTED=$(jq '.imported' /tmp/import_response.json)
    UPDATED=$(jq  '.updated'  /tmp/import_response.json)
    SKIPPED=$(jq  '.skipped'  /tmp/import_response.json)
    FAILED=$(jq   '.failed'   /tmp/import_response.json)

    echo "  OK — imported: $IMPORTED, updated: $UPDATED, skipped: $SKIPPED, failed: $FAILED"

    if [[ "$FAILED" -gt 0 ]]; then
      echo "  Failures:"
      jq -r '.errors[] | "    [\(.index // "?")]  \(.sourceSite // "?")/\(.externalId // "?"): \(.reason)"' \
        /tmp/import_response.json
    fi
  else
    echo "  ERROR — HTTP $HTTP_STATUS"
    cat /tmp/import_response.json
    exit 1
  fi

  OFFSET=$((OFFSET + BATCH_SIZE))
done

echo ""
echo "Done. $BATCH_NUM batch(es) sent."
