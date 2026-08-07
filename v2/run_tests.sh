#!/usr/bin/env bash
#
# Full regression across every phase, from an empty database.
#
#   ./v2/run_tests.sh
#
# Builds a throwaway database, applies the schema, imports the real workbook,
# runs every test suite, then drops the database. Nothing touches your live
# data and nothing is left behind.
#
# Override the server with PGHOST/PGPORT/PGUSER, e.g.
#   PGHOST=/tmp PGPORT=5433 ./v2/run_tests.sh

set -euo pipefail

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
TEST_DB="${TEST_DB:-cmf_regression}"
WORKBOOK="${WORKBOOK:-$ROOT/CMF WIP - Schedule.xlsx}"

PSQL_ARGS=()
[[ -n "${PGHOST:-}"  ]] && PSQL_ARGS+=(-h "$PGHOST")
[[ -n "${PGPORT:-}"  ]] && PSQL_ARGS+=(-p "$PGPORT")
[[ -n "${PGUSER:-}"  ]] && PSQL_ARGS+=(-U "$PGUSER")

say()  { printf '\n\033[1;34m▸ %s\033[0m\n' "$*"; }
fail() { printf '\n\033[1;31m✗ %s\033[0m\n' "$*"; exit 1; }

cleanup() {
  psql "${PSQL_ARGS[@]}" -d postgres -q \
       -c "DROP DATABASE IF EXISTS $TEST_DB" >/dev/null 2>&1 || true
}
trap cleanup EXIT

say "Creating throwaway database: $TEST_DB"
cleanup
psql "${PSQL_ARGS[@]}" -d postgres -q -c "CREATE DATABASE $TEST_DB" \
  || fail "could not create database — is Postgres running?"

say "Applying schema (db/schema.sql)"
psql "${PSQL_ARGS[@]}" -d "$TEST_DB" -v ON_ERROR_STOP=1 -q -f "$ROOT/db/schema.sql" \
  || fail "schema failed to apply"

# Build a libpq URL the app and tests can share.
URL="postgresql://"
[[ -n "${PGUSER:-}" ]] && URL+="${PGUSER}@"
URL+="/${TEST_DB}"
[[ -n "${PGHOST:-}" ]] && URL+="?host=${PGHOST}"
[[ -n "${PGHOST:-}" && -n "${PGPORT:-}" ]] && URL+="&port=${PGPORT}"
export DATABASE_URL="$URL"

if [[ -f "$WORKBOOK" ]]; then
  say "PHASE 1 — importing the real workbook"
  python3 "$ROOT/v2/import_workbook.py" --database-url "$DATABASE_URL" \
    ${HISTORY:+--history "$HISTORY"} 2>&1 \
    | grep -Ev 'UserWarning|warn\(msg\)' || fail "import failed"
else
  say "PHASE 1 — workbook not found, importing nothing"
  echo "  (set WORKBOOK=/path/to/file.xlsx to include the import)"
fi

say "PHASES 1-3 — test suites"
python3 -m pytest "$ROOT/v2/tests" -q --tb=short || fail "tests failed"

if [[ "${E2E:-1}" == "1" ]]; then
  say "END-TO-END — driving a real browser through the full workflow"
  python3 "$ROOT/v2/tests/e2e.py" || fail "end-to-end run failed"
else
  say "END-TO-END — skipped (E2E=0)"
fi

printf '\n\033[1;32m✓ Full regression passed\033[0m\n\n'
