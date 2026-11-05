#!/bin/bash
# data-validation.sh — exercise the full xlsx `validation` (dataValidation)
# feature surface (schemas/help/xlsx/validation.json) using the officecli CLI.
#
# 6 sheets, one validation family each: list (inline + range), number
# (whole/decimal), date & time, text length, custom formula, and the
# input-prompt / error-message / errorStyle surface. CLI twin of
# data-validation.py (officecli SDK); both produce an equivalent
# data-validation.xlsx.
#
# Each rule is one `add --type validation` against the sheet, with type=
# selecting the rule kind and ref= (alias sqref) the target range. The rule
# lands at /SheetName/dataValidation[N].
# NOTE: intentionally NO `set -e`. Like the SDK twin's doc.batch, this script
# tolerates forward-compat 'UNSUPPORTED props' warnings (officecli exit 2) and
# keeps building so the full document is produced.
FILE="$(dirname "$0")/data-validation.xlsx"
echo "Building $FILE ..."
rm -f "$FILE"
officecli create "$FILE"
officecli open "$FILE"

# helper: write a bold-blue header in <cell> then a column of values below it
hdr() { officecli set "$FILE" "/$1/$2" --prop value="$3" --prop font.bold=true --prop fill=1F4E79 --prop font.color=FFFFFF; }
col() {  # col <sheet> <col-letter> <start-row> <v1> <v2> ...
  local sheet="$1" c="$2" r="$3"; shift 3
  for v in "$@"; do officecli set "$FILE" "/$sheet/$c$r" --prop value="$v"; r=$((r+1)); done
}
dv() { officecli add "$FILE" "/$1" --type validation "${@:2}"; }   # dv <sheet> --prop ...
sheet() { officecli add "$FILE" / --type sheet --prop name="$1"; }

# ==========================================================================
# Sheet1: List — inline CSV list AND range-based list (helper column)
# ==========================================================================
# Sheet1 exists by default; rename the entry column headers + a helper list.
hdr Sheet1 A1 "Status (inline)"
hdr Sheet1 B1 "Priority (range)"
hdr Sheet1 H1 "Priorities"
col Sheet1 H 2 Low Medium High Critical           # the range the B-column list points at

# Features: type=list, inline formula1 CSV, inCellDropdown default (dropdown arrow shown)
dv Sheet1 --prop type=list --prop ref=A2:A20 --prop formula1="Draft,Review,Approved,Rejected"
# Features: type=list, range-based formula1 (=$H$2:$H$5), sqref alias for ref, inCellDropdown=false (hide arrow)
dv Sheet1 --prop type=list --prop sqref=B2:B20 --prop 'formula1==$H$2:$H$5' --prop inCellDropdown=false

# ==========================================================================
# Sheet2: Number — whole & decimal with every comparison operator
# ==========================================================================
sheet Number
hdr Number A1 "Qty (whole)"
hdr Number B1 "Discount (decimal)"
hdr Number C1 "Rating (1-5)"
hdr Number D1 "Price (>0)"
hdr Number E1 "Not 13"
col Number A 2 1 5 10 20        # sample entries
col Number B 2 0.05 0.10 0.25
col Number C 2 1 3 5
col Number D 2 9.99 19.5
col Number E 2 12 14

# Features: type=whole, operator=between, formula1+formula2 (both bounds)
dv Number --prop type=whole --prop ref=A2:A50 --prop operator=between --prop formula1=1 --prop formula2=100
# Features: type=decimal, operator=lessThanOrEqual (single bound formula1)
dv Number --prop type=decimal --prop ref=B2:B50 --prop operator=lessThanOrEqual --prop formula1=0.5
# Features: type=whole, operator=greaterThanOrEqual
dv Number --prop type=whole --prop ref=C2:C50 --prop operator=greaterThanOrEqual --prop formula1=1
# Features: type=decimal, operator=greaterThan
dv Number --prop type=decimal --prop ref=D2:D50 --prop operator=greaterThan --prop formula1=0
# Features: type=whole, operator=notEqual (exclude a single value)
dv Number --prop type=whole --prop ref=E2:E50 --prop operator=notEqual --prop formula1=13

# ==========================================================================
# Sheet3: Date & Time — date range, time window, exact-equal date
# ==========================================================================
sheet DateTime
hdr DateTime A1 "Event date (2024)"
hdr DateTime B1 "Shift start (9-17)"
hdr DateTime C1 "Deadline (=EOY)"
hdr DateTime D1 "Ship after"
col DateTime A 2 2024-03-15 2024-07-01
col DateTime B 2 10:30:00 14:00:00
col DateTime D 2 2024-06-01

# Features: type=date, operator=between, formula1+formula2 date bounds (stored as serials)
dv DateTime --prop type=date --prop ref=A2:A50 --prop operator=between --prop formula1=2024-01-01 --prop formula2=2024-12-31
# Features: type=time, operator=between, formula1+formula2 time window (stored as day fractions)
dv DateTime --prop type=time --prop ref=B2:B50 --prop operator=between --prop formula1=09:00:00 --prop formula2=17:00:00
# Features: type=date, operator=equal (exact date)
dv DateTime --prop type=date --prop ref=C2:C50 --prop operator=equal --prop formula1=2024-12-31
# Features: type=date, operator=greaterThan (on/after a date)
dv DateTime --prop type=date --prop ref=D2:D50 --prop operator=greaterThan --prop formula1=2024-01-01

# ==========================================================================
# Sheet4: Text length — bounded, capped, exact, and a notBetween band
# ==========================================================================
sheet TextLength
hdr TextLength A1 "Username (3-16)"
hdr TextLength B1 "Country code (=2)"
hdr TextLength C1 "Tweet (<=280)"
hdr TextLength D1 "PIN (not 5-7)"
col TextLength A 2 alice bob_smith
col TextLength B 2 US GB
col TextLength C 2 "hello world"
col TextLength D 2 1234 12345678

# Features: type=textLength, operator=between, formula1+formula2 length bounds
dv TextLength --prop type=textLength --prop ref=A2:A50 --prop operator=between --prop formula1=3 --prop formula2=16
# Features: type=textLength, operator=equal (exact length)
dv TextLength --prop type=textLength --prop ref=B2:B50 --prop operator=equal --prop formula1=2
# Features: type=textLength, operator=lessThanOrEqual (cap)
dv TextLength --prop type=textLength --prop ref=C2:C50 --prop operator=lessThanOrEqual --prop formula1=280
# Features: type=textLength, operator=notBetween, formula1+formula2 excluded band
dv TextLength --prop type=textLength --prop ref=D2:D50 --prop operator=notBetween --prop formula1=5 --prop formula2=7

# ==========================================================================
# Sheet5: Custom formula — arbitrary boolean expressions
# ==========================================================================
sheet Custom
hdr Custom A1 "Must be number"
hdr Custom B1 "Even only"
hdr Custom C1 "No spaces"
col Custom A 2 42 3.14
col Custom B 2 2 4
col Custom C 2 nospace ok

# Features: type=custom, formula1 = boolean expression (allow only numeric entries)
dv Custom --prop type=custom --prop ref=A2:A50 --prop formula1="ISNUMBER(A2)"
# Features: type=custom, formula1 = MOD expression (allow only even numbers)
dv Custom --prop type=custom --prop ref=B2:B50 --prop formula1="MOD(B2,2)=0"
# Features: type=custom, formula1 = reject any value containing a space
dv Custom --prop type=custom --prop ref=C2:C50 --prop formula1="ISERROR(FIND(\" \",C2))"

# ==========================================================================
# Sheet6: Messages — input prompt, error message, and all three errorStyles
# ==========================================================================
sheet Messages
hdr Messages A1 "Age (stop)"
hdr Messages B1 "Budget (warning)"
hdr Messages C1 "Note (information)"
hdr Messages D1 "Allow blank=false"

# Features: prompt + promptTitle + showInput (input message shown when cell selected)
#           + error + errorTitle + showError + errorStyle=stop (hard block)
dv Messages --prop type=whole --prop ref=A2:A50 --prop operator=between --prop formula1=18 --prop formula2=120 \
  --prop promptTitle="Enter age" --prop prompt="Age must be 18-120" --prop showInput=true \
  --prop errorTitle="Invalid age" --prop error="Please enter a whole number 18-120" --prop showError=true \
  --prop errorStyle=stop
# Features: errorStyle=warning (soft block — user may override)
dv Messages --prop type=decimal --prop ref=B2:B50 --prop operator=lessThanOrEqual --prop formula1=10000 \
  --prop promptTitle="Budget cap" --prop prompt="Suggested max is 10,000" \
  --prop errorTitle="Over budget" --prop error="That exceeds the cap — continue anyway?" \
  --prop errorStyle=warning
# Features: errorStyle=information (advisory only — never blocks)
dv Messages --prop type=textLength --prop ref=C2:C50 --prop operator=lessThanOrEqual --prop formula1=200 \
  --prop promptTitle="Note" --prop prompt="Keep notes under 200 chars" \
  --prop errorTitle="Long note" --prop error="This note is quite long." \
  --prop errorStyle=information
# Features: allowBlank=false (empty cells are themselves invalid) + showInput=false (no prompt bubble)
dv Messages --prop type=whole --prop ref=D2:D50 --prop operator=greaterThan --prop formula1=0 \
  --prop allowBlank=false --prop showInput=false --prop showError=true \
  --prop errorTitle="Required" --prop error="This field is required and must be > 0"

officecli close "$FILE"
officecli validate "$FILE"
echo "Created: $FILE"
