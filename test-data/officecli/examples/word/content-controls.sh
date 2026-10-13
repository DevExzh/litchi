#!/bin/bash
# content-controls.sh — exercise the docx `sdt` (structured document tag /
# content control) property surface (schemas/help/docx/sdt.json) using the
# officecli CLI directly.
#
# An `sdt` is a content control: a bounded region Word treats as a form field.
# Block-level SDTs (added at /body) wrap paragraphs. Every control shares
# alias/tag/lock/placeholder; each variant adds its own props:
#   text      — plain-text field           (text=)
#   richText  — rich formatted field       (text=)
#   dropDown  — pick-one, typing disabled   (items=, dropDown.lastValue=)
#   comboBox  — pick-one or free type       (items=, comboBox.lastValue=)
#   date      — date picker                 (format=, date.fullDate/calendar/lid/storeMappedDataAs=)
#   picture   — image placeholder
#   group     — a locked grouping wrapper
#   checkbox  — a check-box toggle          (checked=)
# CLI twin of content-controls.py (officecli SDK); both produce an equivalent
# content-controls.docx modelling a small employee-intake form.
# NOTE: intentionally NO `set -e`. Like the SDK twin's doc.batch, this script
# tolerates forward-compat 'UNSUPPORTED props' warnings (officecli exit 2) and
# keeps building so the full document is produced.
FILE="$(dirname "$0")/content-controls.docx"
echo "Building $FILE ..."
rm -f "$FILE"
officecli create "$FILE"
officecli open "$FILE"

# --- Title + intro ---
officecli add "$FILE" /body --type paragraph --prop text="Employee Intake Form" --prop style=Title
officecli add "$FILE" /body --type paragraph --prop text="Complete every field. Grey boxes are content controls — click one and Word treats it as a single fillable region." --prop style=Subtitle

# --- 1. plainText — the employee's full name ---
# Features: type=text, alias (Word label), tag (data-binding key),
#           text= (initial/placeholder content), lock, placeholderText.
officecli add "$FILE" /body --type paragraph --prop text="Full name" --prop style=Heading2
officecli add "$FILE" /body --type sdt --prop type=text \
  --prop alias="Full Name" --prop tag=fullName \
  --prop text="[Enter full legal name]" \
  --prop lock=unlocked --prop placeholderText=DefaultPlaceholder

# --- 2. dropDown — department (pick-one, typing disabled) ---
# Features: type=dropdown, items= (comma list), dropDown.lastValue= (current pick).
officecli add "$FILE" /body --type paragraph --prop text="Department" --prop style=Heading2
officecli add "$FILE" /body --type sdt --prop type=dropdown \
  --prop alias="Department" --prop tag=department \
  --prop items="Sales,Engineering,Human Resources,Finance,Operations" \
  --prop dropDown.lastValue=Engineering

# --- 3. comboBox — office location (pick-one OR free-type) ---
# Features: type=combobox, items= with display|value form, comboBox.lastValue=.
officecli add "$FILE" /body --type paragraph --prop text="Primary office" --prop style=Heading2
officecli add "$FILE" /body --type sdt --prop type=combobox \
  --prop alias="Office Location" --prop tag=office \
  --prop items="New York|NYC,London|LON,Singapore|SIN,Remote|REMOTE" \
  --prop comboBox.lastValue=LON

# --- 4. date — start date (calendar picker) ---
# Features: type=date, format= (display mask), date.fullDate (ISO value),
#           date.calendar, date.lid (locale), date.storeMappedDataAs.
officecli add "$FILE" /body --type paragraph --prop text="Start date" --prop style=Heading2
officecli add "$FILE" /body --type sdt --prop type=date \
  --prop alias="Start Date" --prop tag=startDate \
  --prop format="yyyy-MM-dd" \
  --prop date.fullDate=2026-02-01T00:00:00Z \
  --prop date.calendar=gregorian --prop date.lid=en-US \
  --prop date.storeMappedDataAs=dateTime

# --- 5. picture — profile photo placeholder ---
# Features: type=picture. Word shows the image-insert placeholder; alias/tag apply.
officecli add "$FILE" /body --type paragraph --prop text="Profile photo" --prop style=Heading2
officecli add "$FILE" /body --type sdt --prop type=picture \
  --prop alias="Profile Photo" --prop tag=photo

# --- 6. richText — reviewer notes (formatted, multi-run field) ---
# Features: type=richtext, text=. Rich content controls allow bold/colour/etc.
#           inside; lock=contentLocked freezes the content (editable=false).
officecli add "$FILE" /body --type paragraph --prop text="Reviewer notes" --prop style=Heading2
officecli add "$FILE" /body --type sdt --prop type=richtext \
  --prop alias="Reviewer Notes" --prop tag=notes \
  --prop text="Manager may add formatted commentary here." \
  --prop lock=contentLocked

# --- 7. group — a locked wrapper around an approval line ---
# Features: type=group, lock=sdtContentLocked. A group SDT bundles content so
#           the whole region behaves as one unit; sdtContentLocked blocks both
#           deletion of the control and edits to its contents.
officecli add "$FILE" /body --type paragraph --prop text="Approval" --prop style=Heading2
officecli add "$FILE" /body --type sdt --prop type=group \
  --prop alias="Approval Block" --prop tag=approval \
  --prop text="Approved by HR — signature on file." \
  --prop lock=sdtContentLocked

# --- 8. checkbox — the HR approval toggle (real Word check-box control) ---
# Features: type=checkbox, checked= (true → ☒ / false → ☐). Emits a <w14:checkbox>
#           content control; Word renders a clickable box that flips 2612/2610.
officecli add "$FILE" /body --type paragraph --prop text="HR approved" --prop style=Heading2
officecli add "$FILE" /body --type sdt --prop type=checkbox \
  --prop alias="Approved" --prop tag=hrApproved \
  --prop checked=true

# --- Post-add tweak via `set` (alias / tag / lock / text are settable) ---
# Rename the department control's Word label and lock it against deletion.
# (Per-type props like dropDown.lastValue are add/get-only, not settable.)
officecli set "$FILE" /body/sdt[2] --prop alias="Home Department" --prop lock=sdtLocked

officecli close "$FILE"
officecli validate "$FILE"
echo "Created: $FILE"
