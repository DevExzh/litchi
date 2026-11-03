#!/bin/bash
# Native macOS check of the Litchi macOS verification queue (record 0430).
#
# Usage: run_on_mac.sh [--dry-run] [--resume] [--only P0001,P0002] [--app Keynote]
#                      [--timeout SECONDS] [--results FILE] OUT_DIR
#
# OUT_DIR is the directory tools/mac_queue/generate.py wrote (it holds
# queue.tsv, manifest.json and packages/). For every package listed in
# queue.tsv the runner:
#
#   1. verifies the package's SHA-256 against queue.tsv;
#   2. copies it to OUT_DIR/mac-run/<id>/ (the original stays untouched);
#   3. opens the copy in Pages, Numbers or Keynote, saves it as a re-saved
#      copy (Save As, so the application always writes), closes it, reopens
#      the re-saved copy, exports it to PDF and closes it;
#   4. records pass or fail (with the failing stage), the SHA-256 of the
#      re-saved package and of the PDF in the results TSV.
#
# The results TSV (default OUT_DIR/results.tsv) is what
# tools/mac_queue/apply_results.py folds back into
# docs/unified-ops/0363-mac-queue.md.
#
# --dry-run checks the platform, the applications and every package hash,
#   prints the plan and writes result "dry-run" lines; no application starts.
# --resume skips packages that already have a "pass" line in the results.
# --only and --app restrict the run (comma-separated package ids; one app).
#
# Expected applications: Keynote, Pages and Numbers 14.4 (the version of the
# tracked native receipts under test-data/iwork); later 14.x and 15.x are
# accepted and recorded. Written for the stock /bin/bash 3.2 of macOS.

set -u

DRY_RUN=0
RESUME=0
ONLY=""
ONLY_APP=""
TIMEOUT=180
RESULTS=""
OUT=""

usage() {
    sed -n '2,32p' "$0" | sed 's/^# \{0,1\}//'
    exit 2
}

while [ $# -gt 0 ]; do
    case "$1" in
        --dry-run) DRY_RUN=1 ;;
        --resume) RESUME=1 ;;
        --only) shift; ONLY="${1:-}" ;;
        --app) shift; ONLY_APP="${1:-}" ;;
        --timeout) shift; TIMEOUT="${1:-180}" ;;
        --results) shift; RESULTS="${1:-}" ;;
        -h|--help) usage ;;
        -*) echo "unknown option $1" >&2; usage ;;
        *) OUT="$1" ;;
    esac
    shift
done

[ -n "$OUT" ] || usage
OUT=$(cd "$OUT" 2>/dev/null && pwd) || { echo "no such directory" >&2; exit 2; }
QUEUE="$OUT/queue.tsv"
[ -f "$QUEUE" ] || { echo "$QUEUE is missing; run generate.py first" >&2; exit 2; }
[ -n "$RESULTS" ] || RESULTS="$OUT/results.tsv"
WORK="$OUT/mac-run"

HEADER="package_id	package_sha256	app	app_version	macos	result	stage	resaved_sha256	resaved_kind	pdf_sha256	started_utc	message"

sha_file() {
    shasum -a 256 "$1" | awk '{print $1}'
}

# SHA-256 of a package: the file's own hash, or for a package directory the
# hash of the sorted "<sha256>  <relative path>" listing of its files.
sha_package() {
    if [ -d "$1" ]; then
        (cd "$1" && find . -type f -print | LC_ALL=C sort | while IFS= read -r f; do
            printf '%s  %s\n' "$(shasum -a 256 "$f" | awk '{print $1}')" "$f"
        done) | shasum -a 256 | awk '{print $1}'
    else
        sha_file "$1"
    fi
}

app_path() {
    for candidate in "/Applications/$1.app" "/System/Applications/$1.app" "$HOME/Applications/$1.app"; do
        if [ -d "$candidate" ]; then
            echo "$candidate"
            return 0
        fi
    done
    return 1
}

app_version() {
    local path
    path=$(app_path "$1") || { echo "missing"; return 1; }
    defaults read "$path/Contents/Info" CFBundleShortVersionString 2>/dev/null || echo "unknown"
}

macos_version() {
    if command -v sw_vers >/dev/null 2>&1; then
        echo "macOS $(sw_vers -productVersion)"
    else
        echo "not-macos"
    fi
}

# One line of the results TSV; tabs and newlines in the message become spaces.
record() {
    local message
    message=$(printf '%s' "${12}" | tr '\t\n\r' '   ')
    printf '%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\n' \
        "$1" "$2" "$3" "$4" "$5" "$6" "$7" "$8" "$9" "${10}" "${11}" "$message" >> "$RESULTS"
}

already_passed() {
    [ -f "$RESULTS" ] && awk -F'\t' -v id="$1" '$1 == id && $6 == "pass" {found = 1} END {exit !found}' "$RESULTS"
}

selected() {
    local id="$1" app="$2"
    if [ -n "$ONLY_APP" ] && [ "$ONLY_APP" != "$app" ]; then
        return 1
    fi
    if [ -n "$ONLY" ]; then
        case ",$ONLY," in
            *",$id,"*) return 0 ;;
            *) return 1 ;;
        esac
    fi
    return 0
}

# The AppleScript of one check. The application name is substituted into the
# text because `export ... as PDF` needs the application's dictionary at
# compile time. Every stage reports "fail:<stage>:<error>" on failure.
script_for() {
    cat <<APPLESCRIPT
on run argv
    set inPath to item 1 of argv
    set savedPath to item 2 of argv
    set pdfPath to item 3 of argv
    set stage to "open"
    try
        with timeout of $TIMEOUT seconds
            tell application "$1"
                activate
                set theDoc to open ((POSIX file inPath) as alias)
                set stage to "save"
                save theDoc in (POSIX file savedPath)
                set stage to "close"
                close theDoc saving no
                set stage to "reopen"
                set again to open ((POSIX file savedPath) as alias)
                set stage to "export-pdf"
                export again to (POSIX file pdfPath) as PDF
                set stage to "close-reopened"
                close again saving no
            end tell
        end timeout
    on error message number code
        try
            tell application "$1" to close every document saving no
        end try
        return "fail:" & stage & ":" & code & " " & message
    end try
    return "ok"
end run
APPLESCRIPT
}

if [ ! -f "$RESULTS" ]; then
    printf '%s\n' "$HEADER" > "$RESULTS"
fi

MACOS=$(macos_version)
if [ "$DRY_RUN" -eq 0 ] && [ "$MACOS" = "not-macos" ]; then
    echo "this runner needs macOS (osascript); use --dry-run elsewhere" >&2
    exit 2
fi

echo "macOS queue run: $OUT ($MACOS); results $RESULTS"
for app in Keynote Numbers Pages; do
    version=$(app_version "$app")
    echo "  $app: $version (expected 14.4 or later)"
    case "$version" in
        missing|unknown) ;;
        1[4-9].*|[2-9][0-9].*) ;;
        *) echo "  warning: $app $version is older than the 14.4 of the native receipts" >&2 ;;
    esac
done

mkdir -p "$WORK"
total=0
passed=0
failed=0
skipped=0
started=$(date -u +%Y-%m-%dT%H:%M:%SZ)

# Read queue.tsv on descriptor 3 so osascript cannot consume it.
while IFS='	' read -r id app path sha rows <&3; do
    [ "$id" = "package_id" ] && continue
    [ -n "$id" ] || continue
    selected "$id" "$app" || continue
    if [ "$RESUME" -eq 1 ] && already_passed "$id"; then
        skipped=$((skipped + 1))
        continue
    fi
    total=$((total + 1))
    now=$(date -u +%Y-%m-%dT%H:%M:%SZ)
    version=$(app_version "$app")
    source="$OUT/$path"
    if [ ! -f "$source" ]; then
        record "$id" "$sha" "$app" "$version" "$MACOS" "fail" "verify" "-" "-" "-" "$now" "package $path is missing"
        failed=$((failed + 1))
        echo "$id $app FAIL verify (missing)"
        continue
    fi
    actual=$(sha_file "$source")
    if [ "$actual" != "$sha" ]; then
        record "$id" "$sha" "$app" "$version" "$MACOS" "fail" "verify" "-" "-" "-" "$now" "SHA-256 $actual differs from queue.tsv"
        failed=$((failed + 1))
        echo "$id $app FAIL verify (hash)"
        continue
    fi
    extension="${path##*.}"
    if [ "$DRY_RUN" -eq 1 ]; then
        record "$id" "$sha" "$app" "$version" "$MACOS" "dry-run" "-" "-" "-" "-" "$now" "would open $path for $rows queue rows"
        echo "$id $app dry-run: $path ($rows rows)"
        continue
    fi
    if [ "$version" = "missing" ]; then
        record "$id" "$sha" "$app" "$version" "$MACOS" "fail" "app" "-" "-" "-" "$now" "$app is not installed"
        failed=$((failed + 1))
        echo "$id $app FAIL app missing"
        continue
    fi
    dir="$WORK/$id"
    rm -rf "$dir"
    mkdir -p "$dir"
    input="$dir/$id-input.$extension"
    saved="$dir/$id-resaved.$extension"
    pdf="$dir/$id.pdf"
    cp "$source" "$input"
    script="$dir/check.applescript"
    script_for "$app" > "$script"
    outcome=$(osascript "$script" "$input" "$saved" "$pdf" 2>&1)
    status=$?
    if [ $status -ne 0 ]; then
        outcome="fail:osascript:$status $outcome"
    fi
    resaved_sha="-"
    resaved_kind="-"
    pdf_sha="-"
    if [ -e "$saved" ]; then
        resaved_sha=$(sha_package "$saved")
        if [ -d "$saved" ]; then resaved_kind="directory"; else resaved_kind="file"; fi
    fi
    [ -f "$pdf" ] && pdf_sha=$(sha_file "$pdf")
    case "$outcome" in
        ok)
            if [ "$resaved_sha" = "-" ] || [ "$pdf_sha" = "-" ]; then
                record "$id" "$sha" "$app" "$version" "$MACOS" "fail" "artifact" "$resaved_sha" "$resaved_kind" "$pdf_sha" "$now" "the application reported success but an artifact is missing"
                failed=$((failed + 1))
                echo "$id $app FAIL artifact"
            else
                record "$id" "$sha" "$app" "$version" "$MACOS" "pass" "-" "$resaved_sha" "$resaved_kind" "$pdf_sha" "$now" "-"
                passed=$((passed + 1))
                echo "$id $app pass $resaved_sha"
            fi
            ;;
        *)
            stage=$(printf '%s' "$outcome" | cut -d: -f2)
            message=$(printf '%s' "$outcome" | cut -d: -f3-)
            record "$id" "$sha" "$app" "$version" "$MACOS" "fail" "$stage" "$resaved_sha" "$resaved_kind" "$pdf_sha" "$now" "$message"
            failed=$((failed + 1))
            echo "$id $app FAIL $stage: $message"
            ;;
    esac
done 3< "$QUEUE"

echo "done ($started to $(date -u +%Y-%m-%dT%H:%M:%SZ)): $total checked, $passed pass, $failed fail, $skipped skipped (already passed)"
[ "$failed" -eq 0 ]
