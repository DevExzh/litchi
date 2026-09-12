#!/usr/bin/env bash
set -euo pipefail

script_dir=$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)
repo=${SOURCE_CHECKOUT:-$(git -C "$script_dir" rev-parse --show-toplevel)}
candidate_commit=8e3ad310426a534c0bb17a789eb2c42a96e61310
base_commit=2e023cbd253a9b46ec2582e5319cbcbcddac5433

if ! git -C "$repo" cat-file -e "$candidate_commit^{commit}"; then
  echo "candidate commit $candidate_commit is unavailable" >&2
  exit 2
fi
if ! git -C "$repo" diff --quiet "$candidate_commit" -- \
  crates/litchi-odf-common/src/core/mod.rs \
  crates/litchi-odf-common/src/lib.rs \
  crates/litchi-odf-common/src/namespace.rs \
  crates/litchi-ods/src/advanced.rs \
  crates/litchi-ods/src/authoring/builder.rs \
  crates/litchi-ods/src/data_style/mod.rs \
  crates/litchi-ods/src/data_style/source.rs \
  crates/litchi-ods/src/document.rs \
  crates/litchi-ods/src/facade/mod.rs \
  crates/litchi-ods/src/open_parse.rs \
  crates/litchi-ods/src/package/document.rs \
  crates/litchi-ods/src/worksheet/codec.rs \
  crates/litchi-odf-common/src/core/resolved_reader.rs \
  crates/litchi-odf-common/tests/resolved_reader.rs \
  crates/litchi-ods/tests/data_style_vocabulary.rs \
  crates/litchi-odt/src/auto_mark_file/codec.rs \
  crates/litchi-odt/src/bibliography_configuration/codec.rs \
  crates/litchi-odt/src/content_metadata.rs \
  crates/litchi-odt/src/dde_connection/codec.rs \
  crates/litchi-odt/src/elements/xml.rs \
  crates/litchi-odt/src/flat/mod.rs \
  crates/litchi-odt/src/generic/codec.rs \
  crates/litchi-odt/src/variable_declaration/codec.rs \
  crates/litchi-odt/src/variable_declaration/package.rs \
  crates/litchi-odt/src/xforms.rs \
  crates/litchi-odt/tests/flat_text_templates.rs; then
  echo "source files differ from candidate commit $candidate_commit" >&2
  exit 2
fi

manifest="$repo/docs/report/spec-gap-validation-evidence/ods-data-style-source-performance/harness/Cargo.toml"
if [[ ! -f "$manifest" ]]; then
  echo "profile harness is missing: $manifest" >&2
  exit 2
fi

output_parent=${ODS_PROFILE_OUTPUT_PARENT:-/var/tmp}
target_parent=${ODS_PROFILE_TARGET_PARENT:-/var/tmp}
output_dir=$(mktemp -d "$output_parent/ods-data-style-source-profile-replay-output.XXXXXX")
target_dir=$(mktemp -d "$target_parent/ods-data-style-source-profile-replay-target.XXXXXX")
repeats=${ODS_PROFILE_REPEATS:-7}

echo "base_commit=$base_commit"
echo "candidate_commit=$candidate_commit"
echo "source_checkout=$repo"
echo "output_dir=$output_dir"
echo "target_dir=$target_dir"
echo "repeats=$repeats"

ODS_PROFILE_OUTPUT_DIR="$output_dir" \
ODS_PROFILE_REPEATS="$repeats" \
CARGO_TARGET_DIR="$target_dir" \
cargo run --manifest-path "$manifest" --locked --offline --release
