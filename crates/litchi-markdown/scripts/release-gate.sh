#!/bin/sh
set -eu

script_dir=$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)
crate_dir=$(CDPATH= cd -- "$script_dir/.." && pwd)
workspace_dir=$(CDPATH= cd -- "$crate_dir/../.." && pwd)

cd "$workspace_dir"

printf '%s  %s\n' \
  d431b29d97b6f73e69d547109cf5081578fac931e72afe95639ebe766c1b2a20 \
  crates/litchi-markdown/tests/corpus/commonmark-0.31.2/spec.json \
  28a9529c7d0bb4dc51f4bf5c116a3d16ef247a052f7591466768ddf563fd1cf5 \
  crates/litchi-markdown/tests/corpus/commonmark-0.31.2/LICENSE-CC-BY-SA-4.0.txt \
  89cfcb21173de246f141ef6832395b74d45a23595ddf65bf6ffb0334d3e7c651 \
  crates/litchi-markdown/tests/corpus/gfm-0.29.0.gfm.13/spec.json \
  28a9529c7d0bb4dc51f4bf5c116a3d16ef247a052f7591466768ddf563fd1cf5 \
  crates/litchi-markdown/tests/corpus/gfm-0.29.0.gfm.13/LICENSE-CC-BY-SA-4.0.txt \
  c22e885f33b821bddb24cf007145e5540655b6c0f403e49e6c76a93c28e6d9a9 \
  crates/litchi-markdown/tests/corpus/gfm-0.29.0.gfm.13/COPYING-BSD-2-Clause.txt \
  | sha256sum --check --strict

cargo fmt -p litchi-markdown -- --check
cargo test -p litchi-markdown --all-targets
cargo clippy -p litchi-markdown --all-targets -- -D warnings
cargo doc -p litchi-markdown --no-deps
