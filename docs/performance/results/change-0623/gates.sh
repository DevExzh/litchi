#!/usr/bin/env bash
# Change 0623's gates. Each section prints its command, its tail and its exit
# code; the script does not stop on failure so every gate is reported.
set -uo pipefail
cd /home/zhuhe/code/litchi-worktrees/0623
run() {
  local label="$1"; shift
  echo "=============================================================="
  echo "== $label"
  echo "== \$ $*"
  local out
  out="$("$@" 2>&1)"
  local code=$?
  printf '%s\n' "$out" | tail -25
  echo "== exit $code"
  return 0
}
run "formatting"                  cargo fmt --all --check
run "clippy soapberry-zip"        cargo clippy -p soapberry-zip --all-targets
run "clippy litchi-opc"           cargo clippy -p litchi-opc --all-targets
run "clippy litchi-opc, all features" cargo clippy -p litchi-opc --all-targets --all-features
run "rustdoc soapberry-zip"       cargo doc -p soapberry-zip --no-deps
run "rustdoc litchi-opc"          cargo doc -p litchi-opc --no-deps
run "no-default-features check"   cargo check -p litchi-opc --no-default-features
run "tests soapberry-zip"         cargo test -p soapberry-zip
run "tests litchi-opc"            cargo test -p litchi-opc
run "tests litchi-opc, all features" cargo test -p litchi-opc --all-features
run "tests OOXML facades"         cargo test -p litchi-xlsx -p litchi-docx -p litchi-pptx
run "tests litchi facade"         cargo test -p litchi
run "tests litchi-ooxml-common"   cargo test -p litchi-ooxml-common -p litchi-core
echo "=============================================================="
echo "== all gates attempted"
