#!/bin/bash
# Record 0360b: release builds of the probes, base (git archive of the
# merge-tree 684a028b2a = uo2/w4-xlsx 635b884c7c + trunk) and head (uo2/w4b-xlsb).
set -e
source /deleted-coord-artifact
P=/deleted-coord-artifact
build() { # $1 = leg, $2 = source root
  export CARGO_TARGET_DIR=$P/tgt-$1
  cd $2
  cargo build --release -p litchi-ops --features xlsb --example xlsb_w4b_probe
  cargo build --release -p litchi-ops --features xlsx,xlsb,xls,eval --example sheet_dispatch_bench
  cd tools/perf-baseline
  cargo build --release --locked --features xlsb-crud --bin xlsb_crud
}
build ${1:-after} ${2:-/home/zhuhe/code/litchi-worktrees/uo2-w4b-xlsb}
