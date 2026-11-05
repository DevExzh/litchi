#!/bin/bash
# build.sh <tree> <target> <cargo args...>  (release, litchi-ops all features)
set -u
source /deleted-coord-artifact
export CARGO_BUILD_JOBS=6
TREE=$1 TGT=$2; shift 2
cd "$TREE"
CARGO_TARGET_DIR=$TGT cargo build --release --offline "$@"
