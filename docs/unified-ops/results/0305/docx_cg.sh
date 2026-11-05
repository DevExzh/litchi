#!/bin/sh
# usage: docx_cg.sh BINARY TAG -> per-call Ir, I1mr, D1mr, D1mw for leaf and unified
BIN=$1; TAG=$2; S=/home/zhuhe/code/litchi-worktrees/.coord/perf-lazy/cg
cd /home/zhuhe/code/litchi-worktrees/uo2-perf-lazy/crates/litchi-ops
for leg in leaf unified; do for n in 1000 3000; do
  DOCX_FANOUT_SIZES=1 DOCX_FANOUT_PROFILE=$leg DOCX_FANOUT_LOOPS=$n valgrind --tool=callgrind --cache-sim=yes --I1=32768,8,64 --D1=32768,8,64 --LL=33554432,16,64 --callgrind-out-file=$S/docx-$TAG-$leg-$n.out $BIN >/dev/null 2>$S/docx-$TAG-$leg-$n.log &
done; done; wait
python3 - "$S" "$TAG" <<'PY'
import re,sys
S,TAG=sys.argv[1:]
def totals(leg,n):
    s=open(f"{S}/docx-{TAG}-{leg}-{n}.out").read()
    ev=re.search(r'^events: (.*)$',s,re.M).group(1).split()
    t=re.search(r'^totals: (.*)$',s,re.M).group(1).split()
    return dict(zip(ev,map(int,t)))
for leg in ['leaf','unified']:
    a,b=totals(leg,1000),totals(leg,3000)
    print(TAG,leg,' '.join(f"{k}={(b[k]-a[k])/2000:.1f}" for k in ['Ir','I1mr','D1mr','D1mw','Dr','Dw']))
PY
