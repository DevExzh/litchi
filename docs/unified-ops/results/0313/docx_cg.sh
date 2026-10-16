#!/bin/sh
# Record 0313. usage: docx_cg.sh BASE_BIN AFTER_BIN OUT_DIR -> docx-cachesim.txt rows
# Unstripped `profiling` builds of docx_large_bench; per call = (15 loops - 5 loops) / 10.
# Each valgrind runs in its own memory-capped scope (the host has no swap).
B=$1 A=$2 O=$3
mkdir -p "$O"
for d in base after; do bin=$B; [ $d = after ] && bin=$A
  for leg in leaf_set auto_set committed_set eager_set leaf_add eager_add auto_add committed_add fresh_auto_add fresh_eager_add; do for n in 5 15; do
    DOCX_LARGE_PROFILE=1 DOCX_LARGE_ONLY=$leg DOCX_LARGE_LOOPS=$n \
      systemd-run --user --scope --quiet --collect -p OOMPolicy=continue -p MemoryMax=4G -p MemorySwapMax=0 -- \
      valgrind --tool=callgrind --cache-sim=yes --I1=32768,8,64 --D1=32768,8,64 --LL=33554432,16,64 \
      --callgrind-out-file="$O/$d-$leg-$n.out" "$bin" >/dev/null 2>&1 &
  done; done; wait
done
python3 - "$O" <<'PY'
import re,sys
O=sys.argv[1]
def t(f):
    s=open(f).read()
    ev=re.search(r'^events: (.*)$',s,re.M).group(1).split()
    return dict(zip(ev,map(int,re.search(r'^totals: (.*)$',s,re.M).group(1).split())))
print("per call = (15 loops - 5 loops) / 10; docx_large_bench DOCX_LARGE_PROFILE=1")
print("build leg Ir I1mr D1mr D1mw")
for d in ['base','after']:
    for leg in ['leaf_set','auto_set','committed_set','eager_set','leaf_add','eager_add','auto_add','committed_add','fresh_auto_add','fresh_eager_add']:
        a,c=t(f"{O}/{d}-{leg}-5.out"),t(f"{O}/{d}-{leg}-15.out")
        print(d,leg,' '.join(f"{(c[k]-a[k])/10:.0f}" for k in ['Ir','I1mr','D1mr','D1mw']))
PY
