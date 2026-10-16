#!/bin/sh
# Record 0317. usage: cg.sh BIN_DIR OUT_DIR -> cachesim rows
# Unstripped `profiling` builds; per call = (hi loops - lo loops) / (hi - lo).
# Each valgrind runs in its own memory-capped scope (the host has no swap).
D=$1 O=$2
mkdir -p "$O"
DOCX="leaf_set auto_set committed_set eager_set leaf_add eager_add auto_add committed_add fresh_auto_add fresh_eager_add"
PPTX="leaf_set_write unified_set_write leaf_add_write unified_add_write committed_set_write committed_add_write"
for d in base after; do
  for leg in $DOCX; do for n in 5 15; do
    DOCX_LARGE_PROFILE=1 DOCX_LARGE_ONLY=$leg DOCX_LARGE_LOOPS=$n \
      systemd-run --user --scope --quiet --collect -p OOMPolicy=continue -p MemoryMax=4G -p MemorySwapMax=0 -- \
      valgrind --tool=callgrind --cache-sim=yes --I1=32768,8,64 --D1=32768,8,64 --LL=33554432,16,64 \
      --callgrind-out-file="$O/$d-docx-$leg-$n.out" "$D/docx_large_bench-$d-dbg" >/dev/null 2>&1 &
  done; done; wait
  for leg in $PPTX; do for n in 1 3; do
    PPTX_LARGE_PROFILE=1 PPTX_LARGE_ONLY=$leg PPTX_LARGE_LOOPS=$n \
      systemd-run --user --scope --quiet --collect -p OOMPolicy=continue -p MemoryMax=6G -p MemorySwapMax=0 -- \
      valgrind --tool=callgrind --cache-sim=yes --I1=32768,8,64 --D1=32768,8,64 --LL=33554432,16,64 \
      --callgrind-out-file="$O/$d-pptx-$leg-$n.out" "$D/pptx_large_bench-$d-dbg" >/dev/null 2>&1 &
  done; done; wait
done
python3 - "$O" "$DOCX" "$PPTX" <<'PY'
import re,sys
O,DOCX,PPTX=sys.argv[1],sys.argv[2].split(),sys.argv[3].split()
def t(f):
    s=open(f).read()
    ev=re.search(r'^events: (.*)$',s,re.M).group(1).split()
    return dict(zip(ev,map(int,re.search(r'^totals: (.*)$',s,re.M).group(1).split())))
print("per call: docx (15 - 5 loops)/10, pptx (3 - 1 loops)/2")
print("build bench leg Ir I1mr D1mr D1mw")
for d in ['base','after']:
    for b,legs,lo,hi in [('docx',DOCX,5,15),('pptx',PPTX,1,3)]:
        for leg in legs:
            a,c=t(f"{O}/{d}-{b}-{leg}-{lo}.out"),t(f"{O}/{d}-{b}-{leg}-{hi}.out")
            print(d,b,leg,' '.join(f"{(c[k]-a[k])/(hi-lo):.0f}" for k in ['Ir','I1mr','D1mr','D1mw']))
PY
