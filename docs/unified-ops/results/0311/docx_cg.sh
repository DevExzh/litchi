#!/bin/sh
# usage: docx_cg.sh BASE_DIR AFTER_DIR OUT_DIR -> docx-cachesim.txt rows
# Unstripped docx_large_bench builds; per call = (15 loops - 5 loops) / 10.
B=$1 A=$2 O=$3
cd "$(dirname "$0")/../../../../crates/litchi-ops"
for d in base after; do dir=$B; [ $d = after ] && dir=$A
  for leg in leaf_set auto_set committed_set eager_add auto_add; do for n in 5 15; do
    DOCX_LARGE_PROFILE=1 DOCX_LARGE_ONLY=$leg DOCX_LARGE_LOOPS=$n valgrind --tool=callgrind \
      --cache-sim=yes --I1=32768,8,64 --D1=32768,8,64 --LL=33554432,16,64 \
      --callgrind-out-file=$O/$d-$leg-$n.out $dir/docx_large_bench >/dev/null 2>&1 &
  done; done; wait
done
python3 - "$O" <<'PY'
import re,sys
O=sys.argv[1]
def t(f):
    s=open(f).read()
    ev=re.search(r'^events: (.*)$',s,re.M).group(1).split()
    return dict(zip(ev,map(int,re.search(r'^totals: (.*)$',s,re.M).group(1).split())))
print("build leg Ir I1mr D1mr D1mw")
for d in ['base','after']:
    for leg in ['leaf_set','auto_set','committed_set','eager_add','auto_add']:
        a,c=t(f"{O}/{d}-{leg}-5.out"),t(f"{O}/{d}-{leg}-15.out")
        print(d,leg,' '.join(f"{(c[k]-a[k])/10:.0f}" for k in ['Ir','I1mr','D1mr','D1mw']))
PY
