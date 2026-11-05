#!/bin/bash
F=/deleted-coord-artifact
side=$1 scen=$2 rounds=$3
out=$F/cgodf/$side.${scen//\//_}.$rounds
ODF_W6_ONLY=$scen systemd-run --user --scope --quiet --collect -p MemoryMax=20G -p MemorySwapMax=0 -- valgrind --tool=callgrind --callgrind-out-file=$out.cg $F/bin/w6cg-$side $rounds > $out.log 2>&1
echo -e "$side\t$scen\t$rounds\t$(grep -m1 '^totals:' $out.cg | awk '{print $2}')"
