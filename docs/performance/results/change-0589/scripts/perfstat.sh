#!/usr/bin/env bash
# isolation pair: run `low` and `high` iterations, difference, divide by (high-low)
# args: leg bin mode path low high
SC=/tmp/claude-1001/-home-zhuhe-code-litchi/709d31e6-bf68-4968-8b5c-9a2af1c22cc8/scratchpad/agents/0589
leg=$1; bin=$2; mode=$3; path=$4; low=$5; high=$6
name=$(basename "$path")
one() {
  taskset -c 9 perf stat -x, -e cycles,instructions -r 5 "$bin" profile "$mode" "$path" 0 "$1" 2>&1 \
    | awk -F, '$3=="cycles"{c=$1} $3=="instructions"{i=$1} END{printf "%s %s\n", c, i}'
}
read -r c_lo i_lo <<<"$(one "$low")"
read -r c_hi i_hi <<<"$(one "$high")"
python3 - "$leg" "$mode" "$name" "$low" "$high" "$c_lo" "$i_lo" "$c_hi" "$i_hi" <<'PY'
import sys, json
leg, mode, name, low, high, c_lo, i_lo, c_hi, i_hi = sys.argv[1:]
m = int(high) - int(low)
print(json.dumps({
  "leg": leg, "mode": mode, "file": name, "low": int(low), "high": int(high),
  "cycles_per_op": round((float(c_hi)-float(c_lo))/m, 1),
  "instructions_per_op": round((float(i_hi)-float(i_lo))/m, 1),
}))
PY
