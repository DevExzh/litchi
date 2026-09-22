# 0731 prospective public PPT instruction attribution

PPT slide removal has its own append-only writer; DOC render-handoff evidence
does not establish its bottleneck. Profile the existing public workflow without
changing production. A no-inline probe wrapper surrounds only the measured
open/edit/remove/commit/output-copy function, including its local owner drops.
Callgrind collection starts disabled and toggles only in that wrapper. Oracle
setup, negative controls, output verification and returned-output drop remain
outside. One sample/no warmup gives exactly one collected lifecycle per process.
Three native50-sample/3-warmup and three separate allocation1-sample processes
provide current context; three Callgrind processes attribute instructions.
All runs use CPU12, warm OS cache and default features. Native process p50/mean
absolute changes over5% are review flags, not sample-exclusion rules. No runtime
optimization, phase subtraction, broad corpus, RSS or hardware-counter claim.
Keep source/fixture/oracle/binary/command custody and every failed attempt.
