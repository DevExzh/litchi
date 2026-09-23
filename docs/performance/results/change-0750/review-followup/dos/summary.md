# Change 0750 review follow-up: verify_source on adversarial namespace inputs

Release builds of probe/dos against base 3174242282, the first candidate de69fb407d and the fix 0bbcc9bf94 (rustc 1.98.1 for all three),
one process per case and leg pinned to core 24 (scripts/run_dos.sh); each process audits up to 5 times and stops after 60 s,
so the table shows the median of up to 5 audits and '(1 run)' where one audit took more than 60 s.

| case | bytes | base s | first candidate s | fix s | fix / base | verdicts |
| --- | ---: | ---: | ---: | ---: | ---: | --- |
| `r1-one-tag-249990-attributes` | 14,838,818 | 0.0168 | 180.354 (1 run) | 0.0382 | 2.28 | Ok on all legs |
| `r2-one-tag-60000-attributes` | 12,408,948 | 0.0082 | 39.099 | 0.0215 | 2.61 | Ok on all legs |
| `r3-one-tag-60000-attributes-aliased` | 12,408,948 | 0.0083 | 48.598 | 0.0235 | 2.85 | Ok on all legs |
| `r4-50000-tags-two-names` | 8,700,036 | 0.0070 | 3.556 | 0.0169 | 2.41 | Ok on all legs |
| `f1-50000-tags-aliased` | 8,700,036 | 0.0069 | 10.580 | 0.0188 | 2.70 | Ok on all legs |
| `f2-200-levels-redeclaring-100k-name` | 20,184,580 | 0.0104 | 0.025 | 0.0326 | 3.15 | Ok on all legs |
| `f3-100000-prefixes` | 4,066,677 | 0.0099 | 0.023 | 0.0283 | 2.85 | Ok on all legs |
| `f4-6-aliases-of-4mb-name` | 25,160,126 | 0.0158 | 13.606 | 0.0458 | 2.90 | Ok on all legs |
| `w1-window-in-50000-aliased-tags` | 8,700,036 | 0.0074 | 10.263 | 0.0226 | 3.06 | Ok on all legs |
