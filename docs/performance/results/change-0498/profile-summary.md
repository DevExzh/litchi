# Whole-child CPU counters

These six captures include fixture construction, repeated package opens, byte verification, and reporting. They do not isolate the batch operation or explain the superlinear local-read observations. Counts are descriptive shared-host observations; there are no repeated independent perf children for uncertainty estimates.

| Child | Cycles | Instructions | IPC | Branch misses | Cache misses |
| --- | ---: | ---: | ---: | ---: | ---: |
| few-large-instrumented-batch-w1 | 1,867,236,702 | 2,253,044,949 | 1.21 | 2,763,892 | 3,695,492 |
| few-large-instrumented-batch-w4 | 1,749,324,935 | 1,982,776,751 | 1.13 | 2,907,081 | 8,200,189 |
| few-large-instrumented-serial-w1 | 1,874,877,501 | 2,249,527,443 | 1.20 | 2,878,452 | 3,739,126 |
| few-large-owned-batch-w1 | 1,397,061,239 | 1,665,580,744 | 1.19 | 386,550 | 2,437,256 |
| few-large-owned-batch-w4 | 1,282,731,133 | 1,372,692,063 | 1.07 | 147,307 | 2,923,249 |
| few-large-owned-serial-w1 | 1,393,709,279 | 1,665,190,689 | 1.19 | 371,875 | 2,492,467 |
