# 0455 experiment: unchanged ZIP transfer chunks

Compare the current 32 KiB unchanged local-member copy buffer with 64 KiB.
The expected mechanism is fewer caller read requests and sink writes at equal
raw transfer bytes and identical output. A 64 KiB stack buffer adds 32 KiB per
active publication; there is no new heap allocation or parallelism.

Use matched plain and media-rich PPTX copy lifecycles through bytes and the
existing 64 KiB, 200 us/call, 25 MiB/s simulated range provider. Two balanced
A/B/B/A repetitions, 30 samples after three warmups per lane, one worker on
CPU 2. Independent allocation-instrumented bytes lanes are descriptive and
kept out of ordinary timing comparisons. All >5% timing/RSS flags remain
visible. Keep only a useful measured improvement with preservation and resource
gates intact; otherwise revert the production experiment and retain the data.

This is not physical network, cold-cache, native-application, or scaling evidence.
The broader non-iWork goal and CRUD coverage requirements remain open.
