# 0735 PPT record staging without a full editor clone

Both replace_persisted_record and insert_persisted_record validate identity and
record framing before cloning the complete Editor. After the clone they only
insert the owned record into staged_storage and set changed=true, with no
fallible Result-returning work left. Remove the full clone while preserving
validation order, typed refusals, failure atomicity and output identity.
Allocation abort/panic is not a newly recoverable Result and must not be
misrepresented as such. Do not generalize this simplification to mutators that
perform fallible work after mutation.

Capture the unchanged baseline before applying the candidate. Use identical
ordinary public owner probes on 45543.ppt and the first successfully qualified
independent second fixture (fixed order retained by qualify.py). All semantic,
raw CFB directory, stream, slide and survivor witnesses remain exact. Test
both functions' successful staging and every refusal before changing state.

Prospective comparison: three cycles of three process pairs per fixture,
50 native samples and three warmups, followed by three allocation pairs per
fixture with one sample/no warmup. CPU 12 pinned, serialized, pair order rotated.
Keep every observation and flag absolute changes above 5%; assess every positive
regression without selective reruns. Bootstrap across nine process pairs using
10,000 resamples and seed 7335. Expected mechanism is removal of Editor field
copies and clone allocations, not removal of source checks or publication work.
Retain only with practical allocation/work benefit and explicit latency/peak
assessment. No RSS, cold I/O, instructions, concurrency or broad CRUD claim.

The measured public removal workflow calls record replacement. Record insertion
shares the same unnecessary clone pattern and is covered by focused atomicity
and serialization tests, but insertion throughput and multi-record scaling are
not measured by this matrix and must not receive a performance claim.

Before applying the candidate, source inventory and Editor field ownership
predict at least 695,275 primary and 329,614 secondary payload bytes copied
by the full clone (stream vectors plus duplicate Document/Current User vectors).
These are a source-derived lower bound, not measured savings. Path/map/vector
metadata allocations are additional; clone-hypothesis.json records the inputs.
