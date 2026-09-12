# 0536 measured collector mechanism

The candidate changed code layout without materially reducing collector work.
The source-bound assembly indexes show a 1,436 to 1,238 byte collector, a
184 to 152 byte explicit stack reservation, and eight new cold helpers. The
normal executable instead grows by 2,608 bytes. Both binaries keep the checked
bitset update and direct vector append inside the collector.

The [baseline disassembly](baseline/assembly-2.stdout) stores slot, next marker
and next sector at 0x2f294b8, 0x2f29500 and 0x2f2951f. In the first XLS timed
dump those instructions receive 32,928, 32,928 and 32,918 self Ir. The
[candidate disassembly](candidate/assembly-2.stdout) carries the next sector in
a register, but retains a slot spill at 0x2f29573, recomputes the visited-length
address at 0x2f2953b and reloads the allocation-table pointer at 0x2f2959d in
the loop. Its table-bound comparison moves to the loop head; the baseline has
that comparison at the loop footer after its initial check. This is changed
register/control layout, not elimination of the validated walking work.

The [instruction reports](instruction-analysis-comparison.json) map exact
instruction boundaries and keep self costs separate from direct callees and
collection-off call/jump metadata. Collector self Ir changes per five timed
dumps from 5,601,140 to 5,600,880 in each XLS repeat (-0.004642%), and from
5,571,945 to 5,571,845 in each CFB few-large repeat (-0.001795%). XLS constructor
inclusive totals rise slightly in both repeats, failing that independent gate.
These instruction observations do not explain native timing variation.

The [decision](decision.json) rejects the candidate: only four of eight primary
p50 rows meet the frozen 3% threshold. Allocation guards pass, but do not
compensate for failed admission. Production is restored exactly; no cold-helper
optimization is retained. No required check or ownership boundary was removed.
