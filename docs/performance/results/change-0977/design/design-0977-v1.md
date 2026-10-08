# 0977 combined qualification capture design

This is a fresh twelve-child Office packet using the retained 0975 ordinary
binary and FloatingPictures fixture. It is not a retry of 0975 and makes no
claim until the complete raw packet is reduced. Children run serially with
five warmups and eighty measured calls. The fixed cell is `docfloat/e1/default`.

The outer command pins perf to CPU31 and uses `--per-thread`,
`instructions:u` at period 1,000,000, `-m 16M`, and DWARF 32768. A fresh
0977 bridge is a target-generic adaptation of the sealed 0976 handoff: it
verifies the exact perf parent, moves every observed parent perf TID to CPU30,
verifies the unchanged `(tid,starttime)` set, and same-PID `execve`s the
retained 0970 launcher. The launcher then same-PID execs the bound Office
binary. No target-side stdlib checks are used. Bridge masks and bounded
observations are recorded, while the zero-loss raw task stream is the lifetime
authority. Startup and handoff samples remain in the capture.

Admission is fail-closed per child: return code zero, exact report/oracle,
raw lost sum zero, terminal `LOST_SAMPLES` zero, one PID/TID, CPU31, exact
period, positive marker period share, at least 100 marker samples, and zero
empty callchains. A failed child preserves its raw files and stops the cohort;
there are no retries or selective replacement children.

The final-v2 DWFL decoder from 0976 is used only offline after terminal
capture. Each child binds its raw data, ELF/MMAP layout, decoder source/map/SO,
and target binary. `LD_PRELOAD` is absent from the native Office process and is
allowed only for the offline decoder invocation. All native and decoder
artifacts are external to the repository.

The provenance chain is acyclic: helper seal binds helpers; the plan binds the
seal and external reviews; root GO binds the exact plan; the capture manifest
binds the plan and GO. No plan or output path is reused, and no build or
workload is started by preparation or self-tests.

Offline reduction uses only the preregistered corrected final-v2 DWFL path
whose 0976 source/SO, ASLR mapping, and no-op identity reviews are bound in
the plan. Stock or additional decoder modes are not part of this capture.
