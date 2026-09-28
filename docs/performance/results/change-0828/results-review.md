# 0828 result review — aborted before workloads

The frozen symbol admission failed after all five fresh probe quality gates and
both release builds passed. The failing `objdump` command combined
`--demangle=rust` with a mangled `--disassemble` selector. Its output contained
ELF/section headings but no instructions, and the frozen reader rejected it.
`symbols-console.log`, all partial symbol outputs, unchanged drivers and the
original freeze remain retained. No qualification, native or perf lane began:
zero measurement reports and zero timed samples were produced.

A separate post-abort static diagnostic used the two existing executables'
recorded identities and inspected only the frame-pointer binary. It paired exact
raw/demangled `nm` address, size and type rows, then disassembled each exact
address range. All four wrappers retained frame-pointer setup and call
instructions. These six static commands passed. They do not replace the failed
admission, authorize captures, or establish timing or CPU phase attribution.

The pre-cleanup abort validator passed with three retained reader attempts:
pre-freeze preflight, post-quality preflight and abort validation. Each retained
its source snapshot and console log. There are no failed reader attempts at
this point. The one failed execution is the frozen symbol gate. Fresh probe
quality has three passing tests; 0827 quality remains historical reuse and is
verified by committed bytes, source/tool/lock identities and parsed logs.

Production and benchmark runtime source are unchanged. The previous admitted
0827 timing results remain authoritative. The next batch must freeze a corrected
address-bounded symbol driver and run a fresh complete qualification/native/perf
matrix; no old timing or executable may be substituted into that experiment.
Cleanup may remove only the owned 0828 target after exact binary verification.

Post-cleanup final validation also passed (reader attempt 3). All four retained
reader attempts passed; the original symbol-gate failure remains the sole
failed execution. Cleanup removed 3,688 files / 2,054,144,602 logical bytes.
