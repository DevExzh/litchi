# 0828 symbol admission failure review

Status: static failure review. The frozen `decode.py --symbols` gate failed
before any qualification, native, or perf workload. I reviewed the retained
failure and the separate post-abort address-bounded diagnostic; I used read-only file inspection and ran no workload, Cargo command or
profiler for this review; I changed only this file.

The immediate failure is a disassembly selector miss. At
`docs/performance/results/change-0828/decode.py:124-130`, the gate invokes

```text
objdump --demangle=rust --disassemble=<raw nm symbol> --wide --line-numbers <fp>
```

where `<raw nm symbol>` came from `nm -S --defined-only`. The retained
`symbols-console.log` ends at `assert matched["symbol"] in assembly_text`, and
`symbols/edit_region-assembly.txt` is only 194 bytes of ELF and section
headings. The retained output has no matching function body. This supports a
symbol-selection failure; it does not establish that the successful FP binary
lacks the wrapper.

The likely trigger is the new `--demangle=rust` combined with a mangled
selector. The known-good 0822 command at
`docs/performance/results/change-0822/decode.py:110-124` selected the raw name
without enabling demangling. The current command also makes its assertion
brittle even if a body were emitted: demangling can make the printed function
label the Rust path while the assertion requires the raw mangled spelling.
The retained evidence does not establish the internal lookup order, so the
bounded conclusion is that name-based selection is not reliable under this
command, rather than that the binary or compiler output is corrupt.

The post-abort diagnostic applies the minimal safe separation. It leaves the
failed `symbols/` directory untouched, verifies the FP artifact from
`build.json`, and writes to `symbol-diagnostic/` with
`capture_authorized: false` and `workload_executed: false`. It first obtains
raw and demangled `nm -S` tables, requires one exact row for each planned owner,
and joins the rows on `(address, size, type)` (`symbol_diagnostic.py:48-57`). It
then disassembles each numeric range with

```text
objdump --demangle=rust --wide --line-numbers -d \
  --start-address=<address> --stop-address=<address+size> <fp>
```

(`symbol_diagnostic.py:58-66`). This avoids symbol-name lookup while retaining
demangled display names. The diagnostic requires the first instruction at the
`nm` address, every printed instruction within the half-open range, the exact
demangled owner header, an FP prologue (`push %rbp` and `mov %rsp,%rbp`), and a
call (`:62-73`). All four owners pass: `edit_region_0828` (address 1,564,048,
size 276), capture (1,568,336, size 666), set-text (1,567,920, size 406), and
publish (1,583,520, size 1,090). The result records 63, 134, 94, and 207
instructions respectively and no command stderr.

The diagnostic is useful recovery evidence, not a replacement for the frozen
admission gate. `abort.json` records `aborted-before-workload` with zero reports
and zero samples and sets `diagnostic_replaces_admission` to false. The
existing offline analysis contract still expects a `symbols/symbol.json`, exact
raw/demangled rows, and assembly containing the raw symbol
(`analysis.py:1519-1593`, `:1607-1628`). A future profile can therefore either
use the address-bounded command in a new admission driver and update that
reader contract, or retain the old raw-selector command without demangling and
prove its output. It must not silently treat `symbol-diagnostic/result.json`
as the frozen `symbol.json`.

The minimum requirements for a future address-bounded symbol gate are:

- Preserve the failed attempt and recovery in separate append-only paths. Bind
  every receipt and assembly artifact to the exact FP binary path, byte count,
  SHA-256, build/probe descriptors, frozen inputs, and the four planned owner
  names. Keep the no-workload and no-capture markers explicit.
- Require unique exact demangled owner rows and a unique raw row with identical
  `(address, size, type)`. Require a positive text-symbol size, checked
  `end = address + size`, and distinct non-overlapping owner ranges before
  invoking objdump.
- Pass only numeric `--start-address` and `--stop-address` bounds to objdump;
  keep `--demangle=rust` as presentation-only. Retain stdout, stderr, and the
  exact command for every owner and phase. Verify the first instruction address,
  all instruction bounds, exact demangled header, frame-pointer prologue, and
  call boundary.
- Recheck frozen source, tool, lock, corpus, and packet custody around the
  child commands (the post-abort diagnostic checks before and after its static
  lane). Do not authorize qualification or perf capture until all four ranges
  pass and the resulting receipt is admitted by the reader contract.

The failed name lookup is therefore repaired at the evidence boundary. It does
not justify changing the probe, production source, fingerprint validation, or
the workload plan.
