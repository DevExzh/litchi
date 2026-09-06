# Common prelude review resolution

The new opt-in constructor keeps the existing strict constructor unchanged.
Its grammar is declaration?, open ancestors, balanced element-only fixed children,
and a final nonempty trailing sequence of Start events. Fixed children can appear
beneath any open ancestor: `<root><slot><fixed/><inner>` is deliberately valid,
with the insertion at depth three inside inner. There is no separately supplied
first insertion ancestor to infer. Public documentation and focused tests pin
this interpretation. Prefixes ending in Empty or End remain refused.

Fixed character data is deliberately unsupported by this narrow enabler. The
complete fixed shell is audited, including all bytes/events/attributes/depth;
its text-byte count is zero. Character data comes from audited fragments.

Malformed constructor fixtures now require typed InvalidFormat errors; caller
resource ceilings still use typed XML resource/actual/maximum attribution.
Additional tests check a closed fixed child immediately before the trailing
Start sequence and fixed-shell depth dominating a shallow fragment. The full
existing common generated XML tests remain in the validation set.

Strict Clippy stopped at existing package/model.rs ArchiveReaderKind
large_enum_variant debt, also present in the retained before harness diagnostics.
A separate scoped strict command allows only that lint and passes.
