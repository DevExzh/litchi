# 0362ul63a — PPTX raw-properties correction

- Date: 2026-10-25.
- Status: source-ready for independent review; native Cargo validation remains
  with the coordinator's isolated build slot.
- Parent: 1ff8391108 (feat/pptx: bound raw properties and model3d codecs).

This correction closes the independent review of the raw-properties and
model3d slice:

- model3d validation now applies the selected child, raster-child, inert
  fragment depth/node/byte policy before shared validation and preserves typed
  DrawingML LimitKind errors through the PPTX error;
- view-properties MCE processing uses the selected input/output/depth and
  directive bounds, and processed markup is admitted against the output bound;
- view-properties and presentation-properties serialization use an exact
  count pass followed by a fallible, exactly-reserved bounded writer;
- readonly-recommended no-op and source-splice paths reserve their exact
  output with fallible allocation;
- opaque extension/color fragments retain source bytes and add only required
  inherited namespace declarations, without rewriting namespace URI text or
  attributes. The generated typed package roots still select the requested
  transitional or strict namespace.

Focused malformed, bounded, source-span and opaque-fidelity tests were added
for the affected codecs. Source checks performed here are rustfmt,
git diff --check, and focused API/source sweeps. Cargo tests, rustdoc,
clippy and ABBA measurements remain outstanding for the coordinator's
isolated run; no binary or large performance evidence is committed.
