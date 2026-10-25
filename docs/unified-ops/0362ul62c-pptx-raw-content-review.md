# 0362ul62c — PPTX collaboration replay validation correction

Status: source-ready correction to `0362ul62b`, based on the coordinator
replay integration. Native Cargo validation, the workspace gate, and ABBA
remain coordinator work because this leaf has no build slot.

## Correction

Comment-author presence and comment-threading replay now share one bounded raw
scan and one markup-compatibility selection per changed part. The validator
collects only owners containing the exact recognized active extension URI,
then applies the existing typed payload grammar to active branches. It no
longer uses URI substring gates or reparses the complete part once per owner;
inactive MCE fallbacks and unknown opaque extensions remain untouched.

Coverage adds a partial multi-author owner case and an inactive typed-URI case,
alongside the existing tight recipient-string and malformed threading-child
checks. The notes/media aggregate return behavior, selected shape staging, and
no-op replay behavior from `0362ul62b` remain unchanged for coordinator
integration with the shared aggregate accounting/index.

No generated evidence, fixture, executable, or `3rdparty/` path is included.
