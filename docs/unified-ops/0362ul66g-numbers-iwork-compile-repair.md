# 0362ul66g — Numbers/iWork compile repair

Status: source READY for the next scoped F5 retry.

The second F5 run passed all Numbers-wire checks and tests, then stopped at
the first iWork check.  The existing Numbers owner adapter had two stale
qualified/type-alias signatures exposed by the strict package build:

- the private BNC encoder now returns the module's re-exported native `Error`;
- the native comment-reply encoder uses its local `Result<T>` alias, whose
  error type is already `NativeReplyError`.

Both signatures now match their surrounding adapters.  No wire policy,
redaction, or output behavior changes.
