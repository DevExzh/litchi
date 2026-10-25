# 0362ul94r — properties error-contract repair

The isolated `0362ul94r` common-test run from `3774fe4374` stopped at two
properties assertions. The source behavior was already following the shared
error contract: removing an application-properties part that has another
inbound relationship is a valid-operation refusal, while a source graph
relationship remains a package corruption error when the graph itself is
inspected. The application-properties test now checks the typed
`UnsupportedContent` refusal, its stable operation and its content-free
reason instead of expecting the lower-level relationship error from patch
replay.

Application planning also distinguishes an authored XML 1.0 lexical failure
from a malformed retained source or a generated invariant. The lexical error
remains `properties::Error::Invalid` with its typed message; source failures
remain nested package errors and generated failures remain `Internal`. This
preserves the existing classification order without weakening either test.

This is a source-only repair. No Cargo, build, rustdoc, runtime, native, or
performance command was run; the coordinator must rerun the isolated gate
against this commit.
