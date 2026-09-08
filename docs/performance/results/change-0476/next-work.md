# Remaining work after the owned Deflate experiment

Compressor reuse targets repeated allocation requests. It does not remove
the ZIP central directory, archive name storage, streaming duplicate-name
indexes or OPC part-name indexes retained until finalization. The public
PPTX fresh writer therefore still needs an explicit total-memory design.

The next bounded task is to identify which retained structures are required
for successful publication and duplicate detection, then design explicit
window/budget accounting for each owner. A directory spool alone is
insufficient while in-memory name indexes still grow with the member count.
Any external spool or index must have an explicit storage and failure
contract; it must not silently turn the existing in-memory API into a
filesystem-dependent operation.

Preserve the current public ownership boundaries, typed limits, accepted
input/output accounting and transactional publication. Retain deterministic
archive identities and full format-level round-trip oracles. Measure the
complete operation allocation lifetime at several sizes separately from
scratch reservation, cumulative requested bytes and process RSS.

Fresh creation remains separate from logical append to an existing format
structure, adding a package Part, and arbitrary edits followed by repackaging.
The existing native-producer, source-variant, feature-breadth and worker-scaling
gaps remain open. This scoped optimization does not complete the non-iWork goal.
