# 0700 MCE exact-output oracle

This standalone probe uses the shared public MCE API. It emits output SHA-256,
length, Cow ownership and complete Report, or exact Debug error text.
Profiles cover default limits/capabilities, a matching opaque extension,
smaller/larger finite limits, and 4,096 extension names.

Build both revisions with the same retained Cargo.lock using the packet's
build-oracle.py before and after the production edit. No source restoration
or lock refresh is needed for this batch. Run corpus.py with baseline binary,
candidate binary, test-data root and output directory. It compares 192 cases
under all five profiles, rejecting failed or malformed probe invocations.
The optional timing mode is separate from the correctness corpus.
