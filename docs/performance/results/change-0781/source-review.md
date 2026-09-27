# 0781 current-source selection

The previous turn was progress: commits 59bdb64f16/fdca3e6303 retained a source-bound MCE optimization and its explicit lifecycle tradeoff, sealed evidence and completed owned cleanup. Current base fdca3e6303 and the three unrelated main paths were rechecked.

Independent read-only ROI review ranks the fresh PPT conversion first: core/codec.rs clones default plain text at conversion, UserShapeData owns it until the drawing is serialized, and both package.rs fresh output paths call that conversion. The 0753 after-profile attributed about 15% of its payload-heavy timed region to this work. This is historical motivation, not current attribution; new baseline/paired measurements decide retention. Rich paragraphs remain owned because smart-tag conversion mutates copied run identifiers.

The next separate opportunity is phase attribution of 0780's consistent large-lifecycle regression; its aggregate does not identify a cause. Path-specific exact-size OPC ingress reservation is another bounded hypothesis after 0779 rejected geometric growth for memory cost. Both remain deferred, not implemented here. XLS SST per-string UTF-16 allocation and PPTX slide-root recapture are stale queue items, already addressed by 0753/0760.

A parallel coverage audit found an actionable current CI defect: workflow assertions of 37 cases/201 rows conflict with the actual default 41 cases/213 rows, and the active workflow publishes historical CRUD index v1 rather than v2. The repair derives the report contract from the manifest and selects the active index, with executable tests and real harness output validation. This restores goal deliverable 8 evidence flow; it does not promote any correctness-only CRUD row or assert full goal completion.
