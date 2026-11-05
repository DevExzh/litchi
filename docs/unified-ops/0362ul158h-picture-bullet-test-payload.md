# 0362ul158h — exercise nonempty picture-bullet payloads

The approved focused run on `34fc969802232bddb913509f6e56db8148ceb03b`
passes all nine OfficeArt unit tests and compiles the ordinary PPT library.
The new picture-bullet integration target then fails compilation because
its empty-array comparisons are ambiguous with `serde_json`'s numeric
`PartialEq` implementations. No integration result is claimed.

Use an explicit nonempty PNG-signature payload and compare its bytes before
and after serialization. This removes the ambiguity and makes payload
preservation observable, alongside the existing zero-copy ownership and
exact OfficeArt framing checks. The signature is an opaque BLIP test payload,
not a complete PNG image. Production code is unchanged.

The failed run and root's verified terminal receipt remain external under
`/home/zhuhe/code/litchi-worktrees/.coord/ppt-158g-focused-20261026/`.
Root verified command/log/target identities, unchanged source and locks,
and empty owned processes. The next approved run must execute the previously
unrun integration, classification and all-format CRUD rows.
