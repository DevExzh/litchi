# 0362ul158g — keep the PPT fixture budget in tests

The approved focused run on `40eac77e8d29e8f01326a87e4574cf39f361dbd8`
passes 1,249 PPT library tests, with one pre-existing ignored test requiring
a fixture. The next row's ordinary library compilation rejects the obsolete
`MAX_HOST_RECORDS` constant: after 158f, only two unit tests use it.

Move that constant into the test module and migrate both consumers. Runtime
paths continue to use their admitted source limits. This changes no runtime
policy and adds no lint waiver. Rustfmt and `git diff --check` pass.

The failed run remains external at
`/home/zhuhe/code/litchi-worktrees/.coord/ppt-158f-focused-20261026/`.
Root verified the logs, target traces, unchanged source and three locks,
and empty owned processes. Its `root-terminal-verification.json` is 1,475
bytes, SHA-256
`0832e5e348d6077018b56d54bea888e23ea93afa6785fa63ec5c58bc6c2f6378`.
The integration tests, Ops error classification and all-format CRUD checks
remain unrun; a fresh approved run must verify the resulting commit.
