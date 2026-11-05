# 0362ul87ai — private OOXML codec dead-code repair

The coordinator-approved F5 attempt for
`27be5389f73680a2f06883cb9188338130b2a11c` (tree
`df28752955868b5749c95bcc96505d5553f042c4`) passed `fmt-scoped` and stopped
at row 1, `check-common`. The preserved state is:

```text
(deleted-coord-artifact)
```

Cargo reported 22 `unused` errors after the stale `clear` call was repaired:
the private custom-properties, Custom XML, core-properties, application-
properties and graph codec helpers are deliberately retained for their
byte/XML and physical-graph test surfaces, but their former cross-crate
semantic consumers now use the shared Edit/Commit route. The workspace has
`unused = "deny"`, so private helpers used only by unit tests need an explicit
release-build dead-code annotation.

The preserved receipt identities are `result.json` SHA-256
`ffe27fca0bc00dcbb3ff61e95004bfdfc94539b6974aca9e5be96672c261cce9`,
`rows.json` SHA-256
`2417c0b18f8669fbef489d010558948c39f6ffc6e784c1c07569be9a4e4668b6`,
the compiler log SHA-256
`e1ee9fdb52e67b0a4ba8cce71834b1de692eaaa6d0bbf6e0d220ff8b45ef8c4b`,
and the target trace SHA-256
`95da6a0252613e4bd7f3020b2101cdcb8233510bcdf2bc81ce3f19a35a24fcee`.
The target reached 158,578,656 bytes; no resource or ownership guard fired.

The repair adds narrow `#[cfg_attr(not(test), allow(dead_code))]` attributes
only to the compiler-reported private helpers and constants. It does not
restore public or crate-crossing publisher visibility, add aliases, remove
codec/graph capability, or change package behavior. The next F5 attempt must
bind the new source and fixture identities before launch.
