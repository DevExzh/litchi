# 0362ul87am — gate the test-only slot source state

The approved F5 run for `842b9b1042e171f19e27584977c81832aaaeff37`
(tree `fa3c281781c0c321faf88dc84dae1f08528ade38`) passed `fmt-scoped` and
stopped at row 1, `check-common`. The release build reported that
`properties::Slot::source` was never read after `Slot::stage` and `flush` were
correctly gated for tests; the workspace's `-D unused` policy made that a
compile failure. The preserved state is:

```text
(deleted-coord-artifact)
```

The failed receipt identities are `rows.json` SHA-256
`cc87c8a045882fbc133211ba336726d4a94463c74e73350d160e0a94d31b02d4`,
`result.json` SHA-256
`adf8c85dad7fdb31aade4660fa6d9e063bdaba8a84323bca2bdbc046da612555`, and
the `check-common` compiler log SHA-256
`2cba84f0f6dcb1c0da936a29ee610f1d11b490233a2a473d1b0a8647e570bb38`.
The row ran with the approved isolated target and stopped after 33.131
seconds at 158,555,218 target bytes; no later row was launched.

The repair gates `Slot::source` itself and conditionally initializes it only
for test builds. Test staging retains the source snapshot and its lexical
fidelity behavior; release builds retain the public read-only slot projection
without a dead private field. No publisher, alias, or production preflighted
codec changed.

The next F5 attempt must use a fresh target and bind the new source, tree,
fixture, plan, driver, and manifest identities. No rerun is claimed for this
repair.
