# 0362ul66j — iWork test-fixture repair

Status: source READY for the next scoped F5 retry.

The fifth F5 run passed the Numbers-wire rows and Pages compilation, then
stopped while compiling all iWork targets.  The hydrated tracked fixture
`media/litchi_logo.png` was missing from the sparse checkout, and one existing
Numbers editor test helper used `?` in a function returning its concrete test
package type after `bnc_encode` became explicitly fallible.

The tracked fixture is now hydrated for the check, and the helper unwraps its
known-valid minimal BNC fixture at the test boundary.  No production path or
wire output changes.
