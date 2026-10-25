# 0362ul66k — Numbers test compile repair

Status: source READY for the next scoped F5 retry.

The sixth F5 run passed the Numbers-wire rows, Pages, and the hydrated iWork
fixture boundary, then stopped in Numbers package tests.  Two test-only
fixture builders used `?` with the private BNC encoder while returning the
package result type; the package intentionally has no blanket conversion from
the raw-wire error.

The builders now use explicit expectations for their known-valid minimal and
formula cells.  Production adapters continue to map raw errors at their own
boundaries; no public error or wire behavior changes.
