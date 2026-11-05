# 0362ul66h — Pages compile repair

Status: source READY for the next scoped F5 retry.

The third F5 run passed all Numbers-wire rows and stopped at the first iWork
check in Pages.  The raw-profile helper used an unnecessary fully-qualified
`WireLimits` path under `-D warnings`, and the body-table value publisher did
not map the package-owned profile-construction error into its local error
type.

The helper now uses its imported `WireLimits`, and the publisher maps a
profile-construction failure to its existing redacted `InvalidSource` branch.
No wire bytes, limit values, or public error text changed.
