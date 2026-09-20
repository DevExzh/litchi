# Rejected profile preflight

The locked offline dependency build passed after the harness package-name repair. Harness compilation then rejected a dereference of a Boolean returned by value (E0614). Commit f499fe18e6 removes that dereference. No preflight cases or timing samples ran.

Raw logs and target cleanup receipt are retained alongside this note.
