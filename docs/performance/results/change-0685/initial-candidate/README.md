# Initial candidate

All correctness gates and corpus parity passed. Longer controls confirmed tiny
query overhead; restore performs a temporary Arc downgrade only to compare weak
identity. The next candidate removes that unnecessary weak-count increment and
decrement with a non-dereferencing pointer identity comparison; the stored Weak
still keeps the allocation control block alive and prevents ABA reuse.
No causal claim is made for cold-query or open-time code-layout variation.
