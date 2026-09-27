# Offline replay integration notes

No native capture was repeated or discarded during these corrections.

- The validator initially called `observe.analyze()`; the observer implemented
  `observe.observe()`. The call raised a `TypeError` before validation completed.
  Root corrected the API wiring, and full replay passed.
- A cleanup preflight overlapped the requested observer field-label/argument
  validation update, before `observations.json` had been regenerated. It failed
  with `observations.json does not replay exactly`. The guard ran before any
  target removal. Cleanup was retried only after the updated offline observer
  and its generated receipt agreed.

These are validator integration events, not failed native measurements or
changes to the raw measurements.
