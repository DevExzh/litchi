# Deferred umbrella coverage

Sources under this directory are preserved tests whose original APIs belong
to a family that is not yet wired into the dedicated facade, or to a former
cross-family flat-document facade. They remain grouped by semantic owner and
are intentionally outside Cargo's direct integration-test discovery. No
coverage source was deleted; each file is tracked for migration when its
dedicated owner becomes available.
