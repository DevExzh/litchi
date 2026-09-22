The first retention capture completed all 24 processes with exit code zero.
Post-capture checking failed at retention.py line 43 because the checker compared
CLI names default/release with report names default_8mib/release_8mib.
The archived checker preserves that failed expectation. The corrected checker
maps the names and adds analyze mode to process these exact existing captures.
No capture was rerun, discarded or modified.
