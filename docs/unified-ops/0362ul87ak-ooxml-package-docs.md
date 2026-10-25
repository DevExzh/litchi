# 0362ul87ak — OOXML package projection documentation repair

The 87aj semantic-publication closure at `e56c866eada10efef5be3d8ae37dc65f8abeaa6c`
left two private package-module summaries describing mutation even though
those modules now only inspect graphs and read projections. Their headings
now say “Package graph discovery”, matching the remaining implementation and
the shared Edit/Commit publication boundary.

This is a documentation-only follow-up. It adds no publisher, alias,
deprecated twin, or consumer exception. No Cargo, rustdoc, test, native,
performance, or F5 command was run for this follow-up.
