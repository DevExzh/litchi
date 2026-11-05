# 0362ul94j — Custom XML F5 compiler repair

This source-only repair starts from the combined OOXML candidate
`95973e8c6a2bcbb8724d1be50aea4c510fd8cdf3` (tree
`6854b92295c3aa15c97e783e9bee125be8edaeba`) and makes no acceptance claim.
It preserves the failed v2 F5 receipt: rows SHA-256
`cd41583e87c30c34128b7d42e075cc1584a42f502a8f4153a1f3c9a3897b8ae8` and
result SHA-256 `f1cc34a98fd896ba2a8a9b3abcbe74254a1b3eb0e33cd476046c5551df950fce`.

The Custom XML source scanners now read events first and resolve element names
through the immutable namespace resolver only in `Start` and `Empty` arms.
This removes the `NsReader` mutable/immutable borrow conflicts without
allocating a namespace copy per event.  The empty-element replacement-length
check uses the checked `usize` slash offset directly.  The source and test
repairs also remove an unnecessary local `super::errors` qualification,
replace five same-module `codec::read_props` qualifications, and compare
discovered `Props` values by value at the two affected assertions.

Only rustfmt and source diff checks are run for this repair; Cargo, rustdoc,
compiler, native, and performance execution remain pending.
