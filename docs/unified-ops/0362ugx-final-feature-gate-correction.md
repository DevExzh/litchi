# 0362ugx — final feature gate correction

The full 53-step gate on `b6732d02b241705ba153c1a2cb9ff03bf0a13f93`
passed its first 17 steps, then failed the ODG-only strict Clippy profile.
`odf_staging_error` was compiled without a caller in that feature set.
Its production callers are ODC, ODI and ODF formula; its unit tests require
ODB. The helper now has the same feature availability as those callers.
Classification, validation and runtime algorithms are unchanged.

The initial correction `cbbad7a7c02aafa4052d40eb45b541638a46b31a` still
enabled the helper for every test build. ODG's library-test target correctly
rejected that unused helper. Final correction
`e1576fa6dcd9f856ddea0b0ec6b038cb6b076265` restricts the test alternative
to `all(test, feature = "odb")`.

On that clean source, all eight remaining strict feature profiles pass:
ODG, ODC, ODI, ODB, ODF formula, Pages, Keynote and Numbers. They use
`cargo clippy -q -p litchi-ops --no-default-features --features FEATURE
--all-targets -- -D warnings`. The target finishes at 5,613,723,699 bytes,
below the 20,000,000,000-byte limit. Each passed command has an empty log,
no remaining owned builders and a stable target after exit.

[The reduced receipt](results/0362ugx/feature-profiles.json) records every
row from the two failed attempts and the passing eight-profile selection,
including exact commands, source/lock identities, external log sizes and
SHA-256 hashes. The coordinator independently checked those log identities.
These selected passes do not constitute the full gate; a fresh full run on
the next composed candidate remains required.

The first attempt at the final ODG profile was operationally rejected when
an iWork fixture generator launched Cargo from the gate checkout. The
ownership guard remained enabled. After the owner confirmed all jobs had
exited and the coordinator checked for live builders, the same frozen
selection resumed and passed. Native tools and their fixture generators
must run from their own frozen endpoints. The retained iWork owner
validation is content-only diagnostic evidence and explicitly rejects an
acceptance claim; its correction receipt is bound in the reduced gate
record. No native timing result is claimed here.
