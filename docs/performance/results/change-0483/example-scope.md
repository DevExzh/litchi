# Example-target scope

The repository-wide example-target check failed at `example-targets-final-40`
because four example names are shared by the excluded iWork packages
`litchi-keynote`, `litchi-numbers`, and `litchi-pages`. The retained failure is
recorded in `validation/example-targets-final-40.stderr` and the package
inventory is recorded in `workspace-iwork-exclusion.json`.

`check-non-iwork-examples.py` reuses the repository checker's workspace,
manifest, and example-target resolvers. It validates every workspace manifest,
requires the exclusion inventory to match the discovered `litchi-iwa`,
`litchi-keynote`, `litchi-numbers`, and `litchi-pages` package families, and
reports the retained package and target counts. Collisions entirely inside the
excluded iWork family are outside this goal; collisions among retained
packages or across the retained/iWork boundary fail closed. This scoped gate
does not establish that the full workspace has unique example target names or
that the excluded iWork examples compile.
