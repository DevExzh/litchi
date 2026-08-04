# Formula legacy migration

The files beside this document are source-preserving relocations of the old
`litchi-odf::formula` tree. They are intentionally separate from the current
formula migration implementation and are not declared by the active module
tree: the old sources depend on umbrella package and root types.

The active formula crate owns the canonical `model`, `codec`, `package`,
`authoring`, and facade layers. The relocated files remain available as the
behavior-preserving reference for the later migration without adding aliases
or a compatibility path.
