# Deferred ODP test disposition

There are no deferred ODP Rust test sources. Legacy tests were adjudicated against the ADR-aligned public owners rather than restoring the attached mutable model.

| Removed source | Active owner | Disposition |
|---|---|---|
| `odp_drawing_style_resources.rs` | `tests/odp_drawing_style_resources.rs` | Moved to shared ODF drawing-resource codecs and the ODP `Presentation` facade; mutable equality was replaced by immutable presentation reads. |
| `odp_handout_master.rs` | `tests/odp_handout_master.rs` | Migrated all real ODP fixtures and the UTF-8 BOM fragment regression. The vacuous wrong-family ODT test was superseded by valid-ODP absence/error coverage. |
| `odp_image_authoring.rs` | `tests/odp_image_authoring.rs`, `tests/presentation_edit.rs` | Migrated typed creation to the ODP builder and shared image validation/path allocation. Source transactions verify explicit media staging and refuse lossy retained-page regeneration. |
| `odp_layout_master_mutation.rs` | `tests/odp_layout_master.rs` | Migrated reorder, replacement-on-remove, unknown XML preservation, compact output, lexical limits, active-content refusal, and atomic failures. |
| `odp_presentation_settings_references.rs` | `tests/odp_presentation_settings.rs` | Migrated builder reference validation, lexical fixtures, and source-checked page identity/reference transactions. Unrestricted `settings_mut` access remains intentionally unavailable. |
| `odp_save_fidelity.rs` | `tests/odp_save_fidelity.rs`, `tests/presentation_edit.rs` | Exact-byte no-op commits supersede lossy model reserialization; retained-page rewrites are refused unless lossless, while newly appended slides preserve source fragments. |

All referenced XML literals remain compact, and the real fixtures under `test-data/` remain active inputs.
