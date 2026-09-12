//! Deterministic scale and identity metadata for the complete XLDM 140 profile.
//!
//! The correctness runner instantiates every catalog point.  Performance
//! sampling remains a separate, not-yet-authorized phase; the catalog and
//! generated source recipe are still shared so receipts cannot silently
//! invent a scale or relationship topology.

pub const RECIPE_ID: &str = "synthetic-complete-xldm140-table-identity";
pub const RECIPE_VERSION: u32 = 2;
pub const RECIPE_SOURCE: &str = "crates/litchi-xlsb/tests/data_model_identity.rs";
pub const RECIPE_GENERATOR: &str = "docs/report/spec-gap-validation-evidence/xlsb-model-identity-performance/harness/scaled_fixture.rs";

pub const NAME_PROFILES: &[&str] = &[
    "same_length_ascii",
    "shorter",
    "longer",
    "escaped_xml",
    "unicode",
];

pub const ENDPOINT_LAYOUTS: &[&str] = &["selected_table", "distributed"];

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct ScaleCase {
    pub family: &'static str,
    pub tables: usize,
    pub relationships: usize,
}

/// The bounded Cartesian table/relationship matrix from requirements.md.
pub const SCALE_MATRIX: &[ScaleCase] = &[
    ScaleCase {
        family: "tiny",
        tables: 1,
        relationships: 0,
    },
    ScaleCase {
        family: "small",
        tables: 4,
        relationships: 0,
    },
    ScaleCase {
        family: "small",
        tables: 4,
        relationships: 3,
    },
    ScaleCase {
        family: "small",
        tables: 4,
        relationships: 8,
    },
    ScaleCase {
        family: "medium",
        tables: 16,
        relationships: 0,
    },
    ScaleCase {
        family: "medium",
        tables: 16,
        relationships: 15,
    },
    ScaleCase {
        family: "medium",
        tables: 16,
        relationships: 32,
    },
    ScaleCase {
        family: "medium",
        tables: 16,
        relationships: 64,
    },
    ScaleCase {
        family: "large",
        tables: 64,
        relationships: 0,
    },
    ScaleCase {
        family: "large",
        tables: 64,
        relationships: 63,
    },
    ScaleCase {
        family: "large",
        tables: 64,
        relationships: 128,
    },
    ScaleCase {
        family: "large",
        tables: 64,
        relationships: 256,
    },
];

#[must_use]
pub const fn smoke_case(table_count: usize, relationship_count: usize) -> ScaleCase {
    if table_count == 1 && relationship_count == 0 {
        ScaleCase {
            family: "tiny",
            tables: 1,
            relationships: 0,
        }
    } else {
        ScaleCase {
            family: "smoke",
            tables: table_count,
            relationships: relationship_count,
        }
    }
}
