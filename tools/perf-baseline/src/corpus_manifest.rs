//! Additive, content-addressed corpus metadata for performance reports.
//!
//! The report and the `results[*].corpus` object deliberately remain schema 1.
//! This module owns the independently versioned schema-2 catalog.  A schema-2
//! catalog can therefore be attached to a report as a reference or emitted as
//! a sidecar without changing the comparator's legacy case/corpus identity.

use std::{
    collections::{BTreeMap, BTreeSet},
    fmt,
    fs::File,
    io::Write,
    path::Path,
};

use serde::{Deserialize, Serialize};
use serde_json::{Value, json};
use sha2::{Digest, Sha256};

pub(crate) const MANIFEST_VERSION: u32 = 2;
pub(crate) const MANIFEST_KIND: &str = "corpus-catalog";

#[derive(Debug)]
pub(crate) struct ManifestError(String);

impl ManifestError {
    fn new(message: impl Into<String>) -> Self {
        Self(message.into())
    }
}

impl fmt::Display for ManifestError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        self.0.fmt(formatter)
    }
}

impl std::error::Error for ManifestError {}

/// The small reference embedded in an otherwise schema-1 report.
#[derive(Clone, Debug, Serialize, Deserialize, PartialEq, Eq)]
#[serde(deny_unknown_fields)]
pub(crate) struct CatalogReferenceV2 {
    pub(crate) manifest_version: u32,
    pub(crate) catalog_id: String,
    pub(crate) catalog_sha256: String,
    pub(crate) content_set_sha256: String,
}

#[derive(Clone, Debug, Serialize, Deserialize)]
#[serde(deny_unknown_fields)]
pub(crate) struct CorpusCatalogV2 {
    pub(crate) manifest_version: u32,
    pub(crate) manifest_kind: String,
    pub(crate) catalog_id: String,
    pub(crate) canonicalization: CanonicalizationV2,
    pub(crate) catalog_sha256: String,
    pub(crate) content_set_sha256: String,
    pub(crate) build: BuildIdentityV2,
    pub(crate) corpora: Vec<CorpusManifestV2>,
    pub(crate) case_bindings: Vec<CaseCorpusBindingV2>,
}

#[derive(Clone, Debug, Serialize, Deserialize)]
#[serde(deny_unknown_fields)]
pub(crate) struct CanonicalizationV2 {
    pub(crate) algorithm: String,
    pub(crate) hash: String,
}

#[derive(Clone, Debug, Serialize, Deserialize)]
#[serde(deny_unknown_fields)]
pub(crate) struct BuildIdentityV2 {
    pub(crate) tool: String,
    pub(crate) tool_version: String,
    pub(crate) git_revision: Option<String>,
    pub(crate) git_worktree_dirty: Option<bool>,
    pub(crate) source_files: Vec<SourceFileHashV2>,
}

#[derive(Clone, Debug, Serialize, Deserialize)]
#[serde(deny_unknown_fields)]
pub(crate) struct SourceFileHashV2 {
    pub(crate) path: String,
    pub(crate) sha256: String,
}

#[derive(Clone, Debug, Serialize, Deserialize)]
#[serde(deny_unknown_fields)]
pub(crate) struct CorpusManifestV2 {
    pub(crate) id: String,
    pub(crate) name: String,
    pub(crate) legacy_v1: LegacyCorpusManifestV1,
    pub(crate) format: String,
    pub(crate) size_class: String,
    pub(crate) categories: Vec<String>,
    pub(crate) generator: GeneratorV2,
    pub(crate) provenance: ProvenanceV2,
    pub(crate) bytes: ByteSummaryV2,
    pub(crate) shape_parameters: BTreeMap<String, Value>,
    pub(crate) relationships: RelationshipSummaryV2,
    pub(crate) security: SecuritySummaryV2,
    pub(crate) input: InputSummaryV2,
    pub(crate) limits: LimitsSummaryV2,
    pub(crate) members: MemberDigestSetV2,
    pub(crate) targets: Vec<TargetDigestV2>,
    pub(crate) coverage: CoverageV2,
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
#[serde(deny_unknown_fields)]
pub(crate) struct LegacyCorpusManifestV1 {
    pub(crate) name: String,
    pub(crate) generator: String,
    pub(crate) package_format: String,
    pub(crate) shape: String,
    pub(crate) payload_kind: String,
    pub(crate) compression: String,
    pub(crate) entry_count: usize,
    pub(crate) archive_member_count: usize,
    pub(crate) entry_bytes: usize,
    pub(crate) uncompressed_payload_bytes: usize,
    pub(crate) archive_bytes: usize,
    pub(crate) archive_sha256: String,
    pub(crate) target_entry: String,
    pub(crate) target_payload_bytes: usize,
    pub(crate) target_payload_sha256: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) rtf_variant: Option<String>,
    pub(crate) xlsx: Option<LegacyXlsxManifestV1>,
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
#[serde(deny_unknown_fields)]
pub(crate) struct LegacyXlsxManifestV1 {
    pub(crate) sheet_count: usize,
    pub(crate) rows_per_sheet: usize,
    pub(crate) columns_per_sheet: usize,
    pub(crate) one_percent_update_count: usize,
    pub(crate) source_members: LegacyXlsxSourceMembersV1,
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
#[serde(deny_unknown_fields)]
pub(crate) struct LegacyXlsxSourceMembersV1 {
    pub(crate) workbook: String,
    pub(crate) worksheets: Vec<String>,
    pub(crate) shared_strings: Option<String>,
    pub(crate) styles: Option<String>,
}

#[derive(Clone, Debug, Serialize, Deserialize)]
#[serde(deny_unknown_fields)]
pub(crate) struct GeneratorV2 {
    pub(crate) id: String,
    pub(crate) kind: String,
    pub(crate) revision: Option<String>,
    pub(crate) algorithm_id: Option<String>,
    pub(crate) seed_spec: Option<String>,
    pub(crate) parameters: BTreeMap<String, Value>,
}

#[derive(Clone, Debug, Serialize, Deserialize)]
#[serde(deny_unknown_fields)]
pub(crate) struct ProvenanceV2 {
    pub(crate) source_kind: String,
    pub(crate) source_path: Option<String>,
    pub(crate) producer: Option<String>,
    pub(crate) producer_version: Option<String>,
    pub(crate) source_sha256: Option<String>,
    pub(crate) license_spdx: Option<String>,
    pub(crate) license_evidence: Option<String>,
    pub(crate) redistributable: Option<bool>,
}

#[derive(Clone, Debug, Serialize, Deserialize)]
#[serde(deny_unknown_fields)]
pub(crate) struct ByteSummaryV2 {
    pub(crate) archive_bytes: u64,
    pub(crate) archive_sha256: String,
    pub(crate) logical_payload_bytes: u64,
    pub(crate) text_bytes: Option<u64>,
    pub(crate) media_bytes: Option<u64>,
    pub(crate) metadata_bytes: Option<u64>,
}

#[derive(Clone, Debug, Serialize, Deserialize)]
#[serde(deny_unknown_fields)]
pub(crate) struct RelationshipSummaryV2 {
    pub(crate) status: String,
    pub(crate) relationship_count: Option<u64>,
    pub(crate) dependency_closure_nodes: Option<u64>,
    pub(crate) dependency_closure_edges: Option<u64>,
    pub(crate) max_depth: Option<u64>,
    pub(crate) max_out_degree: Option<u64>,
}

#[derive(Clone, Debug, Serialize, Deserialize)]
#[serde(deny_unknown_fields)]
pub(crate) struct SecuritySummaryV2 {
    pub(crate) encryption: SecurityFeatureV2,
    pub(crate) signature: SecurityFeatureV2,
    pub(crate) protection: ProtectionFeatureV2,
    pub(crate) macros: SecurityFeatureV2,
    pub(crate) external_links: ExternalLinksV2,
}

#[derive(Clone, Debug, Serialize, Deserialize)]
#[serde(deny_unknown_fields)]
pub(crate) struct SecurityFeatureV2 {
    pub(crate) state: String,
    pub(crate) kind: Option<String>,
    pub(crate) evidence: String,
}

#[derive(Clone, Debug, Serialize, Deserialize)]
#[serde(deny_unknown_fields)]
pub(crate) struct ProtectionFeatureV2 {
    pub(crate) state: String,
    pub(crate) scopes: Vec<String>,
    pub(crate) evidence: String,
}

#[derive(Clone, Debug, Serialize, Deserialize)]
#[serde(deny_unknown_fields)]
pub(crate) struct ExternalLinksV2 {
    pub(crate) state: String,
    pub(crate) count: Option<u64>,
    pub(crate) targets_sha256: Option<String>,
    pub(crate) evidence: String,
}

#[derive(Clone, Debug, Serialize, Deserialize)]
#[serde(deny_unknown_fields)]
pub(crate) struct InputSummaryV2 {
    pub(crate) validity: String,
    pub(crate) malformation_kind: Option<String>,
    pub(crate) expected_behavior: String,
    pub(crate) within_limits: Option<bool>,
    pub(crate) evidence: String,
}

#[derive(Clone, Debug, Serialize, Deserialize)]
#[serde(deny_unknown_fields)]
pub(crate) struct LimitsSummaryV2 {
    pub(crate) profile_id: Option<String>,
    pub(crate) profile_sha256: Option<String>,
    pub(crate) observed: ObservedLimitsV2,
}

#[derive(Clone, Debug, Serialize, Deserialize)]
#[serde(deny_unknown_fields)]
pub(crate) struct ObservedLimitsV2 {
    pub(crate) input_bytes: Option<u64>,
    pub(crate) members: Option<u64>,
    pub(crate) relationships: Option<u64>,
    pub(crate) materialized_bytes: Option<u64>,
}

#[derive(Clone, Debug, Serialize, Deserialize)]
#[serde(deny_unknown_fields)]
pub(crate) struct MemberDigestSetV2 {
    pub(crate) status: String,
    pub(crate) items: Vec<MemberDigestV2>,
}

#[derive(Clone, Debug, Serialize, Deserialize)]
#[serde(deny_unknown_fields)]
pub(crate) struct MemberDigestV2 {
    pub(crate) ordinal: u64,
    pub(crate) name: String,
    pub(crate) kind: String,
    pub(crate) logical_bytes: u64,
    pub(crate) stored_bytes: Option<u64>,
    pub(crate) sha256: String,
    pub(crate) stored_sha256: Option<String>,
    pub(crate) role: Option<String>,
}

#[derive(Clone, Debug, Serialize, Deserialize)]
#[serde(deny_unknown_fields)]
pub(crate) struct TargetDigestV2 {
    pub(crate) entry: String,
    pub(crate) logical_bytes: u64,
    pub(crate) sha256: String,
}

#[derive(Clone, Debug, Serialize, Deserialize)]
#[serde(deny_unknown_fields)]
pub(crate) struct CoverageV2 {
    pub(crate) timed_cases: Vec<String>,
    pub(crate) guard_cases: Vec<String>,
    pub(crate) inventory_only: bool,
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq, Eq)]
#[serde(deny_unknown_fields)]
pub(crate) struct CaseCorpusBindingV2 {
    pub(crate) case: String,
    pub(crate) corpus_id: String,
    pub(crate) legacy_name: String,
    pub(crate) legacy_archive_sha256: String,
    pub(crate) role: String,
}

/// Owned JSON input keeps this module independent of the private report types.
#[derive(Clone, Debug)]
pub(crate) struct LegacyCaseCorpus {
    pub(crate) case: String,
    pub(crate) corpus: Value,
}

impl CorpusCatalogV2 {
    pub(crate) fn from_legacy_results(
        records: &[LegacyCaseCorpus],
        build: BuildIdentityV2,
    ) -> Result<Self, ManifestError> {
        let mut corpora = BTreeMap::<String, CorpusManifestV2>::new();
        let mut bindings = Vec::with_capacity(records.len());
        let mut binding_keys = BTreeSet::new();

        for record in records {
            let legacy: LegacyCorpusManifestV1 = serde_json::from_value(record.corpus.clone())
                .map_err(|error| ManifestError::new(format!("invalid V1 corpus: {error}")))?;
            let id = content_id(&legacy.package_format, &legacy.archive_sha256);
            let mut corpus = CorpusManifestV2::from_legacy(legacy.clone())?;
            corpus.coverage.timed_cases.push(record.case.clone());
            corpus.coverage.timed_cases.sort();
            corpus.coverage.timed_cases.dedup();

            if let Some(existing) = corpora.get_mut(&id) {
                if existing.legacy_v1 != legacy {
                    return Err(ManifestError::new(format!(
                        "content id {id} maps to conflicting V1 corpus objects"
                    )));
                }
                existing.coverage.timed_cases.push(record.case.clone());
                existing.coverage.timed_cases.sort();
                existing.coverage.timed_cases.dedup();
            } else {
                corpora.insert(id.clone(), corpus);
            }

            let binding_key = (record.case.clone(), id.clone());
            if !binding_keys.insert(binding_key) {
                return Err(ManifestError::new(format!(
                    "duplicate case/corpus binding: {} / {id}",
                    record.case
                )));
            }
            bindings.push(CaseCorpusBindingV2 {
                case: record.case.clone(),
                corpus_id: id,
                legacy_name: legacy.name,
                legacy_archive_sha256: legacy.archive_sha256,
                role: "timed".to_owned(),
            });
        }

        let mut catalog = Self {
            manifest_version: MANIFEST_VERSION,
            manifest_kind: MANIFEST_KIND.to_owned(),
            catalog_id: "litchi-perf-corpus-v2".to_owned(),
            canonicalization: CanonicalizationV2 {
                algorithm: "sorted-json-utf8-compact-v1".to_owned(),
                hash: "sha256".to_owned(),
            },
            catalog_sha256: String::new(),
            content_set_sha256: String::new(),
            build,
            corpora: corpora.into_values().collect(),
            case_bindings: bindings,
        };
        catalog.case_bindings.sort_by(|left, right| {
            left.case
                .cmp(&right.case)
                .then_with(|| left.corpus_id.cmp(&right.corpus_id))
        });
        catalog.refresh_hashes()?;
        catalog.validate()?;
        Ok(catalog)
    }

    pub(crate) fn reference(&self) -> CatalogReferenceV2 {
        CatalogReferenceV2 {
            manifest_version: self.manifest_version,
            catalog_id: self.catalog_id.clone(),
            catalog_sha256: self.catalog_sha256.clone(),
            content_set_sha256: self.content_set_sha256.clone(),
        }
    }

    pub(crate) fn refresh_hashes(&mut self) -> Result<(), ManifestError> {
        self.catalog_sha256.clear();
        self.content_set_sha256.clear();
        self.content_set_sha256 = content_set_sha256(self)?;
        self.catalog_sha256 = catalog_sha256(self)?;
        Ok(())
    }

    pub(crate) fn validate(&self) -> Result<(), ManifestError> {
        if self.manifest_version != MANIFEST_VERSION {
            return Err(ManifestError::new(format!(
                "unsupported corpus manifest version {}",
                self.manifest_version
            )));
        }
        if self.manifest_kind != MANIFEST_KIND {
            return Err(ManifestError::new(format!(
                "unexpected corpus manifest kind {:?}",
                self.manifest_kind
            )));
        }
        if self.canonicalization.algorithm != "sorted-json-utf8-compact-v1"
            || self.canonicalization.hash != "sha256"
        {
            return Err(ManifestError::new("unsupported corpus canonicalization"));
        }
        if !is_sha256(&self.catalog_sha256) || !is_sha256(&self.content_set_sha256) {
            return Err(ManifestError::new(
                "catalog hashes must be lowercase SHA-256",
            ));
        }
        let mut ids = BTreeSet::<String>::new();
        for corpus in &self.corpora {
            if !ids.insert(corpus.id.clone()) {
                return Err(ManifestError::new(format!(
                    "duplicate corpus id {}",
                    corpus.id
                )));
            }
            validate_corpus(corpus)?;
        }
        let mut previous = None;
        for corpus in &self.corpora {
            if let Some(previous) = previous
                && previous >= &corpus.id
            {
                return Err(ManifestError::new("corpora must be sorted by id"));
            }
            previous = Some(&corpus.id);
        }
        let mut binding_keys = BTreeSet::new();
        for binding in &self.case_bindings {
            if !ids.contains(&binding.corpus_id) {
                return Err(ManifestError::new(format!(
                    "binding references unknown corpus {}",
                    binding.corpus_id
                )));
            }
            if !binding_keys.insert((&binding.case, &binding.corpus_id)) {
                return Err(ManifestError::new(format!(
                    "duplicate case/corpus binding {} / {}",
                    binding.case, binding.corpus_id
                )));
            }
        }
        let expected_content_set_sha256 = content_set_sha256(self)?;
        if expected_content_set_sha256 != self.content_set_sha256 {
            return Err(ManifestError::new(
                "content-set SHA-256 does not match catalog",
            ));
        }
        let expected_catalog_sha256 = catalog_sha256(self)?;
        if expected_catalog_sha256 != self.catalog_sha256 {
            return Err(ManifestError::new("catalog SHA-256 does not match catalog"));
        }
        Ok(())
    }
}

impl CorpusManifestV2 {
    pub(crate) fn from_legacy(legacy: LegacyCorpusManifestV1) -> Result<Self, ManifestError> {
        let archive_bytes = u64::try_from(legacy.archive_bytes)
            .map_err(|_| ManifestError::new("archive byte count does not fit u64"))?;
        let archive_member_count = u64::try_from(legacy.archive_member_count).ok();
        let logical_payload_bytes = u64::try_from(legacy.uncompressed_payload_bytes)
            .map_err(|_| ManifestError::new("logical byte count does not fit u64"))?;
        let entry_count = u64::try_from(legacy.entry_count)
            .map_err(|_| ManifestError::new("entry count does not fit u64"))?;
        let entry_bytes = u64::try_from(legacy.entry_bytes)
            .map_err(|_| ManifestError::new("entry byte count does not fit u64"))?;
        let target_payload_bytes = u64::try_from(legacy.target_payload_bytes)
            .map_err(|_| ManifestError::new("target byte count does not fit u64"))?;
        let id = content_id(&legacy.package_format, &legacy.archive_sha256);
        let mut categories = vec!["legacy-migrated".to_owned()];
        if legacy.generator.contains("synthetic") {
            categories.push("synthetic".to_owned());
        }
        if legacy.payload_kind == "compressible" {
            categories.push("highly-compressible".to_owned());
        } else if legacy.payload_kind == "incompressible" {
            categories.push("incompressible".to_owned());
        }
        if legacy.shape == "many-small" {
            categories.push("many-small-parts".to_owned());
        }
        if legacy.shape == "few-large" {
            categories.push("few-large-parts".to_owned());
        }
        categories.sort();
        categories.dedup();

        let mut generator_parameters = BTreeMap::new();
        generator_parameters.insert("legacy_shape".to_owned(), json!(legacy.shape));
        generator_parameters.insert("legacy_payload_kind".to_owned(), json!(legacy.payload_kind));
        let generator_revision = legacy
            .generator
            .rsplit_once("-v")
            .map(|(_, revision)| format!("v{revision}"));
        let generated = legacy.generator.contains("synthetic");

        let mut shape_parameters = BTreeMap::new();
        shape_parameters.insert("entry_count".to_owned(), json!(entry_count));
        shape_parameters.insert("entry_bytes".to_owned(), json!(entry_bytes));
        if let Some(xlsx) = &legacy.xlsx {
            shape_parameters.insert("sheet_count".to_owned(), json!(xlsx.sheet_count));
            shape_parameters.insert("rows_per_sheet".to_owned(), json!(xlsx.rows_per_sheet));
            shape_parameters.insert(
                "columns_per_sheet".to_owned(),
                json!(xlsx.columns_per_sheet),
            );
        }

        let legacy_package_format = legacy.package_format.clone();
        let legacy_shape = legacy.shape.clone();
        let legacy_generator = legacy.generator.clone();
        let legacy_archive_sha256 = legacy.archive_sha256.clone();
        let legacy_target_entry = legacy.target_entry.clone();
        let legacy_target_payload_sha256 = legacy.target_payload_sha256.clone();
        Ok(Self {
            id,
            name: legacy.name.clone(),
            legacy_v1: legacy,
            format: legacy_package_format,
            size_class: match legacy_shape.as_str() {
                "tiny" => "tiny",
                "medium" => "medium",
                "large" => "large",
                _ => "unknown",
            }
            .to_owned(),
            categories,
            generator: GeneratorV2 {
                id: legacy_generator,
                kind: if generated { "synthetic" } else { "unknown" }.to_owned(),
                revision: generator_revision,
                algorithm_id: None,
                seed_spec: None,
                parameters: generator_parameters,
            },
            provenance: ProvenanceV2 {
                source_kind: if generated { "generated" } else { "unknown" }.to_owned(),
                source_path: None,
                producer: if generated {
                    Some("Litchi deterministic generator".to_owned())
                } else {
                    None
                },
                producer_version: None,
                source_sha256: None,
                license_spdx: if generated {
                    Some("Apache-2.0".to_owned())
                } else {
                    None
                },
                license_evidence: if generated {
                    Some("repository-license".to_owned())
                } else {
                    None
                },
                redistributable: if generated { Some(true) } else { None },
            },
            bytes: ByteSummaryV2 {
                archive_bytes,
                archive_sha256: legacy_archive_sha256,
                logical_payload_bytes,
                text_bytes: None,
                media_bytes: None,
                metadata_bytes: None,
            },
            shape_parameters,
            relationships: RelationshipSummaryV2 {
                status: "unknown".to_owned(),
                relationship_count: None,
                dependency_closure_nodes: None,
                dependency_closure_edges: None,
                max_depth: None,
                max_out_degree: None,
            },
            security: SecuritySummaryV2::unknown(),
            input: InputSummaryV2 {
                validity: "unknown".to_owned(),
                malformation_kind: None,
                expected_behavior: "unknown".to_owned(),
                within_limits: None,
                evidence: "not-recorded".to_owned(),
            },
            limits: LimitsSummaryV2 {
                profile_id: None,
                profile_sha256: None,
                observed: ObservedLimitsV2 {
                    input_bytes: Some(archive_bytes),
                    members: archive_member_count,
                    relationships: None,
                    materialized_bytes: None,
                },
            },
            members: MemberDigestSetV2 {
                status: "unavailable".to_owned(),
                items: Vec::new(),
            },
            targets: vec![TargetDigestV2 {
                entry: legacy_target_entry,
                logical_bytes: target_payload_bytes,
                sha256: legacy_target_payload_sha256,
            }],
            coverage: CoverageV2 {
                timed_cases: Vec::new(),
                guard_cases: Vec::new(),
                inventory_only: false,
            },
        })
    }
}

impl SecuritySummaryV2 {
    fn unknown() -> Self {
        Self {
            encryption: SecurityFeatureV2::unknown(),
            signature: SecurityFeatureV2::unknown(),
            protection: ProtectionFeatureV2 {
                state: "unknown".to_owned(),
                scopes: Vec::new(),
                evidence: "not-recorded".to_owned(),
            },
            macros: SecurityFeatureV2::unknown(),
            external_links: ExternalLinksV2 {
                state: "unknown".to_owned(),
                count: None,
                targets_sha256: None,
                evidence: "not-recorded".to_owned(),
            },
        }
    }
}

impl SecurityFeatureV2 {
    fn unknown() -> Self {
        Self {
            state: "unknown".to_owned(),
            kind: None,
            evidence: "not-recorded".to_owned(),
        }
    }
}

fn validate_corpus(corpus: &CorpusManifestV2) -> Result<(), ManifestError> {
    if !is_sha256(&corpus.bytes.archive_sha256)
        || corpus.bytes.archive_sha256 != corpus.legacy_v1.archive_sha256
    {
        return Err(ManifestError::new(format!(
            "archive hash mismatch for {}",
            corpus.id
        )));
    }
    if !corpus.id.ends_with(&corpus.bytes.archive_sha256) {
        return Err(ManifestError::new(format!(
            "content id does not contain archive hash for {}",
            corpus.id
        )));
    }
    if !is_sha256(&corpus.legacy_v1.target_payload_sha256) {
        return Err(ManifestError::new(format!(
            "target hash is not SHA-256 for {}",
            corpus.id
        )));
    }
    let mut categories = corpus.categories.clone();
    categories.sort();
    categories.dedup();
    if categories != corpus.categories {
        return Err(ManifestError::new(format!(
            "categories must be sorted and unique for {}",
            corpus.id
        )));
    }
    Ok(())
}

fn content_id(package_format: &str, archive_sha256: &str) -> String {
    format!("{}:sha256:{archive_sha256}", format_slug(package_format))
}

fn format_slug(value: &str) -> String {
    let mut slug = String::new();
    for character in value.chars() {
        if character.is_ascii_alphanumeric() {
            slug.push(character.to_ascii_lowercase());
        } else if !slug.ends_with('-') {
            slug.push('-');
        }
    }
    slug.trim_matches('-').to_owned()
}

fn catalog_sha256(catalog: &CorpusCatalogV2) -> Result<String, ManifestError> {
    let mut value = serde_json::to_value(catalog)
        .map_err(|error| ManifestError::new(format!("serialize catalog: {error}")))?;
    value
        .as_object_mut()
        .ok_or_else(|| ManifestError::new("catalog did not serialize as an object"))?
        .remove("catalog_sha256");
    Ok(sha256_hex(&canonical_json_bytes(&value)?))
}

fn content_set_sha256(catalog: &CorpusCatalogV2) -> Result<String, ManifestError> {
    let corpora = catalog
        .corpora
        .iter()
        .map(|corpus| {
            json!({
                "id": corpus.id,
                "archive_sha256": corpus.bytes.archive_sha256,
                "members": corpus.members.items.iter().map(|member| {
                    json!({
                        "ordinal": member.ordinal,
                        "name": member.name,
                        "sha256": member.sha256,
                    })
                }).collect::<Vec<_>>(),
            })
        })
        .collect::<Vec<_>>();
    let bindings = catalog
        .case_bindings
        .iter()
        .map(|binding| {
            json!({
                "case": binding.case,
                "corpus_id": binding.corpus_id,
                "role": binding.role,
            })
        })
        .collect::<Vec<_>>();
    let value = json!({
        "corpora": corpora,
        "case_bindings": bindings,
    });
    Ok(sha256_hex(&canonical_json_bytes(&value)?))
}

fn canonical_json_bytes(value: &Value) -> Result<Vec<u8>, ManifestError> {
    let mut output = Vec::new();
    write_canonical_json(value, &mut output)?;
    Ok(output)
}

fn write_canonical_json(value: &Value, output: &mut Vec<u8>) -> Result<(), ManifestError> {
    match value {
        Value::Null | Value::Bool(_) | Value::Number(_) | Value::String(_) => {
            serde_json::to_writer(output, value)
                .map_err(|error| ManifestError::new(format!("serialize JSON scalar: {error}")))?;
        },
        Value::Array(values) => {
            output.push(b'[');
            for (index, value) in values.iter().enumerate() {
                if index != 0 {
                    output.push(b',');
                }
                write_canonical_json(value, output)?;
            }
            output.push(b']');
        },
        Value::Object(values) => {
            let ordered = values.iter().collect::<BTreeMap<_, _>>();
            output.push(b'{');
            for (index, (key, value)) in ordered.into_iter().enumerate() {
                if index != 0 {
                    output.push(b',');
                }
                serde_json::to_writer(&mut *output, key).map_err(|error| {
                    ManifestError::new(format!("serialize JSON object key: {error}"))
                })?;
                output.push(b':');
                write_canonical_json(value, output)?;
            }
            output.push(b'}');
        },
    }
    Ok(())
}

fn is_sha256(value: &str) -> bool {
    value.len() == 64
        && value.bytes().all(|byte| byte.is_ascii_hexdigit())
        && value == value.to_ascii_lowercase()
}

fn sha256_hex(bytes: &[u8]) -> String {
    let digest = Sha256::digest(bytes);
    let mut output = String::with_capacity(digest.len() * 2);
    for byte in digest {
        use fmt::Write as _;
        let _ = write!(output, "{byte:02x}");
    }
    output
}

pub(crate) fn write_catalog(path: &Path, catalog: &CorpusCatalogV2) -> Result<(), ManifestError> {
    let parent = path
        .parent()
        .filter(|parent| !parent.as_os_str().is_empty());
    if let Some(parent) = parent {
        std::fs::create_dir_all(parent)
            .map_err(|error| ManifestError::new(format!("create catalog directory: {error}")))?;
    }
    let mut file = File::create(path)
        .map_err(|error| ManifestError::new(format!("create corpus catalog: {error}")))?;
    serde_json::to_writer_pretty(&mut file, catalog)
        .map_err(|error| ManifestError::new(format!("write corpus catalog: {error}")))?;
    file.write_all(b"\n")
        .map_err(|error| ManifestError::new(format!("finish corpus catalog: {error}")))?;
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    fn legacy() -> LegacyCorpusManifestV1 {
        LegacyCorpusManifestV1 {
            name: "tiny-compressible".to_owned(),
            generator: "litchi-opc-synthetic-v2".to_owned(),
            package_format: "OPC/ZIP".to_owned(),
            shape: "tiny".to_owned(),
            payload_kind: "compressible".to_owned(),
            compression: "deflate".to_owned(),
            entry_count: 3,
            archive_member_count: 5,
            entry_bytes: 512,
            uncompressed_payload_bytes: 1536,
            archive_bytes: 1310,
            archive_sha256: "1e28b8a9049a82f07e8ea88b2d492ef522d2da793d22fa50e2fe7f354dca3e2a"
                .to_owned(),
            target_entry: "benchmark/parts/00001.bin".to_owned(),
            target_payload_bytes: 512,
            target_payload_sha256:
                "3f2d4c8a3b7e1db8c7f35c67c9af1db12d87b4f1c42b4d0a0c5a6f9e8b7c6d5e".to_owned(),
            rtf_variant: None,
            xlsx: None,
        }
    }

    fn build(records: &[&str]) -> CorpusCatalogV2 {
        let records = records
            .iter()
            .map(|case| LegacyCaseCorpus {
                case: (*case).to_owned(),
                corpus: serde_json::to_value(legacy()).unwrap(),
            })
            .collect::<Vec<_>>();
        CorpusCatalogV2::from_legacy_results(
            &records,
            BuildIdentityV2 {
                tool: "litchi-perf-baseline".to_owned(),
                tool_version: "0.1.0".to_owned(),
                git_revision: Some("revision".to_owned()),
                git_worktree_dirty: Some(false),
                source_files: Vec::new(),
            },
        )
        .unwrap()
    }

    #[test]
    fn migration_preserves_legacy_projection() {
        let original = legacy();
        let migrated = CorpusManifestV2::from_legacy(original.clone()).unwrap();
        assert_eq!(migrated.legacy_v1, original);
        assert_eq!(migrated.security.encryption.state, "unknown");
        assert_eq!(migrated.provenance.source_kind, "generated");
        assert_eq!(migrated.limits.profile_id, None);
    }

    #[test]
    fn catalog_hash_is_independent_of_record_order() {
        let first = build(&["zip_index", "zip_read_one"]);
        let second = build(&["zip_read_one", "zip_index"]);
        assert_eq!(first.catalog_sha256, second.catalog_sha256);
        assert_eq!(first.content_set_sha256, second.content_set_sha256);
        assert_eq!(first.case_bindings, second.case_bindings);
    }

    #[test]
    fn duplicate_case_binding_is_rejected() {
        let records = vec![
            LegacyCaseCorpus {
                case: "zip_index".to_owned(),
                corpus: serde_json::to_value(legacy()).unwrap(),
            },
            LegacyCaseCorpus {
                case: "zip_index".to_owned(),
                corpus: serde_json::to_value(legacy()).unwrap(),
            },
        ];
        assert!(
            CorpusCatalogV2::from_legacy_results(
                &records,
                BuildIdentityV2 {
                    tool: "tool".to_owned(),
                    tool_version: "version".to_owned(),
                    git_revision: None,
                    git_worktree_dirty: None,
                    source_files: Vec::new(),
                },
            )
            .is_err()
        );
    }
}
