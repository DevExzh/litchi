//! The parity case: fixture, tier, evidence, the two routes and the reads.

use std::fmt::{Debug, Write as _};
use std::path::{Path, PathBuf};

use crate::archive_diff::{self, MetadataPolicy, first_difference};
use crate::known_fields::KnownFields;

/// Result type of routes, reads and cases.
pub type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

/// The iWork format of a case; selects the leaf reopen and the schemas.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Format {
    /// `.numbers`, leaf `litchi_numbers::Package`.
    Numbers,
    /// `.pages`, leaf `litchi_pages::Package`.
    Pages,
    /// `.key`, leaf `litchi_keynote::Package`.
    Keynote,
}

impl Format {
    /// Strictly reopen `bytes` through the leaf package of this format.
    pub fn reopen(self, bytes: &[u8]) -> TestResult {
        match self {
            Self::Numbers => drop(litchi_numbers::Package::from_bytes(bytes)?),
            Self::Pages => drop(litchi_pages::Package::from_bytes(bytes)?),
            Self::Keynote => drop(litchi_keynote::Package::from_bytes(bytes)?),
        }
        Ok(())
    }
}

/// The parity tier a case proves (record 0363 §4.1).
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Tier {
    /// Host and leaf outputs are byte-identical.
    A,
    /// Untouched members and objects byte-identical; changed objects equal
    /// to the host on every known field.
    B,
}

/// Whether the host route has native Apple evidence (owner decision D1).
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Evidence {
    /// The host route was verified in the native application; the argument
    /// is a tracked path (relative to the repository root) that records it,
    /// usually a `test-data/iwork/**-receipt.json`.
    Native(&'static str),
    /// No native evidence: the migrated route must have a row in
    /// `docs/unified-ops/0363-mac-queue.md` (checked by
    /// `tools/iwa_exit_ledger.py`).
    Queued,
}

/// Repository root.
pub fn repository_root() -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR")).join("../..")
}

/// A tracked input package.
pub struct Fixture {
    /// Path under `test-data/`.
    pub path: &'static str,
    /// Package bytes.
    pub bytes: Vec<u8>,
}

/// Load `test-data/<path>`.
pub fn fixture(path: &'static str) -> TestResult<Fixture> {
    let full = repository_root().join("test-data").join(path);
    let bytes =
        std::fs::read(&full).map_err(|error| format!("fixture {}: {error}", full.display()))?;
    Ok(Fixture { path, bytes })
}

/// What the leaf route returns.
pub struct LeafRun {
    /// The committed package's bytes.
    pub output: Vec<u8>,
    /// The bytes obtained by applying the commit's inverse patch to the
    /// committed package, when the route checks the inverse.
    pub restored: Option<Vec<u8>>,
}

impl LeafRun {
    /// A leaf output without an inverse check.
    pub fn output(output: Vec<u8>) -> Self {
        Self {
            output,
            restored: None,
        }
    }

    /// A leaf output whose inverse must restore the input.
    pub fn with_inverse(output: Vec<u8>, restored: Vec<u8>) -> Self {
        Self {
            output,
            restored: Some(restored),
        }
    }
}

type HostRoute<'a> = Box<dyn FnOnce(&[u8]) -> TestResult<Vec<u8>> + 'a>;
type LeafRoute<'a> = Box<dyn FnOnce(&[u8]) -> TestResult<LeafRun> + 'a>;
type ReadFn<'a> = Box<dyn Fn(&[u8]) -> TestResult<String> + 'a>;

/// One differential case. Build it with [`Case::new`], add the two routes
/// and the reads, then [`Case::run`].
pub struct Case<'a> {
    name: &'static str,
    format: Format,
    tier: Tier,
    evidence: Evidence,
    fixture: &'a Fixture,
    host: Option<HostRoute<'a>>,
    leaf: Option<LeafRoute<'a>>,
    reads: Vec<(&'static str, ReadFn<'a>)>,
    known: KnownFields,
    read_only: bool,
    leaf_drops: &'static [&'static str],
    metadata: MetadataPolicy,
}

/// What a passing case observed.
pub struct Report {
    /// `case name: tier, fixture`.
    pub header: String,
    /// The three-way comparison (notes only when the case passed).
    pub diff: archive_diff::Diff,
}

impl Case<'_> {
    /// Start a case. `name` should be the `#[test]` function name: the
    /// ledger cites it.
    pub fn new<'a>(
        name: &'static str,
        format: Format,
        tier: Tier,
        evidence: Evidence,
        fixture: &'a Fixture,
    ) -> Case<'a> {
        Case {
            name,
            format,
            tier,
            evidence,
            fixture,
            host: None,
            leaf: None,
            reads: Vec::new(),
            known: KnownFields::for_format(format),
            read_only: false,
            leaf_drops: &[],
            metadata: MetadataPolicy::default(),
        }
    }
}

impl<'a> Case<'a> {
    /// The host route: input bytes to output bytes through `litchi_iwa`.
    #[must_use]
    pub fn host(mut self, route: impl FnOnce(&[u8]) -> TestResult<Vec<u8>> + 'a) -> Self {
        self.host = Some(Box::new(route));
        self
    }

    /// The leaf route: input bytes to output bytes through the leaf crate.
    #[must_use]
    pub fn leaf(mut self, route: impl FnOnce(&[u8]) -> TestResult<LeafRun> + 'a) -> Self {
        self.leaf = Some(Box::new(route));
        self
    }

    /// A read applied to both outputs; the two results must be equal.
    /// Add host reads and leaf reads alike.
    #[must_use]
    pub fn read<T: Debug>(
        mut self,
        label: &'static str,
        read: impl Fn(&[u8]) -> TestResult<T> + 'a,
    ) -> Self {
        self.reads.push((
            label,
            Box::new(move |bytes| Ok(format!("{:#?}", read(bytes)?))),
        ));
        self
    }

    /// Mark a read-route case: both routes open the fixture and save it
    /// unchanged (both outputs must equal the input), and the reads carry
    /// the comparison.
    #[must_use]
    pub fn read_only(mut self) -> Self {
        self.read_only = true;
        self
    }

    /// Accept that the leaf deletes Keynote's three root previews
    /// (`preview.jpg`, `preview-micro.jpg`, `preview-web.jpg`) where the
    /// host keeps them stale. The leaf's rendering invalidation is its
    /// documented, natively verified policy (ADR 0028: "Each Rust candidate
    /// deleted all three root previews; Keynote regenerated them on
    /// resave"); every other member and object is still compared.
    #[must_use]
    pub fn leaf_invalidates_root_previews(mut self) -> Self {
        self.leaf_drops = &["preview.jpg", "preview-micro.jpg", "preview-web.jpg"];
        self
    }

    /// Accept the leaf structural workspace's package-metadata policies
    /// (record 0298 §2.1): the save token advances by one, and created
    /// objects get deterministic UUIDs where the host draws random ones.
    /// Both are checked for their documented shape and then compared as
    /// equal; see [`MetadataPolicy`].
    #[must_use]
    pub fn leaf_structural_metadata(mut self) -> Self {
        self.metadata = MetadataPolicy {
            save_token: true,
            fresh_uuids: true,
        };
        self
    }

    /// Replace the known-field schema table.
    #[must_use]
    pub fn known_fields(mut self, known: impl FnOnce(KnownFields) -> KnownFields) -> Self {
        self.known = known(self.known);
        self
    }

    /// Run both routes and assert the declared tier. Returns the report
    /// (also printed to stderr) or an error listing every failure.
    pub fn run(self) -> TestResult<Report> {
        let input = self.fixture.bytes.as_slice();
        let header = format!(
            "parity case {} [{:?}, tier {:?}, evidence {:?}] on {}",
            self.name, self.format, self.tier, self.evidence, self.fixture.path
        );
        let mut failures = Vec::new();
        if let Evidence::Native(path) = self.evidence
            && !repository_root().join(path).is_file()
        {
            failures.push(format!("native evidence {path} does not exist"));
        }
        let host = self.host.ok_or("case has no host route")?;
        let leaf = self.leaf.ok_or("case has no leaf route")?;
        let host_output =
            host(input).map_err(|error| format!("{header}: host route failed: {error}"))?;
        let leaf_run =
            leaf(input).map_err(|error| format!("{header}: leaf route failed: {error}"))?;
        let leaf_output = leaf_run.output.as_slice();
        if self.read_only {
            for (route, bytes) in [("host", host_output.as_slice()), ("leaf", leaf_output)] {
                if bytes != input {
                    failures.push(format!(
                        "read-only case: the {route} save is not byte-identical to the input (first difference at {:?})",
                        first_difference(bytes, input)
                    ));
                }
            }
        } else if host_output.as_slice() == input {
            failures.push("the host route produced the input bytes (no-op); pick an operation that changes the package, or mark the case read_only".to_owned());
        }

        for (route, bytes) in [("host", host_output.as_slice()), ("leaf", leaf_output)] {
            if let Err(error) = self.format.reopen(bytes) {
                failures.push(format!(
                    "strict leaf reopen of the {route} output failed: {error}"
                ));
            }
        }
        for (label, read) in &self.reads {
            match (read(&host_output), read(leaf_output)) {
                (Ok(host_read), Ok(leaf_read)) if host_read == leaf_read => {},
                (Ok(host_read), Ok(leaf_read)) => failures.push(format!(
                    "read `{label}` differs:\n{}",
                    archive_diff::debug_diff(&host_read, &leaf_read)
                )),
                (host_read, leaf_read) => failures.push(format!(
                    "read `{label}` failed (host {:?}, leaf {:?})",
                    host_read.err().map(|e| e.to_string()),
                    leaf_read.err().map(|e| e.to_string())
                )),
            }
        }
        if let Some(restored) = &leaf_run.restored
            && restored.as_slice() != input
        {
            failures.push(format!(
                "leaf inverse does not restore the input ({} vs {} bytes, first difference at {:?})",
                restored.len(),
                input.len(),
                first_difference(restored, input)
            ));
        }

        let diff = archive_diff::compare(
            input,
            &host_output,
            leaf_output,
            &self.known,
            self.leaf_drops,
            self.metadata,
        )?;
        match self.tier {
            Tier::A => {
                if host_output != leaf_output {
                    failures.push(format!(
                        "tier A: outputs differ ({} vs {} bytes, first difference at {:?})\n{}",
                        host_output.len(),
                        leaf_output.len(),
                        first_difference(&host_output, leaf_output),
                        diff.render()
                    ));
                }
            },
            Tier::B => {
                if !diff.violations.is_empty() {
                    failures.push(format!("tier B violations:\n{}", diff.render()));
                }
            },
        }

        if failures.is_empty() {
            let mut text = format!("{header}: passed\n");
            text.push_str(&diff.render());
            eprint!("{text}");
            Ok(Report { header, diff })
        } else {
            let mut text = format!("{header}: FAILED\n");
            for failure in failures {
                let _ = writeln!(text, "- {failure}");
            }
            Err(text.into())
        }
    }
}

/// Write a leaf package's bytes through any `write_to(&mut Vec<u8>)`.
#[macro_export]
macro_rules! leaf_bytes {
    ($package:expr) => {{
        let mut bytes = Vec::new();
        $package
            .write_to(&mut bytes)
            .map_err(|error| format!("leaf write failed: {error:?}"))?;
        bytes
    }};
}
