//! Arm definitions, JSON emission and the repeat loop.

use std::{fs, path::Path, sync::Arc, time::Instant};

use crate::counting::{CountingSource, Request, Transport};

/// One fixture in the corpus manifest.
pub struct Fixture {
    pub id: &'static str,
    pub path: &'static str,
    pub format: &'static str,
    pub role: &'static str,
    pub bytes: Vec<u8>,
    pub sha256: String,
}

impl Fixture {
    pub fn load(repo: &Path, id: &'static str, rel: &'static str, format: &'static str, role: &'static str) -> Self {
        use sha2::{Digest, Sha256};
        let bytes = fs::read(repo.join(rel)).unwrap_or_else(|e| panic!("{rel}: {e}"));
        let sha256 = Sha256::digest(&bytes).iter().map(|b| format!("{b:02x}")).collect::<String>();
        Self { id, path: rel, format, role, bytes, sha256 }
    }
}

/// The three read-ahead policies under test.
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub enum Policy {
    Exact,
    ForwardStart(usize),
}

impl Policy {
    pub fn label(self) -> String {
        match self {
            Self::Exact => "exact".to_string(),
            Self::ForwardStart(n) => format!("forward_start({n})"),
        }
    }
    pub fn to_source_read_policy(self) -> litchi_opc::SourceReadPolicy {
        match self {
            Self::Exact => litchi_opc::SourceReadPolicy::exact(),
            Self::ForwardStart(n) => {
                litchi_opc::SourceReadPolicy::forward_start(n).expect("valid window")
            }
        }
    }
}

/// The transport arms.
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub enum TransportArm {
    ZeroDelay,
    Delayed0493,
    ZeroDelayCapped,
}

impl TransportArm {
    pub fn label(self) -> &'static str {
        match self {
            Self::ZeroDelay => "zero_delay",
            Self::Delayed0493 => "delayed_1ms_100MiBps_64KiB",
            Self::ZeroDelayCapped => "zero_delay_64KiB_cap",
        }
    }
    pub fn transport(self) -> Transport {
        match self {
            Self::ZeroDelay => Transport::control(),
            Self::Delayed0493 => Transport::delayed_0493(),
            Self::ZeroDelayCapped => Transport::capped_control(),
        }
    }
}

/// How the package is constructed: the native leaf constructor, or the
/// two-step `SourceBackedPackage`-then-adopt route.
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub enum Route {
    NativeLeaf,
    PackageThenAdopt,
}

impl Route {
    pub fn label(self) -> &'static str {
        match self {
            Self::NativeLeaf => "native_leaf",
            Self::PackageThenAdopt => "package_then_adopt",
        }
    }
}

/// One repeat's outcome.
pub struct Repeat {
    pub requests: Vec<Request>,
    pub elapsed_ns: u128,
    pub open_ns: u128,
    /// Requests issued by construction alone, before the read scenario runs.
    pub open_requests: u64,
    pub observation: String,
}

/// What one scenario reports back: the requests the open alone cost, the time
/// the open alone cost, and a fingerprint proving the read actually happened.
pub struct Outcome {
    pub open_requests: u64,
    pub open_ns: u128,
    pub observation: String,
}

/// Runs one scenario closure `repeats` times against a freshly built source.
pub fn run_repeats<F>(
    fixture: &Fixture,
    transport: TransportArm,
    repeats: usize,
    mut scenario: F,
) -> Vec<Repeat>
where
    F: FnMut(Arc<CountingSource>) -> Outcome,
{
    let mut out = Vec::with_capacity(repeats);
    for _ in 0..repeats {
        let source = Arc::new(CountingSource::new(fixture.bytes.clone(), transport.transport()));
        let handle = Arc::clone(&source);
        let started = Instant::now();
        let outcome = scenario(handle);
        let elapsed_ns = started.elapsed().as_nanos();
        out.push(Repeat {
            requests: source.take_log(),
            elapsed_ns,
            open_ns: outcome.open_ns,
            open_requests: outcome.open_requests,
            observation: outcome.observation,
        });
    }
    out
}

/// Serialises one arm into the capture document.
pub fn arm_json(
    scenario: &str,
    fixture: &Fixture,
    policy: Policy,
    route: Route,
    transport: TransportArm,
    repeats: &[Repeat],
) -> serde_json::Value {
    serde_json::json!({
        "scenario": scenario,
        "fixture": fixture.id,
        "policy": policy.label(),
        "route": route.label(),
        "transport": transport.label(),
        "repeats": repeats.iter().map(|r| serde_json::json!({
            "elapsed_ns": r.elapsed_ns,
            "open_ns": r.open_ns,
            "open_requests": r.open_requests,
            "observation": r.observation,
            "requests": r.requests.iter().map(|q| serde_json::json!({
                "offset": q.offset, "length": q.length, "returned": q.returned,
            })).collect::<Vec<_>>(),
        })).collect::<Vec<_>>(),
    })
}
