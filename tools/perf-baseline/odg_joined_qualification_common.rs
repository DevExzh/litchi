//! Shared, source-neutral protocol helpers for the paired ODG joined probes.
//!
//! This file is copied beside each source-specific example at build time. It
//! deliberately uses only the API which existed at the pinned baseline: the
//! endpoint adapters own the old `Package::apply_joined` and new
//! `Edit::join(...).commit()` lifecycles separately.

#![allow(dead_code)]
#![allow(unexpected_cfgs)]

use std::{error::Error, fs, io::Write, path::PathBuf, time::Instant};

use litchi_core::{CompositionLimits, DiagnosticFingerprint};
use litchi_odf_common::core::{PackageWriter, StreamingArchiveLimits};
use litchi_odg::{Package, edit};
use serde_json::{Value, json};

pub(crate) const ROUTES: &[&str] = &[
    "package_replace_one",
    "semantic_two_disjoint",
    "semantic_four_disjoint",
    "empty_join",
];
pub(crate) const SCOPES: &[&str] = &["join_apply", "join_apply_publish"];
pub(crate) const MIMETYPE: &str = "application/vnd.oasis.opendocument.graphics";
pub(crate) const DEFAULT_MAX_ITERATIONS: usize = 1_000_000;
pub(crate) const REQUIRED_WARMUPS: usize = 20;
pub(crate) const REQUIRED_SAMPLES: usize = 1;
pub(crate) const MAX_FIXTURE_BYTES: usize = 4 * 1024 * 1024;
pub(crate) const MAX_SUB_EDIT_WIRE_BYTES: usize = 64 * 1024;
pub(crate) const LIMITS: CompositionLimits = CompositionLimits::new(8, 8, 32, 16);

pub(crate) type Result<T> = std::result::Result<T, Box<dyn Error>>;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum Mode {
    Verify,
    Calibrate,
    Timed,
    Allocation,
}

impl Mode {
    fn parse(value: &str) -> Result<Self> {
        match value {
            "verify" => Ok(Self::Verify),
            "calibrate" => Ok(Self::Calibrate),
            "timed" => Ok(Self::Timed),
            "allocation" => Ok(Self::Allocation),
            other => Err(format!("unknown mode {other:?}").into()),
        }
    }

    pub(crate) const fn as_str(self) -> &'static str {
        match self {
            Self::Verify => "verify",
            Self::Calibrate => "calibrate",
            Self::Timed => "timed",
            Self::Allocation => "allocation",
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum Scope {
    JoinApply,
    JoinApplyPublish,
}

impl Scope {
    pub(crate) fn parse(value: &str) -> Result<Self> {
        match value {
            "join_apply" => Ok(Self::JoinApply),
            "join_apply_publish" => Ok(Self::JoinApplyPublish),
            other => Err(format!("unknown scope {other:?}").into()),
        }
    }

    pub(crate) const fn as_str(self) -> &'static str {
        match self {
            Self::JoinApply => "join_apply",
            Self::JoinApplyPublish => "join_apply_publish",
        }
    }
}

#[derive(Debug)]
pub(crate) struct Args {
    pub(crate) mode: Mode,
    pub(crate) fixture: PathBuf,
    pub(crate) route: Option<String>,
    pub(crate) scope: Option<Scope>,
    pub(crate) iterations: usize,
    pub(crate) samples: usize,
    pub(crate) warmups: usize,
    pub(crate) target_ns: u128,
    pub(crate) max_iterations: usize,
}

impl Args {
    pub(crate) fn parse() -> Result<Self> {
        let mut arguments = std::env::args_os();
        let program = arguments
            .next()
            .unwrap_or_else(|| "litchi-perf-odg-joined-qualification".into());
        let mut mode = None;
        let mut fixture = None;
        let mut route = None;
        let mut scope = None;
        let mut iterations = 1_usize;
        let mut samples = 1_usize;
        let mut warmups = 0_usize;
        let mut target_ns = 500_000_000_u128;
        let mut max_iterations = DEFAULT_MAX_ITERATIONS;

        while let Some(argument) = arguments.next() {
            let argument = argument
                .into_string()
                .map_err(|_| format!("{program:?}: arguments must be UTF-8"))?;
            let mut value = |name: &str| {
                arguments
                    .next()
                    .ok_or_else(|| format!("{program:?}: {name} requires a value"))
                    .and_then(|value| {
                        value
                            .into_string()
                            .map_err(|_| format!("{program:?}: {name} must be UTF-8"))
                    })
            };
            match argument.as_str() {
                "--mode" => mode = Some(Mode::parse(&value("--mode")?)?),
                "--fixture" => fixture = Some(PathBuf::from(value("--fixture")?)),
                "--route" => route = Some(value("--route")?),
                "--scope" => scope = Some(Scope::parse(&value("--scope")?)?),
                "--iterations" => {
                    iterations = parse_positive(&value("--iterations")?, "iterations")?
                },
                "--samples" => samples = parse_positive(&value("--samples")?, "samples")?,
                "--warmups" => warmups = parse_nonnegative(&value("--warmups")?, "warmups")?,
                "--target-ns" => {
                    target_ns = parse_positive_u128(&value("--target-ns")?, "target-ns")?
                },
                "--max-iterations" => {
                    max_iterations = parse_positive(&value("--max-iterations")?, "max-iterations")?
                },
                "--help" | "-h" => return Err(usage(&program.to_string_lossy()).into()),
                other => return Err(format!("{program:?}: unknown argument {other:?}").into()),
            }
        }

        let mode = mode.ok_or_else(|| format!("{program:?}: --mode is required"))?;
        let fixture = fixture.ok_or_else(|| format!("{program:?}: --fixture is required"))?;
        if max_iterations == 0 {
            return Err("max-iterations must be positive".into());
        }
        if iterations > max_iterations || samples > max_iterations || warmups > max_iterations {
            return Err("timing controls exceed max-iterations".into());
        }
        if matches!(mode, Mode::Verify) && (iterations != 1 || samples != 1 || warmups != 0) {
            return Err("verify does not accept timing controls".into());
        }
        if matches!(mode, Mode::Verify) && (route.is_some() || scope.is_some()) {
            return Err("verify does not accept route or scope".into());
        }
        if !matches!(mode, Mode::Verify) && route.is_none() {
            return Err("calibrate, timed, and allocation require --route".into());
        }
        if matches!(mode, Mode::Timed | Mode::Calibrate | Mode::Allocation) && scope.is_none() {
            return Err("calibrate, timed, and allocation require --scope".into());
        }
        if matches!(mode, Mode::Allocation) && iterations != 1 {
            return Err("allocation measures one operation per sample".into());
        }
        if matches!(mode, Mode::Calibrate | Mode::Timed | Mode::Allocation)
            && warmups != REQUIRED_WARMUPS
        {
            return Err(format!("timing modes require exactly {REQUIRED_WARMUPS} warmups").into());
        }
        if matches!(mode, Mode::Timed | Mode::Allocation) && samples != REQUIRED_SAMPLES {
            return Err(format!("timing modes require exactly {REQUIRED_SAMPLES} sample").into());
        }
        if let Some(route) = route.as_deref() {
            if !ROUTES.contains(&route) {
                return Err(format!("unknown route {route:?}").into());
            }
        }
        Ok(Self {
            mode,
            fixture,
            route,
            scope,
            iterations,
            samples,
            warmups,
            target_ns,
            max_iterations,
        })
    }
}

pub(crate) fn usage(program: &str) -> String {
    format!(
        "usage: {program} --mode verify|calibrate|timed|allocation --fixture PATH \
         [--route ROUTE --scope join_apply|join_apply_publish] \
         [--iterations N --samples N --warmups N --target-ns N --max-iterations N]"
    )
}

fn parse_positive(value: &str, name: &str) -> Result<usize> {
    let parsed = value
        .parse::<usize>()
        .map_err(|error| format!("{name} must be a positive integer: {error}"))?;
    if parsed == 0 {
        return Err(format!("{name} must be positive").into());
    }
    Ok(parsed)
}

fn parse_nonnegative(value: &str, name: &str) -> Result<usize> {
    value
        .parse::<usize>()
        .map_err(|error| format!("{name} must be a nonnegative integer: {error}").into())
}

fn parse_positive_u128(value: &str, name: &str) -> Result<u128> {
    let parsed = value
        .parse::<u128>()
        .map_err(|error| format!("{name} must be a positive integer: {error}"))?;
    if parsed == 0 {
        return Err(format!("{name} must be positive").into());
    }
    Ok(parsed)
}

pub(crate) fn digest(bytes: &[u8]) -> Value {
    json!({"bytes": bytes.len(), "sha256": DiagnosticFingerprint::of(bytes).as_hex()})
}

pub(crate) fn fixture_package(path: &PathBuf) -> Result<(Vec<u8>, Package)> {
    let content = fs::read(path)?;
    if content.is_empty() || content.len() > MAX_FIXTURE_BYTES {
        return Err(format!("fixture content length is invalid: {}", content.len()).into());
    }
    let package = package_from_content(&content)?;
    Ok((content, package))
}

pub(crate) fn package_from_content(content: &[u8]) -> Result<Package> {
    let mut writer = PackageWriter::new(StreamingArchiveLimits::default());
    writer.set_mimetype(MIMETYPE)?;
    writer.add_file("content.xml", content)?;
    let bytes = writer.finish_to_bytes()?;
    Ok(Package::from_bytes(&bytes)?)
}

#[derive(Clone)]
pub(crate) struct PreparedSubEdit {
    pub(crate) id: String,
    pub(crate) writes: Vec<String>,
    pub(crate) wire: Vec<u8>,
}

pub(crate) struct PreparedRoute {
    pub(crate) route: &'static str,
    pub(crate) sub_edits: Vec<PreparedSubEdit>,
}

fn encode_sub_edit(sub_edit: &edit::SubEdit) -> Result<PreparedSubEdit> {
    Ok(PreparedSubEdit {
        id: sub_edit.id().to_owned(),
        writes: sub_edit.effects().writes().map(str::to_owned).collect(),
        wire: sub_edit.payload().to_deterministic_json()?,
    })
}

fn prepare_one<F>(source: &Package, id: &str, mutate: F) -> Result<PreparedSubEdit>
where
    F: FnOnce(&mut edit::Edit) -> Result<()>,
{
    let mut edit = source.edit();
    mutate(&mut edit)?;
    let commit = edit.commit()?;
    Ok(encode_sub_edit(&commit.patch().prepare(id, LIMITS)?)?)
}

pub(crate) fn prepare_route(source: &Package, route: &str) -> Result<PreparedRoute> {
    let route = ROUTES
        .iter()
        .copied()
        .find(|candidate| *candidate == route)
        .ok_or_else(|| format!("unknown route {route:?}"))?;
    let sub_edits = match route {
        "package_replace_one" => vec![prepare_one(source, "package-replace", |edit| {
            edit.add_page(litchi_odg::page::Page::new("Replacement Page"))?;
            Ok(())
        })?],
        "semantic_two_disjoint" => vec![
            prepare_one(source, "text-page-one", |edit| {
                edit.set_shape_text(0, 0, "Joined page one")?;
                Ok(())
            })?,
            prepare_one(source, "name-page-two", |edit| {
                edit.set_shape_name(1, 1, "Joined page two")?;
                Ok(())
            })?,
        ],
        "semantic_four_disjoint" => vec![
            prepare_one(source, "text-page-one", |edit| {
                edit.set_shape_text(0, 0, "Joined page one")?;
                Ok(())
            })?,
            prepare_one(source, "name-page-one", |edit| {
                edit.set_shape_name(0, 1, "Joined page one name")?;
                Ok(())
            })?,
            prepare_one(source, "text-page-two", |edit| {
                edit.set_shape_text(1, 0, "Joined page two")?;
                Ok(())
            })?,
            prepare_one(source, "name-page-two", |edit| {
                edit.set_shape_name(1, 1, "Joined page two name")?;
                Ok(())
            })?,
        ],
        "empty_join" => Vec::new(),
        other => return Err(format!("unknown route {other:?}").into()),
    };
    Ok(PreparedRoute { route, sub_edits })
}

pub(crate) fn prepare_all(source: &Package) -> Result<Vec<PreparedRoute>> {
    ROUTES
        .iter()
        .copied()
        .map(|route| prepare_route(source, route))
        .collect()
}

pub(crate) fn rebuild_plan(source: &Package, route: &PreparedRoute) -> Result<edit::JoinedEdits> {
    if route.sub_edits.len() > LIMITS.max_sub_edits() {
        return Err("joined route exceeds max_sub_edits".into());
    }
    let mut joined = source.joined_edits(LIMITS);
    let lineage = joined.lineage().clone();
    for prepared in &route.sub_edits {
        if prepared.wire.len() > MAX_SUB_EDIT_WIRE_BYTES {
            return Err(format!(
                "joined sub-edit {} exceeds the bounded wire input",
                prepared.id
            )
            .into());
        }
        let payload = edit::DurablePatch::from_deterministic_json(&prepared.wire)?;
        let sub_edit = edit::SubEdit::new(
            lineage.clone(),
            LIMITS,
            prepared.id.clone(),
            Vec::<String>::new(),
            prepared.writes.clone(),
            payload,
        )?;
        joined
            .join(sub_edit)
            .map_err(|error| format!("joined plan reconstruction failed: {error:?}"))?;
    }
    Ok(joined)
}

pub(crate) fn publish(package: &Package) -> Result<(Vec<u8>, Package)> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    let readback = Package::from_bytes(&bytes)?;
    Ok((bytes, readback))
}

pub(crate) fn summary(package: &Package) -> Value {
    let pages = package
        .pages()
        .iter()
        .map(|page| {
            let shapes = page
                .shapes()
                .iter()
                .map(|shape| json!({"name": shape.name(), "text": shape.text()}))
                .collect::<Vec<_>>();
            json!({"name": page.name(), "shapes": shapes})
        })
        .collect::<Vec<_>>();
    json!({"pages": pages})
}

pub(crate) fn foreign_package(content: &[u8]) -> Result<Package> {
    let mut changed = content.to_vec();
    // Keep the replacement the same length as the quoted attribute value.
    // The previous `Other 1` replacement was one byte longer and overwrote
    // the following quote, making this refusal proof a malformed-XML panic
    // rather than a changed-recipient check.
    let marker = b"Page 1";
    let replacement = b"Page 2";
    let position = changed
        .windows(marker.len())
        .position(|window| window == marker)
        .ok_or("fixture marker was not found")?;
    changed[position..position + replacement.len()].copy_from_slice(replacement);
    package_from_content(&changed)
}

pub(crate) fn run_once<F>(
    source: &Package,
    route: &PreparedRoute,
    scope: Scope,
    mut apply: F,
) -> Result<(u128, Value)>
where
    F: FnMut(edit::JoinedEdits) -> Result<Package>,
{
    let started = Instant::now();
    let package = apply(rebuild_plan(source, route)?)?;
    // `join_apply` measures the apply operation only.  Materializing bytes
    // for its proof happens after the clock closes; publication deliberately
    // remains inside the `join_apply_publish` scope because write/readback is
    // the operation being compared there.
    let (elapsed_ns, output) = match scope {
        Scope::JoinApply => {
            let elapsed_ns = started.elapsed().as_nanos();
            (elapsed_ns, package.as_bytes().to_vec())
        },
        Scope::JoinApplyPublish => {
            let output = publish(&package)?.0;
            (started.elapsed().as_nanos(), output)
        },
    };
    if elapsed_ns == 0 {
        return Err("timed joined operation completed in zero nanoseconds".into());
    };
    Ok((
        elapsed_ns,
        json!({"output": digest(&output), "semantic": summary(&package)}),
    ))
}

#[cfg(odg_joined_allocator)]
pub(crate) fn allocation_once<F>(
    source: &Package,
    route: &PreparedRoute,
    scope: Scope,
    mut apply: F,
) -> Result<(Value, Value)>
where
    F: FnMut(edit::JoinedEdits) -> Result<Package>,
{
    // The allocator region begins before durable plan reconstruction.  This
    // is the operation boundary: preparation, join/application, and (for the
    // publication scope) write/readback are all accounted for.  Proof
    // serialization is intentionally after the region closes.
    let region = crate::allocation::begin();
    let package = apply(rebuild_plan(source, route)?)?;
    let (allocation, output) = match scope {
        Scope::JoinApply => (region.finish(), package.as_bytes().to_vec()),
        Scope::JoinApplyPublish => {
            let output = publish(&package)?.0;
            (region.finish(), output)
        },
    };
    Ok((
        allocation,
        json!({"output": digest(&output), "semantic": summary(&package)}),
    ))
}

pub(crate) fn run_batch<F>(
    source: &Package,
    route: &PreparedRoute,
    scope: Scope,
    iterations: usize,
    mut apply: F,
) -> Result<(u128, Value)>
where
    F: FnMut(edit::JoinedEdits) -> Result<Package>,
{
    let mut elapsed_ns = 0_u128;
    let mut proof = None;
    for _ in 0..iterations {
        let (elapsed, current) = run_once(source, route, scope, &mut apply)?;
        if let Some(previous) = &proof {
            if previous != &current {
                return Err("joined route output proof changed within one batch".into());
            }
        } else {
            proof = Some(current);
        }
        elapsed_ns = elapsed_ns.saturating_add(elapsed);
    }
    Ok((elapsed_ns, proof.ok_or("empty timed batch")?))
}

pub(crate) fn write_json(value: &Value) -> Result<()> {
    let mut output = std::io::BufWriter::new(std::io::stdout().lock());
    serde_json::to_writer(&mut output, value)?;
    output.write_all(b"\n")?;
    output.flush()?;
    Ok(())
}
