//! Owned chart snapshots, source-checked edits, and reversible package patches.

use super::codec::{content_xml, locate_pages, page_index};
use super::model::{Chart, Limits, Location, Page, Part, Selector, Storage};
use super::package;
use crate::Presentation;
use litchi_core::{Error, Result};
use std::sync::Arc;
use xml_minifier::audit;

const MAX_PACKAGE_BYTES: usize = 128 * 1024 * 1024;
const MAX_XML_PARTS: usize = 65_536;

/// Immutable embedded-chart inventory tied to exact ODP package bytes.
#[derive(Clone, Debug)]
pub struct Snapshot {
    source: Arc<Vec<u8>>,
    limits: Limits,
    charts: Arc<[Chart]>,
}

impl Snapshot {
    /// Parse an owned ODP package under the default chart resource budget.
    ///
    /// # Errors
    ///
    /// Returns an error when the package is oversized, malformed, or is not an ODP package.
    pub fn from_bytes(source: Vec<u8>) -> Result<Self> {
        Self::from_bytes_with(source, Limits::default())
    }

    /// Parse an owned ODP package with an explicit chart resource budget.
    ///
    /// # Errors
    ///
    /// Returns an error when the package or chart inventory violates the supplied limits.
    pub fn from_bytes_with(source: Vec<u8>, limits: Limits) -> Result<Self> {
        Self::from_shared_bytes(Arc::new(source), limits)
    }

    pub(crate) fn from_shared_bytes(source: Arc<Vec<u8>>, limits: Limits) -> Result<Self> {
        if source.len() > MAX_PACKAGE_BYTES {
            return invalid("ODP chart snapshot exceeds the 128 MiB package limit");
        }
        let presentation = Presentation::from_shared_bytes(Arc::clone(&source))?;
        let charts = presentation.charts_with(limits)?.charts;
        Ok(Self {
            source,
            limits,
            charts: Arc::from(charts),
        })
    }

    /// Borrow the exact complete ODP package bytes.
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        self.source.as_slice()
    }

    /// Return the chart resource budget retained by this snapshot.
    #[must_use]
    pub const fn limits(&self) -> Limits {
        self.limits
    }

    /// Borrow embedded charts in presentation drawing order.
    #[must_use]
    pub fn charts(&self) -> &[Chart] {
        &self.charts
    }

    /// Select one chart by exact frame name or checked discovery position.
    ///
    /// # Errors
    ///
    /// Returns an error when an exact frame name is ambiguous.
    pub fn get<'a, S>(&self, selector: S) -> Result<Option<&Chart>>
    where
        S: Into<Selector<'a>>,
    {
        select(&self.charts, selector.into())
            .map(|index| index.map(|selected| &self.charts[selected]))
    }

    /// Start an isolated, source-checked embedded-chart edit.
    #[must_use]
    pub fn edit(&self) -> Edit {
        Edit {
            before: self.clone(),
            draft: self.charts.to_vec(),
            next_add_token: 0,
        }
    }

    /// Materialize this snapshot as the ordinary presentation facade.
    ///
    /// # Errors
    ///
    /// Returns an error when the retained package can no longer be parsed.
    pub fn to_presentation(&self) -> Result<Presentation> {
        Presentation::from_shared_bytes(Arc::clone(&self.source))
    }
}

/// Clone-staged embedded-chart edit over one immutable ODP package.
#[derive(Clone, Debug)]
pub struct Edit {
    before: Snapshot,
    draft: Vec<Chart>,
    next_add_token: usize,
}

impl Edit {
    /// Borrow the current candidate chart inventory.
    #[must_use]
    pub fn charts(&self) -> &[Chart] {
        &self.draft
    }

    /// Replace one chart part, including occurrences sharing its package part.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing or ambiguous selector or a resource-limit breach.
    #[allow(
        clippy::needless_pass_by_value,
        reason = "replacement takes ownership of an authored chart part and may install it at multiple shared occurrences"
    )]
    pub fn replace<'a, S>(&mut self, selector: S, part: Part) -> Result<()>
    where
        S: Into<Selector<'a>>,
    {
        validate_part(&part, self.before.limits)?;
        let index = select(&self.draft, selector.into())?
            .ok_or_else(|| invalid_error("ODP embedded chart selector did not match"))?;
        let selected = self.draft[index].clone();
        for chart in &mut self.draft {
            if chart.same_identity(&selected) || chart.shares_storage(&selected) {
                *chart = chart.with_part(part.clone());
            }
        }
        validate_total(&self.draft, self.before.limits)
    }

    /// Remove one selected chart occurrence from its drawing page.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing or ambiguous selector.
    pub fn remove<'a, S>(&mut self, selector: S) -> Result<Chart>
    where
        S: Into<Selector<'a>>,
    {
        let index = select(&self.draft, selector.into())?
            .ok_or_else(|| invalid_error("ODP embedded chart selector did not match"))?;
        Ok(self.draft.remove(index))
    }

    /// Add a named chart frame to a selected presentation page.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid name, page selector, or resource-limit breach.
    pub fn add<'a, P, N>(&mut self, page: P, name: N, storage: Storage, part: Part) -> Result<usize>
    where
        P: Into<Page<'a>>,
        N: Into<String>,
    {
        validate_part(&part, self.before.limits)?;
        let frame_name = name.into();
        validate_name(&frame_name)?;
        if self
            .draft
            .iter()
            .any(|chart| chart.name() == Some(frame_name.as_str()))
        {
            return invalid("ODP chart frame names must be unique for selector stability");
        }
        if self.draft.len() >= self.before.limits.max_charts() {
            return invalid("ODP staged chart count exceeds the chart limit");
        }

        let presentation = self.before.to_presentation()?;
        let content = content_xml(presentation.owned_package())?;
        let selected_page = page_index(&content, page.into())?;
        let page_name = locate_pages(&content)?
            .into_iter()
            .find(|candidate| candidate.index == selected_page)
            .and_then(|candidate| candidate.name);
        let index = self.draft.len();
        let token = self.next_add_token;
        self.next_add_token = token
            .checked_add(1)
            .ok_or_else(|| invalid_error("ODP chart insertion token overflow"))?;
        self.draft.push(Chart {
            frame: Some(litchi_odf_common::drawing::Frame {
                name: Some(frame_name),
                page_name,
                ..Default::default()
            }),
            storage,
            part,
            location: Location::Added {
                page_index: selected_page,
                token,
            },
        });
        validate_total(&self.draft, self.before.limits)?;
        Ok(index)
    }

    /// Restore the exact source chart inventory in this draft.
    pub fn rollback(&mut self) {
        self.draft = self.before.charts.to_vec();
        self.next_add_token = 0;
    }

    /// Rebuild, audit, reopen, and semantically verify the complete package.
    ///
    /// # Errors
    ///
    /// Returns an error when serialization, compactness auditing, package reopening, or typed
    /// readback fails. The source snapshot is never changed.
    pub fn commit(self) -> Result<Commit> {
        let changed = !charts_semantically_equal(&self.before.charts, &self.draft);
        let snapshot = if changed {
            let presentation = self.before.to_presentation()?;
            let bytes = package::apply(
                presentation.owned_package(),
                &self.before.charts,
                &self.draft,
            )?;
            let target = Arc::new(bytes);
            validate_compact_package(&target)?;
            let reopened = Snapshot::from_shared_bytes(target, self.before.limits)?;
            if !charts_semantically_equal(&reopened.charts, &self.draft) {
                return invalid("ODP chart commit failed typed readback");
            }
            reopened
        } else {
            self.before.clone()
        };
        let diagnostics = Diagnostics::between(&self.before, &snapshot, changed);
        Ok(Commit {
            changed,
            patch: Patch {
                before: self.before,
                after: snapshot.clone(),
            },
            snapshot,
            diagnostics,
        })
    }
}

/// Exact-byte-source-checked reversible ODP chart patch.
#[derive(Clone, Debug)]
pub struct Patch {
    before: Snapshot,
    after: Snapshot,
}

impl Patch {
    /// Apply this patch only to its exact complete-package source snapshot.
    ///
    /// # Errors
    ///
    /// Returns an error when `source` is not byte-for-byte identical to the accepted source.
    pub fn apply(&self, source: &Snapshot) -> Result<Commit> {
        if !self.is_applicable_to(source) {
            return invalid("stale ODP chart patch source");
        }
        let snapshot =
            Snapshot::from_shared_bytes(Arc::clone(&self.after.source), self.after.limits)?;
        let changed = !self.is_noop();
        Ok(Commit {
            changed,
            diagnostics: Diagnostics::between(&self.before, &snapshot, changed),
            snapshot,
            patch: self.clone(),
        })
    }

    /// Return the patch that restores the exact accepted source package.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }

    /// Return whether this patch preserves the exact package bytes.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        same_source(&self.before, &self.after)
    }

    /// Return whether this patch accepts the exact supplied source snapshot.
    #[must_use]
    pub fn is_applicable_to(&self, source: &Snapshot) -> bool {
        same_source(&self.before, source)
    }
}

/// Content-free diagnostics for one chart publication.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct Diagnostics {
    charts_before: usize,
    charts_after: usize,
    changed: bool,
}

impl Diagnostics {
    fn between(before: &Snapshot, after: &Snapshot, changed: bool) -> Self {
        Self {
            charts_before: before.charts.len(),
            charts_after: after.charts.len(),
            changed,
        }
    }

    /// Return the chart occurrence count in the accepted source package.
    #[must_use]
    pub const fn charts_before(self) -> usize {
        self.charts_before
    }

    /// Return the chart occurrence count in the published package.
    #[must_use]
    pub const fn charts_after(self) -> usize {
        self.charts_after
    }

    /// Return whether complete-package bytes changed.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }
}

/// Fully rehydrated chart publication with its reversible patch.
#[derive(Clone, Debug)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
    diagnostics: Diagnostics,
    changed: bool,
}

impl Commit {
    /// Return whether complete-package bytes changed.
    #[must_use]
    pub const fn changed(&self) -> bool {
        self.changed
    }

    /// Borrow the resulting immutable chart snapshot.
    #[must_use]
    pub const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Borrow the exact-source-checked reversible patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Return content-free publication diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> Diagnostics {
        self.diagnostics
    }

    /// Consume this publication into its immutable chart snapshot.
    #[must_use]
    pub fn into_snapshot(self) -> Snapshot {
        self.snapshot
    }
}

pub(super) fn validate_compact_package(bytes: &[u8]) -> Result<()> {
    let presentation = Presentation::from_bytes(bytes.to_vec())?;
    let package = presentation.owned_package();
    let mut part_count = 0usize;
    let mut aggregate_bytes = 0usize;
    for path in package.files()? {
        if !path
            .rsplit_once('.')
            .is_some_and(|(_, extension)| extension.eq_ignore_ascii_case("xml"))
        {
            continue;
        }
        let payload = package.get_file(&path)?;
        if payload.windows(3).any(|window| window == b"> <") {
            return Err(Error::Unsupported(format!(
                "ODP chart XML part '{path}' contains inter-element spacing"
            )));
        }
        part_count = part_count
            .checked_add(1)
            .ok_or_else(|| invalid_error("ODP chart XML part count overflow"))?;
        aggregate_bytes = aggregate_bytes
            .checked_add(payload.len())
            .ok_or_else(|| invalid_error("ODP chart aggregate XML size overflow"))?;
        if part_count > MAX_XML_PARTS || aggregate_bytes > MAX_PACKAGE_BYTES {
            return invalid("ODP chart XML package audit exceeds its aggregate limit");
        }
        let limits = audit::Limits::new(
            MAX_PACKAGE_BYTES,
            512,
            1_000_000,
            250_000,
            audit::Limits::TOKEN_BYTE_CEILING,
            MAX_PACKAGE_BYTES,
        )
        .map_err(|source| invalid_error(format!("invalid ODP chart XML audit limits: {source}")))?;
        let _report = audit::verify(&payload, limits).map_err(|source| match source {
            audit::Error::NotCompact(_) => Error::Unsupported(format!(
                "ODP chart XML part '{path}' is not compact: {source}"
            )),
            audit::Error::Limit { .. }
            | audit::Error::Encoding { .. }
            | audit::Error::Malformed { .. }
            | audit::Error::Doctype { .. }
            | audit::Error::Allocation
            | _ => Error::InvalidFormat(format!(
                "ODP chart XML part '{path}' failed audit: {source}"
            )),
        })?;
    }
    Ok(())
}

pub(super) fn charts_semantically_equal(left: &[Chart], right: &[Chart]) -> bool {
    left.len() == right.len()
        && left.iter().zip(right).all(|(left_chart, right_chart)| {
            left_chart.frame == right_chart.frame
                && left_chart.storage == right_chart.storage
                && left_chart.part.chart == right_chart.part.chart
        })
}

fn validate_part(part: &Part, limits: Limits) -> Result<()> {
    if part.xml().is_empty() || part.xml().len() > limits.max_part_bytes() {
        return invalid("ODP replacement chart exceeds the part-byte limit");
    }
    Ok(())
}

fn validate_total(charts: &[Chart], limits: Limits) -> Result<()> {
    let total = charts.iter().try_fold(0usize, |total, chart| {
        total
            .checked_add(chart.part().xml().len())
            .ok_or_else(|| invalid_error("ODP chart byte count overflow"))
    })?;
    if total > limits.max_total_bytes() {
        return invalid("ODP staged chart content exceeds the total-byte limit");
    }
    Ok(())
}

fn select(charts: &[Chart], selector: Selector<'_>) -> Result<Option<usize>> {
    match selector {
        Selector::Index(index) => Ok((index < charts.len()).then_some(index)),
        Selector::Name(name) => {
            let mut selected = None;
            for (index, chart) in charts.iter().enumerate() {
                if chart.name() == Some(name) {
                    if selected.is_some() {
                        return invalid("ODP embedded chart name is ambiguous");
                    }
                    selected = Some(index);
                }
            }
            Ok(selected)
        },
    }
}

fn same_source(left: &Snapshot, right: &Snapshot) -> bool {
    Arc::ptr_eq(&left.source, &right.source) || left.source == right.source
}

fn validate_name(name: &str) -> Result<()> {
    if name.is_empty() || name.len() > 64 * 1024 || name.chars().any(char::is_control) {
        return invalid("ODP chart frame name is empty or invalid");
    }
    Ok(())
}

fn invalid<T>(message: impl Into<String>) -> Result<T> {
    Err(invalid_error(message))
}

fn invalid_error(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}
