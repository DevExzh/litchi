//! Concise family entry points.

use litchi_core::{Error, Metadata, Result};
use litchi_odf_common::chart::{ChartClass, Element, Legend, PlotArea, read};
use std::{path::Path, sync::Arc};

pub use crate::authoring::Builder;
use crate::authoring::Definition;

/// Immutable document snapshot.
#[derive(Clone)]
pub struct Chart(Arc<State>);

struct State {
    package: crate::package::Snapshot,
    chart: Element,
}

impl Chart {
    /// Build a chart snapshot from a typed definition.
    ///
    /// # Errors
    ///
    /// Returns an error if the definition fails to build into a valid chart
    /// package.
    pub fn from_definition(definition: Definition) -> Result<Self> {
        Self::from_definition_with_limits(definition, crate::Limits::default())
    }

    /// Build a chart snapshot under caller-selected retained limits.
    ///
    /// # Errors
    ///
    /// Returns an error when validation or package publication fails.
    pub fn from_definition_with_limits(
        definition: Definition,
        limits: crate::Limits,
    ) -> Result<Self> {
        Self::from_bytes_with_limits(
            Builder::new()
                .with_limits(limits)
                .with_definition(definition)
                .build()?,
            limits,
        )
    }

    /// Open a chart package from a file path.
    ///
    /// # Errors
    ///
    /// Returns an error if the package cannot be read or the chart content
    /// cannot be parsed.
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        let package = crate::package::Snapshot::open(path)?;
        Self::from_package(package)
    }

    /// Open a chart package under caller-selected retained limits.
    ///
    /// # Errors
    ///
    /// Returns an error when the file or chart package violates the limits.
    pub fn open_with_limits(path: impl AsRef<Path>, limits: crate::Limits) -> Result<Self> {
        Self::from_package(crate::package::Snapshot::open_with_limits(path, limits)?)
    }

    /// Open a chart package from in-memory bytes.
    ///
    /// # Errors
    ///
    /// Returns an error if the package cannot be read or the chart content
    /// cannot be parsed.
    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        let package = crate::package::Snapshot::from_bytes(bytes)?;
        Self::from_package(package)
    }

    /// Open chart package bytes under caller-selected retained limits.
    ///
    /// # Errors
    ///
    /// Returns an error when the package or chart content is invalid.
    pub fn from_bytes_with_limits(bytes: Vec<u8>, limits: crate::Limits) -> Result<Self> {
        Self::from_package(crate::package::Snapshot::from_bytes_with_limits(
            bytes, limits,
        )?)
    }

    fn from_package(package: crate::package::Snapshot) -> Result<Self> {
        let chart = read(package.content_xml())?;
        Ok(Self(Arc::new(State { package, chart })))
    }

    #[must_use]
    pub fn content_xml(&self) -> &str {
        self.0.package.content_xml()
    }

    #[must_use]
    pub fn chart(&self) -> &Element {
        &self.0.chart
    }

    /// Return the typed root `chart:class` without normalizing its `QName`.
    ///
    /// # Errors
    ///
    /// Returns an invalid-format error if the retained chart has no valid
    /// `chart:class` value.
    pub fn class(&self) -> Result<ChartClass> {
        self.0.chart.chart_class()
    }

    #[must_use]
    pub fn plot_area(&self) -> Option<PlotArea<'_>> {
        self.0.chart.plot_area()
    }

    #[must_use]
    pub fn legend(&self) -> Option<Legend<'_>> {
        self.0.chart.legend()
    }

    #[must_use]
    pub fn styles_xml(&self) -> Option<&str> {
        self.0.package.styles_xml()
    }

    #[must_use]
    pub fn metadata(&self) -> Option<&Metadata> {
        self.0.package.metadata()
    }

    #[must_use]
    pub fn as_bytes(&self) -> &[u8] {
        self.0.package.as_bytes()
    }

    /// Return the limits retained for edits and patch application.
    #[must_use]
    pub fn limits(&self) -> crate::Limits {
        self.0.package.limits()
    }

    /// List the file names stored in the package.
    ///
    /// # Errors
    ///
    /// Returns an error if the package entries cannot be enumerated.
    pub fn files(&self) -> Result<Vec<String>> {
        self.0.package.files()
    }

    /// Return inert non-core package resources without fetching links.
    #[must_use]
    pub fn resources(&self) -> &[crate::Resource] {
        self.0.package.resources()
    }

    /// Read one package-local resource by inventory index.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid selector or unreadable package member.
    pub fn resource_bytes(&self, index: usize) -> Result<Vec<u8>> {
        self.0.package.resource_bytes(index)
    }

    /// Starts a source-bound package axis transaction.
    ///
    /// The transaction edits only `chart:name` attributes on existing axes.
    /// Untouched package members and unmodeled chart XML remain payload-exact.
    #[must_use]
    pub fn edit(&self) -> Edit<'_> {
        Edit {
            source: self,
            transaction: self.0.package.content_snapshot().edit(),
            replacement: None,
            styles_replacement: None,
            resource_edits: Vec::new(),
        }
    }

    #[must_use]
    pub fn into_bytes(self) -> Vec<u8> {
        match Arc::try_unwrap(self.0) {
            Ok(state) => state.package.into_bytes(),
            Err(state) => state.package.as_bytes().to_vec(),
        }
    }
}

/// A source-bound package chart transaction.
pub struct Edit<'a> {
    source: &'a Chart,
    transaction: crate::FlatChartEdit,
    replacement: Option<Definition>,
    styles_replacement: Option<Option<String>>,
    resource_edits: Vec<ResourceEdit>,
}

impl Edit<'_> {
    /// Stages an axis-name update by zero-based plot-area axis position.
    ///
    /// # Errors
    ///
    /// Returns an error when the selector is out of bounds or the update is
    /// outside the bounded, losslessly editable axis-name surface.
    pub fn update_axis(&mut self, index: usize, update: crate::AxisUpdate) -> Result<()> {
        if self.replacement.is_some() {
            return Err(Error::InvalidFormat(
                "an ODC whole-chart replacement is already staged".to_string(),
            ));
        }
        self.transaction.update_axis(index, update)
    }

    /// Stages explicit replacement of the complete chart content definition.
    ///
    /// This deliberately replaces the selected chart subtree, including any
    /// unknown XML within it. Other package member payloads remain preserved.
    /// A replacement supersedes axis updates staged earlier in this edit.
    ///
    /// # Errors
    ///
    /// Returns an error when the detached definition violates the bounded
    /// typed chart invariants.
    pub fn replace_chart(&mut self, definition: &Definition) -> Result<()> {
        crate::validation::validate_definition(definition, self.source.limits())?;
        self.replacement = Some(definition.clone());
        Ok(())
    }

    /// Add or replace the complete validated `styles.xml` package part.
    pub fn set_styles_xml(&mut self, styles_xml: impl Into<String>) {
        self.styles_replacement = Some(Some(styles_xml.into()));
    }

    /// Remove `styles.xml` from the package.
    pub fn remove_styles_xml(&mut self) {
        self.styles_replacement = Some(None);
    }

    /// Add an inert package-local resource.
    ///
    /// # Errors
    ///
    /// Returns an error for duplicate paths, invalid media types, or limits.
    pub fn add_resource(
        &mut self,
        path: impl Into<String>,
        media_type: impl Into<String>,
        bytes: impl Into<Vec<u8>>,
    ) -> Result<()> {
        let path = path.into();
        let media_type = media_type.into();
        validate_media_type(&media_type)?;
        if self
            .source
            .resources()
            .iter()
            .any(|resource| resource.path() == path)
            || self.resource_edits.iter().any(|edit| edit.path == path)
        {
            return Err(Error::InvalidFormat(
                "ODC resource path already exists in the transaction".into(),
            ));
        }
        let additions = self
            .resource_edits
            .iter()
            .filter(|edit| edit.before_bytes.is_none())
            .count();
        if self.source.resources().len().saturating_add(additions)
            >= self.source.limits().max_resources()
        {
            return Err(Error::InvalidFormat(
                "ODC resource count exceeds the caller-selected limit".into(),
            ));
        }
        self.resource_edits.push(ResourceEdit {
            path,
            before_media_type: None,
            after_media_type: Some(media_type),
            before_bytes: None,
            after_bytes: Some(bytes.into()),
        });
        Ok(())
    }

    /// Replace one inventoried package-local resource.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid selector or media type.
    pub fn update_resource(
        &mut self,
        index: usize,
        media_type: impl Into<String>,
        bytes: impl Into<Vec<u8>>,
    ) -> Result<()> {
        let media_type = media_type.into();
        validate_media_type(&media_type)?;
        self.stage_existing_resource(index, Some(media_type), Some(bytes.into()))
    }

    /// Remove one inventoried package-local resource.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid selector.
    pub fn remove_resource(&mut self, index: usize) -> Result<()> {
        self.stage_existing_resource(index, None, None)
    }

    fn stage_existing_resource(
        &mut self,
        index: usize,
        after_media_type: Option<String>,
        after_bytes: Option<Vec<u8>>,
    ) -> Result<()> {
        let resource =
            self.source.resources().get(index).ok_or_else(|| {
                Error::InvalidFormat("ODC resource selector is out of bounds".into())
            })?;
        let before_bytes = self.source.resource_bytes(index)?;
        if let Some(edit) = self
            .resource_edits
            .iter_mut()
            .find(|edit| edit.path == resource.path())
        {
            edit.after_media_type = after_media_type;
            edit.after_bytes = after_bytes;
        } else {
            self.resource_edits.push(ResourceEdit {
                path: resource.path().to_string(),
                before_media_type: resource.media_type().map(str::to_owned),
                after_media_type,
                before_bytes: Some(before_bytes.clone()),
                after_bytes,
            });
        }
        self.resource_edits.retain(|edit| {
            edit.path != resource.path()
                || edit.before_media_type != edit.after_media_type
                || edit.before_bytes != edit.after_bytes
        });
        Ok(())
    }

    /// Atomically validates, rebuilds, and publishes the package edit.
    ///
    /// # Errors
    ///
    /// Returns an error when the source cannot be losslessly rewritten, when
    /// the package is signed, or when full reopen and typed readback fail.
    #[allow(
        clippy::too_many_lines,
        reason = "publication keeps content, styles, resources, and readback in one atomic boundary"
    )]
    pub fn commit(mut self) -> Result<Commit> {
        let content_commit = self.transaction.commit()?;
        let replacement = self.replacement.take();
        let changes = if replacement.is_some() {
            Vec::new()
        } else {
            content_commit.patch().changes().to_vec()
        };
        let content = if let Some(definition) = replacement.as_ref() {
            crate::serialize_content_with_limits(definition, self.source.limits())?
        } else {
            std::str::from_utf8(content_commit.snapshot().as_bytes())
                .map_err(|_utf8_error| {
                    Error::InvalidFormat("edited ODC content is not UTF-8".to_string())
                })?
                .to_string()
        };
        let replaces_chart = replacement.is_some() && content != self.source.content_xml();
        let style_change = self.styles_replacement.as_ref().and_then(|after| {
            let before = self.source.styles_xml();
            (before != after.as_deref()).then(|| StylesChange {
                before_size: before.map(str::len),
                after_size: after.as_deref().map(str::len),
            })
        });
        self.resource_edits
            .sort_unstable_by(|left, right| left.path.cmp(&right.path));
        let resource_changes = self
            .resource_edits
            .iter()
            .map(ResourceEdit::change)
            .collect::<Vec<_>>();
        let no_content_change = content == self.source.content_xml();
        let snapshot = if no_content_change && style_change.is_none() && resource_changes.is_empty()
        {
            self.source.clone()
        } else {
            let styles = match self.styles_replacement.as_ref() {
                None => crate::package::StylesReplacement::Unchanged,
                Some(Some(xml)) => crate::package::StylesReplacement::Replace(xml),
                Some(None) => crate::package::StylesReplacement::Remove,
            };
            let replacements = self
                .resource_edits
                .iter()
                .map(|edit| crate::package::ResourceReplacement {
                    path: &edit.path,
                    media_type: edit.after_media_type.as_deref().unwrap_or_default(),
                    bytes: edit.after_bytes.as_deref(),
                })
                .collect::<Vec<_>>();
            let package = self
                .source
                .0
                .package
                .rebuild(&content, styles, &replacements)?;
            Chart::from_package(package)?
        };
        if let Some(definition) = replacement.as_ref() {
            if snapshot.content_xml() != content || snapshot.class()? != definition.class {
                return Err(Error::InvalidFormat(
                    "ODC chart replacement failed typed readback".to_string(),
                ));
            }
        } else {
            for change in &changes {
                let actual = snapshot
                    .plot_area()
                    .and_then(|plot_area| plot_area.axes().nth(change.index()))
                    .ok_or_else(|| {
                        Error::InvalidFormat(
                            "edited ODC axis disappeared during readback".to_string(),
                        )
                    })?;
                if actual.name() != change.after() {
                    return Err(Error::InvalidFormat(
                        "ODC package edit failed semantic readback".to_string(),
                    ));
                }
            }
        }
        if let Some(after) = self.styles_replacement.as_ref()
            && snapshot.styles_xml() != after.as_deref()
        {
            return Err(Error::InvalidFormat(
                "ODC styles edit failed package readback".into(),
            ));
        }
        for edit in &self.resource_edits {
            let actual = snapshot
                .resources()
                .iter()
                .position(|resource| resource.path() == edit.path)
                .map(|index| snapshot.resource_bytes(index))
                .transpose()?;
            if actual.as_deref() != edit.after_bytes.as_deref() {
                return Err(Error::InvalidFormat(
                    "ODC resource edit failed byte readback".into(),
                ));
            }
        }
        Ok(Commit {
            snapshot: snapshot.clone(),
            patch: Patch {
                source: self.source.clone(),
                target: snapshot,
                changes,
                replaces_chart,
                style_change,
                resource_changes,
            },
        })
    }
}

struct ResourceEdit {
    path: String,
    before_media_type: Option<String>,
    after_media_type: Option<String>,
    before_bytes: Option<Vec<u8>>,
    after_bytes: Option<Vec<u8>>,
}

impl ResourceEdit {
    fn change(&self) -> ResourceChange {
        ResourceChange {
            path: self.path.clone(),
            before_media_type: self.before_media_type.clone(),
            after_media_type: self.after_media_type.clone(),
            before_size: self.before_bytes.as_ref().map(Vec::len),
            after_size: self.after_bytes.as_ref().map(Vec::len),
        }
    }
}

/// A committed immutable chart package and its reversible patch.
pub struct Commit {
    snapshot: Chart,
    patch: Patch,
}

impl Commit {
    /// Returns whether package bytes changed.
    #[must_use]
    pub fn changed(&self) -> bool {
        self.patch.source.as_bytes() != self.patch.target.as_bytes()
    }

    /// Returns the committed chart snapshot.
    #[must_use]
    pub fn chart(&self) -> &Chart {
        &self.snapshot
    }

    /// Returns the exact-source reversible patch.
    #[must_use]
    pub fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Consumes the commit into its immutable chart snapshot.
    #[must_use]
    pub fn into_chart(self) -> Chart {
        self.snapshot
    }
}

/// A source-checked reversible package chart patch.
#[derive(Clone)]
pub struct Patch {
    source: Chart,
    target: Chart,
    changes: Vec<crate::AxisChange>,
    replaces_chart: bool,
    style_change: Option<StylesChange>,
    resource_changes: Vec<ResourceChange>,
}

impl Patch {
    /// Returns whether this patch applies to the supplied exact package bytes.
    #[must_use]
    pub fn is_applicable_to(&self, source: &Chart) -> bool {
        self.source.as_bytes() == source.as_bytes()
    }

    /// Applies this patch only to its exact immutable source.
    ///
    /// # Errors
    ///
    /// Returns an error when the supplied package differs byte-for-byte from
    /// the source against which this patch was committed.
    pub fn apply(&self, source: &Chart) -> Result<Chart> {
        if !self.is_applicable_to(source) {
            return Err(Error::InvalidFormat(
                "ODC package patch source does not match its expected snapshot".to_string(),
            ));
        }
        Ok(self.target.clone())
    }

    /// Returns the semantic axis changes in transaction order.
    #[must_use]
    pub fn changes(&self) -> &[crate::AxisChange] {
        &self.changes
    }

    /// Returns whether this patch replaces the complete chart definition.
    #[must_use]
    pub fn replaces_chart(&self) -> bool {
        self.replaces_chart
    }

    /// Return the optional whole styles-part change.
    #[must_use]
    pub const fn style_change(&self) -> Option<&StylesChange> {
        self.style_change.as_ref()
    }

    /// Return package-resource changes in stable path order.
    #[must_use]
    pub fn resource_changes(&self) -> &[ResourceChange] {
        &self.resource_changes
    }

    /// Compose two contiguous exact-package patches.
    ///
    /// # Errors
    ///
    /// Returns an error when the first target is not the second source.
    pub fn compose(&self, next: &Self) -> Result<Self> {
        if self.target.as_bytes() != next.source.as_bytes() {
            return Err(Error::InvalidFormat(
                "ODC package patches are not contiguous".into(),
            ));
        }
        let mut changes = self.changes.clone();
        changes.extend_from_slice(&next.changes);
        let mut resource_changes = self.resource_changes.clone();
        resource_changes.extend_from_slice(&next.resource_changes);
        Ok(Self {
            source: self.source.clone(),
            target: next.target.clone(),
            changes,
            replaces_chart: self.replaces_chart || next.replaces_chart,
            style_change: compose_style_change(
                self.style_change.as_ref(),
                next.style_change.as_ref(),
            ),
            resource_changes,
        })
    }

    /// Returns a patch that restores the exact source package.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            source: self.target.clone(),
            target: self.source.clone(),
            changes: self
                .changes
                .iter()
                .map(crate::AxisChange::new_inverse)
                .collect(),
            replaces_chart: self.replaces_chart,
            style_change: self.style_change.as_ref().map(StylesChange::inverse),
            resource_changes: self
                .resource_changes
                .iter()
                .rev()
                .map(ResourceChange::inverse)
                .collect(),
        }
    }
}

/// A whole `styles.xml` creation, replacement, or removal.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct StylesChange {
    before_size: Option<usize>,
    after_size: Option<usize>,
}

impl StylesChange {
    #[must_use]
    pub const fn before_size(&self) -> Option<usize> {
        self.before_size
    }

    #[must_use]
    pub const fn after_size(&self) -> Option<usize> {
        self.after_size
    }

    fn inverse(&self) -> Self {
        Self {
            before_size: self.after_size,
            after_size: self.before_size,
        }
    }
}

/// One package-local resource creation, replacement, or removal.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ResourceChange {
    path: String,
    before_media_type: Option<String>,
    after_media_type: Option<String>,
    before_size: Option<usize>,
    after_size: Option<usize>,
}

impl ResourceChange {
    #[must_use]
    pub fn path(&self) -> &str {
        &self.path
    }

    #[must_use]
    pub fn before_media_type(&self) -> Option<&str> {
        self.before_media_type.as_deref()
    }

    #[must_use]
    pub fn after_media_type(&self) -> Option<&str> {
        self.after_media_type.as_deref()
    }

    #[must_use]
    pub const fn before_size(&self) -> Option<usize> {
        self.before_size
    }

    #[must_use]
    pub const fn after_size(&self) -> Option<usize> {
        self.after_size
    }

    fn inverse(&self) -> Self {
        Self {
            path: self.path.clone(),
            before_media_type: self.after_media_type.clone(),
            after_media_type: self.before_media_type.clone(),
            before_size: self.after_size,
            after_size: self.before_size,
        }
    }
}

fn compose_style_change(
    first: Option<&StylesChange>,
    second: Option<&StylesChange>,
) -> Option<StylesChange> {
    match (first, second) {
        (None, None) => None,
        (Some(change), None) | (None, Some(change)) => Some(change.clone()),
        (Some(first), Some(second)) => Some(StylesChange {
            before_size: first.before_size,
            after_size: second.after_size,
        }),
    }
}

fn validate_media_type(media_type: &str) -> Result<()> {
    if media_type.is_empty()
        || media_type.len() > 1_024
        || !media_type.is_ascii()
        || !media_type.contains('/')
        || media_type
            .bytes()
            .any(|byte| byte.is_ascii_control() || byte.is_ascii_whitespace())
    {
        return Err(Error::InvalidFormat(
            "ODC resource media type is invalid".into(),
        ));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    #![allow(
        clippy::unwrap_used,
        reason = "tests are expected to panic on unexpected errors"
    )]

    use super::{Builder, Chart};

    #[test]
    fn builder_opens_as_validated_snapshot() {
        let bytes = Builder::new().build().unwrap();
        let document = Chart::from_bytes(bytes).unwrap();
        assert!(document.content_xml().contains("<office:chart"));
        assert!(document.plot_area().is_some());
        assert!(!document.as_bytes().is_empty());
    }
}
