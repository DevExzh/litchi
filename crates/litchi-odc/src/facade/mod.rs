//! Concise family entry points.

use litchi_core::{Error, Metadata, Result};
use litchi_odf_common::chart::{ChartClass, Element, Legend, PlotArea, read};
use std::{collections::BTreeSet, path::Path, sync::Arc};

pub use crate::authoring::Builder;
use crate::authoring::Definition;

enum StagedStyles {
    Replace(String),
    Remove,
}

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

    /// Project this opened canonical chart into the granular typed edit model.
    ///
    /// Lossless projection succeeds only when deterministic reserialization is
    /// byte-identical to `content.xml`; producer extensions or lexical choices
    /// outside that surface are refused.
    ///
    /// # Errors
    ///
    /// Returns an unsupported error when projection would discard content.
    pub fn definition(&self) -> Result<Definition> {
        crate::project::definition(self)
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

    /// Return whether package signature metadata makes ordinary edits unsafe.
    #[must_use]
    pub fn is_signed(&self) -> bool {
        self.0.package.is_signed()
    }

    /// Return whether the manifest reports encrypted package members.
    #[must_use]
    pub fn is_encrypted(&self) -> bool {
        self.0.package.is_encrypted()
    }

    /// Read one package-local resource by inventory index.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid selector or unreadable package member.
    pub fn resource_bytes(&self, index: usize) -> Result<Vec<u8>> {
        self.0.package.resource_bytes(index)
    }

    /// Starts a source-bound package chart transaction.
    ///
    /// The exact surface covers selected chart, plot, series, and axis
    /// attributes; canonical definitions additionally support structural edits.
    /// Untouched package members and unmodeled chart XML remain payload-exact.
    #[must_use]
    pub fn edit(&self) -> Edit<'_> {
        Edit {
            source: self,
            transaction: self.0.package.content_snapshot().edit(),
            replacement: None,
            typed_transaction: None,
            flat_content_staged: false,
            staged_styles: None,
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
    typed_transaction: Option<crate::DefinitionEdit>,
    flat_content_staged: bool,
    staged_styles: Option<StagedStyles>,
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
        if self.replacement.is_some() || self.typed_transaction.is_some() {
            return Err(Error::InvalidFormat(
                "an ODC typed chart edit is already staged".to_string(),
            ));
        }
        self.transaction.update_axis(index, update)?;
        self.flat_content_staged = true;
        Ok(())
    }

    /// Stage a namespace-resolved, exact-span chart attribute edit.
    ///
    /// This surface accepts only the documented target/attribute combinations
    /// and reparses the complete candidate before publication.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid selector, attribute pairing, value, or
    /// a change that cannot preserve source XML outside the checked tag span.
    pub fn update_exact(
        &mut self,
        target: crate::ExactTarget,
        attribute: crate::ExactAttribute,
        after: Option<String>,
    ) -> Result<()> {
        if self.replacement.is_some() || self.typed_transaction.is_some() {
            return Err(Error::InvalidFormat(
                "an ODC typed chart edit is already staged".to_string(),
            ));
        }
        self.transaction.update_exact(target, attribute, after)?;
        self.flat_content_staged = true;
        Ok(())
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
        self.typed_transaction = None;
        Ok(())
    }

    /// Return a granular typed transaction for this opened package chart.
    ///
    /// The projection is lazy and lossless-or-refuse. A typed edit publishes
    /// as one atomic chart replacement while styles and resources remain in
    /// the same package commit.
    ///
    /// # Errors
    ///
    /// Returns an error after a flat axis edit or when projection is lossy.
    pub fn definition_edit(&mut self) -> Result<&mut crate::DefinitionEdit> {
        if self.flat_content_staged {
            return Err(Error::InvalidFormat(
                "ODC granular definition edits cannot follow a flat axis edit".into(),
            ));
        }
        if self.typed_transaction.is_none() {
            let definition = self
                .replacement
                .clone()
                .map_or_else(|| self.source.definition(), Ok)?;
            let snapshot = crate::DefinitionSnapshot::new(definition, self.source.limits())?;
            self.typed_transaction = Some(snapshot.edit());
        }
        self.typed_transaction
            .as_mut()
            .ok_or_else(|| Error::InvalidFormat("ODC definition edit was not initialized".into()))
    }

    /// Insert an axis into the opened package chart.
    ///
    /// # Errors
    ///
    /// Returns an error when projection or the insertion selector fails.
    pub fn insert_axis(&mut self, index: usize, axis: crate::AxisSpec) -> Result<()> {
        self.definition_edit()?.insert_axis(index, axis)
    }

    /// Replace an axis in the opened package chart.
    ///
    /// # Errors
    ///
    /// Returns an error when projection or the axis selector fails.
    pub fn replace_axis(&mut self, index: usize, axis: crate::AxisSpec) -> Result<()> {
        self.definition_edit()?.update_axis(index, axis)
    }

    /// Remove an axis from the opened package chart.
    ///
    /// # Errors
    ///
    /// Returns an error when projection or the axis selector fails.
    pub fn remove_axis(&mut self, index: usize) -> Result<crate::AxisSpec> {
        self.definition_edit()?.remove_axis(index)
    }

    /// Insert a series into the opened package chart.
    ///
    /// # Errors
    ///
    /// Returns an error when projection or the insertion selector fails.
    pub fn insert_series(&mut self, index: usize, series: crate::SeriesSpec) -> Result<()> {
        self.definition_edit()?.insert_series(index, series)
    }

    /// Replace a series in the opened package chart.
    ///
    /// # Errors
    ///
    /// Returns an error when projection or the series selector fails.
    pub fn replace_series(&mut self, index: usize, series: crate::SeriesSpec) -> Result<()> {
        self.definition_edit()?.update_series(index, series)
    }

    /// Remove a series from the opened package chart.
    ///
    /// # Errors
    ///
    /// Returns an error when projection or the series selector fails.
    pub fn remove_series(&mut self, index: usize) -> Result<crate::SeriesSpec> {
        self.definition_edit()?.remove_series(index)
    }

    /// Insert a data point into one opened-package series.
    ///
    /// # Errors
    ///
    /// Returns an error when projection or either selector fails.
    pub fn insert_data_point(
        &mut self,
        series: usize,
        index: usize,
        point: crate::DataPointSpec,
    ) -> Result<()> {
        self.definition_edit()?
            .insert_data_point(series, index, point)
    }

    /// Replace a data point in one opened-package series.
    ///
    /// # Errors
    ///
    /// Returns an error when projection or either selector fails.
    pub fn replace_data_point(
        &mut self,
        series: usize,
        index: usize,
        point: crate::DataPointSpec,
    ) -> Result<()> {
        self.definition_edit()?
            .update_data_point(series, index, point)
    }

    /// Remove a data point from one opened-package series.
    ///
    /// # Errors
    ///
    /// Returns an error when projection or either selector fails.
    pub fn remove_data_point(
        &mut self,
        series: usize,
        index: usize,
    ) -> Result<crate::DataPointSpec> {
        self.definition_edit()?.remove_data_point(series, index)
    }

    /// Set or remove a style reference at a typed opened-package site.
    ///
    /// # Errors
    ///
    /// Returns an error when projection or the style target fails.
    pub fn set_style(
        &mut self,
        target: crate::StyleTarget,
        style_name: Option<String>,
    ) -> Result<()> {
        self.definition_edit()?.set_style(target, style_name)
    }

    /// Create, replace, or remove the opened chart's cached table.
    ///
    /// # Errors
    ///
    /// Returns an error when lossless typed projection fails.
    pub fn set_cached_table(&mut self, table: Option<crate::CachedTable>) -> Result<()> {
        self.definition_edit()?.set_cached_table(table);
        Ok(())
    }

    /// Add or replace the complete validated `styles.xml` package part.
    pub fn set_styles_xml(&mut self, styles_xml: impl Into<String>) {
        self.staged_styles = Some(StagedStyles::Replace(styles_xml.into()));
    }

    /// Remove `styles.xml` from the package.
    pub fn remove_styles_xml(&mut self) {
        self.staged_styles = Some(StagedStyles::Remove);
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
        let resource_path = path.into();
        let resource_media_type = media_type.into();
        validate_media_type(&resource_media_type)?;
        if self
            .source
            .resources()
            .iter()
            .any(|resource| resource.path() == resource_path)
            || self
                .resource_edits
                .iter()
                .any(|edit| edit.path == resource_path)
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
            path: resource_path,
            before_media_type: None,
            after_media_type: Some(resource_media_type),
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
        let resource_media_type = media_type.into();
        validate_media_type(&resource_media_type)?;
        self.stage_existing_resource(index, Some(resource_media_type), Some(bytes.into()))
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
        let mut definition_changes = Vec::new();
        if let Some(typed_transaction) = self.typed_transaction.take() {
            let definition_commit = typed_transaction.commit()?;
            definition_changes = definition_commit.patch().changes().to_vec();
            self.replacement = Some(definition_commit.into_snapshot().definition().clone());
        }
        let content_commit = self.transaction.commit()?;
        let replacement = self.replacement.take();
        let changes = if replacement.is_some() {
            Vec::new()
        } else {
            content_commit.patch().changes().to_vec()
        };
        let exact_changes = if replacement.is_some() {
            Vec::new()
        } else {
            content_commit.patch().exact_changes().to_vec()
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
        let style_change = self.staged_styles.as_ref().and_then(|staged| {
            let before = self.source.styles_xml();
            let after = match staged {
                StagedStyles::Replace(xml) => Some(xml.as_str()),
                StagedStyles::Remove => None,
            };
            (before != after).then(|| StylesChange {
                before_size: before.map(str::len),
                after_size: after.map(str::len),
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
            let styles = match self.staged_styles.as_ref() {
                None => crate::package::StylesReplacement::Unchanged,
                Some(StagedStyles::Replace(xml)) => crate::package::StylesReplacement::Replace(xml),
                Some(StagedStyles::Remove) => crate::package::StylesReplacement::Remove,
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
            let package = self.source.0.package.rebuild(
                &content,
                replacement.is_none().then_some(content_commit.patch()),
                styles,
                &replacements,
            )?;
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
        if let Some(staged) = self.staged_styles.as_ref()
            && snapshot.styles_xml()
                != match staged {
                    StagedStyles::Replace(xml) => Some(xml.as_str()),
                    StagedStyles::Remove => None,
                }
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
                exact_changes,
                replaces_chart,
                style_change,
                resource_changes,
                definition_changes,
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
    exact_changes: Vec<crate::ExactChange>,
    replaces_chart: bool,
    style_change: Option<StylesChange>,
    resource_changes: Vec<ResourceChange>,
    definition_changes: Vec<crate::DefinitionChange>,
}

impl Patch {
    const WIRE_MAGIC: &'static [u8] = b"LITCHI-ODC-PATCH\0\x01";

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

    /// Serialize this exact-source patch deterministically for durable storage.
    ///
    /// The wire form contains only validated source and target package bytes;
    /// semantic summaries are deterministically reconstructed on decode.
    #[must_use]
    pub fn to_bytes(&self) -> Vec<u8> {
        let source = self.source.as_bytes();
        let target = self.target.as_bytes();
        let mut output = Vec::with_capacity(
            Self::WIRE_MAGIC
                .len()
                .saturating_add(16)
                .saturating_add(source.len())
                .saturating_add(target.len()),
        );
        output.extend_from_slice(Self::WIRE_MAGIC);
        output.extend_from_slice(
            &u64::try_from(source.len())
                .unwrap_or(u64::MAX)
                .to_le_bytes(),
        );
        output.extend_from_slice(
            &u64::try_from(target.len())
                .unwrap_or(u64::MAX)
                .to_le_bytes(),
        );
        output.extend_from_slice(source);
        output.extend_from_slice(target);
        output
    }

    /// Decode and fully reopen a deterministic durable patch.
    ///
    /// # Errors
    ///
    /// Returns an error for invalid framing, limits, or either package.
    pub fn from_bytes(bytes: &[u8], limits: crate::Limits) -> Result<Self> {
        let header = Self::WIRE_MAGIC.len().saturating_add(16);
        if bytes.len() < header || !bytes.starts_with(Self::WIRE_MAGIC) {
            return Err(Error::InvalidFormat("invalid ODC patch wire header".into()));
        }
        let mut cursor = Self::WIRE_MAGIC.len();
        let source_len = read_wire_length(bytes, &mut cursor)?;
        let target_len = read_wire_length(bytes, &mut cursor)?;
        if source_len > limits.max_package_bytes() || target_len > limits.max_package_bytes() {
            return Err(Error::InvalidFormat(
                "ODC patch package exceeds caller-selected limits".into(),
            ));
        }
        let source_end = cursor
            .checked_add(source_len)
            .ok_or_else(|| Error::InvalidFormat("ODC patch source length overflow".into()))?;
        let target_end = source_end
            .checked_add(target_len)
            .ok_or_else(|| Error::InvalidFormat("ODC patch target length overflow".into()))?;
        if target_end != bytes.len() {
            return Err(Error::InvalidFormat(
                "ODC patch wire lengths do not match its payload".into(),
            ));
        }
        let source = Chart::from_bytes_with_limits(bytes[cursor..source_end].to_vec(), limits)?;
        let target = Chart::from_bytes_with_limits(bytes[source_end..target_end].to_vec(), limits)?;
        Ok(patch_between(source, target))
    }

    /// Returns the semantic axis changes in transaction order.
    #[must_use]
    pub fn changes(&self) -> &[crate::AxisChange] {
        &self.changes
    }

    /// Returns controlled exact-span chart, plot-area, and series changes.
    #[must_use]
    pub fn exact_changes(&self) -> &[crate::ExactChange] {
        &self.exact_changes
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

    /// Return granular typed definition changes committed with this package.
    #[must_use]
    pub fn definition_changes(&self) -> &[crate::DefinitionChange] {
        &self.definition_changes
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
        let mut exact_changes = self.exact_changes.clone();
        exact_changes.extend_from_slice(&next.exact_changes);
        let mut resource_changes = self.resource_changes.clone();
        resource_changes.extend_from_slice(&next.resource_changes);
        Ok(Self {
            source: self.source.clone(),
            target: next.target.clone(),
            changes,
            exact_changes,
            replaces_chart: self.replaces_chart || next.replaces_chart,
            style_change: compose_style_change(
                self.style_change.as_ref(),
                next.style_change.as_ref(),
            ),
            resource_changes,
            definition_changes: self
                .definition_changes
                .iter()
                .chain(&next.definition_changes)
                .cloned()
                .collect(),
        })
    }

    /// Three-way join two package patches that share an exact source.
    ///
    /// The source and both targets remain immutable. Canonical typed content,
    /// styles, and resource paths merge independently; divergent values are
    /// returned as deterministic conflicts.
    ///
    /// # Errors
    ///
    /// Returns an error only when constructing or validating a conflict-free
    /// merged package fails.
    pub fn join(&self, other: &Self) -> Result<PackageMerge> {
        if self.source.as_bytes() != other.source.as_bytes() {
            return Ok(PackageMerge::new(
                None,
                vec![crate::Conflict::new("package.source")],
            ));
        }
        let mut conflicts = Vec::new();
        let mut edit = self.source.edit();
        merge_package_content(self, other, &mut edit, &mut conflicts)?;
        merge_package_styles(self, other, &mut edit, &mut conflicts);
        merge_package_resources(self, other, &mut edit, &mut conflicts)?;
        if !conflicts.is_empty() {
            return Ok(PackageMerge::new(None, conflicts));
        }
        let commit = edit.commit()?;
        Ok(PackageMerge::new(Some(commit.patch().clone()), Vec::new()))
    }

    /// Transfer this package patch onto another canonical chart snapshot.
    ///
    /// Typed chart data and style references, `styles.xml`, and inert package
    /// resources are merged against this patch's source. Inputs remain
    /// immutable and conflicts use stable semantic paths.
    ///
    /// # Errors
    ///
    /// Returns an error when reading resources or publishing a conflict-free
    /// transferred package fails.
    pub fn transfer_to(&self, destination: &Chart) -> Result<PackageMerge> {
        let mut conflicts = Vec::new();
        let mut edit = destination.edit();
        transfer_package_content(self, destination, &mut edit, &mut conflicts)?;
        transfer_package_styles(self, destination, &mut edit, &mut conflicts);
        transfer_package_resources(self, destination, &mut edit, &mut conflicts)?;
        if !conflicts.is_empty() {
            return Ok(PackageMerge::new(None, conflicts));
        }
        let commit = edit.commit()?;
        Ok(PackageMerge::new(Some(commit.patch().clone()), Vec::new()))
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
            exact_changes: self
                .exact_changes
                .iter()
                .map(crate::ExactChange::new_inverse)
                .collect(),
            replaces_chart: self.replaces_chart,
            style_change: self.style_change.as_ref().map(StylesChange::inverse),
            resource_changes: self
                .resource_changes
                .iter()
                .rev()
                .map(ResourceChange::inverse)
                .collect(),
            definition_changes: self
                .definition_changes
                .iter()
                .rev()
                .map(crate::transaction::inverse_change)
                .collect(),
        }
    }
}

/// Result of a non-mutating package patch join.
pub struct PackageMerge {
    patch: Option<Patch>,
    conflicts: Vec<crate::Conflict>,
}

impl PackageMerge {
    fn new(patch: Option<Patch>, mut conflicts: Vec<crate::Conflict>) -> Self {
        conflicts.sort();
        conflicts.dedup();
        Self { patch, conflicts }
    }

    #[must_use]
    pub const fn patch(&self) -> Option<&Patch> {
        self.patch.as_ref()
    }

    #[must_use]
    pub fn into_patch(self) -> Option<Patch> {
        self.patch
    }

    #[must_use]
    pub fn conflicts(&self) -> &[crate::Conflict] {
        &self.conflicts
    }

    #[must_use]
    pub fn is_merged(&self) -> bool {
        self.patch.is_some()
    }
}

/// Commit-coupled bounded package undo/redo history.
pub struct History {
    current: Chart,
    undo: Vec<Patch>,
    redo: Vec<Patch>,
}

impl History {
    /// Start history at one immutable package snapshot.
    #[must_use]
    pub fn new(chart: Chart) -> Self {
        Self {
            current: chart,
            undo: Vec::new(),
            redo: Vec::new(),
        }
    }

    /// Return the current immutable package snapshot.
    #[must_use]
    pub fn current(&self) -> &Chart {
        &self.current
    }

    /// Record one contiguous published commit and clear the redo branch.
    ///
    /// # Errors
    ///
    /// Returns an error for a stale commit or exhausted retained history limit.
    pub fn record(&mut self, commit: &Commit) -> Result<()> {
        let patch = commit.patch();
        if !patch.is_applicable_to(&self.current) {
            return Err(Error::InvalidFormat(
                "ODC history commit is not contiguous".into(),
            ));
        }
        if !commit.changed() {
            return Ok(());
        }
        if self.undo.len() >= self.current.limits().max_history() {
            return Err(Error::InvalidFormat(
                "ODC package history exceeds the caller-selected limit".into(),
            ));
        }
        self.current = commit.chart().clone();
        self.undo.push(patch.clone());
        self.redo.clear();
        Ok(())
    }

    /// Restore the exact previous package when present.
    ///
    /// # Errors
    ///
    /// Returns an error if retained history is internally non-contiguous.
    pub fn undo(&mut self) -> Result<bool> {
        let Some(patch) = self.undo.pop() else {
            return Ok(false);
        };
        self.current = patch.inverse().apply(&self.current)?;
        self.redo.push(patch);
        Ok(true)
    }

    /// Reapply the next exact package when present.
    ///
    /// # Errors
    ///
    /// Returns an error if retained history is internally non-contiguous.
    pub fn redo(&mut self) -> Result<bool> {
        let Some(patch) = self.redo.pop() else {
            return Ok(false);
        };
        self.current = patch.apply(&self.current)?;
        self.undo.push(patch);
        Ok(true)
    }

    #[must_use]
    pub fn undo_len(&self) -> usize {
        self.undo.len()
    }

    #[must_use]
    pub fn redo_len(&self) -> usize {
        self.redo.len()
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
    earlier: Option<&StylesChange>,
    later: Option<&StylesChange>,
) -> Option<StylesChange> {
    match (earlier, later) {
        (None, None) => None,
        (Some(change), None) | (None, Some(change)) => Some(change.clone()),
        (Some(earlier_change), Some(later_change)) => Some(StylesChange {
            before_size: earlier_change.before_size,
            after_size: later_change.after_size,
        }),
    }
}

fn merge_package_content(
    left: &Patch,
    right: &Patch,
    edit: &mut Edit<'_>,
    conflicts: &mut Vec<crate::Conflict>,
) -> Result<()> {
    if left.target.content_xml() == left.source.content_xml()
        && right.target.content_xml() == right.source.content_xml()
    {
        return Ok(());
    }
    let definitions = (
        left.source.definition(),
        left.target.definition(),
        right.target.definition(),
    );
    let (Ok(base_definition), Ok(left_definition), Ok(right_definition)) = definitions else {
        if flat_summary_matches(left) && flat_summary_matches(right) {
            merge_flat_content(left, right, edit, conflicts)?;
        } else {
            conflicts.push(crate::Conflict::new("chart.content"));
        }
        return Ok(());
    };
    let limits = left.source.limits();
    let base_snapshot = crate::DefinitionSnapshot::new(base_definition, limits)?;
    let left_snapshot = crate::DefinitionSnapshot::new(left_definition, limits)?;
    let right_snapshot = crate::DefinitionSnapshot::new(right_definition, limits)?;
    let left_patch = crate::DefinitionPatch {
        source: base_snapshot.clone(),
        target: left_snapshot,
        changes: left.definition_changes.clone(),
    };
    let right_patch = crate::DefinitionPatch {
        source: base_snapshot,
        target: right_snapshot,
        changes: right.definition_changes.clone(),
    };
    let merged = left_patch.join(&right_patch);
    conflicts.extend_from_slice(merged.conflicts());
    if let Some(patch) = merged.patch() {
        edit.replace_chart(patch.target.definition())?;
    }
    Ok(())
}

fn merge_package_styles(
    left: &Patch,
    right: &Patch,
    edit: &mut Edit<'_>,
    conflicts: &mut Vec<crate::Conflict>,
) {
    let base = left.source.styles_xml().map(str::to_owned);
    let left_value = left.target.styles_xml().map(str::to_owned);
    let right_value = right.target.styles_xml().map(str::to_owned);
    match merge_package_value(&base, &left_value, &right_value) {
        Some(Some(styles)) if Some(styles.as_str()) != left.source.styles_xml() => {
            edit.set_styles_xml(styles);
        },
        Some(None) if left.source.styles_xml().is_some() => edit.remove_styles_xml(),
        Some(_) => {},
        None => conflicts.push(crate::Conflict::new("package.styles")),
    }
}

fn transfer_package_content(
    patch: &Patch,
    destination: &Chart,
    edit: &mut Edit<'_>,
    conflicts: &mut Vec<crate::Conflict>,
) -> Result<()> {
    if patch.target.content_xml() == patch.source.content_xml() {
        return Ok(());
    }
    let definitions = (
        patch.source.definition(),
        patch.target.definition(),
        destination.definition(),
    );
    let (Ok(source_definition), Ok(target_definition), Ok(destination_definition)) = definitions
    else {
        if flat_summary_matches(patch) {
            transfer_flat_content(patch, destination, edit, conflicts)?;
        } else {
            conflicts.push(crate::Conflict::new("chart.content"));
        }
        return Ok(());
    };
    let source = crate::DefinitionSnapshot::new(source_definition, patch.source.limits())?;
    let target = crate::DefinitionSnapshot::new(target_definition, patch.target.limits())?;
    let destination_snapshot =
        crate::DefinitionSnapshot::new(destination_definition, destination.limits())?;
    let definition_patch = crate::DefinitionPatch {
        source,
        target,
        changes: patch.definition_changes.clone(),
    };
    let transferred = definition_patch.transfer_to(&destination_snapshot);
    conflicts.extend_from_slice(transferred.conflicts());
    if let Some(transferred_patch) = transferred.patch() {
        edit.replace_chart(transferred_patch.target.definition())?;
    }
    Ok(())
}

fn flat_summary_matches(patch: &Patch) -> bool {
    let replay = || -> Result<bool> {
        let source = patch.source.0.package.content_snapshot();
        let mut edit = source.edit();
        for change in &patch.changes {
            edit.update_axis(
                change.index(),
                crate::AxisUpdate {
                    name: (change.before() != change.after())
                        .then(|| change.after().map(str::to_owned)),
                    style_name: (change.before_style_name() != change.after_style_name())
                        .then(|| change.after_style_name().map(str::to_owned)),
                },
            )?;
        }
        for change in &patch.exact_changes {
            edit.update_exact(
                change.target(),
                change.attribute(),
                change.after().map(str::to_owned),
            )?;
        }
        Ok(edit.commit()?.snapshot().as_bytes() == patch.target.content_xml().as_bytes())
    };
    replay().unwrap_or(false)
}

fn merge_flat_content(
    left: &Patch,
    right: &Patch,
    edit: &mut Edit<'_>,
    conflicts: &mut Vec<crate::Conflict>,
) -> Result<()> {
    for left_change in &left.exact_changes {
        if let Some(right_change) = right.exact_changes.iter().find(|candidate| {
            candidate.target() == left_change.target()
                && candidate.attribute() == left_change.attribute()
        }) && left_change.after() != right_change.after()
        {
            conflicts.push(crate::Conflict::new(exact_change_path(left_change)));
        }
    }
    for left_change in &left.changes {
        if let Some(right_change) = right
            .changes
            .iter()
            .find(|candidate| candidate.index() == left_change.index())
        {
            if left_change.before() != left_change.after()
                && right_change.before() != right_change.after()
                && left_change.after() != right_change.after()
            {
                conflicts.push(crate::Conflict::new(format!(
                    "chart.plot.axes[{}].name",
                    left_change.index()
                )));
            }
            if left_change.before_style_name() != left_change.after_style_name()
                && right_change.before_style_name() != right_change.after_style_name()
                && left_change.after_style_name() != right_change.after_style_name()
            {
                conflicts.push(crate::Conflict::new(format!(
                    "chart.plot.axes[{}].style-name",
                    left_change.index()
                )));
            }
        }
    }
    if !conflicts.is_empty() {
        return Ok(());
    }
    for change in left.exact_changes.iter().chain(&right.exact_changes) {
        edit.update_exact(
            change.target(),
            change.attribute(),
            change.after().map(str::to_owned),
        )?;
    }
    for change in left.changes.iter().chain(&right.changes) {
        edit.update_axis(
            change.index(),
            crate::AxisUpdate {
                name: (change.before() != change.after())
                    .then(|| change.after().map(str::to_owned)),
                style_name: (change.before_style_name() != change.after_style_name())
                    .then(|| change.after_style_name().map(str::to_owned)),
            },
        )?;
    }
    Ok(())
}

fn transfer_flat_content(
    patch: &Patch,
    destination: &Chart,
    edit: &mut Edit<'_>,
    conflicts: &mut Vec<crate::Conflict>,
) -> Result<()> {
    let destination_flat = destination.0.package.content_snapshot();
    for change in &patch.exact_changes {
        let Ok(current) = destination_flat.exact_value(change.target(), change.attribute()) else {
            conflicts.push(crate::Conflict::new(exact_change_path(change)));
            continue;
        };
        if current != change.before() && current != change.after() {
            conflicts.push(crate::Conflict::new(exact_change_path(change)));
        } else if current == change.before() {
            edit.update_exact(
                change.target(),
                change.attribute(),
                change.after().map(str::to_owned),
            )?;
        }
    }
    for change in &patch.changes {
        let axis = destination.plot_area().and_then(|plot| {
            plot.axes().nth(change.index()).map(|axis| {
                (
                    axis.name().map(str::to_owned),
                    axis.style_name().map(str::to_owned),
                )
            })
        });
        let Some((axis_name, axis_style_name)) = axis else {
            conflicts.push(crate::Conflict::new(format!(
                "chart.plot.axes[{}]",
                change.index()
            )));
            continue;
        };
        let name_changed = change.before() != change.after();
        let style_changed = change.before_style_name() != change.after_style_name();
        let name_conflict = name_changed
            && axis_name.as_deref() != change.before()
            && axis_name.as_deref() != change.after();
        let style_conflict = style_changed
            && axis_style_name.as_deref() != change.before_style_name()
            && axis_style_name.as_deref() != change.after_style_name();
        if name_conflict {
            conflicts.push(crate::Conflict::new(format!(
                "chart.plot.axes[{}].name",
                change.index()
            )));
        }
        if style_conflict {
            conflicts.push(crate::Conflict::new(format!(
                "chart.plot.axes[{}].style-name",
                change.index()
            )));
        }
        if !name_conflict && !style_conflict {
            edit.update_axis(
                change.index(),
                crate::AxisUpdate {
                    name: (name_changed && axis_name.as_deref() == change.before())
                        .then(|| change.after().map(str::to_owned)),
                    style_name: (style_changed
                        && axis_style_name.as_deref() == change.before_style_name())
                    .then(|| change.after_style_name().map(str::to_owned)),
                },
            )?;
        }
    }
    Ok(())
}

fn exact_change_path(change: &crate::ExactChange) -> String {
    format!(
        "chart.exact[{:?}].{:?}",
        change.target(),
        change.attribute()
    )
}

fn transfer_package_styles(
    patch: &Patch,
    destination: &Chart,
    edit: &mut Edit<'_>,
    conflicts: &mut Vec<crate::Conflict>,
) {
    let base = patch.source.styles_xml().map(str::to_owned);
    let changed = patch.target.styles_xml().map(str::to_owned);
    let destination_value = destination.styles_xml().map(str::to_owned);
    match merge_package_value(&base, &changed, &destination_value) {
        Some(value) if value == destination_value => {},
        Some(Some(styles)) => edit.set_styles_xml(styles),
        Some(None) => edit.remove_styles_xml(),
        None => conflicts.push(crate::Conflict::new("package.styles")),
    }
}

fn merge_package_resources(
    left: &Patch,
    right: &Patch,
    edit: &mut Edit<'_>,
    conflicts: &mut Vec<crate::Conflict>,
) -> Result<()> {
    let paths = left
        .source
        .resources()
        .iter()
        .chain(left.target.resources())
        .chain(right.target.resources())
        .map(crate::Resource::path)
        .collect::<BTreeSet<_>>();
    for path in paths {
        let base = resource_state(&left.source, path);
        let left_value = resource_state(&left.target, path);
        let right_value = resource_state(&right.target, path);
        let Some(merged) = merge_package_value(&base, &left_value, &right_value) else {
            conflicts.push(crate::Conflict::new(format!("package.resource[{path}]")));
            continue;
        };
        if merged == base {
            continue;
        }
        stage_resource(edit, &left.source, path, merged)?;
    }
    Ok(())
}

fn transfer_package_resources(
    patch: &Patch,
    destination: &Chart,
    edit: &mut Edit<'_>,
    conflicts: &mut Vec<crate::Conflict>,
) -> Result<()> {
    let paths = patch
        .source
        .resources()
        .iter()
        .chain(patch.target.resources())
        .chain(destination.resources())
        .map(crate::Resource::path)
        .collect::<BTreeSet<_>>();
    for path in paths {
        let base = resource_state(&patch.source, path);
        let changed = resource_state(&patch.target, path);
        let destination_value = resource_state(destination, path);
        let Some(merged) = merge_package_value(&base, &changed, &destination_value) else {
            conflicts.push(crate::Conflict::new(format!("package.resource[{path}]")));
            continue;
        };
        if merged == destination_value {
            continue;
        }
        stage_resource(edit, destination, path, merged)?;
    }
    Ok(())
}

fn stage_resource(
    edit: &mut Edit<'_>,
    source: &Chart,
    path: &str,
    value: Option<(Option<String>, Vec<u8>)>,
) -> Result<()> {
    let source_index = source
        .resources()
        .iter()
        .position(|resource| resource.path() == path);
    match (source_index, value) {
        (Some(index), Some((optional_media_type, bytes))) => {
            let required_media_type = optional_media_type.ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "ODC merged resource '{path}' has no manifest media type"
                ))
            })?;
            edit.update_resource(index, required_media_type, bytes)?;
        },
        (Some(index), None) => edit.remove_resource(index)?,
        (None, Some((optional_media_type, bytes))) => {
            let required_media_type = optional_media_type.ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "ODC merged resource '{path}' has no manifest media type"
                ))
            })?;
            edit.add_resource(path, required_media_type, bytes)?;
        },
        (None, None) => {},
    }
    Ok(())
}

fn merge_package_value<T: Clone + PartialEq>(base: &T, left: &T, right: &T) -> Option<T> {
    if left == right {
        Some(left.clone())
    } else if left == base {
        Some(right.clone())
    } else if right == base {
        Some(left.clone())
    } else {
        None
    }
}

fn read_wire_length(bytes: &[u8], cursor: &mut usize) -> Result<usize> {
    let end = cursor
        .checked_add(8)
        .ok_or_else(|| Error::InvalidFormat("ODC patch length offset overflow".into()))?;
    let raw: [u8; 8] = bytes
        .get(*cursor..end)
        .ok_or_else(|| Error::InvalidFormat("ODC patch length is truncated".into()))?
        .try_into()
        .map_err(|error| Error::InvalidFormat(format!("invalid ODC patch length: {error}")))?;
    *cursor = end;
    usize::try_from(u64::from_le_bytes(raw))
        .map_err(|error| Error::InvalidFormat(format!("ODC patch length is too large: {error}")))
}

fn patch_between(source: Chart, target: Chart) -> Patch {
    let before_axes = axis_states(&source);
    let after_axes = axis_states(&target);
    let changes = before_axes
        .iter()
        .zip(&after_axes)
        .enumerate()
        .filter(|(_, (before, after))| before != after)
        .map(|(index, (before, after))| {
            crate::AxisChange::new(index, before.0.clone(), after.0.clone())
                .with_style(before.1.clone(), after.1.clone())
        })
        .collect();
    let style_change = (source.styles_xml() != target.styles_xml()).then(|| StylesChange {
        before_size: source.styles_xml().map(str::len),
        after_size: target.styles_xml().map(str::len),
    });
    let resource_changes = resource_changes_between(&source, &target);
    let definition_changes = match (source.definition(), target.definition()) {
        (Ok(before), Ok(after)) if before != after => {
            vec![crate::DefinitionChange::DefinitionUpdated]
        },
        _ => Vec::new(),
    };
    let replaces_chart = source.content_xml() != target.content_xml();
    let exact_changes = crate::flat::exact_changes_between(
        &source.0.package.content_snapshot(),
        &target.0.package.content_snapshot(),
    );
    Patch {
        source,
        target,
        changes,
        exact_changes,
        replaces_chart,
        style_change,
        resource_changes,
        definition_changes,
    }
}

fn axis_states(chart: &Chart) -> Vec<(Option<String>, Option<String>)> {
    chart
        .plot_area()
        .map(|plot| {
            plot.axes()
                .map(|axis| {
                    (
                        axis.name().map(str::to_owned),
                        axis.style_name().map(str::to_owned),
                    )
                })
                .collect()
        })
        .unwrap_or_default()
}

fn resource_changes_between(source: &Chart, target: &Chart) -> Vec<ResourceChange> {
    let paths = source
        .resources()
        .iter()
        .chain(target.resources())
        .map(crate::Resource::path)
        .collect::<BTreeSet<_>>();
    paths
        .into_iter()
        .filter_map(|path| {
            let before = resource_state(source, path);
            let after = resource_state(target, path);
            (before != after).then(|| ResourceChange {
                path: path.to_owned(),
                before_media_type: before.as_ref().and_then(|value| value.0.clone()),
                after_media_type: after.as_ref().and_then(|value| value.0.clone()),
                before_size: before.as_ref().map(|value| value.1.len()),
                after_size: after.as_ref().map(|value| value.1.len()),
            })
        })
        .collect()
}

fn resource_state(chart: &Chart, path: &str) -> Option<(Option<String>, Vec<u8>)> {
    chart
        .resources()
        .iter()
        .position(|resource| resource.path() == path)
        .and_then(|index| {
            chart.resource_bytes(index).ok().map(|bytes| {
                (
                    chart.resources()[index].media_type().map(str::to_owned),
                    bytes,
                )
            })
        })
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
