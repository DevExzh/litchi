//! Immutable XLSB worksheet-property writes and publication.
//!
//! The first XLSB unified write is deliberately narrow: only the canonical
//! `sheet.name` property is admitted. Target resolution and property
//! coercion happen before the native workbook editor is materialized. The
//! native `Package` transaction remains responsible for exact package
//! replay, binary validation, signature policy, and the durable patch.

use std::collections::HashSet;
use std::io::{self, Write};
use std::path::Path;

use caseless::Caseless as _;
use litchi_core::FileFormat;
use litchi_xlsb::cell_values::{TransferLimits, structure::Patch};
use unicode_normalization::UnicodeNormalization;

use super::node::Node;
use super::spec::{Kind, SHEET_WRITE_TABLE, SheetSetter};
use super::{Nav, Xlsb, XlsbEditor, XlsbOwner};
use crate::error::{Code, Error, Result};
use crate::props::{IntoProps, Props};
use crate::registry::Operation;
use crate::save::{Overwrite, SaveOptions};
use crate::target::TargetArg;
use crate::transaction::{Error as TransactionError, LimitKind, SemanticLimits};
use crate::value::Scalar;

/// A committed immutable XLSB write.
pub(crate) struct Commit {
    after: Xlsb,
    patch: Patch,
}

impl Commit {
    /// Consume the result for the format-neutral one-shot setter.
    pub(crate) fn into_file(self) -> Xlsb {
        self.after
    }

    /// Consume the result for explicit transaction history.
    pub(crate) fn into_parts(self) -> (Xlsb, Patch) {
        (self.after, self.patch)
    }

    fn from_native(
        base: &Xlsb,
        native: litchi_xlsb::cell_values::structure::Commit,
    ) -> Result<Self> {
        base.check_source()?;
        let result = (|| {
            let (package, patch) = native.into_parts();
            // A native edit may retain a semantic operation whose replay is
            // an exact image no-op. Preserve the original immutable handle in
            // that case so source-backed publication never canonicalizes a
            // byte-identical no-op through a new package owner.
            let after = if patch.is_noop() {
                base.clone()
            } else {
                base.adopt_package(package)?
            };
            Ok(Self { after, patch })
        })();
        base.check_source()?;
        result
    }
}

#[derive(Debug, Default)]
struct SheetAssignments {
    name: Option<String>,
}

const WRITE_ENTRY_OVERHEAD: usize = 64;

/// The XLSB one-shot setter entry point.
pub(crate) fn set_with_limits<'p, 'v>(
    file: &Xlsb,
    target: TargetArg<'p>,
    props: impl IntoProps<'v>,
    limits: &SemanticLimits,
) -> Result<Commit> {
    let props = Props::from_props(props)?;
    let positions = resolve_sheet_targets(file, &target, limits)?;
    let assignments = admit_sheet_properties(&props, limits)?;
    check_limits(positions.len(), &assignments, limits)?;

    if assignments.name.is_some() && positions.len() > 1 {
        return Err(Error::new(
            Code::InvalidArgument,
            "renaming a selector with multiple sheet matches is ambiguous",
        )
        .with_format(FileFormat::Xlsb)
        .with_property("name"));
    }

    if let Some(name) = assignments.name.as_deref() {
        validate_sheet_name(file, &positions, name)?;
    }

    // All target and property admission above is intentionally before this
    // editor construction. Signed and retention policy are still checked by
    // `Xlsb::check_source` at the outer fence and by `editor` here.
    let editor = file.editor()?;
    let package = package_from_editor(editor);
    let native_limits = TransferLimits::new(
        limits.max_snapshot_bytes(),
        limits.max_operations(),
        1,
        limits.max_snapshot_bytes(),
    );
    let mut edit = package
        .edit_with_limits(native_limits)
        .map_err(native_error)?;
    if let Some(name) = assignments.name {
        // A name assignment is restricted to one target above, so cloning the
        // owned value does not multiply authored payload retention.
        edit.rename_sheet(positions[0], name)
            .map_err(rename_error)?;
    }
    let native = edit.commit().map_err(native_error)?;
    let commit = Commit::from_native(file, native)?;
    check_snapshot_limits(&commit.patch, limits)?;
    Ok(commit)
}

fn package_from_editor(editor: XlsbEditor<'_>) -> litchi_xlsb::Package {
    match editor {
        XlsbEditor::Eager(workbook) => workbook.to_package(),
        XlsbEditor::Materialized(workbook) => workbook.into_package(),
    }
}

/// Resolve only workbook sheet nodes. A cell, range, table, or any other
/// element is rejected before an XLSB native editor is built.
fn resolve_sheet_targets(
    file: &Xlsb,
    target: &TargetArg<'_>,
    limits: &SemanticLimits,
) -> Result<Vec<usize>> {
    let mut positions = Vec::new();
    let mut seen = HashSet::new();

    if let Some(path) = target.path() {
        let node = crate::path::resolve::get(&Nav::new(file), path)?;
        append_sheet(&mut positions, &mut seen, node, limits)?;
    } else if let Some(query) = target.query() {
        let walk = crate::query::run(Nav::new(file), query.clone())?;
        if !walk.plan().diagnostics().is_empty() {
            return Err(Error::new(
                Code::InvalidSelector,
                "the mutation selector contains an unsupported predicate",
            )
            .with_format(FileFormat::Xlsb));
        }
        let mut matched = false;
        for node in walk {
            matched = true;
            append_sheet(&mut positions, &mut seen, node?, limits)?;
        }
        file.check_source()?;
        if !matched {
            return Err(Error::new(
                Code::ZeroMatches,
                "the mutation selector matched no element",
            )
            .with_format(FileFormat::Xlsb));
        }
    } else {
        return Err(
            Error::new(Code::InvalidArgument, "the mutation target is empty")
                .with_format(FileFormat::Xlsb),
        );
    }

    file.check_source()?;
    Ok(positions)
}

fn append_sheet(
    positions: &mut Vec<usize>,
    seen: &mut HashSet<usize>,
    node: Node<'_>,
    limits: &SemanticLimits,
) -> Result<()> {
    if node.element() != Kind::Sheet {
        return Err(Error::new(
            Code::UnsupportedElement,
            "the XLSB mutation target must be a sheet",
        )
        .with_format(FileFormat::Xlsb)
        .with_suggestion(format!("resolved target kind: {}", node.kind())));
    }
    let Some(position) = node.worksheet_position() else {
        return Err(Error::new(
            Code::UnsupportedType,
            "the selected XLSB tab is not a worksheet",
        )
        .with_format(FileFormat::Xlsb));
    };
    if !seen.insert(position) {
        return Ok(());
    }
    let observed = positions
        .len()
        .checked_add(1)
        .ok_or_else(|| expanded_limit(usize::MAX, limits.max_expanded_targets()))?;
    if observed > limits.max_expanded_targets() {
        return Err(expanded_limit(observed, limits.max_expanded_targets()));
    }
    positions
        .try_reserve(1)
        .map_err(|_| allocation("XLSB sheet write targets"))?;
    positions.push(position);
    Ok(())
}

fn admit_sheet_properties(props: &Props<'_>, limits: &SemanticLimits) -> Result<SheetAssignments> {
    let mut assignments = SheetAssignments::default();
    for (key, value) in props.iter() {
        let (_, entry) = SHEET_WRITE_TABLE
            .resolve(key.as_str())
            .map_err(|error| error.with_property(key.as_str().to_owned()))?;
        let entry = *entry;
        entry
            .prepare(Operation::Set, value, |setter, value| match setter {
                SheetSetter::Name => {
                    if assignments.name.is_some() {
                        return Err(duplicate_property());
                    }
                    let Some(Scalar::Text(value)) = value else {
                        return Err(Error::new(Code::InvalidValue, "sheet name must be text")
                            .with_format(FileFormat::Xlsb)
                            .with_property("name"));
                    };
                    let observed = value.len().saturating_add(WRITE_ENTRY_OVERHEAD);
                    if observed > limits.max_staged_bytes() {
                        return Err(staged_limit(observed, limits.max_staged_bytes()));
                    }
                    assignments.name = Some(value.to_owned());
                    Ok(())
                },
            })
            .map_err(|error| error.with_format(FileFormat::Xlsb))?;
    }
    Ok(assignments)
}

fn validate_sheet_name(file: &Xlsb, targets: &[usize], name: &str) -> Result<()> {
    let units = name.encode_utf16().count();
    if units == 0
        || units > 31
        || name.contains([':', '\\', '*', '?', '/', '[', ']'])
        || name
            .chars()
            .any(|character| !is_xml_10_character(character))
        || name.starts_with('\'')
        || name.ends_with('\'')
    {
        return Err(Error::new(Code::InvalidValue, "the sheet name is invalid")
            .with_format(FileFormat::Xlsb)
            .with_property("name"));
    }
    let catalog = file.catalog()?;
    let key = caseless_key(name);
    if (0..catalog.len())
        .filter_map(|index| catalog.get(index))
        .any(|tab| {
            let selected = tab
                .worksheet_position()
                .is_some_and(|position| targets.contains(&position.get()));
            !selected && caseless_key(tab.name()) == key
        })
    {
        return Err(
            Error::new(Code::DuplicateName, "a sheet name is already in use")
                .with_format(FileFormat::Xlsb)
                .with_property("name"),
        );
    }
    Ok(())
}

const fn is_xml_10_character(character: char) -> bool {
    matches!(character, '\u{9}' | '\u{a}' | '\u{d}')
        || (character >= '\u{20}' && character <= '\u{d7ff}')
        || (character >= '\u{e000}' && character <= '\u{fffd}')
        || (character >= '\u{10000}' && character <= '\u{10ffff}')
}

fn caseless_key(value: &str) -> String {
    value.nfd().default_case_fold().nfd().collect()
}

fn check_limits(
    targets: usize,
    assignments: &SheetAssignments,
    limits: &SemanticLimits,
) -> Result<()> {
    if limits.max_operations() < 1 {
        return Err(operation_limit(1, limits.max_operations()));
    }
    if targets > limits.max_expanded_targets() {
        return Err(expanded_limit(targets, limits.max_expanded_targets()));
    }
    let per_target = assignments
        .name
        .as_ref()
        .map_or(0, String::len)
        .saturating_add(WRITE_ENTRY_OVERHEAD);
    let staged = targets.saturating_mul(per_target);
    if staged > limits.max_staged_bytes() {
        return Err(staged_limit(staged, limits.max_staged_bytes()));
    }
    Ok(())
}

fn check_snapshot_limits(patch: &Patch, limits: &SemanticLimits) -> Result<()> {
    let observed = patch
        .before()
        .len()
        .checked_add(patch.after().len())
        .ok_or_else(|| snapshot_limit(usize::MAX, limits.max_snapshot_bytes()))?;
    if observed > limits.max_snapshot_bytes() {
        return Err(snapshot_limit(observed, limits.max_snapshot_bytes()));
    }
    Ok(())
}

fn operation_limit(observed: usize, maximum: usize) -> Error {
    transaction_limit(LimitKind::Operations, observed, maximum)
}

fn expanded_limit(observed: usize, maximum: usize) -> Error {
    transaction_limit(LimitKind::ExpandedTargets, observed, maximum)
}

fn staged_limit(observed: usize, maximum: usize) -> Error {
    transaction_limit(LimitKind::StagedBytes, observed, maximum)
}

fn snapshot_limit(observed: usize, maximum: usize) -> Error {
    transaction_limit(LimitKind::SnapshotBytes, observed, maximum)
}

fn transaction_limit(kind: LimitKind, observed: usize, maximum: usize) -> Error {
    Error::from(TransactionError::Limit {
        kind,
        observed,
        maximum,
    })
    .with_format(FileFormat::Xlsb)
}

fn duplicate_property() -> Error {
    Error::new(
        Code::InvalidArgument,
        "the same XLSB sheet property was supplied more than once",
    )
    .with_format(FileFormat::Xlsb)
    .with_property("name")
}

fn allocation(resource: &'static str) -> Error {
    Error::new(Code::InternalError, "could not allocate mutation targets")
        .with_format(FileFormat::Xlsb)
        .with_suggestion(resource)
}

fn native_error(error: litchi_xlsb::package::PackageError) -> Error {
    Error::from(error).with_format(FileFormat::Xlsb)
}

fn rename_error(error: litchi_xlsb::package::PackageError) -> Error {
    let duplicate = matches!(
        &error,
        litchi_xlsb::package::PackageError::InvalidFormat(message)
            if message.starts_with("duplicate sheet name")
    );
    if duplicate {
        return Error::new(Code::DuplicateName, "a sheet name is already in use")
            .with_format(FileFormat::Xlsb)
            .with_property("name");
    }
    native_error(error).with_property("name")
}

/// Stream a finalized XLSB snapshot. Source-backed owners stream their
/// retained artifact directly; eager owners serialize their validated package.
pub(crate) fn write_to<W: Write>(file: &Xlsb, mut sink: W) -> Result<()> {
    file.0.owner.finish(|| {
        if file.signature() == super::Signature::RemovalAuthorized {
            return file
                .unsigned_workbook()?
                .into_package()
                .write_to(&mut sink)
                .map_err(native_error);
        }
        match &file.0.owner {
            XlsbOwner::Source(workbook) => workbook
                .workbook()
                .source_backed_package()
                .source_artifact()
                .write_to_stream(&mut sink)
                .map_err(Error::from),
            XlsbOwner::Eager(workbook) => workbook
                .to_package()
                .write_to(&mut sink)
                .map_err(native_error),
        }
    })
}

/// Publish an XLSB snapshot using the shared guarded atomic writer. The
/// source freshness fence runs before replacement; no post-replacement fence
/// is run against a path-backed source whose file has just been replaced.
pub(crate) fn save_to(file: &Xlsb, path: &Path, options: &SaveOptions) -> Result<u64> {
    #[cfg(feature = "encryption")]
    if options.reencrypt.is_some() {
        return Err(Error::new(
            Code::UnsupportedCapability,
            "XLSB re-encryption is not supported",
        )
        .with_format(FileFormat::Xlsb));
    }
    file.0.owner.ensure_current()?;
    let context = match &file.0.owner {
        XlsbOwner::Source(workbook) => workbook
            .workbook()
            .source_backed_package()
            .execution_context(),
        XlsbOwner::Eager(_) => None,
    };
    let mut bytes = 0_u64;
    let write = |sink: &mut litchi_opc::atomic::AtomicSink<'_>| -> Result<()> {
        let mut counted = Counted { sink, bytes: 0 };
        write_to(file, &mut counted)?;
        bytes = counted.bytes;
        Ok(())
    };
    let fence = || file.0.owner.ensure_current();
    match options.overwrite {
        Overwrite::Replace => litchi_opc::atomic::replace_guarded_with_durability(
            path,
            options.durability,
            context.as_ref(),
            write,
            fence,
        )?,
        Overwrite::Refuse => litchi_opc::atomic::replace_guarded_with_durability_noclobber(
            path,
            options.durability,
            context.as_ref(),
            write,
            fence,
        )?,
    }
    Ok(bytes)
}

struct Counted<W> {
    sink: W,
    bytes: u64,
}

impl<W: Write> Write for Counted<W> {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        let count = self.sink.write(bytes)?;
        self.bytes = self
            .bytes
            .checked_add(count as u64)
            .ok_or_else(|| io::Error::other("output byte count exceeds u64"))?;
        Ok(count)
    }

    fn flush(&mut self) -> io::Result<()> {
        self.sink.flush()
    }
}
