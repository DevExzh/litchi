//! The staged part-level engine behind the package-scoped edit families
//! (unified-ops records 0320 and 0360).
//!
//! [`Staged`] owns one immutable [`Package`] (a reference-count clone) and,
//! from the first change on, a staged copy of its OPC graph: starting an edit
//! copies nothing. A family stages whole-part replacements, additions and
//! removals, or runs typed graph operations with [`Staged::update`], and
//! commits: [`PartPatch`] captures the exact before/after state (content
//! type, payload and relationships) of each changed part and of the
//! package-level relationships, and an unchanged staging returns the source
//! package itself, so an exact no-op keeps every byte. The source package
//! never changes.
//!
//! The families built on it ([`crate::conditional_formatting::edit`],
//! [`crate::data_validation::edit`], [`crate::auto_filter::edit`],
//! [`crate::workbook::comments::edit`], [`crate::drawing::edit`],
//! [`crate::opc`] and the package-level families) expose the ADR 0034 shape
//! through [`family!`]; the families whose replay payload is their own record
//! patch use [`record_family!`].

use std::collections::BTreeSet;
use std::sync::Arc;

use litchi_opc::{BlobPart, OpcPackage, PackURI, Part, Relationships, TargetMode};

use crate::Package;
use crate::error::{Error, Result, invalid};

#[derive(Debug, Clone, PartialEq, Eq)]
struct Relationship {
    id: String,
    kind: String,
    target: String,
    external: bool,
}

fn relationships(graph: &Relationships) -> Vec<Relationship> {
    let mut states: Vec<_> = graph
        .iter()
        .map(|relationship| Relationship {
            id: relationship.r_id().to_owned(),
            kind: relationship.reltype().to_owned(),
            target: relationship.target_ref().to_owned(),
            external: relationship.is_external(),
        })
        .collect();
    states.sort_unstable_by(|left, right| left.id.cmp(&right.id));
    states
}

fn add_relationships(graph: &mut Relationships, states: &[Relationship]) -> Result<()> {
    for relationship in states {
        graph.try_add_relationship(
            relationship.kind.clone(),
            relationship.target.clone(),
            relationship.id.clone(),
            if relationship.external {
                TargetMode::External
            } else {
                TargetMode::Internal
            },
        )?;
    }
    Ok(())
}

/// Whether two relationship collections hold the same relationships, without
/// allocating.
fn relationships_equal(left: &Relationships, right: &Relationships) -> bool {
    left.len() == right.len()
        && left.iter().all(|relationship| {
            right.get(relationship.r_id()).is_some_and(|candidate| {
                candidate.reltype() == relationship.reltype()
                    && candidate.target_ref() == relationship.target_ref()
                    && candidate.is_external() == relationship.is_external()
            })
        })
}

/// One part's exact state: its content type, payload and relationships.
#[derive(Debug, Clone)]
struct State {
    content_type: String,
    blob: Arc<Vec<u8>>,
    relationships: Vec<Relationship>,
}

impl PartialEq for State {
    fn eq(&self, other: &Self) -> bool {
        self.content_type == other.content_type
            && (Arc::ptr_eq(&self.blob, &other.blob) || self.blob == other.blob)
            && self.relationships == other.relationships
    }
}

impl Eq for State {}

impl State {
    fn read(part: &dyn Part) -> Self {
        Self {
            content_type: part.content_type().to_owned(),
            blob: part.blob_arc(),
            relationships: relationships(part.rels()),
        }
    }

    fn of(package: &OpcPackage, name: &PackURI) -> Option<Self> {
        package.get_part(name).ok().map(Self::read)
    }

    /// Whether `part` holds exactly this state, without copying its payload.
    fn matches(&self, part: &dyn Part) -> bool {
        self.content_type == part.content_type()
            && {
                let blob = part.blob();
                std::ptr::eq(self.blob.as_slice(), blob) || self.blob.as_slice() == blob
            }
            && self.relationships == relationships(part.rels())
    }

    fn to_part(&self, name: PackURI) -> Result<BlobPart> {
        let mut part =
            BlobPart::new_shared(name, self.content_type.clone(), Arc::clone(&self.blob));
        add_relationships(part.rels_mut(), &self.relationships)?;
        Ok(part)
    }
}

fn same_part(before: &dyn Part, after: &dyn Part) -> bool {
    before.content_type() == after.content_type()
        && {
            let (left, right) = (before.blob(), after.blob());
            std::ptr::eq(left, right) || left == right
        }
        && relationships_equal(before.rels(), after.rels())
}

fn optional_part<'a>(package: &'a OpcPackage, name: &PackURI) -> Result<Option<&'a dyn Part>> {
    match package.get_part(name) {
        Ok(part) => Ok(Some(part)),
        Err(litchi_opc::OpcError::PartNotFound(_)) => Ok(None),
        Err(error) => Err(error.into()),
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct Delta {
    name: PackURI,
    before: Option<State>,
    after: Option<State>,
}

/// The exact, reversible part-level change of one staged edit: the changed
/// parts and, when they changed, the package-level relationships.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub(crate) struct PartPatch {
    deltas: Vec<Delta>,
    root: Option<(Vec<Relationship>, Vec<Relationship>)>,
}

impl PartPatch {
    pub(crate) fn is_noop(&self) -> bool {
        self.deltas.is_empty() && self.root.is_none()
    }

    /// The number of parts (and the package-level relationship set) the
    /// patch changes.
    pub(crate) fn parts(&self) -> usize {
        self.deltas.len() + usize::from(self.root.is_some())
    }

    pub(crate) fn inverse(&self) -> Self {
        Self {
            deltas: self
                .deltas
                .iter()
                .map(|delta| Delta {
                    name: delta.name.clone(),
                    before: delta.after.clone(),
                    after: delta.before.clone(),
                })
                .collect(),
            root: self
                .root
                .as_ref()
                .map(|(before, after)| (after.clone(), before.clone())),
        }
    }

    fn root(
        before: &OpcPackage,
        after: &OpcPackage,
    ) -> Option<(Vec<Relationship>, Vec<Relationship>)> {
        if relationships_equal(before.rels(), after.rels()) {
            None
        } else {
            Some((relationships(before.rels()), relationships(after.rels())))
        }
    }

    /// Capture the parts named in `names` that differ between `before` and
    /// `after`.
    fn capture(before: &OpcPackage, after: &OpcPackage, names: &BTreeSet<String>) -> Result<Self> {
        let mut deltas = Vec::new();
        for name in names {
            let name = PackURI::new(name.as_str()).map_err(|error| invalid(error.to_string()))?;
            let before = State::of(before, &name);
            let after = State::of(after, &name);
            if before != after {
                deltas.push(Delta {
                    name,
                    before,
                    after,
                });
            }
        }
        Ok(Self {
            deltas,
            root: Self::root(before, after),
        })
    }

    /// Capture every part that differs between `before` and `after`.
    ///
    /// The staged graph is a clone of the source, so an untouched part shares
    /// its payload allocation and compares by pointer; only changed parts are
    /// compared by value.
    pub(crate) fn capture_all(before: &OpcPackage, after: &OpcPackage) -> Result<Self> {
        // One commit changes at most as many parts as the package's source
        // limits admit.
        let maximum = before.read_limits().max_parts();
        let mut deltas = Vec::new();
        let mut push = |delta: Delta| -> Result<()> {
            if delta.before == delta.after {
                return Ok(());
            }
            if deltas.len() >= maximum {
                return Err(Error::limit(
                    crate::opc::LimitKind::Parts,
                    deltas.len() + 1,
                    maximum,
                ));
            }
            deltas.push(delta);
            Ok(())
        };
        for part in before.try_iter_parts() {
            let part = part?;
            let name = part.partname();
            match optional_part(after, name)? {
                Some(other) if same_part(part, other) => {},
                other => push(Delta {
                    name: name.clone(),
                    before: Some(State::read(part)),
                    after: other.map(State::read),
                })?,
            }
        }
        for part in after.try_iter_parts() {
            let part = part?;
            if optional_part(before, part.partname())?.is_none() {
                push(Delta {
                    name: part.partname().clone(),
                    before: None,
                    after: Some(State::read(part)),
                })?;
            }
        }
        deltas.sort_unstable_by(|left, right| left.name.as_str().cmp(right.name.as_str()));
        Ok(Self {
            deltas,
            root: Self::root(before, after),
        })
    }

    /// Replay onto `package`, refusing a package that is not the patch's
    /// before state.
    pub(crate) fn apply(&self, package: &Package, operation: &'static str) -> Result<Package> {
        if self.is_noop() {
            return Ok(package.clone());
        }
        package.workbook_ref().ensure_mutation_allowed(operation)?;
        let source = package.opc();
        if let Some((before, _)) = &self.root
            && relationships(source.rels()) != *before
        {
            return Err(Error::PatchConflict {
                part: "/_rels/.rels".to_owned(),
            });
        }
        for delta in &self.deltas {
            let current = optional_part(source, &delta.name)?;
            let matches = match (&delta.before, current) {
                (None, None) => true,
                (Some(state), Some(part)) => state.matches(part),
                _ => false,
            };
            if !matches {
                return Err(Error::PatchConflict {
                    part: delta.name.as_str().to_owned(),
                });
            }
        }
        crate::signature_policy::ensure_unsigned(source)?;
        let mut target = source.clone();
        for delta in &self.deltas {
            target.remove_part(&delta.name);
            if let Some(after) = &delta.after {
                target.add_part(Box::new(after.to_part(delta.name.clone())?));
            }
        }
        if let Some((before, after)) = &self.root {
            let graph = target.rels_mut();
            for relationship in before {
                graph.remove(&relationship.id);
            }
            add_relationships(graph, after)?;
        }
        package.derive(target)
    }
}

/// A staged part-level edit of one [`Package`].
///
/// Starting one clones the package (a reference-count increment); the OPC
/// graph is copied on the first change only.
#[derive(Debug, Clone)]
pub(crate) struct Staged {
    operation: &'static str,
    source: Package,
    staged: Option<OpcPackage>,
    touched: BTreeSet<String>,
    /// Whether a whole-graph operation ran, so the commit compares every
    /// part instead of the touched ones.
    all: bool,
}

impl Staged {
    /// Start staging; refuses packages whose mutations are not allowed.
    pub(crate) fn new(package: &Package, operation: &'static str) -> Result<Self> {
        package.workbook_ref().ensure_mutation_allowed(operation)?;
        Ok(Self {
            operation,
            source: package.clone(),
            staged: None,
            touched: BTreeSet::new(),
            all: false,
        })
    }

    /// Start staging without the mutation-policy check (the caller made it).
    #[allow(dead_code, reason = "used by the package-level families")]
    pub(crate) fn raw(package: &Package, operation: &'static str) -> Self {
        Self {
            operation,
            source: package.clone(),
            staged: None,
            touched: BTreeSet::new(),
            all: false,
        }
    }

    /// The source package.
    #[allow(dead_code, reason = "read by the package-level families")]
    pub(crate) const fn source(&self) -> &Package {
        &self.source
    }

    /// The staged OPC graph.
    pub(crate) fn opc(&self) -> &OpcPackage {
        self.staged.as_ref().unwrap_or_else(|| self.source.opc())
    }

    fn detach(&mut self) -> &mut OpcPackage {
        let source = &self.source;
        self.staged.get_or_insert_with(|| source.opc().clone())
    }

    /// The staged OPC graph for a change: refuses a signed package first.
    /// The caller records every part it changes with [`Self::touch`].
    pub(crate) fn opc_mut(&mut self) -> Result<&mut OpcPackage> {
        crate::signature_policy::ensure_unsigned(self.opc())?;
        Ok(self.detach())
    }

    pub(crate) fn touch(&mut self, name: &PackURI) {
        self.touched.insert(name.as_str().to_owned());
    }

    /// Run one whole-graph operation on a copy of the staged graph; the copy
    /// replaces the staged graph only when the operation succeeds, so an
    /// error leaves the staging as it was. The commit compares every part.
    /// The operation itself refuses a changing operation on a signed graph.
    pub(crate) fn update<T>(
        &mut self,
        operation: impl FnOnce(&mut OpcPackage) -> Result<T>,
    ) -> Result<T> {
        let mut candidate = self.opc().clone();
        let value = operation(&mut candidate)?;
        self.staged = Some(candidate);
        self.all = true;
        Ok(value)
    }

    /// Whether no change detached the staged graph yet.
    #[cfg(test)]
    pub(crate) const fn staged_is_shared(&self) -> bool {
        self.staged.is_none()
    }

    /// Discard every staged change.
    pub(crate) fn reset(&mut self) {
        self.staged = None;
        self.touched.clear();
        self.all = false;
    }

    /// The part name of the worksheet at catalog position `sheet`; refuses a
    /// missing sheet and a chart or dialog sheet.
    pub(crate) fn worksheet(&self, sheet: usize) -> Result<PackURI> {
        let worksheet = self
            .source
            .workbook_ref()
            .sheet(sheet)?
            .ok_or(Error::refused(
                litchi_core::edit::RefusalKind::NotFound,
                self.operation,
                "the worksheet does not exist",
            ))?;
        if worksheet.kind() != crate::workbook::WorksheetKind::Worksheet {
            return Err(Error::NotWorksheet {
                sheet: worksheet.name().to_owned(),
            });
        }
        Ok(worksheet.part_uri().clone())
    }

    /// The staged payload of `name`.
    pub(crate) fn blob(&self, name: &PackURI) -> Result<&[u8]> {
        Ok(self.opc().get_part(name)?.blob())
    }

    /// Replace the payload of an existing part; identical bytes are a no-op.
    pub(crate) fn replace_blob(&mut self, name: &PackURI, bytes: Vec<u8>) -> Result<bool> {
        if self.opc().get_part(name)?.blob() == bytes.as_slice() {
            return Ok(false);
        }
        self.opc_mut()?.get_part_mut(name)?.set_blob(bytes);
        self.touch(name);
        Ok(true)
    }

    /// The first free part name `{prefix}{N}{suffix}`.
    pub(crate) fn free_name(&self, prefix: &str, suffix: &str) -> Result<PackURI> {
        for index in 1_u32..=1_000_000 {
            let candidate = PackURI::new(format!("{prefix}{index}{suffix}"))
                .map_err(|error| invalid(error.to_string()))?;
            if self.opc().get_part(&candidate).is_err() {
                return Ok(candidate);
            }
        }
        Err(invalid("no free part name is left"))
    }

    /// The staged styles part, when the workbook has one.
    pub(crate) fn styles(&self) -> Option<PackURI> {
        self.source.workbook_ref().styles_part_uri().cloned()
    }

    /// Capture the patch and publish the staged package; an unchanged
    /// staging returns the source package. A changed staging of a signed
    /// package is refused with [`Error::Signed`] (ADR 0006).
    pub(crate) fn commit(self) -> Result<(Package, PartPatch)> {
        let Some(staged) = self.staged else {
            return Ok((self.source, PartPatch::default()));
        };
        let patch = if self.all {
            PartPatch::capture_all(self.source.opc(), &staged)?
        } else {
            PartPatch::capture(self.source.opc(), &staged, &self.touched)?
        };
        if patch.is_noop() {
            return Ok((self.source, patch));
        }
        crate::signature_policy::ensure_unsigned(&staged)?;
        let package = self.source.derive(staged)?;
        Ok((package, patch))
    }
}

/// The direct children of `worksheet`, in ECMA-376 `CT_Worksheet` order.
const WORKSHEET_CHILDREN: &[&[u8]] = &[
    b"sheetPr",
    b"dimension",
    b"sheetViews",
    b"sheetFormatPr",
    b"cols",
    b"sheetData",
    b"sheetCalcPr",
    b"sheetProtection",
    b"protectedRanges",
    b"scenarios",
    b"autoFilter",
    b"sortState",
    b"dataConsolidate",
    b"customSheetViews",
    b"mergeCells",
    b"phoneticPr",
    b"conditionalFormatting",
    b"dataValidations",
    b"hyperlinks",
    b"printOptions",
    b"pageMargins",
    b"pageSetup",
    b"headerFooter",
    b"rowBreaks",
    b"colBreaks",
    b"customProperties",
    b"cellWatches",
    b"ignoredErrors",
    b"smartTags",
    b"drawing",
    b"legacyDrawing",
    b"legacyDrawingHF",
    b"picture",
    b"oleObjects",
    b"controls",
    b"webPublishItems",
    b"tableParts",
    b"extLst",
];

const RELATIONSHIP_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_RELATIONSHIP_NS: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";
const STRICT_MAIN_NS: &[u8] = b"http://purl.oclc.org/ooxml/spreadsheetml/main";

/// One direct child of a worksheet root: its local name, its byte span and
/// its `r:id` (any prefix bound to a relationships namespace).
#[derive(Debug, Clone)]
pub(crate) struct WorksheetChild {
    pub(crate) name: Vec<u8>,
    pub(crate) span: std::ops::Range<usize>,
    pub(crate) relationship_id: Option<String>,
}

/// The layout of a worksheet part's direct children.
#[derive(Debug)]
pub(crate) struct WorksheetLayout {
    /// The root element's prefix with its colon (`x:`), or empty.
    pub(crate) prefix: String,
    /// Whether the root binds the Strict main namespace.
    pub(crate) strict: bool,
    pub(crate) children: Vec<WorksheetChild>,
    /// The offset of the root's end tag.
    pub(crate) root_close: usize,
}

impl WorksheetLayout {
    /// Scan a worksheet part. Markup-compatibility branches among the direct
    /// children are refused: their effective content cannot be spliced.
    pub(crate) fn scan(xml: &[u8]) -> Result<Self> {
        use quick_xml::events::Event;
        let mut reader = quick_xml::Reader::from_reader(xml);
        reader.config_mut().check_end_names = true;
        let mut depth = 0usize;
        let mut prefix = String::new();
        let mut strict = false;
        let mut children = Vec::new();
        let mut open: Option<(Vec<u8>, usize, Option<String>)> = None;
        loop {
            let start = position(&reader)?;
            let event = reader
                .read_event()
                .map_err(|error| invalid(format!("worksheet markup: {error}")))?;
            let end = position(&reader)?;
            let (element, empty) = match event {
                Event::Start(element) => (Some(element), false),
                Event::Empty(element) => (Some(element), true),
                other => {
                    match other {
                        Event::End(_) => {
                            depth = depth.saturating_sub(1);
                            if depth == 1
                                && let Some((name, begin, relationship_id)) = open.take()
                            {
                                children.push(WorksheetChild {
                                    name,
                                    span: begin..end,
                                    relationship_id,
                                });
                            }
                            if depth == 0 {
                                return Ok(Self {
                                    prefix,
                                    strict,
                                    children,
                                    root_close: start,
                                });
                            }
                        },
                        Event::Eof => return Err(invalid("worksheet part has no closed root")),
                        _ => {},
                    }
                    (None, false)
                },
            };
            if let Some(element) = element {
                {
                    if depth == 0 {
                        let name = element.name();
                        if let Some(at) = name.as_ref().iter().position(|byte| *byte == b':') {
                            prefix = String::from_utf8_lossy(&name.as_ref()[..=at]).into_owned();
                        }
                        strict = litchi_ooxml_common::xml::attributes::first_wins(&element)
                            .flatten()
                            .any(|attribute| {
                                attribute.key.as_ref().starts_with(b"xmlns")
                                    && attribute.value.as_ref() == STRICT_MAIN_NS
                            });
                    } else if depth == 1 {
                        let local = element.local_name().as_ref().to_vec();
                        if local == b"AlternateContent" {
                            return Err(invalid(
                                "worksheet markup-compatibility branches cannot be edited",
                            ));
                        }
                        let relationship_id =
                            litchi_ooxml_common::xml::attributes::first_wins(&element)
                                .flatten()
                                .find_map(|attribute| {
                                    (attribute.key.local_name().as_ref() == b"id"
                                        && attribute.key.prefix().is_some())
                                    .then(|| String::from_utf8_lossy(&attribute.value).into_owned())
                                });
                        if empty {
                            children.push(WorksheetChild {
                                name: local,
                                span: start..end,
                                relationship_id,
                            });
                        } else {
                            open = Some((local, start, relationship_id));
                        }
                    }
                    if !empty {
                        depth += 1;
                    }
                }
            }
        }
    }

    /// The first direct child named `name`.
    pub(crate) fn child(&self, name: &str) -> Option<&WorksheetChild> {
        self.children
            .iter()
            .find(|child| child.name == name.as_bytes())
    }

    /// The relationships namespace this worksheet's flavor uses.
    pub(crate) const fn relationship_namespace(&self) -> &'static str {
        if self.strict {
            STRICT_RELATIONSHIP_NS
        } else {
            RELATIONSHIP_NS
        }
    }

    /// Insert a direct relationship child (`<drawing r:id/>`,
    /// `<legacyDrawing r:id/>`) at its schema position.
    pub(crate) fn insert_relationship_child(
        &self,
        xml: &[u8],
        name: &str,
        relationship_id: &str,
    ) -> Result<Vec<u8>> {
        if self.child(name).is_some() {
            return Err(invalid(format!(
                "the worksheet already has a {name} element"
            )));
        }
        let rank = WORKSHEET_CHILDREN
            .iter()
            .position(|child| *child == name.as_bytes())
            .ok_or_else(|| invalid("unknown worksheet child"))?;
        let at = self
            .children
            .iter()
            .find(|child| {
                WORKSHEET_CHILDREN
                    .iter()
                    .position(|known| *known == child.name.as_slice())
                    .is_some_and(|position| position > rank)
            })
            .map_or(self.root_close, |child| child.span.start);
        let fragment = format!(
            "<{}{name} xmlns:r=\"{}\" r:id=\"{}\"/>",
            self.prefix,
            self.relationship_namespace(),
            litchi_core::xml::escape_xml(relationship_id)
        );
        let mut out = Vec::with_capacity(xml.len() + fragment.len());
        out.extend_from_slice(&xml[..at]);
        out.extend_from_slice(fragment.as_bytes());
        out.extend_from_slice(&xml[at..]);
        Ok(out)
    }

    /// Rewrite a direct child `name` written in its own `x:` prefix with a
    /// local declaration into the root's prefix, so every worksheet writer
    /// reads it as a direct child. Other children keep their bytes.
    pub(crate) fn in_root_prefix(&self, xml: &[u8], name: &str) -> Vec<u8> {
        let Some(child) = self.child(name) else {
            return xml.to_vec();
        };
        let fragment = &xml[child.span.clone()];
        let namespace = if self.strict {
            STRICT_MAIN_NS
        } else {
            b"http://schemas.openxmlformats.org/spreadsheetml/2006/main".as_slice()
        };
        let mut declaration = b" xmlns:x=\"".to_vec();
        declaration.extend_from_slice(namespace);
        declaration.push(b'"');
        if !fragment.starts_with(b"<x:") || self.prefix == "x:" {
            return xml.to_vec();
        }
        let prefix = self.prefix.as_bytes();
        let mut out = Vec::with_capacity(xml.len());
        out.extend_from_slice(&xml[..child.span.start]);
        let mut index = 0;
        while index < fragment.len() {
            let rest = &fragment[index..];
            if rest.starts_with(&declaration) {
                index += declaration.len();
            } else if rest.starts_with(b"</x:") {
                out.extend_from_slice(b"</");
                out.extend_from_slice(prefix);
                index += 4;
            } else if rest.starts_with(b"<x:") {
                out.push(b'<');
                out.extend_from_slice(prefix);
                index += 3;
            } else {
                out.push(fragment[index]);
                index += 1;
            }
        }
        out.extend_from_slice(&xml[child.span.end..]);
        out
    }

    /// Rewrite a direct child `name` written unprefixed, in a default
    /// namespace declared on itself, into the root's prefix (declared on the
    /// child itself instead), so a prefixed worksheet keeps every direct
    /// child in its own prefix (record 0320 gap 7, fixed in record 0322).
    /// Prefixed descendants and every other child keep their bytes.
    pub(crate) fn unprefixed_in_root_prefix(&self, xml: &[u8], name: &str) -> Vec<u8> {
        let Some(child) = self.child(name) else {
            return xml.to_vec();
        };
        let fragment = &xml[child.span.clone()];
        let open = format!("<{name}");
        if self.prefix.is_empty()
            || !fragment.starts_with(open.as_bytes())
            || !matches!(
                fragment.get(open.len()),
                Some(b' ' | b'>' | b'/' | b'\t' | b'\r' | b'\n')
            )
        {
            return xml.to_vec();
        }
        let namespace = if self.strict {
            STRICT_MAIN_NS
        } else {
            b"http://schemas.openxmlformats.org/spreadsheetml/2006/main".as_slice()
        };
        let mut declaration = b" xmlns=\"".to_vec();
        declaration.extend_from_slice(namespace);
        declaration.push(b'"');
        let prefix = self.prefix.as_bytes();
        // The owner declares the root's prefix itself, so a reader that
        // parses the owner on its own still resolves it.
        let mut local = b" xmlns:".to_vec();
        local.extend_from_slice(&prefix[..prefix.len() - 1]);
        local.extend_from_slice(b"=\"");
        local.extend_from_slice(namespace);
        local.push(b'"');
        let unprefixed = |rest: &[u8]| {
            rest.first()
                .is_some_and(|byte| byte.is_ascii_alphabetic() || *byte == b'_')
                && !rest
                    .iter()
                    .take_while(|byte| !matches!(byte, b' ' | b'>' | b'/' | b'\t' | b'\r' | b'\n'))
                    .any(|byte| *byte == b':')
        };
        let mut out = Vec::with_capacity(xml.len() + 64);
        out.extend_from_slice(&xml[..child.span.start]);
        let mut index = 0;
        while index < fragment.len() {
            let rest = &fragment[index..];
            if rest.starts_with(&declaration) {
                out.extend_from_slice(&local);
                index += declaration.len();
            } else if rest.starts_with(b"</") && unprefixed(&rest[2..]) {
                out.extend_from_slice(b"</");
                out.extend_from_slice(prefix);
                index += 2;
            } else if rest.starts_with(b"<") && unprefixed(&rest[1..]) {
                out.push(b'<');
                out.extend_from_slice(prefix);
                index += 1;
            } else {
                out.push(fragment[index]);
                index += 1;
            }
        }
        out.extend_from_slice(&xml[child.span.end..]);
        out
    }

    /// Remove the first direct child named `name`; `None` when it is absent.
    pub(crate) fn remove_child(&self, xml: &[u8], name: &str) -> Option<Vec<u8>> {
        let child = self.child(name)?;
        let mut out = xml.to_vec();
        out.drain(child.span.clone());
        Some(out)
    }
}

/// The text of a predefined or character entity reference (`amp`, `#x41`).
pub(crate) fn entity(name: &[u8]) -> Option<String> {
    let name = std::str::from_utf8(name).ok()?;
    Some(match name {
        "amp" => "&".to_owned(),
        "lt" => "<".to_owned(),
        "gt" => ">".to_owned(),
        "apos" => "'".to_owned(),
        "quot" => "\"".to_owned(),
        _ => {
            let code = name.strip_prefix('#')?;
            let value = match code.strip_prefix('x') {
                Some(hex) => u32::from_str_radix(hex, 16).ok()?,
                None => code.parse().ok()?,
            };
            char::from_u32(value)?.to_string()
        },
    })
}

fn position(reader: &quick_xml::Reader<&[u8]>) -> Result<usize> {
    usize::try_from(reader.buffer_position()).map_err(|_| invalid("XML offset overflow"))
}

/// The `Patch`, `Commit` and `Diagnostics` types of one family, and its
/// `Error` re-export (ADR 0034 §3).
macro_rules! family {
    ($what:literal) => {
        pub use crate::Error;

        #[doc = concat!("The exact, reversible part-level change of a ", $what, " edit.")]
        #[derive(Debug, Clone, Default, PartialEq, Eq)]
        pub struct Patch(pub(crate) crate::part_edit::PartPatch);

        impl Patch {
            /// Whether the patch changes nothing.
            #[must_use]
            pub fn is_noop(&self) -> bool {
                self.0.is_noop()
            }

            /// The reverse patch.
            #[must_use]
            pub fn inverse(&self) -> Self {
                Self(self.0.inverse())
            }
        }

        #[doc = concat!("What a committed ", $what, " edit changed.")]
        #[derive(Debug, Clone, Copy, PartialEq, Eq)]
        pub struct Diagnostics {
            parts: usize,
        }

        impl Diagnostics {
            /// Whether the edit changed the package.
            #[must_use]
            pub const fn changed(&self) -> bool {
                self.parts > 0
            }

            /// The number of package parts the edit added, changed or
            /// removed.
            #[must_use]
            pub const fn parts(&self) -> usize {
                self.parts
            }
        }

        #[doc = concat!("The result of a ", $what, " edit's `commit`.")]
        #[derive(Debug)]
        pub struct Commit {
            package: crate::Package,
            patch: Patch,
            diagnostics: Diagnostics,
        }

        impl Commit {
            pub(crate) fn new(
                (package, patch): (crate::Package, crate::part_edit::PartPatch),
            ) -> Self {
                let diagnostics = Diagnostics {
                    parts: patch.parts(),
                };
                Self {
                    package,
                    patch: Patch(patch),
                    diagnostics,
                }
            }

            /// The package after the edit.
            #[must_use]
            pub const fn package(&self) -> &crate::Package {
                &self.package
            }

            /// Consume the commit into the package after the edit.
            #[must_use]
            pub fn into_package(self) -> crate::Package {
                self.package
            }

            /// The exact, reversible patch.
            #[must_use]
            pub const fn patch(&self) -> &Patch {
                &self.patch
            }

            /// What the edit changed.
            #[must_use]
            pub const fn diagnostics(&self) -> &Diagnostics {
                &self.diagnostics
            }

            /// Consume the commit into the package and the patch.
            #[must_use]
            pub fn into_parts(self) -> (crate::Package, Patch) {
                (self.package, self.patch)
            }
        }
    };
}

pub(crate) use family;
