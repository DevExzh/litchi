//! The staged part-level engine behind the worksheet feature edit families
//! (unified-ops record 0320).
//!
//! [`Staged`] holds one immutable [`Package`] and a staged copy of its OPC
//! graph. A family stages whole-part replacements, additions and removals,
//! records every part it touched, and commits: [`PartPatch`] captures the
//! exact before/after state (content type, payload and relationships) of each
//! touched part, and an unchanged staging returns the source package itself,
//! so an exact no-op keeps every byte. The source package never changes.
//!
//! The families built on it ([`crate::conditional_formatting::edit`],
//! [`crate::data_validation::edit`], [`crate::auto_filter::edit`],
//! [`crate::page_breaks::edit`], [`crate::workbook::comments::edit`] and
//! [`crate::drawing::edit`]) expose the ADR 0034 transaction shape through
//! [`family!`].

use std::collections::BTreeSet;
use std::sync::Arc;

use litchi_opc::{BlobPart, OpcPackage, PackURI, Part, TargetMode};

use crate::Package;
use crate::error::{Error, Result, invalid};

#[derive(Debug, Clone, PartialEq, Eq)]
struct Relationship {
    id: String,
    kind: String,
    target: String,
    external: bool,
}

/// One part's exact state: its content type, payload and relationships.
#[derive(Debug, Clone, PartialEq, Eq)]
struct State {
    content_type: String,
    blob: Arc<Vec<u8>>,
    relationships: Vec<Relationship>,
}

impl State {
    fn read(part: &dyn Part) -> Self {
        let mut relationships: Vec<_> = part
            .rels()
            .iter()
            .map(|relationship| Relationship {
                id: relationship.r_id().to_owned(),
                kind: relationship.reltype().to_owned(),
                target: relationship.target_ref().to_owned(),
                external: relationship.is_external(),
            })
            .collect();
        relationships.sort_unstable_by(|left, right| left.id.cmp(&right.id));
        Self {
            content_type: part.content_type().to_owned(),
            blob: part.blob_arc(),
            relationships,
        }
    }

    fn of(package: &OpcPackage, name: &PackURI) -> Option<Self> {
        package.get_part(name).ok().map(Self::read)
    }

    fn to_part(&self, name: PackURI) -> Result<BlobPart> {
        let mut part =
            BlobPart::new_shared(name, self.content_type.clone(), Arc::clone(&self.blob));
        for relationship in &self.relationships {
            part.rels_mut().try_add_relationship(
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
        Ok(part)
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct Delta {
    name: PackURI,
    before: Option<State>,
    after: Option<State>,
}

/// The exact, reversible part-level change of one staged edit.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub(crate) struct PartPatch {
    deltas: Vec<Delta>,
}

impl PartPatch {
    pub(crate) fn is_noop(&self) -> bool {
        self.deltas.is_empty()
    }

    pub(crate) fn parts(&self) -> usize {
        self.deltas.len()
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
        }
    }

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
        Ok(Self { deltas })
    }

    /// Replay onto `package`, refusing a package that is not the patch's
    /// before state.
    pub(crate) fn apply(&self, package: &Package, operation: &'static str) -> Result<Package> {
        if self.is_noop() {
            return Ok(package.clone());
        }
        package.workbook_ref().ensure_mutation_allowed(operation)?;
        crate::signature_policy::ensure_unsigned(package.opc())?;
        let mut target = package.opc().clone();
        for delta in &self.deltas {
            if State::of(&target, &delta.name) != delta.before {
                return Err(Error::PatchConflict {
                    part: delta.name.as_str().to_owned(),
                });
            }
        }
        for delta in &self.deltas {
            target.remove_part(&delta.name);
            if let Some(after) = &delta.after {
                target.add_part(Box::new(after.to_part(delta.name.clone())?));
            }
        }
        package
            .workbook_ref()
            .adopt_published_package(target)
            .map(Package::from_workbook)
    }
}

/// A staged part-level edit of one [`Package`].
#[derive(Debug)]
pub(crate) struct Staged<'a> {
    package: &'a Package,
    staged: OpcPackage,
    touched: BTreeSet<String>,
}

impl<'a> Staged<'a> {
    /// Start staging; refuses packages whose mutations are not allowed.
    pub(crate) fn new(package: &'a Package, operation: &'static str) -> Result<Self> {
        package.workbook_ref().ensure_mutation_allowed(operation)?;
        Ok(Self {
            package,
            staged: package.opc().clone(),
            touched: BTreeSet::new(),
        })
    }

    /// The staged OPC graph.
    pub(crate) const fn opc(&self) -> &OpcPackage {
        &self.staged
    }

    /// The staged OPC graph for a change: refuses a signed package first.
    /// The caller records every part it changes with [`Self::touch`].
    pub(crate) fn opc_mut(&mut self) -> Result<&mut OpcPackage> {
        crate::signature_policy::ensure_unsigned(&self.staged)?;
        Ok(&mut self.staged)
    }

    pub(crate) fn touch(&mut self, name: &PackURI) {
        self.touched.insert(name.as_str().to_owned());
    }

    /// The part name of the worksheet at catalog position `sheet`; refuses a
    /// missing sheet and a chart or dialog sheet.
    pub(crate) fn worksheet(&self, sheet: usize) -> Result<PackURI> {
        let worksheet = self
            .package
            .workbook_ref()
            .sheet(sheet)?
            .ok_or_else(|| invalid(format!("worksheet {sheet} does not exist")))?;
        if worksheet.kind() != crate::workbook::WorksheetKind::Worksheet {
            return Err(Error::NotWorksheet {
                sheet: worksheet.name().to_owned(),
            });
        }
        Ok(worksheet.part_uri().clone())
    }

    /// The staged payload of `name`.
    pub(crate) fn blob(&self, name: &PackURI) -> Result<&[u8]> {
        Ok(self.staged.get_part(name)?.blob())
    }

    /// Replace the payload of an existing part; identical bytes are a no-op.
    pub(crate) fn replace_blob(&mut self, name: &PackURI, bytes: Vec<u8>) -> Result<bool> {
        if self.staged.get_part(name)?.blob() == bytes.as_slice() {
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
            if self.staged.get_part(&candidate).is_err() {
                return Ok(candidate);
            }
        }
        Err(invalid("no free part name is left"))
    }

    /// The staged styles part, when the workbook has one.
    pub(crate) fn styles(&self) -> Option<PackURI> {
        self.package.workbook_ref().styles_part_uri().cloned()
    }

    /// Capture the patch and publish the staged package; an unchanged
    /// staging returns the source package.
    pub(crate) fn commit(self) -> Result<(Package, PartPatch)> {
        let patch = PartPatch::capture(self.package.opc(), &self.staged, &self.touched)?;
        if patch.is_noop() {
            return Ok((self.package.clone(), patch));
        }
        let package = self
            .package
            .workbook_ref()
            .adopt_published_package(self.staged)
            .map(Package::from_workbook)?;
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
