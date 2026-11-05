//! Structural additions delegate authoring and preservation to the native owner.
use super::node::Item;
use super::spec::{SLIDE_WRITE_TABLE, SlideProperty};
use super::write::{
    Commit, Proof, native_snapshot_budget, opened_edit_options, source_edit_options,
};
use super::{Editable, Nav, Pptx, PptxOwner};
use crate::path::PathArg;
use crate::registry::Operation;
use crate::transaction::{Error as TxError, LimitKind, SemanticLimits};
use crate::value::Scalar;
use crate::{Code, Error, NewNode, Placement, Result};
use litchi_core::selector::{Path, Predicate, Step};
use litchi_core::{FileFormat, Position};
use std::sync::Arc;

impl Pptx {
    /// Whether this owner can add an empty slide through the source-native
    /// transaction.  Source owners retain the original package and can stage
    /// the presentation binding plus a fresh slide/layout relationship.
    pub(crate) fn supports_slide_add(&self) -> bool {
        self.0.retain_source
            && self.signature() != super::Signature::Signed
            && (matches!(self.0.owner, PptxOwner::Source(_))
                || self.signature() == super::Signature::RemovalAuthorized
                || matches!(self.0.owner, PptxOwner::Eager(_)))
    }

    /// Whether this owner can add a top-level text-box shape.  Keep this
    /// separate from slide insertion because the two native transactions
    /// have different admission rules (groups, masters and layouts remain
    /// outside this capability).
    pub(crate) fn supports_shape_add(&self) -> bool {
        self.0.retain_source
            && self.signature() != super::Signature::Signed
            && (matches!(self.0.owner, PptxOwner::Source(_) | PptxOwner::Eager(_))
                || self.signature() == super::Signature::RemovalAuthorized)
    }
    pub(crate) fn add(
        &self,
        parent: PathArg<'_>,
        node: NewNode<'_>,
    ) -> Result<(Commit, Path<'static>)> {
        self.add_with_limits(parent, node, &SemanticLimits::default())
    }
    pub(crate) fn add_with_limits(
        &self,
        parent: PathArg<'_>,
        node: NewNode<'_>,
        limits: &SemanticLimits,
    ) -> Result<(Commit, Path<'static>)> {
        self.check_source()?;
        let result = self
            .add_inner(parent, node, limits)
            .map_err(|e| e.with_format(FileFormat::Pptx));
        self.check_source()?;
        result
    }
    fn add_inner(
        &self,
        parent: PathArg<'_>,
        node: NewNode<'_>,
        limits: &SemanticLimits,
    ) -> Result<(Commit, Path<'static>)> {
        // The element kind first (record 0269, finding X2): a kind the deck
        // cannot add is `unsupported_element` whatever its parent; an
        // addable kind under the wrong parent is `unsupported_type`.
        if !ADDABLE
            .iter()
            .any(|kind| node.kind().eq_ignore_ascii_case(kind))
        {
            return Err(error(
                Code::UnsupportedElement,
                "this element kind cannot be added to a PPTX presentation yet",
            )
            .with_suggestion(
                "addable: slide, shape (textbox), picture, chart, table (and its rows and columns), smartart (and its nodes), hyperlink, media, comment, modernComment, placeholder, notes, paragraph and run; capabilities() lists the gaps",
            ));
        }
        if ["picture", "image", "img"]
            .iter()
            .any(|kind| node.kind().eq_ignore_ascii_case(kind))
        {
            return self.add_picture(parent, node, limits);
        }
        if node.kind().eq_ignore_ascii_case("chart") {
            return self.add_chart(parent, node, limits);
        }
        if node.kind().eq_ignore_ascii_case("smartart") {
            return self.add_smartart(parent, node, limits);
        }
        if node.kind().eq_ignore_ascii_case("node") {
            return self.add_smartart_node(parent, node, limits);
        }
        if ["video", "audio", "media"]
            .iter()
            .any(|kind| node.kind().eq_ignore_ascii_case(kind))
        {
            return self.add_media(parent, node, limits);
        }
        if node.kind().eq_ignore_ascii_case("hyperlink")
            || node.kind().eq_ignore_ascii_case("hlink")
        {
            return self.add_hyperlink(parent, node, limits);
        }
        if node.kind().eq_ignore_ascii_case("col") || node.kind().eq_ignore_ascii_case("column") {
            return self.add_column(parent, node, limits);
        }
        if node.kind().eq_ignore_ascii_case("comment") {
            return self.add_comment(parent, node, limits);
        }
        if MODERN_KINDS
            .iter()
            .any(|kind| node.kind().eq_ignore_ascii_case(kind))
        {
            return self.add_modern_comment(parent, node, limits);
        }
        // Placeholders and speaker notes (record 0346).
        if node.kind().eq_ignore_ascii_case("placeholder") || node.kind().eq_ignore_ascii_case("ph")
        {
            return self.add_placeholder(parent, node, limits);
        }
        if node.kind().eq_ignore_ascii_case("notes") {
            return self.add_notes(parent, node, limits);
        }
        if node.kind().eq_ignore_ascii_case("tr") || node.kind().eq_ignore_ascii_case("row") {
            return self.add_row(parent, node, limits);
        }
        if node.kind().eq_ignore_ascii_case("table") {
            return self.add_table(parent, node, limits);
        }
        if ["paragraph", "p", "para"]
            .iter()
            .any(|kind| node.kind().eq_ignore_ascii_case(kind))
        {
            return self.add_paragraph(parent, node, limits);
        }
        if node.kind().eq_ignore_ascii_case("run") || node.kind().eq_ignore_ascii_case("r") {
            return self.add_run(parent, node, limits);
        }
        if node.kind().eq_ignore_ascii_case("shape") || node.kind().eq_ignore_ascii_case("textbox")
        {
            // `add_text_box` constructs its bounded native edit on every
            // success path, which performs strict admission once (record
            // 0281).
            let source = self.editable_owner()?;
            return self.add_text_box(parent, node, source, limits);
        }
        let nav = Nav::new(self);
        if !matches!(
            crate::path::resolve::get(&nav, &parent)?.item(),
            Item::Presentation
        ) {
            return Err(error(
                Code::UnsupportedType,
                "slides require the presentation root as parent",
            ));
        }
        debug_assert!(node.kind().eq_ignore_ascii_case("slide"));
        let mut layout = None;
        let mut name = None;
        let mut hidden = None;
        // OfficeCLI's title and body shorthand (record 0346).
        let mut title = None;
        let mut body = None;
        let mut assigned = 0_u8;
        let mut staged_bytes = 0_usize;
        for (key, value) in node.props().iter() {
            let (index, entry) = SLIDE_WRITE_TABLE
                .resolve(key.as_str())
                .map_err(|e| e.with_property(key.as_str().to_owned()))?;
            let bit = 1_u8 << index;
            if assigned & bit != 0 {
                return Err(error(
                    Code::InvalidArgument,
                    "slide property was assigned more than once",
                )
                .with_property(entry.spec.name));
            }
            assigned |= bit;
            entry.prepare(Operation::Add, value, |property, value| {
                match property {
                    SlideProperty::Layout
                    | SlideProperty::Name
                    | SlideProperty::Title
                    | SlideProperty::Text => {
                        let value = match value {
                            Some(Scalar::Text(value)) => value,
                            None => "",
                            _ => unreachable!("slide text property coercion"),
                        };
                        staged_bytes = staged_bytes.saturating_add(value.len());
                        bound(
                            LimitKind::StagedBytes,
                            staged_bytes,
                            limits.max_staged_bytes(),
                        )?;
                        if !matches!(property, SlideProperty::Layout) {
                            super::write::validate_text(value)
                                .map_err(|error| error.with_property(entry.spec.name))?;
                        }
                        let mut owned = String::new();
                        owned
                            .try_reserve_exact(value.len())
                            .map_err(|source| Error::from(TxError::Allocation { source }))?;
                        owned.push_str(value);
                        match property {
                            SlideProperty::Layout => layout = Some(owned),
                            SlideProperty::Name => name = Some(owned),
                            SlideProperty::Title => title = Some(owned),
                            SlideProperty::Text => body = Some(owned),
                            SlideProperty::Hidden
                            | SlideProperty::Notes
                            | SlideProperty::Transition => {
                                unreachable!("text property")
                            },
                        }
                    },
                    SlideProperty::Notes => unreachable!("notes are not an add setter"),
                    SlideProperty::Transition => unreachable!("transitions are not an add setter"),
                    SlideProperty::Hidden => {
                        let Some(Scalar::Bool(value)) = value else {
                            unreachable!("slide visibility coercion")
                        };
                        staged_bytes = staged_bytes.saturating_add(size_of::<bool>());
                        bound(
                            LimitKind::StagedBytes,
                            staged_bytes,
                            limits.max_staged_bytes(),
                        )?;
                        hidden = Some(value);
                    },
                }
                Ok(())
            })?;
        }
        let layout = layout.ok_or_else(|| {
            error(
                Code::MissingProperty,
                "slide insertion requires an exact layout name",
            )
            .with_property("layout")
        })?;
        bound(LimitKind::ExpandedTargets, 1, limits.max_expanded_targets())?;
        let count = self.slides_in_catalog()?;
        let relative = |path: &PathArg<'_>| -> Result<usize> {
            let node = crate::path::resolve::get(&nav, path)?;
            match node.item() {
                Item::Slide(slide) => Ok(slide.position as usize),
                _ => Err(error(
                    Code::UnsupportedType,
                    "slide placement requires a slide sibling",
                )),
            }
        };
        let index = match node.placement() {
            None | Some(Placement::End) => count,
            Some(Placement::Start) => 0,
            Some(Placement::Position(position)) => position.get(),
            Some(Placement::Before(path)) => relative(path)?,
            Some(Placement::After(path)) => relative(path)?
                .checked_add(1)
                .ok_or_else(|| error(Code::OutOfRange, "slide position overflow"))?,
            Some(Placement::Into(path, position)) => {
                if !matches!(
                    crate::path::resolve::get(&nav, path)?.item(),
                    Item::Presentation
                ) {
                    return Err(error(
                        Code::UnsupportedType,
                        "slide placement requires the presentation root",
                    ));
                }
                position.map_or(count, Position::get)
            },
        };
        let source = self.editable_owner()?;
        // Placeholder text is staged by the eager editor: a source owner is
        // promoted for a slide added with `title` or `text`, as comment and
        // notes writes promote it (record 0346).
        let shorthand = title.is_some() || body.is_some();
        let native_limit = native_snapshot_budget(limits);
        let (target, proof, after) = match &source {
            Editable::Source(_) | Editable::Eager(_) if shorthand => {
                let package = self.eager_base(&source)?;
                let mut edit = package
                    .edit_with_options(opened_edit_options(native_limit))
                    .map_err(super::write::native_error)?;
                let _ = edit.insert_slide(Position::new(index), &layout)?;
                edit.set_slide_properties(index, name.as_deref(), hidden)?;
                if let Some(title) = &title {
                    let _ = edit.put_placeholder_text(
                        index,
                        litchi_pptx::PlaceholderKind::Title,
                        None,
                        title,
                    )?;
                }
                if let Some(body) = &body {
                    // OfficeCLI's content slot: `type="body" idx="1"`.
                    let _ = edit.put_placeholder_text(
                        index,
                        litchi_pptx::PlaceholderKind::Body,
                        Some(1),
                        body,
                    )?;
                }
                let (target, patch) = edit.commit()?.into_parts();
                let target = Arc::new(target);
                let after = self.adopt_eager(Arc::clone(&target));
                (Editable::Eager(target), Proof::Eager(patch), after)
            },
            Editable::Source(package) => {
                let mut edit = package
                    .edit_with_options(source_edit_options(native_limit))
                    .map_err(super::write::native_error)?;
                let _ = edit.insert_slide(Position::new(index), &layout)?;
                edit.set_slide_properties(Position::new(index), name.as_deref(), hidden)?;
                let (target, patch) = edit.commit()?.into_parts();
                let after = self.adopt_source(target.clone())?;
                (Editable::Source(target), Proof::Source(patch), after)
            },
            Editable::Eager(package) => {
                let mut edit = package
                    .edit_with_options(opened_edit_options(native_limit))
                    .map_err(super::write::native_error)?;
                let _ = edit.insert_slide(Position::new(index), &layout)?;
                edit.set_slide_properties(Position::new(index), name.as_deref(), hidden)?;
                let (target, patch) = edit.commit()?.into_parts();
                let target = Arc::new(target);
                let after = self.adopt_eager(Arc::clone(&target));
                (Editable::Eager(target), Proof::Eager(patch), after)
            },
        };
        let path =
            Path::from_steps([Step::new("slide").with(Predicate::Index(Position::new(index)))]);
        Ok((Commit::new(after, source, target, proof, 1), path))
    }
}
/// The modern comment kind and its aliases (record 0346).
const MODERN_KINDS: &[&str] = &[
    "moderncomment",
    "modern-comment",
    "thread",
    "threadedcomment",
];

/// The element kinds a PPTX `add` creates (with their aliases).
const ADDABLE: &[&str] = &[
    "slide",
    "shape",
    "textbox",
    "picture",
    "image",
    "img",
    "chart",
    "comment",
    "moderncomment",
    "modern-comment",
    "thread",
    "threadedcomment",
    "placeholder",
    "ph",
    "notes",
    "smartart",
    "node",
    "col",
    "column",
    "tr",
    "row",
    "table",
    "hyperlink",
    "hlink",
    "video",
    "audio",
    "media",
    "paragraph",
    "p",
    "para",
    "run",
    "r",
];

pub(super) fn bound(kind: LimitKind, observed: usize, maximum: usize) -> Result<()> {
    if observed > maximum {
        return Err(TxError::Limit {
            kind,
            observed,
            maximum,
        }
        .into());
    }
    Ok(())
}
pub(super) fn error(code: Code, message: &'static str) -> Error {
    Error::new(code, message).with_format(FileFormat::Pptx)
}
