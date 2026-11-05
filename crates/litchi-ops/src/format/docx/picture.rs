//! Inline picture insertion (record 0251d).
//!
//! `add(<paragraph>, picture, {src, width, height, alt, name})` reads the
//! image from a base64 `data:` URI or a local file path, bounded by the
//! default staged-byte limit, and stages it through the leaf
//! `insert_picture_at`: a media part, its relationship and a run holding
//! the inline drawing, in one commit.
//!
//! `set` (record 0335) writes `width`, `height` (the drawing's and the
//! shape's extent; one side keeps the other), `alt`, `name` and `src` (a
//! new image part the picture then names) through the leaf
//! `set_picture_at`; `remove` drops the drawing, and its run when the run
//! holds nothing else, through `remove_picture_at`.

use litchi_core::{FileFormat, Position};
use litchi_docx::document::{ElementPath, NewPicture, PictureChange};

use super::element::{element_path, refusal};
use super::node::{Node, Value};
use super::setters::Setter;
use super::spec::Kind;
use super::write::{Commit, WorkEdit};
use super::{Docx, DocxEditor};
use crate::error::{Code, Error, Result};
use crate::props::Props;
use crate::registry::Operation;
use crate::transaction::{LimitKind, SemanticLimits};
use crate::value::Scalar;

fn error(code: Code, message: &'static str) -> Error {
    Error::new(code, message)
        .with_format(FileFormat::Docx)
        .with_property("src")
}

fn bound(observed: usize, maximum: usize) -> Result<()> {
    if observed > maximum {
        return Err(crate::transaction::Error::Limit {
            kind: LimitKind::StagedBytes,
            observed,
            maximum,
        }
        .into());
    }
    Ok(())
}

/// Read a picture source: a base64 `data:` URI or a local file path.
fn read_source(src: &str, maximum: usize) -> Result<Vec<u8>> {
    if let Some(rest) = src.strip_prefix("data:") {
        let (header, payload) = rest.split_once(',').ok_or_else(|| {
            error(
                Code::InvalidValue,
                "a data URI needs a comma before its payload",
            )
        })?;
        if !header.ends_with(";base64") {
            return Err(error(
                Code::UnsupportedCapability,
                "only base64 data URIs are accepted",
            ));
        }
        // Base64 expands by 4/3; refuse before decoding.
        bound(payload.len() / 4 * 3, maximum)?;
        use base64::Engine as _;
        return base64::engine::general_purpose::STANDARD
            .decode(payload.trim())
            .map_err(|_| error(Code::InvalidValue, "the data URI payload is not base64"));
    }
    let length = std::fs::metadata(src)
        .map_err(|source| Error::from(source).with_property("src"))?
        .len();
    bound(usize::try_from(length).unwrap_or(usize::MAX), maximum)?;
    use std::io::Read as _;
    let mut data = Vec::new();
    std::fs::File::open(src)
        .and_then(|file| file.take(maximum as u64 + 1).read_to_end(&mut data))
        .map_err(|source| Error::from(source).with_property("src"))?;
    bound(data.len(), maximum)?;
    Ok(data)
}

/// A new picture from its add properties: `src` is required.
pub(super) fn new_picture(props: &Props<'_>) -> Result<NewPicture> {
    let mut src = None::<String>;
    let mut width = None;
    let mut height = None;
    let mut alt = None::<String>;
    let mut name = None::<String>;
    for (key, value) in props.iter() {
        let (_, entry) = super::spec::PICTURE_REGISTRY
            .resolve(key.as_str())
            .map_err(|error| error.with_property(key.as_str().to_owned()))?;
        entry
            .prepare(Operation::Add, value, |setter, value| {
                let text = |value: Option<Scalar<'_>>| match value {
                    Some(Scalar::Text(text) | Scalar::Token(text)) => text.to_owned(),
                    Some(scalar) => scalar.to_string(),
                    None => String::new(),
                };
                let emu = |value: Option<Scalar<'_>>| match value {
                    Some(Scalar::Length(length)) => Ok(length.emu()),
                    _ => Err(Error::new(
                        Code::InvalidValue,
                        "a picture extent is a length",
                    )),
                };
                match setter {
                    Setter::PictureSrc => src = Some(text(value)),
                    Setter::PictureWidth => width = Some(emu(value)?),
                    Setter::PictureHeight => height = Some(emu(value)?),
                    Setter::PictureAlt => alt = Some(text(value)),
                    Setter::PictureName => name = Some(text(value)),
                    _ => {
                        return Err(Error::new(
                            Code::UnsupportedProperty,
                            "the property does not seed a DOCX picture",
                        ));
                    },
                }
                Ok(())
            })
            .map_err(|error| {
                error
                    .with_property(key.as_str().to_owned())
                    .with_format(FileFormat::Docx)
            })?;
    }
    let Some(src) = src else {
        return Err(Error::new(
            Code::MissingProperty,
            "a new DOCX picture needs its image source",
        )
        .with_property("src")
        .with_format(FileFormat::Docx));
    };
    let data = read_source(&src, SemanticLimits::default().max_staged_bytes())?;
    if litchi_docx::ImageFormat::detect_from_bytes(&data).is_none() {
        return Err(error(
            Code::InvalidValue,
            "the picture source is not a recognized image",
        ));
    }
    let mut picture = NewPicture::new(data).with_size(width, height);
    if let Some(alt) = alt {
        picture = picture.with_description(alt);
    }
    if let Some(name) = name {
        picture = picture.with_name(name);
    }
    Ok(picture)
}

/// The admitted values of a picture `set` (record 0335).
fn picture_change(props: &Props<'_>) -> Result<(PictureChange, (Option<i64>, Option<i64>))> {
    let mut change = PictureChange::new();
    let mut size = (None, None);
    for (key, value) in props.iter() {
        let (_, entry) = super::spec::PICTURE_REGISTRY
            .resolve(key.as_str())
            .map_err(|error| error.with_property(key.as_str().to_owned()))?;
        let mut source = None;
        entry
            .prepare(Operation::Set, value, |setter, value| {
                let text = |value: Option<Scalar<'_>>| match value {
                    Some(Scalar::Text(text) | Scalar::Token(text)) => text.to_owned(),
                    Some(scalar) => scalar.to_string(),
                    None => String::new(),
                };
                let emu = |value: Option<Scalar<'_>>| match value {
                    Some(Scalar::Length(length)) => Ok(length.emu()),
                    _ => Err(Error::new(
                        Code::InvalidValue,
                        "a picture extent is a length",
                    )),
                };
                match setter {
                    Setter::PictureSrc => source = Some(text(value)),
                    Setter::PictureWidth => size.0 = Some(emu(value)?),
                    Setter::PictureHeight => size.1 = Some(emu(value)?),
                    Setter::PictureAlt => {
                        change = std::mem::take(&mut change).with_description(text(value));
                    },
                    Setter::PictureName => {
                        change = std::mem::take(&mut change).with_name(text(value));
                    },
                    _ => {
                        return Err(Error::new(
                            Code::UnsupportedProperty,
                            "the property does not set a DOCX picture",
                        ));
                    },
                }
                Ok(())
            })
            .map_err(|error| {
                error
                    .with_property(key.as_str().to_owned())
                    .with_format(FileFormat::Docx)
            })?;
        if let Some(source) = source {
            let data = read_source(&source, SemanticLimits::default().max_staged_bytes())?;
            if litchi_docx::ImageFormat::detect_from_bytes(&data).is_none() {
                return Err(error(
                    Code::InvalidValue,
                    "the picture source is not a recognized image",
                ));
            }
            change = change.with_image(data);
        }
    }
    Ok((change, size))
}

/// The paragraph path and picture position of a picture node.
fn picture_target(docx: &Docx, node: &Node<'_>) -> Result<(ElementPath, Position)> {
    let paragraph = super::nested::parent_node(docx, node)?;
    let ordinal = node
        .address
        .steps()
        .last()
        .map(|step| Position::new(step.ordinal as usize))
        .ok_or_else(|| error(Code::InternalError, "the picture node has no address"))?;
    Ok((element_path(&paragraph)?, ordinal))
}

impl Docx {
    /// A `set` whose targets are pictures (record 0335); `Err` hands the
    /// editor back for other targets.
    pub(super) fn try_set_picture(
        &self,
        editor: DocxEditor,
        nodes: &[Node<'_>],
        props: &Props<'_>,
    ) -> Result<std::result::Result<Commit, DocxEditor>> {
        if nodes.is_empty() || nodes.iter().any(|node| node.kind != Kind::Picture) {
            return Ok(Err(editor));
        }
        let (change, size) = picture_change(props)?;
        let mut targets = Vec::with_capacity(nodes.len());
        for node in nodes {
            let Value::Picture(picture) = &node.value else {
                return Err(error(
                    Code::InternalError,
                    "the picture node holds no picture",
                ));
            };
            let (paragraph, ordinal) = picture_target(self, node)?;
            let mut change = change.clone();
            if size.0.is_some() || size.1.is_some() {
                change = change.with_size(
                    size.0.unwrap_or(picture.width_emu()),
                    size.1.unwrap_or(picture.height_emu()),
                );
            }
            targets.push((paragraph, ordinal, change));
        }
        self.owner().check_source()?;
        let mut edit = WorkEdit::from(self.eager_editor(editor)?)?;
        let WorkEdit::Eager(eager) = &mut edit else {
            return Err(refusal("pictures need the owning package editor"));
        };
        for (paragraph, ordinal, change) in &targets {
            if !change.is_empty() {
                eager
                    .set_picture_at(paragraph, *ordinal, change)
                    .map_err(Error::from)?;
            }
        }
        edit.commit(self).map(Ok)
    }

    /// Remove pictures selected by their paths (record 0335).
    pub(super) fn remove_pictures(&self, editor: DocxEditor, nodes: &[Node<'_>]) -> Result<Commit> {
        let mut targets = Vec::with_capacity(nodes.len());
        for node in nodes {
            targets.push(picture_target(self, node)?);
        }
        // Later pictures first, so earlier positions stay valid.
        targets.sort_by(|left, right| {
            super::nested::order_key(&right.0)
                .cmp(&super::nested::order_key(&left.0))
                .then(right.1.cmp(&left.1))
        });
        targets.dedup();
        self.owner().check_source()?;
        let mut edit = WorkEdit::from(self.eager_editor(editor)?)?;
        let WorkEdit::Eager(eager) = &mut edit else {
            return Err(refusal("pictures need the owning package editor"));
        };
        for (paragraph, ordinal) in &targets {
            eager
                .remove_picture_at(paragraph, *ordinal)
                .map_err(Error::from)?;
        }
        edit.commit(self)
    }
}
