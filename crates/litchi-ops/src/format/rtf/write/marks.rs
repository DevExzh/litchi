//! Unified writes to RTF sections, bookmarks, fields, tables of contents
//! and tab stops (record 0328).
//!
//! Each selected node names an entry of the native collections; the whole
//! selection is staged as one native mark edit
//! (`litchi_rtf::edit::Edit::set_field_instruction` and its siblings),
//! which splices only the control words or destination groups that spell
//! the entry and proves the candidate by reparsing it.

use litchi_core::FileFormat;
use litchi_core::selector::{Path, Predicate, Step};
use litchi_rtf::edit::{EditOptions, SectionValue, SemanticLimits as NativeLimits};
use litchi_rtf::{
    PageOrientation, SectionBreakType, TabAlignment, TabLeader, TabStop, VerticalAlignment,
};

use super::super::node::Item;
use super::super::{Kind, Node, Rtf, spec};
use super::{Commit, patch_snapshot_bytes, snapshot_limit, staged_limit};
use crate::error::{Code, Error, Result};
use crate::node::spec::PropSpec;
use crate::props::{PropValue, Props};
use crate::transaction::{Error as TransactionError, LimitKind, SemanticLimits};
use crate::value::Scalar;

/// One native entry selected by a mutation.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord)]
pub(super) enum Target {
    Section(usize),
    Bookmark(usize),
    Field(usize),
    Toc(usize),
    /// A tab stop: body paragraph and tab index.
    Tab(usize, usize),
}

impl Target {
    /// The native entry `node` stands for, if any.
    pub(super) fn of(node: &Node<'_>) -> Option<Self> {
        Some(match node.item {
            Item::Section { index, .. } => Self::Section(index),
            Item::Bookmark { index, .. } => Self::Bookmark(index),
            Item::Field { index, .. } if node.element_kind() == Kind::Toc => Self::Toc(index),
            Item::Field { index, .. } => Self::Field(index),
            Item::Tab { index, .. } => Self::Tab(node.paragraph_position()?.get(), index),
            _ => return None,
        })
    }

    const fn family(self) -> u8 {
        match self {
            Self::Section(_) => 0,
            Self::Bookmark(_) => 1,
            Self::Field(_) | Self::Toc(_) => 2,
            Self::Tab(..) => 3,
        }
    }
}

pub(super) fn append(targets: &mut Vec<Target>, node: &Node<'_>, maximum: usize) -> Result<()> {
    let target = Target::of(node).ok_or_else(|| {
        Error::new(
            Code::UnsupportedType,
            "RTF mark mutations target sections, bookmarks, fields, tables of contents and tab stops",
        )
        .with_format(FileFormat::Rtf)
    })?;
    if let Some(first) = targets.first()
        && first.family() != target.family()
    {
        return Err(Error::new(
            Code::UnsupportedType,
            "RTF mutations cannot mix section, bookmark, field and tab targets",
        )
        .with_format(FileFormat::Rtf));
    }
    if targets.len() >= maximum {
        return Err(TransactionError::Limit {
            kind: LimitKind::ExpandedTargets,
            observed: targets.len().checked_add(1).unwrap_or(usize::MAX),
            maximum,
        }
        .into());
    }
    targets.push(target);
    Ok(())
}

fn invalid(key: &str, message: &'static str) -> Error {
    Error::new(Code::InvalidValue, message)
        .with_property(key.to_owned())
        .with_format(FileFormat::Rtf)
}

/// Refuse a key `specs` does not name as writable by `set` (or by `add`).
fn admit_keys(props: &Props<'_>, specs: &[PropSpec], add: bool) -> Result<()> {
    if props.is_empty() {
        return Err(
            Error::new(Code::InvalidArgument, "the RTF write names no property")
                .with_format(FileFormat::Rtf),
        );
    }
    for (key, _) in props.iter() {
        let writable = specs.iter().any(|spec| {
            spec.name.eq_ignore_ascii_case(key.as_str())
                && if add {
                    spec.access.add
                } else {
                    spec.access.set
                }
        });
        if !writable {
            return Err(Error::new(
                Code::UnsupportedCapability,
                "the requested RTF property is not writable",
            )
            .with_property(key.as_str().to_owned())
            .with_format(FileFormat::Rtf));
        }
    }
    Ok(())
}

/// `value` coerced as `key` of `specs` declares it, passed to `read`.
fn scalar<T>(
    specs: &[PropSpec],
    key: &'static str,
    value: &PropValue<'_>,
    read: impl for<'s> FnOnce(Scalar<'s>) -> Option<T>,
) -> Result<T> {
    let ty = specs
        .iter()
        .find(|spec| spec.name == key)
        .map(|spec| spec.ty)
        .ok_or_else(|| invalid(key, "the RTF property has no declared type"))?;
    value
        .with_coerced(ty, |value| {
            value
                .and_then(read)
                .ok_or_else(|| invalid(key, "the RTF property value is out of range"))
        })
        .map_err(|error| error.with_property(key).with_format(FileFormat::Rtf))
}

fn twips(specs: &[PropSpec], key: &'static str, value: &PropValue<'_>) -> Result<i32> {
    scalar(specs, key, value, |value| match value {
        Scalar::Length(length) => length.checked_twips(),
        _ => None,
    })
}

fn boolean(specs: &[PropSpec], key: &'static str, value: &PropValue<'_>) -> Result<bool> {
    scalar(specs, key, value, |value| match value {
        Scalar::Bool(value) => Some(value),
        _ => None,
    })
}

fn token(specs: &[PropSpec], key: &'static str, value: &PropValue<'_>) -> Result<&'static str> {
    scalar(specs, key, value, |value| match value {
        Scalar::Token(value) => Some(value),
        _ => None,
    })
}

fn integer(specs: &[PropSpec], key: &'static str, value: &PropValue<'_>) -> Result<i64> {
    scalar(specs, key, value, |value| match value {
        Scalar::Int(value) => Some(value),
        _ => None,
    })
}

fn text<'a>(props: &'a Props<'_>, key: &'static str) -> Result<Option<&'a str>> {
    props
        .get(key)
        .map(|value| {
            value
                .as_text()
                .ok_or_else(|| invalid(key, "the RTF property requires a text value"))
        })
        .transpose()
}

/// The page-setup values `props` sets on a section.
fn section_values(props: &Props<'_>) -> Result<Vec<SectionValue>> {
    let specs = spec::SECTION_PROPS;
    admit_keys(props, specs, false)?;
    let mut values = Vec::new();
    for (key, value) in props.iter() {
        let key = specs
            .iter()
            .find(|spec| spec.name.eq_ignore_ascii_case(key.as_str()))
            .map_or("", |spec| spec.name);
        values.push(match key {
            "type" => SectionValue::BreakType(match token(specs, key, value)? {
                "continuous" => SectionBreakType::Continuous,
                "nextColumn" => SectionBreakType::Column,
                "evenPage" => SectionBreakType::EvenPage,
                "oddPage" => SectionBreakType::OddPage,
                _ => SectionBreakType::Page,
            }),
            "pageWidth" => SectionValue::PageWidth(twips(specs, key, value)?),
            "pageHeight" => SectionValue::PageHeight(twips(specs, key, value)?),
            "orientation" => SectionValue::Orientation(match token(specs, key, value)? {
                "landscape" => PageOrientation::Landscape,
                _ => PageOrientation::Portrait,
            }),
            "marginTop" => SectionValue::MarginTop(twips(specs, key, value)?),
            "marginBottom" => SectionValue::MarginBottom(twips(specs, key, value)?),
            "marginLeft" => SectionValue::MarginLeft(twips(specs, key, value)?),
            "marginRight" => SectionValue::MarginRight(twips(specs, key, value)?),
            "marginHeader" => SectionValue::HeaderDistance(twips(specs, key, value)?),
            "marginFooter" => SectionValue::FooterDistance(twips(specs, key, value)?),
            "marginGutter" => SectionValue::MarginGutter(twips(specs, key, value)?),
            "columns" => SectionValue::Columns(
                u16::try_from(integer(specs, key, value)?)
                    .map_err(|_error| invalid(key, "RTF section columns run from 1 to 64"))?,
            ),
            "columnSpace" => SectionValue::ColumnSpacing(twips(specs, key, value)?),
            "titlePage" => SectionValue::TitlePage(boolean(specs, key, value)?),
            "vAlign" => SectionValue::VerticalAlignment(match token(specs, key, value)? {
                "center" => VerticalAlignment::Center,
                "both" => VerticalAlignment::Justify,
                "bottom" => VerticalAlignment::Bottom,
                _ => VerticalAlignment::Top,
            }),
            _ => {
                return Err(invalid(key, "the RTF section property is not writable"));
            },
        });
    }
    Ok(values)
}

/// The words of a field instruction, a quoted argument kept whole.
fn instruction_words(instruction: &str) -> Vec<String> {
    let mut words = Vec::new();
    let mut current = String::new();
    let mut quoted = false;
    for character in instruction.chars() {
        match character {
            '"' => {
                quoted = !quoted;
                current.push(character);
            },
            character if character.is_whitespace() && !quoted => {
                if !current.is_empty() {
                    words.push(std::mem::take(&mut current));
                }
            },
            character => current.push(character),
        }
    }
    if !current.is_empty() {
        words.push(current);
    }
    words
}

/// Whether `word` is the switch `\name` (case-insensitive).
fn is_switch(word: &str, name: &str) -> bool {
    word.strip_prefix('\\')
        .is_some_and(|switch| switch.eq_ignore_ascii_case(name))
}

/// The `TOC` instruction `instruction` with `props` applied to its
/// switches: `levels` sets `\o "N-M"`, `hyperlinks` adds or drops `\h`,
/// `pageNumbers` drops or adds a bare `\n`; `instruction` replaces it
/// whole.
fn toc_instruction(instruction: &str, props: &Props<'_>) -> Result<String> {
    let specs = spec::TOC_PROPS;
    admit_keys(props, specs, false)?;
    if let Some(whole) = text(props, "instruction")? {
        if props.len() > 1 {
            return Err(Error::new(
                Code::InvalidArgument,
                "an RTF table of contents takes `instruction` alone or its switch properties",
            )
            .with_format(FileFormat::Rtf));
        }
        return Ok(whole.to_owned());
    }
    let mut words = instruction_words(instruction);
    if let Some(levels) = text(props, "levels")? {
        let valid = levels.split_once('-').is_some_and(|(low, high)| {
            matches!((low.parse::<u8>(), high.parse::<u8>()), (Ok(low), Ok(high)) if (1..=9).contains(&low) && (low..=9).contains(&high))
        });
        if !valid {
            return Err(invalid(
                "levels",
                "TOC levels are a range such as 1-3 within 1-9",
            ));
        }
        let argument = format!("\"{levels}\"");
        match words.iter().position(|word| is_switch(word, "o")) {
            Some(at) => {
                let next = at + 1;
                if words.get(next).is_some_and(|word| word.starts_with('"')) {
                    words[next] = argument;
                } else {
                    words.insert(next, argument);
                }
            },
            None => {
                words.push("\\o".to_owned());
                words.push(argument);
            },
        }
    }
    if let Some(value) = props.get("hyperlinks") {
        let on = boolean(specs, "hyperlinks", value)?;
        words.retain(|word| !is_switch(word, "h"));
        if on {
            words.push("\\h".to_owned());
        }
    }
    if let Some(value) = props.get("pageNumbers") {
        let on = boolean(specs, "pageNumbers", value)?;
        let mut kept = Vec::with_capacity(words.len());
        let mut iter = words.into_iter().peekable();
        while let Some(word) = iter.next() {
            if is_switch(&word, "n") && !iter.peek().is_some_and(|next| next.starts_with('"')) {
                continue;
            }
            kept.push(word);
        }
        words = kept;
        if !on {
            words.push("\\n".to_owned());
        }
    }
    Ok(words.join(" "))
}

/// A tab stop with `props` applied (`pos`, `val`, `leader`).
fn tab_stop(mut stop: TabStop, props: &Props<'_>, add: bool) -> Result<TabStop> {
    let specs = spec::TAB_PROPS;
    admit_keys(props, specs, add)?;
    if let Some(value) = props.get("pos") {
        stop.position = i32::try_from(integer(specs, "pos", value)?)
            .ok()
            .filter(|position| (0..=31_680).contains(position))
            .ok_or_else(|| invalid("pos", "an RTF tab stop position runs from 0 to 31680 twips"))?;
    }
    if let Some(value) = props.get("val") {
        stop.alignment = match token(specs, "val", value)? {
            "center" => TabAlignment::Center,
            "right" => TabAlignment::Right,
            "decimal" => TabAlignment::Decimal,
            "bar" => TabAlignment::Bar,
            _ => TabAlignment::Left,
        };
    }
    if let Some(value) = props.get("leader") {
        stop.leader = match token(specs, "leader", value)? {
            "dot" => TabLeader::Dot,
            "middleDot" => TabLeader::MiddleDot,
            "hyphen" => TabLeader::Hyphen,
            "underscore" => TabLeader::Underscore,
            "heavy" => TabLeader::ThickLine,
            "equal" => TabLeader::Equal,
            _ => TabLeader::None,
        };
    }
    Ok(stop)
}

fn tab_stops(file: &Rtf, paragraph: usize) -> Result<Vec<TabStop>> {
    file.0
        .package
        .body()
        .paragraphs()
        .nth(paragraph)
        .map(|value| value.format().tab_stops().to_vec())
        .ok_or_else(|| {
            Error::new(Code::PathNotFound, "the RTF paragraph does not exist")
                .with_format(FileFormat::Rtf)
        })
}

/// The tab stops of each paragraph the targets select, in target order.
fn by_paragraph(targets: &[Target]) -> Vec<(usize, Vec<usize>)> {
    let mut out: Vec<(usize, Vec<usize>)> = Vec::new();
    for target in targets {
        if let Target::Tab(paragraph, index) = *target {
            match out.iter_mut().find(|(held, _)| *held == paragraph) {
                Some((_, indexes)) => indexes.push(index),
                None => out.push((paragraph, vec![index])),
            }
        }
    }
    out
}

fn native_options(limits: &SemanticLimits) -> Result<EditOptions> {
    let native = NativeLimits::builder()
        .max_operations(limits.max_operations())
        .map_err(Error::from)?
        .max_replacement_bytes(limits.max_staged_bytes())
        .map_err(Error::from)?
        .build()
        .map_err(Error::from)?;
    Ok(EditOptions::new().semantic_limits(native))
}

/// Stage with `stage` on a fresh native edit and publish it.
fn publish(
    file: &Rtf,
    operations: usize,
    limits: &SemanticLimits,
    stage: impl FnOnce(&mut litchi_rtf::edit::Edit) -> std::result::Result<(), litchi_rtf::edit::Error>,
) -> Result<Commit> {
    if operations > limits.max_operations() {
        return Err(TransactionError::Limit {
            kind: LimitKind::Operations,
            observed: operations,
            maximum: limits.max_operations(),
        }
        .into());
    }
    let mut edit = file.0.package.edit_with_options(native_options(limits)?);
    stage(&mut edit).map_err(Error::from)?;
    let native = edit.commit().map_err(Error::from)?;
    let changed = native.diagnostics().changed();
    let operation_count = native.diagnostics().operation_count();
    let (package, patch) = native.into_parts();
    let snapshot_bytes = patch_snapshot_bytes(&patch);
    if snapshot_bytes > limits.max_snapshot_bytes() {
        return Err(snapshot_limit(snapshot_bytes, limits.max_snapshot_bytes()));
    }
    let before = file.clone();
    let after = if changed {
        file.with_package(package)
    } else {
        file.clone()
    };
    file.check_source()?;
    Ok(Commit {
        before,
        after,
        patch,
        operations: operation_count,
        snapshot_bytes,
    })
}

/// Set `props` on every selected entry in one native edit.
pub(super) fn set(
    file: &Rtf,
    targets: &[Target],
    props: &Props<'_>,
    limits: &SemanticLimits,
) -> Result<Commit> {
    let package = &file.0.package;
    let staged = props
        .iter()
        .map(|(_, value)| value.as_text().map_or(16, str::len))
        .sum::<usize>()
        .saturating_mul(targets.len());
    if staged > limits.max_staged_bytes() {
        return Err(staged_limit(staged, limits.max_staged_bytes()));
    }
    match targets.first() {
        Some(Target::Tab(..)) => {
            let groups = by_paragraph(targets);
            let mut plans = Vec::with_capacity(groups.len());
            for (paragraph, indexes) in groups {
                let mut stops = tab_stops(file, paragraph)?;
                for index in indexes {
                    let stop = stops.get(index).copied().ok_or_else(|| {
                        Error::new(Code::PathNotFound, "the RTF tab stop does not exist")
                            .with_format(FileFormat::Rtf)
                    })?;
                    stops[index] = tab_stop(stop, props, false)?;
                }
                plans.push((paragraph, stops));
            }
            publish(file, plans.len(), limits, |edit| {
                for (paragraph, stops) in &plans {
                    edit.set_paragraph_tab_stops(*paragraph, stops)?;
                }
                Ok(())
            })
        },
        Some(Target::Section(_)) => {
            let values = section_values(props)?;
            publish(file, targets.len(), limits, |edit| {
                for target in targets {
                    if let Target::Section(index) = target {
                        edit.set_section_values(*index, &values)?;
                    }
                }
                Ok(())
            })
        },
        Some(Target::Bookmark(_)) => {
            admit_keys(props, spec::BOOKMARK_PROPS, false)?;
            let name = text(props, "name")?
                .ok_or_else(|| invalid("name", "an RTF bookmark name is text"))?;
            if targets.len() > 1 {
                return Err(Error::new(
                    Code::InvalidArgument,
                    "RTF bookmark names are unique; rename one bookmark at a time",
                )
                .with_format(FileFormat::Rtf));
            }
            let bookmarks = package.raw().bookmarks().bookmarks();
            if targets.iter().any(|target| {
                bookmarks
                    .iter()
                    .enumerate()
                    .any(|(index, held)| *target != Target::Bookmark(index) && held.name == name)
            }) {
                return Err(invalid(
                    "name",
                    "another RTF bookmark already has this name",
                ));
            }
            publish(file, 1, limits, |edit| {
                for target in targets {
                    if let Target::Bookmark(index) = target {
                        edit.rename_bookmark(*index, name)?;
                    }
                }
                Ok(())
            })
        },
        Some(Target::Field(_) | Target::Toc(_)) => {
            let mut plans = Vec::with_capacity(targets.len());
            for target in targets {
                let (index, instruction) = match *target {
                    Target::Field(index) => {
                        admit_keys(props, spec::FIELD_PROPS, false)?;
                        let instruction = text(props, "instruction")?.ok_or_else(|| {
                            invalid("instruction", "an RTF field instruction is text")
                        })?;
                        (index, instruction.to_owned())
                    },
                    Target::Toc(index) => {
                        let current = package
                            .fields()
                            .get(index)
                            .map(|field| field.instruction.trim().to_owned())
                            .unwrap_or_default();
                        (index, toc_instruction(&current, props)?)
                    },
                    _ => continue,
                };
                plans.push((index, instruction));
            }
            publish(file, plans.len(), limits, |edit| {
                for (index, instruction) in &plans {
                    edit.set_field_instruction(*index, instruction.as_str())?;
                }
                Ok(())
            })
        },
        None => Err(Error::new(
            Code::ZeroMatches,
            "the RTF mutation selector matched nothing",
        )
        .with_format(FileFormat::Rtf)),
    }
}

/// Remove every selected bookmark or tab stop in one native edit.
pub(super) fn remove(file: &Rtf, targets: &[Target], limits: &SemanticLimits) -> Result<Commit> {
    match targets.first() {
        Some(Target::Tab(..)) => {
            let groups = by_paragraph(targets);
            let mut plans = Vec::with_capacity(groups.len());
            for (paragraph, indexes) in groups {
                let stops: Vec<TabStop> = tab_stops(file, paragraph)?
                    .into_iter()
                    .enumerate()
                    .filter(|(index, _)| !indexes.contains(index))
                    .map(|(_, stop)| stop)
                    .collect();
                plans.push((paragraph, stops));
            }
            publish(file, plans.len(), limits, |edit| {
                for (paragraph, stops) in &plans {
                    edit.set_paragraph_tab_stops(*paragraph, stops)?;
                }
                Ok(())
            })
        },
        Some(Target::Bookmark(_)) => publish(file, targets.len(), limits, |edit| {
            for target in targets {
                if let Target::Bookmark(index) = target {
                    edit.remove_bookmark(*index)?;
                }
            }
            Ok(())
        }),
        Some(_) => Err(Error::new(
            Code::UnsupportedCapability,
            "RTF sections, fields and tables of contents cannot be removed; set their properties instead",
        )
        .with_format(FileFormat::Rtf)),
        None => Err(Error::new(
            Code::ZeroMatches,
            "the RTF mutation selector matched nothing",
        )
        .with_format(FileFormat::Rtf)),
    }
}

/// Add a tab stop under body paragraph `paragraph`, kept in position
/// order; returns the commit and the new stop's path.
pub(super) fn add_tab(
    file: &Rtf,
    paragraph: usize,
    props: &Props<'_>,
    limits: &SemanticLimits,
) -> Result<(Commit, Path<'static>)> {
    if props.get("pos").is_none() {
        return Err(
            Error::new(Code::MissingProperty, "an RTF tab stop needs a `pos`")
                .with_property("pos")
                .with_format(FileFormat::Rtf),
        );
    }
    let stop = tab_stop(TabStop::new(0), props, true)?;
    let mut stops = tab_stops(file, paragraph)?;
    if stops.iter().any(|held| held.position == stop.position) {
        return Err(invalid(
            "pos",
            "the RTF paragraph already has a tab stop at this position",
        ));
    }
    let at = stops
        .iter()
        .position(|held| held.position > stop.position)
        .unwrap_or(stops.len());
    stops.insert(at, stop);
    let commit = publish(file, 1, limits, |edit| {
        edit.set_paragraph_tab_stops(paragraph, &stops)?;
        Ok(())
    })?;
    let path = Path::from_steps([
        Step::new("body"),
        Step::new("p").with(Predicate::Index(litchi_core::Position::new(paragraph))),
        Step::new("tab").with(Predicate::Index(litchi_core::Position::new(at))),
    ]);
    Ok((commit, path))
}
