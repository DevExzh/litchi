//! The non-batch routes of record 0351 §5 per format: open (default and
//! eager unified owner), first and second save of one freshly opened owner,
//! metadata title set, DOCX paragraph add, DOC paragraph get and the PPTX
//! structure edits. The fixture is the document the format's batch builder
//! made (`fixture` targets), published once by its unified handle.

#![allow(unused_imports, unused_variables)]

use std::time::Duration;

use crate::{Case, Checked, Equal, Result, compare, run_timed};
use litchi_core::Position;
use litchi_ops::{File, OpenOptions, Props, SourceMode};

const TITLE: &str = "Quarterly review";

fn leak<T>(value: T) -> &'static T {
    Box::leak(Box::new(value))
}

type Open<P> = &'static (dyn Fn(&[u8]) -> Result<P> + Sync);
type Save<P> = &'static (dyn Fn(&mut P) -> Result<Vec<u8>> + Sync);
type LeafEdit<P> = &'static (dyn Fn(&P) -> Result<P> + Sync);
type UnifiedEdit = &'static (dyn Fn(&File) -> Result<File> + Sync);

/// One leaf owner type and how the unified side opens the same bytes.
struct Owner<P: 'static> {
    format: &'static str,
    bytes: &'static [u8],
    /// The unified handle of the save and edit routes is eager (as the batch
    /// cases' handles are); the `open` route uses the default open.
    eager: bool,
    open: Open<P>,
    save: Save<P>,
}

impl<P: 'static> Clone for Owner<P> {
    fn clone(&self) -> Self {
        Self { ..*self }
    }
}

impl<P: 'static> Copy for Owner<P> {}

fn eager_options() -> OpenOptions {
    OpenOptions::new().source_mode(SourceMode::Eager)
}

fn unified_open(bytes: &[u8], eager: bool) -> Result<File> {
    Ok(if eager {
        litchi_ops::open_bytes_with(bytes.to_vec(), &eager_options())?
    } else {
        litchi_ops::open_bytes(bytes.to_vec())?
    })
}

/// Times `call` on an untimed input it borrows; input and output drop after
/// the clock stops.
fn timed_ref<I, T>(
    input: impl FnOnce() -> Result<I>,
    call: impl FnOnce(&mut I) -> Result<T>,
) -> Result<Duration> {
    let mut input = input()?;
    let (output, elapsed) = run_timed(|| call(&mut input))?;
    drop(output);
    drop(input);
    Ok(elapsed)
}

fn soft(leaf: &[u8], unified: &[u8]) -> Equal {
    compare(leaf, unified).unwrap_or(Equal::Differs)
}

fn checked(equal: Equal) -> Checked {
    Checked {
        equal,
        leaf_commits: 1,
        unified_commits: 1,
    }
}

fn case(
    owner_format: &'static str,
    scenario: &'static str,
    owner: &'static str,
    check: Box<dyn FnMut() -> Result<Checked>>,
    leaf: crate::Leg<'static>,
    unified: crate::Leg<'static>,
) -> Case<'static> {
    Case {
        format: owner_format,
        scenario,
        owner,
        count: 1,
        check,
        leaf,
        unified,
    }
}

/// `open` (default and eager unified owner).
fn owner_cases<P: 'static>(o: Owner<P>) -> Vec<Case<'static>> {
    let mut cases = Vec::new();
    cases.push(case(
        o.format,
        "open",
        "default",
        Box::new(move || {
            (o.open)(o.bytes)?;
            litchi_ops::open_bytes(o.bytes.to_vec())?;
            Ok(checked(Equal::NotApplicable))
        }),
        Box::new(move || timed_ref(|| Ok(()), |_| (o.open)(o.bytes))),
        Box::new(move || {
            timed_ref(
                || Ok(Some(o.bytes.to_vec())),
                |b| Ok(litchi_ops::open_bytes(b.take().ok_or("bytes")?)?),
            )
        }),
    ));
    if o.eager {
        cases.push(case(
            o.format,
            "open",
            "eager",
            Box::new(move || {
                unified_open(o.bytes, true)?;
                Ok(checked(Equal::NotApplicable))
            }),
            Box::new(move || timed_ref(|| Ok(()), |_| (o.open)(o.bytes))),
            Box::new(move || {
                timed_ref(
                    || Ok(Some(o.bytes.to_vec())),
                    |b| {
                        Ok(litchi_ops::open_bytes_with(
                            b.take().ok_or("bytes")?,
                            &eager_options(),
                        )?)
                    },
                )
            }),
        ));
    }
    cases
}

/// `save_first` and `save_second`: the first and the second save of one
/// freshly opened owner after one edit (`edit` on the leaf package,
/// `unified` on the unified handle), so the save publishes a changed part
/// and passes the untouched ones through.
fn save_cases<P: 'static>(
    o: Owner<P>,
    edit: LeafEdit<P>,
    unified: UnifiedEdit,
) -> Vec<Case<'static>> {
    let mut cases = Vec::new();
    let owner = if o.eager { "eager" } else { "default" };
    cases.push(case(
        o.format,
        "save_first",
        owner,
        Box::new(move || {
            let leaf = (o.save)(&mut edit(&(o.open)(o.bytes)?)?)?;
            let unified = unified(&unified_open(o.bytes, o.eager)?)?.to_bytes()?;
            Ok(checked(soft(&leaf, &unified)))
        }),
        Box::new(move || timed_ref(|| edit(&(o.open)(o.bytes)?), |p| (o.save)(p))),
        Box::new(move || {
            timed_ref(
                || unified(&unified_open(o.bytes, o.eager)?),
                |f| Ok(f.to_bytes()?),
            )
        }),
    ));
    cases.push(case(
        o.format,
        "save_second",
        owner,
        Box::new(move || {
            let mut package = edit(&(o.open)(o.bytes)?)?;
            let (first, second) = ((o.save)(&mut package)?, (o.save)(&mut package)?);
            if first != second {
                return Err(format!("{}: leaf second save differs", o.format).into());
            }
            let file = unified(&unified_open(o.bytes, o.eager)?)?;
            let (first, unified) = (file.to_bytes()?, file.to_bytes()?);
            if first != unified {
                return Err(format!("{}: unified second save differs", o.format).into());
            }
            Ok(checked(soft(&second, &unified)))
        }),
        Box::new(move || {
            timed_ref(
                || {
                    let mut package = edit(&(o.open)(o.bytes)?)?;
                    (o.save)(&mut package)?;
                    Ok(package)
                },
                |p| (o.save)(p),
            )
        }),
        Box::new(move || {
            timed_ref(
                || {
                    let file = unified(&unified_open(o.bytes, o.eager)?)?;
                    file.to_bytes()?;
                    Ok(file)
                },
                |f| Ok(f.to_bytes()?),
            )
        }),
    ));
    cases
}

/// One leaf edit (edit + commit) against one unified call on prepared owners.
fn edit_case<P: 'static>(
    o: Owner<P>,
    scenario: &'static str,
    leaf: LeafEdit<P>,
    unified: UnifiedEdit,
) -> Result<Case<'static>> {
    let package: &'static P = leak((o.open)(o.bytes)?);
    let file: &'static File = leak(unified_open(o.bytes, o.eager)?);
    let owner = if o.eager { "eager" } else { "default" };
    Ok(case(
        o.format,
        scenario,
        owner,
        Box::new(move || {
            let leaf = (o.save)(&mut leaf(package)?)?;
            let unified = unified(file)?.to_bytes()?;
            Ok(checked(soft(&leaf, &unified)))
        }),
        Box::new(move || timed_ref(|| Ok(()), |_| leaf(package))),
        Box::new(move || timed_ref(|| Ok(()), |_| unified(file))),
    ))
}

fn title() -> Props<'static> {
    Props::new().with("title", TITLE)
}

/// The OLE summary-information title (`SummaryProperties::set`); the
/// baseline's `properties_mut` returned the properties directly.
#[cfg(not(feature = "base"))]
macro_rules! ole_title {
    ($edit:ident) => {
        $edit.properties_mut()?.set("title", TITLE)?
    };
}
#[cfg(feature = "base")]
macro_rules! ole_title {
    ($edit:ident) => {
        $edit.properties_mut().set("title", TITLE)?
    };
}

macro_rules! owner {
    ($format:expr, $bytes:expr, $eager:expr, $ty:ty) => {
        Owner::<$ty> {
            format: $format,
            bytes: $bytes,
            eager: $eager,
            open: leak(|b: &[u8]| Ok(<$ty>::from_bytes(b)?)),
            save: leak(|p: &mut $ty| Ok(p.to_bytes()?)),
        }
    };
}

/// The `extra` cases of one format on its fixture bytes.
pub fn cases(format: &str, bytes: &'static [u8], fixture: usize) -> Result<Vec<Case<'static>>> {
    let mut cases = Vec::new();
    match format {
        #[cfg(feature = "docx")]
        "docx" => {
            let o = owner!("docx", bytes, true, litchi_docx::Package);
            cases.extend(owner_cases(o));
            // The eager unified open keeps the eager owner's validation text
            // pass (`Docx::from_owned_package_with_origin`); the paired leaf
            // open does the same leaf work, as the perf-baseline
            // `docx_file_leaf_eager_*` legs do (M0a step S8b).
            cases.push(case(
                "docx",
                "open_validated",
                "eager",
                Box::new(move || Ok(checked(Equal::NotApplicable))),
                Box::new(move || {
                    timed_ref(
                        || Ok(()),
                        |_| {
                            let package = litchi_docx::Package::from_bytes(bytes)?;
                            package.document()?.text()?;
                            Ok(package)
                        },
                    )
                }),
                Box::new(move || {
                    timed_ref(
                        || Ok(Some(bytes.to_vec())),
                        |b| {
                            Ok(litchi_ops::open_bytes_with(
                                b.take().ok_or("bytes")?,
                                &eager_options(),
                            )?)
                        },
                    )
                }),
            ));
            let leaf: LeafEdit<_> = leak(|p: &litchi_docx::Package| {
                let mut edit = p.edit_metadata();
                edit.update(|m| m.title = Some(TITLE.to_owned()));
                Ok(edit.commit()?.into_package())
            });
            let unified: UnifiedEdit = leak(|f: &File| Ok(f.set("/", title())?));
            cases.push(edit_case(o, "metadata_title_set", leaf, unified)?);
            cases.extend(save_cases(o, leaf, unified));
            let at: &'static usize = leak(fixture);
            cases.push(edit_case(
                o,
                "paragraph_add",
                leak(move |p: &litchi_docx::Package| {
                    let mut edit = p.edit()?;
                    edit.insert_paragraph(Position::new(*at), "added")?;
                    Ok(edit.commit()?.into_package())
                }),
                leak(|f: &File| {
                    Ok(f.add(
                        "/body",
                        litchi_ops::NewNode::new("paragraph").prop("text", "added"),
                    )?
                    .file)
                }),
            )?);
        },
        #[cfg(feature = "pptx")]
        "pptx" => {
            let o = owner!("pptx", bytes, true, litchi_pptx::Package);
            cases.extend(owner_cases(o));
            let leaf: LeafEdit<_> = leak(|p: &litchi_pptx::Package| {
                let mut edit = p.edit_metadata()?;
                edit.update(|m| m.title = Some(TITLE.to_owned()))?;
                Ok(edit.commit()?.into_package())
            });
            let unified: UnifiedEdit = leak(|f: &File| Ok(f.set("/metadata", title())?));
            cases.push(edit_case(o, "metadata_title_set", leaf, unified)?);
            cases.extend(save_cases(o, leaf, unified));
            // Structure edits on a 20-slide deck of five text boxes each.
            let deck: &'static [u8] = Box::leak(pptx_deck(20, 5)?.into_boxed_slice());
            let d = owner!("pptx", deck, true, litchi_pptx::Package);
            cases.push(edit_case(
                d,
                "slide_remove",
                leak(|p: &litchi_pptx::Package| {
                    let mut edit = p.edit()?;
                    edit.remove_slide(Position::new(9))?;
                    Ok(edit.commit()?.into_package())
                }),
                leak(|f: &File| Ok(f.remove("/slide[10]")?)),
            )?);
            cases.push(edit_case(
                d,
                "slide_move",
                leak(|p: &litchi_pptx::Package| {
                    let mut edit = p.edit()?;
                    edit.move_slide(Position::new(0), Position::new(1))?;
                    Ok(edit.commit()?.into_package())
                }),
                leak(|f: &File| {
                    Ok(f.move_to("/slide[1]", litchi_ops::Placement::after("/slide[2]")?)?)
                }),
            )?);
            cases.push(edit_case(
                d,
                "shape_remove",
                leak(|p: &litchi_pptx::Package| {
                    let mut edit = p.edit()?;
                    edit.remove_shape(9, 2_usize)?;
                    Ok(edit.commit()?.into_package())
                }),
                leak(|f: &File| Ok(f.remove("/slide[10]/shape[3]")?)),
            )?);
        },
        #[cfg(feature = "xlsx")]
        "xlsx" => {
            let o = owner!("xlsx", bytes, true, litchi_xlsx::Package);
            cases.extend(owner_cases(o));
            let leaf: LeafEdit<_> = leak(|p: &litchi_xlsx::Package| {
                let mut edit = p.edit_metadata()?;
                edit.update(|m| m.title = Some(TITLE.to_owned()))?;
                Ok(edit.commit()?.into_package())
            });
            let unified: UnifiedEdit = leak(|f: &File| Ok(f.set("/metadata", title())?));
            cases.push(edit_case(o, "metadata_title_set", leaf, unified)?);
            cases.extend(save_cases(o, leaf, unified));
        },
        #[cfg(feature = "xlsb")]
        "xlsb" => {
            let o = owner!("xlsb", bytes, true, litchi_xlsb::Package);
            cases.extend(owner_cases(o));
            let leaf: LeafEdit<_> = leak(|p: &litchi_xlsb::Package| {
                let mut edit = p.edit_metadata()?;
                edit.update(|m| m.title = Some(TITLE.to_owned()))?;
                Ok(edit.commit()?.into_package())
            });
            let unified: UnifiedEdit = leak(|f: &File| Ok(f.set("/metadata", title())?));
            cases.push(edit_case(o, "metadata_title_set", leaf, unified)?);
            cases.extend(save_cases(o, leaf, unified));
        },
        #[cfg(feature = "doc")]
        "doc" => {
            let o = owner!("doc", bytes, true, litchi_doc::Package);
            cases.extend(owner_cases(o));
            let leaf: LeafEdit<_> = leak(|p: &litchi_doc::Package| {
                let mut edit = p.edit_metadata()?;
                ole_title!(edit);
                Ok(edit.commit()?.into_package())
            });
            let unified: UnifiedEdit = leak(|f: &File| Ok(f.set("/metadata", title())?));
            cases.push(edit_case(o, "metadata_title_set", leaf, unified)?);
            cases.extend(save_cases(o, leaf, unified));
            let package: &'static litchi_doc::Package =
                leak(litchi_doc::Package::from_bytes(bytes)?);
            let file: &'static File = leak(unified_open(bytes, true)?);
            let leaf = || -> Result<Option<String>> {
                let document = package.document();
                Ok(document
                    .body_paragraph_text(Position::new(499))?
                    .map(str::to_owned))
            };
            let unified = || -> Result<Option<String>> {
                let node = file.get("/body/p[500]")?;
                Ok(node.view().text()?.map(std::borrow::Cow::into_owned))
            };
            cases.push(case(
                "doc",
                "paragraph_text_get",
                "eager",
                Box::new(move || {
                    if leaf()? != unified()? {
                        return Err("doc paragraph get: legs differ".into());
                    }
                    Ok(checked(Equal::Exact))
                }),
                Box::new(move || timed_ref(|| Ok(()), |_| leaf())),
                Box::new(move || timed_ref(|| Ok(()), |_| unified())),
            ));
        },
        #[cfg(feature = "xls")]
        "xls" => {
            let o = owner!("xls", bytes, true, litchi_xls::Package);
            cases.extend(owner_cases(o));
            let leaf: LeafEdit<_> = leak(|p: &litchi_xls::Package| {
                let mut edit = p.edit_metadata()?;
                ole_title!(edit);
                Ok(edit.commit()?.into_package())
            });
            let unified: UnifiedEdit = leak(|f: &File| Ok(f.set("/metadata", title())?));
            cases.push(edit_case(o, "metadata_title_set", leaf, unified)?);
            cases.extend(save_cases(o, leaf, unified));
        },
        #[cfg(feature = "ppt")]
        "ppt" => {
            let o = owner!("ppt", bytes, true, litchi_ppt::Package);
            cases.extend(owner_cases(o));
            let leaf: LeafEdit<_> = leak(|p: &litchi_ppt::Package| {
                let mut edit = p.edit_metadata()?;
                ole_title!(edit);
                Ok(edit.commit()?.into_package())
            });
            let unified: UnifiedEdit = leak(|f: &File| Ok(f.set("/metadata", title())?));
            cases.push(edit_case(o, "metadata_title_set", leaf, unified)?);
            cases.extend(save_cases(o, leaf, unified));
        },
        #[cfg(feature = "rtf")]
        "rtf" => cases.extend(owner_cases(owner!(
            "rtf",
            bytes,
            false,
            litchi_rtf::Package
        ))),
        #[cfg(feature = "odt")]
        "odt" => cases.extend(owner_cases(owner!(
            "odt",
            bytes,
            false,
            litchi_odt::Package
        ))),
        #[cfg(feature = "ods")]
        "ods" => cases.extend(owner_cases(owner!(
            "ods",
            bytes,
            false,
            litchi_ods::Package
        ))),
        #[cfg(feature = "odp")]
        "odp" => cases.extend(owner_cases(owner!(
            "odp",
            bytes,
            false,
            litchi_odp::Package
        ))),
        #[cfg(feature = "odg")]
        "odg" => cases.extend(owner_cases(owner!(
            "odg",
            bytes,
            false,
            litchi_odg::Package
        ))),
        #[cfg(feature = "odc")]
        "odc" => cases.extend(owner_cases(owner!(
            "odc",
            bytes,
            false,
            litchi_odc::Package
        ))),
        #[cfg(feature = "pages")]
        "pages" => cases.extend(owner_cases(owner!(
            "pages",
            bytes,
            false,
            litchi_pages::Package
        ))),
        #[cfg(feature = "keynote")]
        "keynote" => cases.extend(owner_cases(Owner::<litchi_keynote::Package> {
            format: "keynote",
            bytes,
            eager: false,
            open: leak(|b: &[u8]| Ok(litchi_keynote::Package::from_bytes(b)?)),
            save: leak(|p: &mut litchi_keynote::Package| {
                let mut out = Vec::new();
                p.write_to(&mut out)?;
                Ok(out)
            }),
        })),
        #[cfg(feature = "numbers")]
        "numbers" => cases.extend(owner_cases(Owner::<litchi_numbers::Package> {
            format: "numbers",
            bytes,
            eager: false,
            open: leak(|b: &[u8]| Ok(litchi_numbers::Package::from_bytes(b)?)),
            save: leak(|p: &mut litchi_numbers::Package| {
                let mut out = Vec::new();
                p.write_to(&mut out)?;
                Ok(out)
            }),
        })),
        _ => {},
    }
    Ok(cases)
}

/// A deck of `slides` slides with `shapes` text boxes each.
#[cfg(feature = "pptx")]
fn pptx_deck(slides: usize, shapes: usize) -> Result<Vec<u8>> {
    let mut builder = litchi_pptx::builder::Builder::new()?;
    for s in 0..slides {
        let slide = builder.presentation()?.add_slide()?;
        for i in 0..shapes {
            slide.add_text_box(
                &format!("slide {s} box {i}"),
                (i as i64) * 100,
                0,
                914_400,
                914_400,
            );
        }
    }
    #[allow(unused_mut, reason = "the baseline's `to_bytes` takes `&mut self`")]
    let mut package = builder.build()?;
    Ok(package.to_bytes()?)
}
