//! PresentationML package-graph adapters for slides, layouts, and masters.
//!
//! Relationship cardinality, internal-target checks, content-type checks,
//! strict/transitional relationship families, and XML-owned layout ordering
//! remain enforced by the validated [`crate::parts::SlidePart`] family and
//! shared part helpers. These adapters add only contextual companion-part
//! resolution; they never materialize or reinterpret unknown DrawingML.

use litchi_opc::OpcPackage;
use litchi_opc::constants::{content_type as ct, relationship_type as rt};

use crate::Result;
use crate::parts::{SlideLayoutPart, SlideMasterPart, SlidePart, related_part_by_type};

pub(crate) fn slide_tags<'a>(
    package: &'a OpcPackage,
    part: &SlidePart<'a>,
) -> Result<Option<crate::tag::List>> {
    Ok(crate::tag::load(package, part.part().partname())?.map(crate::tag::Source::into_list))
}

pub(crate) fn slide_shape_tags<'a, 'k>(
    package: &'a OpcPackage,
    part: &SlidePart<'a>,
    shape: impl Into<crate::shape::Key<'k>>,
) -> Result<Option<crate::tag::List>> {
    Ok(
        crate::tag::shape::load(package, part.part().partname(), shape)?
            .map(crate::tag::Source::into_list),
    )
}

pub(crate) fn slide_tag_inventory<'a>(
    package: &'a OpcPackage,
    part: &SlidePart<'a>,
) -> Result<Vec<crate::tag::Source>> {
    crate::tag::discover(part.part(), package).map_err(Into::into)
}

pub(crate) fn slide_charts<'a>(
    package: &'a OpcPackage,
    part: &SlidePart<'a>,
) -> Result<Vec<crate::chart::Part<'a>>> {
    part.charts(package)
}

pub(crate) fn slide_chart_extensions<'a>(
    package: &'a OpcPackage,
    part: &SlidePart<'a>,
) -> Result<Vec<crate::chart::extension::Part<'a>>> {
    part.chart_extensions(package)
}

pub(crate) fn slide_comments<'a>(
    package: &'a OpcPackage,
    part: &SlidePart<'a>,
) -> Result<Option<crate::comments::ListPart<'a>>> {
    part.comments(package)
}

pub(crate) fn slide_zooms<'a>(
    package: &'a OpcPackage,
    part: &SlidePart<'a>,
) -> Result<crate::shape::zoom::Owner> {
    crate::shape::zoom::package::load_part(package, part)
}

pub(crate) fn slide_sync<'a>(
    package: &'a OpcPackage,
    part: &SlidePart<'a>,
) -> Result<Option<crate::presentation_properties::metadata::slide_sync::Properties>> {
    let part_name = part.part().partname();
    let mut matches = crate::presentation_properties::metadata::slide_sync::load(package)?
        .into_iter()
        .filter(|entry| entry.slide_part_name == *part_name);
    Ok(matches.next().map(|entry| entry.properties))
}

pub(crate) fn slide_layout<'a>(
    package: &'a OpcPackage,
    part: &SlidePart<'a>,
) -> Result<Option<SlideLayoutPart<'a>>> {
    related_part_by_type(
        package,
        part.part(),
        rt::SLIDE_LAYOUT,
        "slideLayout",
        ct::PML_SLIDE_LAYOUT,
    )?
    .map(SlideLayoutPart::from_part)
    .transpose()
}

pub(crate) fn layout_theme_override<'a>(
    package: &'a OpcPackage,
    part: &SlideLayoutPart<'a>,
) -> Result<Option<crate::shape::theme::Override>> {
    crate::shape::theme::package::load_override(package, part.part().partname().as_str())
}

pub(crate) fn layout_master<'a>(
    package: &'a OpcPackage,
    part: &SlideLayoutPart<'a>,
) -> Result<SlideMasterPart<'a>> {
    part.master(package)
}

pub(crate) fn master_theme<'a>(
    package: &'a OpcPackage,
    part: &SlideMasterPart<'a>,
) -> Result<Option<crate::shape::theme::ThemeSummary>> {
    part.theme(package)
}

pub(crate) fn master_layouts<'a>(
    package: &'a OpcPackage,
    part: &SlideMasterPart<'a>,
) -> Result<Vec<SlideLayoutPart<'a>>> {
    part.layouts(package)
}
