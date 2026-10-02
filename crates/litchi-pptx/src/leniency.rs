//! Repair reporting for non-structural defects in `PresentationML` packages.
//!
//! Two kinds of repair can happen while a presentation is read, and both are
//! *recorded*, so a caller can enumerate exactly what changed through
//! [`ToleranceReport`]. Nothing is swallowed silently, and nothing outside
//! the closed enumerations changes: ZIP and XML well-formedness, OPC package
//! grammar, bounded resource limits, and encryption remain hard errors under
//! every variant.
//!
//! **`PowerPoint`-compatible repairs** are applied under every [`Leniency`]
//! variant, including the default [`Leniency::Strict`]. They cover markup
//! ECMA-376 leaves unconstrained but real producers emit anyway — markup
//! `PowerPoint` and LibreOffice read past without complaint, so rejecting it
//! would cost the caller the whole slide or timing tree over metadata that
//! cannot change what a conforming deck means. Each substitutes the value
//! documented on the corresponding defect variant ([`TimingDefect`],
//! [`ShapeDefect`]). Every `PresentationML` repair implemented so far is in
//! this class; the first members are the timing-tree cross-reference repairs
//! of audit finding 6.1 (orphaned animation/build targets and duplicated
//! shape IDs).
//!
//! **Opt-in repairs** trade strict conformance for reach and stay gated by
//! [`Leniency`] variants. No `PresentationML` repair currently needs that
//! judgment call, so [`Leniency::Strict`] is the only variant; the enum is
//! where such a gate would appear.
//!
//! # Where the report lives
//!
//! This crate parses lazily: opening a package validates the OPC catalog and
//! the presentation root, and slide payloads — including the `p:timing`
//! animation tree — parse only when a caller asks for them, through `&self`
//! views and free entry points such as [`crate::animations::Sequence`]'s XML
//! parsers. The report is therefore aggregated at two levels:
//!
//! * The owned document handles, [`crate::Package`] and
//!   [`crate::SourceBackedPresentation`], carry the report assembled while
//!   the package was opened and expose it as `tolerance_report()`. Opening
//!   reads the presentation slide list, so its [`PresentationDefect`]
//!   repairs (audit finding 6.7) are recorded here, stamped with the
//!   presentation part name. For [`crate::Package`] the open-time catalog
//!   read is advisory: a slide list it cannot read stays the lazy
//!   accessors' error, exactly as before.
//! * The packaging and metadata reports of the lower layers are folded in
//!   rather than kept alongside. The OPC report (`litchi_opc`'s
//!   relationship and content-type repairs, audit findings 8.5–8.9) leads
//!   both open-time reports as [`Defect::Package`]. [`crate::Package`] also
//!   reads the core and custom properties at open — advisorily, a part they
//!   reject stays its accessor's error — and appends their repairs
//!   (`litchi_ooxml_common`, findings 8.1–8.4) as [`Defect::Properties`].
//!   [`crate::SourceBackedPresentation`] reads only the presentation root at
//!   open, so its core-properties repairs are collected by
//!   [`crate::SourceBackedPresentation::properties_with_tolerance`]. A folded
//!   entry keeps its lower-level location: [`ToleratedDefect::part`] plus the
//!   [`ToleratedDefect::item`] and [`ToleratedDefect::value`] strings, and
//!   the lower report's unrecorded count carries over.
//! * A lazy parse carries the repairs it performed itself, the way
//!   `PapBinTable::tolerance_report` lets litchi-doc aggregate part reports
//!   after the fact. [`crate::Slide::animations`] (and the free
//!   [`crate::animations::parse_package_slide`]) returns a
//!   [`crate::animations::Sequence`] whose `tolerance_report()` lists every
//!   finding-6.1 repair, each stamped with the slide's part name
//!   ([`ToleratedDefect::part`]) and the shape ID concerned
//!   ([`ToleratedDefect::context`]). The parse entry points have
//!   `*_with_tolerance` companions taking `&mut ToleranceReport` (explicit
//!   threading, no interior mutability) so a caller can merge per-slide
//!   reports into one deck-level aggregate.
//! * A shape scene ([`crate::shape::Scene`], returned by
//!   [`crate::Slide::shapes`] and the layout/master `shapes()` views) carries
//!   its per-shape repairs in [`crate::shape::Scene::tolerance_report`],
//!   stamped with the owner part name when read through a part view.
//! * A slide transition ([`crate::transition::Transition`], from
//!   [`crate::transition::read`] or
//!   [`crate::SourceBackedSlideSnapshot::transition`]) carries the repairs
//!   its read applied in
//!   [`crate::transition::Transition::tolerance_report`]; the snapshot
//!   stamps its slide part name. The report shares the value's retained
//!   read-state allocation, so a conforming transition stays one cache line.
//!
//! # Timing-tree repairs (audit finding 6.1)
//!
//! [`crate::animations::Sequence::parse_slide_xml`] — the read default
//! behind every package-aware entry point — no longer fails the whole slide
//! timing tree over a stale cross-reference:
//!
//! * an animation, interactive trigger, or build-list entry (`p:bldP`,
//!   `p:bldDgm`, `p:bldGraphic`, `p:bldOleChart`) whose `spid` names no shape
//!   on the slide, or a shape unable to host it, is skipped and recorded as
//!   [`TimingDefect::OrphanedTarget`]; the rest of the tree is kept, click
//!   order is renumbered contiguously, and a build group (`grpId`) a skipped
//!   entry carried is left out of the effect/build group-pairing check, whose
//!   pairing the stale reference itself broke. The verbatim timing
//!   subtree is not preserved for a repaired slide, because it still contains
//!   the skipped entries; serialization regenerates from the repaired model;
//! * a `p:cNvPr@id` repeated on the slide resolves to the first shape
//!   carrying it and is recorded as [`TimingDefect::DuplicateShapeId`] once
//!   per repeat; the package-relationship host scan applies the same
//!   first-wins rule.
//!
//! XML well-formedness, the timing-tree grammar (balanced nodes, required
//! attributes, valid token values, duplicate build pairs), the size/depth/
//! count limits, and host-relationship validation of the surviving build
//! entries stay hard errors. Callers that must prove a timing tree is fully
//! self-consistent before a lossless rewrite keep the old contract through
//! [`crate::animations::Sequence::parse_slide_xml_strict`], which rejects
//! any slide whose parse needed a repair.
//!
//! # Shape-scene repairs (audit finding 6.2)
//!
//! [`crate::shape::Scene::read`] — behind `shapes()` on slides, layouts, and
//! masters — no longer fails every shape on the owner over one malformed
//! transform. Only an `a:off`/`a:ext` whose parent is the shape's own
//! `a:xfrm` (or a graphic frame's `p:xfrm`) is read as geometry; when the
//! first offset lacks `x`/`y`, the first extent lacks `cx`/`cy`, a value is
//! not an integer, or an extent is negative, that shape keeps its place,
//! identity, placeholder, and text but reports no
//! [`crate::shape::Bounds`], and the repair is recorded as
//! [`ShapeDefect::MalformedTransform`] with the shape ID. The reader drops
//! the geometry rather than substituting LibreOffice's `0` defaults, because
//! a `0` position or size would be content the producer never wrote; absent
//! geometry is already how the scene represents a shape that inherits its
//! placement. XML well-formedness, the scene's structural grammar, and its
//! size/depth/count limits stay hard errors.
//!
//! # Placeholder extension repairs (audit finding 6.4)
//!
//! A placeholder's `p:extLst` no longer fails the scene over extension
//! arity. `CT_Extension` is `xsd:any processContents="lax"` with any number
//! of children, so a `p:ext` whose URI is not the p232 type-extension owner
//! is skipped without inspection — zero or many children are legal and
//! record nothing. When such an opaque owner itself breaks `CT_Extension`
//! (missing or blank `uri`, an extra attribute, direct character data) it is
//! still skipped, and recorded once as
//! [`ShapeDefect::PlaceholderExtensionOwner`]. The strict grammar applies
//! only to the p232 owner URI, and even there a violation (extra owner
//! children or attributes, character data, a missing or duplicated token, a
//! repeated `p:extLst`) degrades to "no type extension" for that
//! placeholder, recorded as [`ShapeDefect::PlaceholderTypeExtension`]; the
//! placeholder, shape, and scene survive. The lossless placeholder-type
//! *edit* planner in `master_layout` keeps its strict owner validation.
//!
//! # Transition repairs (audit finding 6.5)
//!
//! A wheel (`p:wheel`) or reverse wheel (`p14:wheelReverse`) whose `spokes`
//! is an `xsd:unsignedInt` outside {1, 2, 3, 4, 8} no longer fails the
//! transition. [`crate::transition::Spokes`] cannot represent the count, so
//! the effect element is kept verbatim as [`crate::transition::Kind::Raw`]
//! — the posture the reader already applies to an effect it does not know,
//! and lossless on write — while speed, timing, and sibling children read
//! normally; the repair is recorded as [`TransitionDefect::SpokeCount`]
//! with the count. Mapping the count to the nearest supported value (as
//! LibreOffice's renderer does) was rejected because it would silently
//! rewrite the producer's value on save. A `spokes` that is not an
//! `unsignedInt` at all is not this repair. The source-backed transition
//! edit planner still refuses to splice over a transition whose read needed
//! a repair.
//!
//! # Transition token domains (audit finding 6.6)
//!
//! An attribute value outside its domain no longer fails the whole
//! transition. On a known effect element (`dir`, `orient`, `pattern`, morph
//! `option` — unknown or missing — boolean flags, preset names) the effect
//! is kept verbatim as [`crate::transition::Kind::Raw`] and recorded as
//! [`TransitionDefect::EffectToken`]: the crate already keeps unknown effect
//! *elements* raw, and this extends that posture one level down without
//! inventing a direction the producer never wrote. On `p:transition` itself
//! an unknown `spd`, a non-boolean `advClick`, or an unparseable `advTm`,
//! `dur`, or `p14:dur` is treated as absent (LibreOffice's attribute
//! defaults) and recorded as [`TransitionDefect::AttributeToken`]. An
//! integral `advTm`/`dur` above [`crate::transition::MAX_MS`] — the
//! `2^31..2^32` range is schema-valid `xsd:unsignedInt` — saturates to that
//! maximum and is recorded as [`TransitionDefect::TimeClamped`] with the
//! original value. Duplicate attributes, nested markup inside typed
//! effects, XML well-formedness, and the read limits stay hard errors.
//!
//! # Slide-list repairs (audit finding 6.7)
//!
//! [`crate::parts::PresentationPart::slide_references`] — behind
//! [`crate::Presentation::slides`] and the source-backed catalog — keeps the
//! first `p:sldId` carrying a slide ID, relationship ID, or target part and
//! skips a later repeat ([`PresentationDefect::DuplicateSlideId`],
//! [`PresentationDefect::DuplicateSlideRelationship`],
//! [`PresentationDefect::DuplicateSlideTarget`], context = the skipped
//! entry's slide ID), mirroring LibreOffice's tolerant
//! `presentationfragmenthandler` import. Foreign children (their subtrees are
//! skipped), attributes, and character data inside `p:sldIdLst` are ignored
//! and recorded as [`PresentationDefect::SlideListContent`]. An entry
//! missing `id`/`r:id`, an ID outside `ST_SlideId` (256..2^31), a stray
//! `p:sldId` outside the list, repeated lists, XML well-formedness, and the
//! slide-count limit stay hard errors. The lossless edit planners —
//! [`crate::SourceBackedPresentationEditor`], cross-presentation copy, and
//! the opened-presentation transactions — keep the strict contract and
//! refuse a slide list that needed a repair.
//!
//! # Processing instructions (audit finding 6.3)
//!
//! The slide-owner read surfaces skip processing instructions instead of
//! failing the part: the shape scene, semantic slide text, the timing
//! host scan, the layout-ID scan, and typed transition effects. The repair
//! is recorded once per part, as [`MarkupDefect::ProcessingInstruction`],
//! by the part's shape scene ([`crate::shape::Scene::tolerance_report`]),
//! so a caller enumerating `shapes()` reports sees every part that carried
//! one; the other readers of the same part skip them consistently rather
//! than record duplicates. DOCTYPE declarations stay rejected everywhere.
//!
//! The feature readers that carry a report skip them too — removed from the
//! semantic input before markup-compatibility processing, which itself
//! refuses them — and record the skip once per part:
//!
//! * view properties
//!   ([`crate::view_properties::load_from_package_with_tolerance`],
//!   [`crate::view_properties::ViewProperties::parse_with_tolerance`]);
//! * presentation properties
//!   ([`crate::presentation_properties::load_from_package_with_tolerance`],
//!   [`crate::presentation_properties::Properties::parse_with_tolerance`]);
//!   the math edit transactions load through
//!   [`crate::presentation_properties::Properties::parse_strict`] and refuse
//!   such a part;
//! * the chart inspection ([`crate::chart::Info::tolerance_report`]), which
//!   previously ignored them without a record;
//! * legacy comments and comment authors
//!   ([`crate::comments::load_presentation_comments_with_tolerance`],
//!   [`crate::comments::parse_slide_comments_with_tolerance`]), which
//!   previously stripped them without a record. The strict comment load
//!   behind the comment edits keeps accepting them, as before: the edits
//!   regenerate the part from the model.
//!
//! The remaining rejection sites stay strict, for one of two reasons.
//!
//! *Source-bound planners and span scanners* account for every byte of the
//! source they splice, patch, or prove unchanged, so an unmodelled
//! instruction must stay a refusal: the opened-presentation transactions
//! (`opened/xml.rs`, `opened/copy_plan.rs`, `opened/remove_plan.rs`),
//! cross-presentation copy (`presentation/source_cross_copy.rs`), slide
//! reordering (`presentation/order.rs`), source-backed picture, SVG, and
//! slide-root queries (`presentation/source.rs`,
//! `presentation/source/svg_lifecycle*`), master/layout edits
//! (`master_layout/codec.rs`, `master_layout/package.rs`), fragment
//! insertion for media, laser traces, ink, OLE frames, and content parts
//! (`media_parts/package.rs`, `laser/package.rs`,
//! `presentation/embedded/{ink/package.rs, ole/authoring.rs,
//! content_parts/codec.rs}`, `presentation_properties/metadata/slide_patch`),
//! font-embedding edits (`font/package.rs`), tag owner maps and shape spans
//! (`tag/package/codec.rs`, `tag/shape/codec.rs`, `table/shape.rs` span
//! scan), designer and classification sources (`shape/designer/*`,
//! `shape/classification/codec.rs`), change tracking
//! (`change_tracking/codec.rs`), the source-bound metadata layouts
//! (`comments/collaboration/codec.rs`, `presentation_properties/
//! {readonly_recommended, package.rs root-dialect probe}`,
//! `presentation_properties/metadata/{guides, events, custom_show,
//! structure layout}`, `modern_comments/wire/*`), the transition and
//! animation-template validators used by edits
//! (`presentation/transition.rs`, `animations/codec/validation.rs`), and
//! the package validator (`validation.rs`).
//!
//! *Feature readers whose values have no report carrier yet* keep rejecting
//! rather than skip silently: media parts (`media_parts/codec.rs`), table
//! properties, text, and styles (`table/shape.rs`, `table/style/*`), tags
//! (`tag/codec.rs`), notes (`notes/codec.rs`), embedded-font lists
//! (`font/codec.rs`), laser traces (`laser/codec.rs`), 3D models
//! (`model3d/codec.rs`), zoom (`shape/zoom/codec.rs`), `ChartEx`
//! (`chart/extension/codec/xml.rs`), modern comments
//! (`modern_comments/codec*`), presentation metadata parts
//! (`presentation_properties/metadata/{revision, changes, structure,
//! sections, designer_tags, tracks, protection}`,
//! `presentation_properties/math/codec.rs`), and embedded ink, ink actions,
//! controls, OLE objects, and content-part scans
//! (`presentation/embedded/*`). Each needs a report on its returned value
//! (or a `*_with_tolerance` companion) before it can skip them.
//!
//! # Semantic slide text (audit finding 6.9)
//!
//! Slide text extraction ([`crate::Slide::text`], the source-backed slide
//! text, and the presentation text sink) no longer fails over markup that
//! carries no DrawingML text. An element named `t` outside the DrawingML
//! namespace (`p:t`, an extension's `x:t`) is skipped with its subtree and
//! recorded as [`MarkupDefect::ForeignTextElement`]; the `*_with_tolerance`
//! text companions ([`crate::Slide::text_with_tolerance`],
//! [`crate::parts::SlidePart::text_with_tolerance`],
//! [`crate::SourceSlide::text_with_tolerance`]) collect it, stamped with the
//! slide part. CDATA sections and entity references outside `a:t` are
//! ordinary XML and are validated and ignored like a text node there —
//! nothing is repaired, so nothing is recorded. XML well-formedness,
//! undefined entities, character data outside the document root, nested or
//! unbalanced DrawingML `a:t`, and the size/depth/count limits stay hard
//! errors.
//!
//! # Chart inspection (audit finding 6.10)
//!
//! [`crate::chart::Part::chart_info`] no longer fails the inspection over
//! chart structure LibreOffice reads past. A missing or unknown `c:barDir`
//! value reads as a column chart ([`ChartDefect::BarDirection`]); a repeated
//! `c:chart`, `c:title`, `c:legend`, or `c:plotArea` keeps the first
//! occurrence ([`ChartDefect::DuplicateElement`]); a chart without
//! `c:chart` or `c:plotArea` reports [`crate::chart::Type::Unknown`]
//! ([`ChartDefect::MissingElement`]). The repairs are carried by the
//! returned [`crate::chart::Info::tolerance_report`], stamped with the chart
//! part. XML well-formedness, a missing or foreign `c:chartSpace` root, and
//! nested title text stay hard errors.
//!
//! # View properties (audit finding 6.11)
//!
//! [`crate::view_properties::ViewProperties::parse`] and
//! [`crate::view_properties::load_from_package`] no longer fail
//! `viewProps.xml` over an attribute or child their model does not know
//! ([`ViewDefect::UnknownAttribute`], [`ViewDefect::UnknownElement`]; a new
//! Office release adding one attribute would otherwise make every file it
//! saves unreadable), over children out of schema order (read by name,
//! [`ViewDefect::ElementOrder`]), or over a repeated child (first wins,
//! [`ViewDefect::DuplicateElement`]). The `*_with_tolerance` companions
//! collect the repairs, stamped with the view-properties part by
//! [`crate::view_properties::load_from_package_with_tolerance`]. A required
//! child that is missing, an invalid value (boolean, integer, enumeration,
//! zero scale denominator), text in typed content, XML well-formedness, and
//! the size/depth/count limits stay hard errors.
//! [`crate::view_properties::ViewProperties::parse_strict`] keeps the old
//! contract and backs the slide-order edit planner's metadata inspection.
//!
//! # Shape metadata (audit finding 6.8)
//!
//! A `p:cNvPr@id` that is not an `unsignedInt` leaves the shape without an
//! ID ([`ShapeDefect::InvalidShapeId`]), and an unparseable `p:ph@idx`
//! falls back to the schema default `0`
//! ([`ShapeDefect::InvalidPlaceholderIndex`]); the scene, the shape, and its
//! siblings are kept.
//!
//! # Legacy comments (audit finding 6.12)
//!
//! [`crate::comments::load_presentation_comments`] and
//! [`crate::comments::parse_slide_comments`] no longer fail the comment
//! graph over a comment with `idx="0"` ([`CommentDefect::ZeroIndex`]) or a
//! comment repeating an earlier `(authorId, idx)` key on the same slide or
//! across the presentation ([`CommentDefect::DuplicateKey`]). Both keep the
//! comment, in document order — LibreOffice imports every `p:cm` as written
//! — so no comment is lost on read. [`crate::comments::Comments`] stays a
//! plain public-field value (callers build it to store a graph), so the
//! report travels beside it:
//! [`crate::comments::load_presentation_comments_with_tolerance`] and
//! [`crate::comments::parse_slide_comments_with_tolerance`] append the
//! repairs (context = author ID, [`ToleratedDefect::item`] = index, part =
//! comment part). Every comment *edit* function loads through
//! [`crate::comments::load_presentation_comments_strict`] and refuses a
//! graph that needed a repair, so an edit never rewrites — or silently
//! drops — such comments; writers validate strictly as before. Unknown
//! attributes or children, non-numeric indexes, invalid dates, and the
//! author-reference and `lastIdx` checks stay hard errors.
//!
//! This mirrors the contract `crates/litchi-doc/src/leniency.rs` establishes
//! for legacy Word documents so the readers behave the same way.

use std::sync::Arc;

/// Upper bound on individually recorded defects.
///
/// A hostile or merely enormous presentation must not be able to make the
/// report grow without limit. Defects beyond this bound are counted, not
/// stored; see [`ToleranceReport::unrecorded`].
const MAX_RECORDED_DEFECTS: usize = 1024;

/// How the reader treats non-structural presentation defects.
///
/// Structural defects — malformed ZIP or XML framing, a broken OPC package
/// grammar, unsupported/absent decryption material, or an exceeded resource
/// limit — are hard errors under every variant. The `PowerPoint`-compatible
/// repairs listed by [`TimingDefect`] are applied and recorded under every
/// variant, including [`Leniency::Strict`].
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum Leniency {
    /// Reject every defect whose repair requires opting in.
    ///
    /// This is the default and currently the only variant: every
    /// `PresentationML` repair on the leniency roadmap is
    /// `PowerPoint`-compatible, so the [`TimingDefect`] repairs are applied
    /// and recorded here too, and no opt-in gate exists yet.
    #[default]
    Strict,
}

/// The class of slide timing-tree defect the reader repairs the way
/// `PowerPoint` repairs it.
///
/// This enumeration is closed by design: it is the exhaustive contract of the
/// repairs a read applies under every [`Leniency`] variant, including
/// [`Leniency::Strict`]. Each one matches what `PowerPoint` and LibreOffice
/// do with markup ECMA-376 does not constrain, so applying it never changes
/// what a conforming presentation means.
///
/// The variants are the recording contract of leniency-audit finding 6.1: the
/// slide timing parser records them, and the repairs above are applied.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum TimingDefect {
    /// An animation, an interactive-sequence trigger, or a build-list entry
    /// targeted a shape ID that no shape on the slide carries, or whose shape
    /// cannot host the entry (for example a chart build aimed at a picture).
    ///
    /// ECMA-376-1 §19.5 (`p:spTgt`, `p:bldP`) defines no cross-reference
    /// validity requirement: orphaned timing entries are representable, and
    /// `PowerPoint` routinely leaves them behind when shapes are deleted.
    /// `PowerPoint` and LibreOffice skip the dead entries and play the rest,
    /// so the reader skips the orphaned animation or build entry and keeps
    /// the rest of the timing tree.
    OrphanedTarget,
    /// Two non-visual shape property elements (`p:cNvPr`) on one slide carry
    /// the same shape ID.
    ///
    /// `ST_DrawingElementId` uniqueness is a schema expectation, not a rule
    /// consumers enforce; duplicate IDs survive copy/paste and third-party
    /// round-trips. `PowerPoint` resolves the ID to the first shape carrying
    /// it, so the reader keeps the first occurrence: target resolution sees
    /// the ID once, as the first shape.
    DuplicateShapeId,
}

impl TimingDefect {
    /// A short, stable description of what the repair did.
    #[must_use]
    pub const fn repair(self) -> &'static str {
        match self {
            Self::OrphanedTarget => {
                "skipped the orphaned animation or build entry; the rest of the timing tree was kept"
            },
            Self::DuplicateShapeId => {
                "kept the first shape carrying the duplicated ID; later duplicates were ignored for target resolution"
            },
        }
    }
}

/// One class of defect repaired while a presentation was read.
///
/// This enumeration is closed by design: it is the exhaustive contract of
/// what a read may change about a presentation. [`Defect::Timing`] classes
/// are repaired under every [`Leniency`] variant.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Defect {
    /// A slide timing-tree defect repaired the way `PowerPoint` repairs it.
    Timing(TimingDefect),
    /// A shape-tree defect repaired on the one shape carrying it.
    Shape(ShapeDefect),
    /// A slide-transition defect repaired on the one transition carrying it.
    Transition(TransitionDefect),
    /// A presentation slide-list defect repaired when the package is opened.
    Presentation(PresentationDefect),
    /// An XML-markup defect skipped by the reader.
    Markup(MarkupDefect),
    /// An OPC packaging defect (relationship or content-types manifest)
    /// repaired by `litchi_opc` while the package was opened, folded into
    /// the presentation's open-time report.
    Package(litchi_opc::leniency::Defect),
    /// A legacy slide-comment defect repaired on the comment carrying it.
    Comment(CommentDefect),
    /// A chart-part defect repaired by the chart inspection.
    Chart(ChartDefect),
    /// A view-properties defect repaired by the view-properties reader.
    View(ViewDefect),
    /// A core- or custom-properties defect repaired by
    /// `litchi_ooxml_common` while the document metadata was read, folded
    /// into the presentation's report.
    Properties(litchi_ooxml_common::leniency::Defect),
}

impl Defect {
    /// A short, stable description of what the repair did.
    #[must_use]
    pub const fn repair(self) -> &'static str {
        match self {
            Self::Timing(defect) => defect.repair(),
            Self::Shape(defect) => defect.repair(),
            Self::Transition(defect) => defect.repair(),
            Self::Presentation(defect) => defect.repair(),
            Self::Markup(defect) => defect.repair(),
            Self::Package(defect) => defect.repair(),
            Self::Comment(defect) => defect.repair(),
            Self::Chart(defect) => defect.repair(),
            Self::View(defect) => defect.repair(),
            Self::Properties(defect) => defect.repair(),
        }
    }
}

/// The class of shape-tree defect the scene reader repairs per shape.
///
/// Like [`TimingDefect`], this enumeration is closed by design and every
/// variant is repaired under every [`Leniency`] variant. Each repair is
/// confined to the one shape that carries the defect: the shape keeps its
/// place, name, ID, placeholder, and text in the [`crate::shape::Scene`],
/// and its well-formed siblings are unaffected.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum ShapeDefect {
    /// The shape's transform (`a:xfrm`/`p:xfrm`) carried an offset
    /// (`a:off`) without `x`/`y` or with a non-integer coordinate, or an
    /// extent (`a:ext`) without `cx`/`cy`, with a non-integer value, or
    /// with a negative one.
    ///
    /// ECMA-376-1 §20.1.7.3/§20.1.7.4 type these as `CT_Point2D` and
    /// `CT_PositiveSize2D`, so the markup is non-conformant, but nothing
    /// mandates refusing the subtree: LibreOffice's `oox` transform context
    /// reads them with `getInteger(…, 0)` defaults and `PowerPoint` opens
    /// such decks. Substituting `0` would invent a position the producer
    /// never wrote, so the reader degrades only this shape's geometry to
    /// absent ([`crate::shape::Shape::bounds`] is `None`, as for a shape that
    /// inherits its placement) and keeps the rest of the scene.
    MalformedTransform,
    /// A placeholder extension owner (`p:ph/p:extLst/p:ext`) that this
    /// reader does not model broke `CT_Extension`'s own rules: its `uri` was
    /// missing or blank, it carried another non-namespace attribute, or it
    /// held direct character data.
    ///
    /// `CT_Extension` content is `xsd:any processContents="lax"` with any
    /// arity, and the Markup Compatibility model requires unrecognized
    /// extension content to be ignorable; `PowerPoint` ignores extensions it
    /// does not understand. The reader skips the owner without inspecting
    /// its subtree, exactly as it skips a conforming unknown extension.
    PlaceholderExtensionOwner,
    /// The placeholder's p232 type extension (`p:ext` with the
    /// `urn:litchi:pptx:p232:phTypeExt` owner URI wrapping
    /// `p232:phTypeExt/p232:type/{cameo|unknown}`) broke its grammar: extra
    /// owner attributes or children, character data, a missing or duplicated
    /// token, attributes on the typed elements, or a repeated `p:extLst`.
    ///
    /// The placeholder itself is well defined without its type extension,
    /// so the reader keeps the placeholder and reports no type extension
    /// ([`crate::shape::Placeholder::type_extension`] is `None`), skipping
    /// the rest of its extension list.
    PlaceholderTypeExtension,
    /// A non-visual shape ID (`p:cNvPr@id`) was not an `unsignedInt`
    /// (`ST_DrawingElementId`), for example empty or 64-bit.
    ///
    /// The ID is non-structural metadata the scene already models as
    /// optional; `PowerPoint` re-assigns broken IDs on load and LibreOffice
    /// never parses them eagerly. The reader reports the shape with no ID
    /// ([`crate::shape::Shape::id`] is `None`). The context is `0`, since
    /// the shape has no usable ID.
    InvalidShapeId,
    /// A placeholder index (`p:ph@idx`) was not an `unsignedInt`.
    ///
    /// The reader applies the schema default `0`, as for an omitted `idx`.
    /// The context is the shape ID (or `0`).
    InvalidPlaceholderIndex,
}

impl ShapeDefect {
    /// A short, stable description of what the repair did.
    #[must_use]
    pub const fn repair(self) -> &'static str {
        match self {
            Self::MalformedTransform => {
                "dropped the shape's malformed transform; its geometry is absent and the rest of the scene was kept"
            },
            Self::PlaceholderExtensionOwner => {
                "skipped the malformed placeholder extension owner without inspecting it"
            },
            Self::PlaceholderTypeExtension => {
                "dropped the placeholder's malformed p232 type extension; the placeholder was kept"
            },
            Self::InvalidShapeId => "reported the shape with no ID in place of its unparseable ID",
            Self::InvalidPlaceholderIndex => {
                "used the default placeholder index 0 in place of the unparseable index"
            },
        }
    }
}

/// The class of slide-transition defect the transition reader repairs.
///
/// Closed by design, and repaired under every [`Leniency`] variant. The
/// transition stays readable: an effect `PowerPoint` cannot model is kept
/// verbatim as [`crate::transition::Kind::Raw`] (so a write reproduces the
/// producer's bytes), the same posture the reader already applies to an
/// effect element it does not know.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum TransitionDefect {
    /// A wheel (`p:wheel`) or reverse-wheel (`p14:wheelReverse`) effect
    /// carried a `spokes` count outside the {1, 2, 3, 4, 8} that
    /// `PowerPoint` offers.
    ///
    /// ECMA-376-1 `CT_WheelTransition` types `spokes` as plain
    /// `xsd:unsignedInt`, so every such count is schema-valid; LibreOffice
    /// imports it with `getUnsigned(XML_spokes, 4)` and never fails. The
    /// typed [`crate::transition::Spokes`] cannot represent the count, and
    /// rounding it to a supported one would rewrite the producer's value on
    /// save, so the reader keeps the effect element verbatim as
    /// [`crate::transition::Kind::Raw`]. The context is the spoke count.
    SpokeCount,
    /// A known effect element carried an attribute value outside its
    /// enumeration or lexical domain: a `dir`/`orient` token the effect does
    /// not define, a glitter/shred `pattern`, a morph `option` that is
    /// unknown or missing, a non-boolean flag such as `thruBlk`, or a preset
    /// name the reader cannot carry.
    ///
    /// These are schema enumerations, so the markup is non-conformant, but
    /// LibreOffice defaults unknown tokens (`getToken(XML_dir, …)`) and
    /// `PowerPoint` repairs them. The typed effect cannot hold the value, so
    /// the reader keeps the element verbatim as
    /// [`crate::transition::Kind::Raw`], exactly as it keeps an effect
    /// element it does not know.
    EffectToken,
    /// An attribute of `p:transition` itself carried a value outside its
    /// domain: an unknown `spd` token, a non-boolean `advClick`, or an
    /// `advTm`, `dur`, or `p14:dur` that is not a time value.
    ///
    /// LibreOffice reads these with defaults (`getToken(XML_spd, …)`,
    /// `getBool(XML_advClick, true)`, `getInteger(XML_advTm, -1)`). The
    /// reader treats the attribute as absent: speed [`Speed::Medium`],
    /// advance on click, no automatic advance, no explicit duration.
    ///
    /// [`Speed::Medium`]: crate::transition::Speed::Medium
    AttributeToken,
    /// An integral transition time (`advTm`, or the legacy unqualified
    /// `dur`) exceeded [`crate::transition::MAX_MS`] (`i32::MAX`).
    ///
    /// `advTm` is `xsd:unsignedInt`, so values up to `u32::MAX` are
    /// schema-valid even though `PowerPoint` documents the narrower range.
    /// The reader saturates the time to [`crate::transition::MAX_MS`] — the
    /// longest delay `PowerPoint` can play — and the context is the original
    /// value (saturated to `u32::MAX`).
    TimeClamped,
}

impl TransitionDefect {
    /// A short, stable description of what the repair did.
    #[must_use]
    pub const fn repair(self) -> &'static str {
        match self {
            Self::SpokeCount => {
                "kept the wheel effect with an unsupported spoke count verbatim as a raw effect"
            },
            Self::EffectToken => {
                "kept the effect with an out-of-domain attribute value verbatim as a raw effect"
            },
            Self::AttributeToken => "treated the out-of-domain transition attribute as absent",
            Self::TimeClamped => "saturated the transition time to the PowerPoint maximum",
        }
    }
}

/// The class of presentation slide-list defect the catalog reader repairs.
///
/// Closed by design, and repaired under every [`Leniency`] variant. The
/// repairs apply to `p:sldIdLst` in the presentation part and are recorded
/// when the package is opened.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum PresentationDefect {
    /// A `p:sldId` repeated the slide ID (`@id`) of an earlier entry.
    ///
    /// `CT_SlideIdList` states no uniqueness constraint; LibreOffice reads
    /// every entry (overwriting its ID map on repeats) and `PowerPoint` opens
    /// decks with duplicated IDs, which script generators produce. The
    /// reader keeps the first entry and skips the repeat. The context is the
    /// repeated slide ID.
    DuplicateSlideId,
    /// A `p:sldId` repeated the relationship ID (`r:id`) of an earlier
    /// entry, so it would list the same slide twice. The first entry wins;
    /// the context is the skipped entry's slide ID.
    DuplicateSlideRelationship,
    /// A `p:sldId` resolved, through a distinct relationship, to the same
    /// slide part as an earlier entry. The first entry wins; the context is
    /// the skipped entry's slide ID.
    DuplicateSlideTarget,
    /// `p:sldIdLst` carried something other than `p:sldId` children and
    /// namespace declarations: a foreign child element (its subtree is
    /// skipped), a non-namespace attribute, or non-whitespace character
    /// data. The schema admits none of it, and `PowerPoint` ignores what it
    /// does not understand; the reader ignores it. The context is `0`.
    SlideListContent,
}

impl PresentationDefect {
    /// A short, stable description of what the repair did.
    #[must_use]
    pub const fn repair(self) -> &'static str {
        match self {
            Self::DuplicateSlideId => "skipped the slide-list entry repeating an earlier slide ID",
            Self::DuplicateSlideRelationship => {
                "skipped the slide-list entry repeating an earlier relationship ID"
            },
            Self::DuplicateSlideTarget => {
                "skipped the slide-list entry repeating an earlier slide part"
            },
            Self::SlideListContent => "ignored foreign content inside the slide list",
        }
    }
}

/// The class of XML-markup defect a reader repairs.
///
/// Closed by design, and repaired under every [`Leniency`] variant.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum MarkupDefect {
    /// The part contained a processing instruction (`<?target …?>`).
    ///
    /// Processing instructions are legal XML 1.0 (§2.6) and OOXML forbids
    /// none; XSLT publishing pipelines, CMS re-serializers, and archival
    /// tools inject them, and `PowerPoint` and LibreOffice skip them through
    /// ordinary XML parsers. The reader skips them too. It is recorded once
    /// per part — for slide-owner parts by the part's shape scene, for the
    /// feature parts listed in the module documentation by that part's
    /// reader; the context is `0`. DOCTYPE declarations remain rejected as a
    /// security posture.
    ProcessingInstruction,
    /// Semantic slide text met an element named `t` outside the DrawingML
    /// namespace — a PresentationML `p:t`, or an extension payload's `x:t`
    /// (`CT_ExtensionList` admits arbitrary foreign content).
    ///
    /// Text extraction reads only DrawingML `a:t`; `PowerPoint` and the
    /// crate's own shape-text extractor ignore foreign lookalikes. The
    /// reader skips the element and its subtree. Recorded once per skipped
    /// element by the text read ([`crate::Slide::text_with_tolerance`]); the
    /// context is `0`.
    ForeignTextElement,
}

impl MarkupDefect {
    /// A short, stable description of what the repair did.
    #[must_use]
    pub const fn repair(self) -> &'static str {
        match self {
            Self::ProcessingInstruction => "skipped the part's processing instructions",
            Self::ForeignTextElement => {
                "skipped a non-DrawingML text element and its subtree during text extraction"
            },
        }
    }
}

/// The class of legacy slide-comment defect (`p:cmLst`, audit finding 6.12)
/// the comment reader repairs.
///
/// Closed by design, and repaired under every [`Leniency`] variant. Each
/// repair keeps the comment: nothing a producer wrote is dropped. The
/// context is the comment's `authorId` and [`ToleratedDefect::item`] its
/// `idx` in decimal; the part is the comment part.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum CommentDefect {
    /// A comment carried `idx="0"`.
    ///
    /// `CT_Comment@idx` is `ST_Index`, a plain `xsd:unsignedInt`, so zero is
    /// schema-valid; only `PowerPoint`'s authoring numbers from 1.
    /// LibreOffice imports the index as an opaque string. The reader keeps
    /// the comment with index `0`.
    ZeroIndex,
    /// A comment repeated the `(authorId, idx)` key of an earlier comment on
    /// the same slide or, when the whole presentation graph is loaded, on
    /// another slide.
    ///
    /// The schema states no uniqueness constraint; duplicates appear after
    /// co-authoring merges and producer round-trips. LibreOffice imports
    /// every `p:cm` in order and `PowerPoint` renumbers rather than discard
    /// them, so the reader keeps both comments in document order; a lookup by
    /// key resolves to the first.
    DuplicateKey,
}

impl CommentDefect {
    /// A short, stable description of what the repair did.
    #[must_use]
    pub const fn repair(self) -> &'static str {
        match self {
            Self::ZeroIndex => "kept the comment with index 0 as written",
            Self::DuplicateKey => {
                "kept the comment repeating an earlier author/index key after the first"
            },
        }
    }
}

/// The class of chart-part defect the chart inspection
/// ([`crate::chart::Part::chart_info`], audit finding 6.10) repairs.
///
/// Closed by design, and repaired under every [`Leniency`] variant. The
/// context is `0`; the part is the chart part.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum ChartDefect {
    /// The primary bar chart's `c:barDir` had no `val` or a value outside
    /// `ST_BarDir` (`bar`, `col`).
    ///
    /// LibreOffice reads it with `getToken(XML_val, XML_col)`; the reader
    /// reports a column chart, LibreOffice's default.
    BarDirection,
    /// `c:chart`, or the chart's `c:title`, `c:legend`, or `c:plotArea`,
    /// occurred more than once.
    ///
    /// The schema allows each once. The first occurrence is inspected and
    /// later ones are ignored, as LibreOffice's model keeps the first.
    DuplicateElement,
    /// The chart space had no `c:chart`, or the chart no `c:plotArea`, both
    /// required by the schema.
    ///
    /// No plot can be found, so the chart type is
    /// [`crate::chart::Type::Unknown`]; a title or legend that was present
    /// is still reported.
    MissingElement,
}

impl ChartDefect {
    /// A short, stable description of what the repair did.
    #[must_use]
    pub const fn repair(self) -> &'static str {
        match self {
            Self::BarDirection => "read the missing or unknown bar direction as a column chart",
            Self::DuplicateElement => "inspected the first occurrence of a repeated chart element",
            Self::MissingElement => {
                "reported an unknown chart type for a chart without a chart or plot area"
            },
        }
    }
}

/// The class of view-properties defect (`viewProps.xml`, audit finding
/// 6.11) the reader repairs.
///
/// Closed by design, and repaired under every [`Leniency`] variant.
/// `viewProps.xml` is non-authoring metadata (window panes, zoom, guides),
/// and `PowerPoint` round-trips view-property attributes it does not know,
/// so the reader ignores what its model does not cover instead of failing.
/// The context is `0`; the part is the view-properties part.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum ViewDefect {
    /// An element carried an attribute the model does not know (such as one
    /// added by a newer `PowerPoint` release). The attribute is ignored.
    UnknownAttribute,
    /// An element carried a child the model does not know, or one in a
    /// foreign namespace. The child and its subtree are ignored.
    UnknownElement,
    /// Known children appeared out of schema order. They are read by name.
    ElementOrder,
    /// A known child occurred more than once. The first occurrence is read
    /// and later ones are ignored.
    DuplicateElement,
}

impl ViewDefect {
    /// A short, stable description of what the repair did.
    #[must_use]
    pub const fn repair(self) -> &'static str {
        match self {
            Self::UnknownAttribute => "ignored an unknown view-properties attribute",
            Self::UnknownElement => "ignored an unknown view-properties element and its subtree",
            Self::ElementOrder => "read out-of-order view-properties children by name",
            Self::DuplicateElement => "read the first of a repeated view-properties element",
        }
    }
}

impl From<ViewDefect> for Defect {
    fn from(defect: ViewDefect) -> Self {
        Self::View(defect)
    }
}

impl From<ChartDefect> for Defect {
    fn from(defect: ChartDefect) -> Self {
        Self::Chart(defect)
    }
}

impl From<CommentDefect> for Defect {
    fn from(defect: CommentDefect) -> Self {
        Self::Comment(defect)
    }
}

impl From<litchi_opc::leniency::Defect> for Defect {
    fn from(defect: litchi_opc::leniency::Defect) -> Self {
        Self::Package(defect)
    }
}

impl From<litchi_ooxml_common::leniency::Defect> for Defect {
    fn from(defect: litchi_ooxml_common::leniency::Defect) -> Self {
        Self::Properties(defect)
    }
}

impl From<MarkupDefect> for Defect {
    fn from(defect: MarkupDefect) -> Self {
        Self::Markup(defect)
    }
}

impl From<PresentationDefect> for Defect {
    fn from(defect: PresentationDefect) -> Self {
        Self::Presentation(defect)
    }
}

impl From<TransitionDefect> for Defect {
    fn from(defect: TransitionDefect) -> Self {
        Self::Transition(defect)
    }
}

impl From<ShapeDefect> for Defect {
    fn from(defect: ShapeDefect) -> Self {
        Self::Shape(defect)
    }
}

impl From<TimingDefect> for Defect {
    fn from(defect: TimingDefect) -> Self {
        Self::Timing(defect)
    }
}

/// One repaired defect, with enough context to locate it in the package.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct ToleratedDefect {
    /// Which defect class was repaired.
    pub defect: Defect,
    /// Where the defect was found within its part. Its meaning is fixed per
    /// defect class:
    ///
    /// * [`Defect::Timing`] — the shape ID (`p:cNvPr@id`,
    ///   `ST_DrawingElementId`) the defect concerns: the orphaned `spid` for
    ///   [`TimingDefect::OrphanedTarget`], the duplicated `id` for
    ///   [`TimingDefect::DuplicateShapeId`];
    /// * [`Defect::Shape`] — the shape ID of the repaired shape, or `0` when
    ///   the shape carries no usable ID (`0` is outside the IDs `PowerPoint`
    ///   assigns to shapes);
    /// * [`Defect::Transition`] — a part holds at most one transition, so
    ///   the part locates it; the context is the offending spoke count for
    ///   [`TransitionDefect::SpokeCount`], the original time (saturated to
    ///   `u32::MAX`) for [`TransitionDefect::TimeClamped`], and `0`
    ///   otherwise;
    /// * [`Defect::Presentation`] — the slide ID (`p:sldId@id`) of the
    ///   skipped slide-list entry, or `0` for
    ///   [`PresentationDefect::SlideListContent`];
    /// * [`Defect::Markup`] — `0`; the part locates it;
    /// * [`Defect::Comment`] — the comment's `authorId`; its `idx` is
    ///   [`Self::item()`];
    /// * [`Defect::Package`] and [`Defect::Properties`] — `0`; the folded
    ///   defect is located by [`Self::part()`], [`Self::item()`], and
    ///   [`Self::value()`], carried over from the lower-level report.
    pub context: u32,
    /// The package part the defect was found in, such as
    /// `/ppt/slides/slide3.xml`.
    ///
    /// Package-aware entry points ([`crate::Slide::animations`],
    /// [`crate::animations::parse_package_slide`]) always stamp it. A parse of
    /// bare slide XML ([`crate::animations::Sequence::parse_slide_xml`]) has
    /// no part to name, so it leaves this `None`.
    pub part: Option<Arc<str>>,
    /// The keyed item within the part, for defects folded from a lower-level
    /// report: a relationship `Id`, a content-type extension or part name,
    /// or a property name (see `litchi_opc::leniency::ToleratedDefect::item`
    /// and `litchi_ooxml_common::leniency::ToleratedDefect::item`); for a
    /// [`Defect::Comment`], the comment index (`idx`) in decimal. `None` for
    /// this crate's other defect classes, which locate themselves through
    /// [`Self::context`].
    pub item: Option<Box<str>>,
    /// The raw value a folded lower-level repair replaced or discarded, as
    /// documented on the folded defect variant. `None` for this crate's own
    /// defect classes.
    pub value: Option<Box<str>>,
}

impl ToleratedDefect {
    /// The package part the defect was found in, when the parse knew it.
    #[inline]
    #[must_use]
    pub fn part(&self) -> Option<&str> {
        self.part.as_deref()
    }

    /// The keyed item a folded lower-level defect concerns (a relationship
    /// `Id`, content-type key, or property name), if any.
    #[inline]
    #[must_use]
    pub fn item(&self) -> Option<&str> {
        self.item.as_deref()
    }

    /// The raw value a folded lower-level repair replaced or discarded.
    #[inline]
    #[must_use]
    pub fn value(&self) -> Option<&str> {
        self.value.as_deref()
    }
}

/// The empty report handed out by values that carry no repairs.
pub(crate) static EMPTY_REPORT: ToleranceReport = ToleranceReport {
    defects: Vec::new(),
    unrecorded: 0,
};

/// Everything a read repaired, in the order the defects were found.
///
/// An empty report means the presentation conformed and no substitution was
/// made. After a [`Leniency::Strict`] read the report can still list
/// [`Defect::Timing`] repairs, which `PowerPoint` itself applies.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct ToleranceReport {
    defects: Vec<ToleratedDefect>,
    unrecorded: u32,
}

impl ToleranceReport {
    /// Whether the presentation parsed without any repair.
    #[inline]
    #[must_use]
    pub fn is_clean(&self) -> bool {
        self.defects.is_empty() && self.unrecorded == 0
    }

    /// The individually recorded defects, in discovery order.
    #[inline]
    #[must_use]
    pub fn defects(&self) -> &[ToleratedDefect] {
        &self.defects
    }

    /// Defects that were repaired but not stored, once the bound was reached.
    #[inline]
    #[must_use]
    pub fn unrecorded(&self) -> u32 {
        self.unrecorded
    }

    /// Total repairs performed, recorded or not.
    #[inline]
    #[must_use]
    pub fn total(&self) -> u64 {
        self.defects.len() as u64 + u64::from(self.unrecorded)
    }

    /// Record one repair, saturating at [`MAX_RECORDED_DEFECTS`] stored entries.
    pub(crate) fn record(&mut self, defect: impl Into<Defect>, context: u32) {
        self.push(ToleratedDefect {
            defect: defect.into(),
            context,
            part: None,
            item: None,
            value: None,
        });
    }

    /// Record one comment repair: the context is the author ID and the item
    /// the comment index.
    pub(crate) fn record_comment(&mut self, defect: CommentDefect, author_id: u32, index: u32) {
        self.push(ToleratedDefect {
            defect: defect.into(),
            context: author_id,
            part: None,
            item: Some(index.to_string().into_boxed_str()),
            value: None,
        });
    }

    /// Fold an OPC packaging report into this one, keeping its discovery
    /// order, locations, and unrecorded count.
    pub(crate) fn append_package(&mut self, package: &litchi_opc::leniency::ToleranceReport) {
        for tolerated in package.defects() {
            self.push(ToleratedDefect {
                defect: Defect::Package(tolerated.defect()),
                context: 0,
                part: Some(Arc::from(tolerated.part())),
                item: tolerated.item().map(Box::from),
                value: tolerated.value().map(Box::from),
            });
        }
        self.unrecorded = self.unrecorded.saturating_add(package.unrecorded());
    }

    /// Fold a core- or custom-properties report into this one, keeping its
    /// discovery order, locations, and unrecorded count.
    pub(crate) fn append_properties(
        &mut self,
        properties: &litchi_ooxml_common::leniency::ToleranceReport,
    ) {
        for tolerated in properties.defects() {
            self.push(ToleratedDefect {
                defect: Defect::Properties(tolerated.defect()),
                context: 0,
                part: Some(Arc::from(tolerated.part())),
                item: tolerated.item().map(Box::from),
                value: tolerated.value().map(Box::from),
            });
        }
        self.unrecorded = self.unrecorded.saturating_add(properties.unrecorded());
    }

    fn push(&mut self, tolerated: ToleratedDefect) {
        if self.defects.len() < MAX_RECORDED_DEFECTS {
            self.defects.push(tolerated);
        } else {
            self.unrecorded = self.unrecorded.saturating_add(1);
        }
    }

    /// Stamp `part` on every recorded defect that does not name its part yet.
    pub(crate) fn locate_in_part(&mut self, part: &str) {
        if self
            .defects
            .iter()
            .all(|tolerated| tolerated.part.is_some())
        {
            return;
        }
        let part: Arc<str> = Arc::from(part);
        for tolerated in &mut self.defects {
            if tolerated.part.is_none() {
                tolerated.part = Some(Arc::clone(&part));
            }
        }
    }

    /// Merge another parse's report into this one, keeping discovery order.
    ///
    /// The stored-entry bound applies to the merged report as a whole; entries
    /// that no longer fit join the unrecorded count.
    pub(crate) fn append(&mut self, other: &ToleranceReport) {
        for tolerated in other.defects() {
            self.push(tolerated.clone());
        }
        self.unrecorded = self.unrecorded.saturating_add(other.unrecorded);
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn strict_is_the_default() {
        assert_eq!(Leniency::default(), Leniency::Strict);
    }

    #[test]
    fn a_fresh_report_is_clean() {
        let report = ToleranceReport::default();
        assert!(report.is_clean());
        assert_eq!(report.total(), 0);
        assert!(report.defects().is_empty());
    }

    #[test]
    fn recorded_defects_are_returned_in_discovery_order() {
        let mut report = ToleranceReport::default();
        report.record(TimingDefect::OrphanedTarget, 7);
        report.record(TimingDefect::DuplicateShapeId, 42);
        report.record(TimingDefect::OrphanedTarget, 9);

        assert!(!report.is_clean());
        assert_eq!(report.total(), 3);
        assert_eq!(
            report
                .defects()
                .iter()
                .map(|d| (d.defect, d.context))
                .collect::<Vec<_>>(),
            vec![
                (Defect::Timing(TimingDefect::OrphanedTarget), 7),
                (Defect::Timing(TimingDefect::DuplicateShapeId), 42),
                (Defect::Timing(TimingDefect::OrphanedTarget), 9),
            ]
        );
    }

    /// A hostile timing tree must not be able to grow the report without limit.
    #[test]
    fn defects_past_the_bound_are_counted_not_stored() {
        let mut report = ToleranceReport::default();
        for index in 0..(MAX_RECORDED_DEFECTS + 5) {
            report.record(TimingDefect::OrphanedTarget, index as u32);
        }
        for _ in 0..3 {
            report.record(TimingDefect::DuplicateShapeId, 2);
        }

        assert_eq!(report.defects().len(), MAX_RECORDED_DEFECTS);
        assert_eq!(report.unrecorded(), 8);
        assert_eq!(report.total(), MAX_RECORDED_DEFECTS as u64 + 8);
    }

    #[test]
    fn appended_reports_merge_within_the_bound() {
        let mut base = ToleranceReport::default();
        base.record(TimingDefect::DuplicateShapeId, 7);

        let mut parsed = ToleranceReport::default();
        parsed.record(TimingDefect::OrphanedTarget, 12);
        for _ in 1..MAX_RECORDED_DEFECTS {
            parsed.record(TimingDefect::OrphanedTarget, 13);
        }
        assert_eq!(parsed.unrecorded(), 0);

        base.append(&parsed);
        assert_eq!(base.defects().len(), MAX_RECORDED_DEFECTS);
        assert_eq!(base.unrecorded(), 1);
        assert_eq!(base.total(), MAX_RECORDED_DEFECTS as u64 + 1);
        assert_eq!(
            base.defects()[0],
            ToleratedDefect {
                defect: Defect::Timing(TimingDefect::DuplicateShapeId),
                context: 7,
                part: None,
                item: None,
                value: None,
            }
        );

        let mut clean = ToleranceReport::default();
        clean.append(&ToleranceReport::default());
        assert!(clean.is_clean());
    }

    #[test]
    fn every_defect_class_describes_its_repair() {
        let defects = [
            Defect::from(TimingDefect::OrphanedTarget),
            Defect::from(TimingDefect::DuplicateShapeId),
            Defect::from(ShapeDefect::MalformedTransform),
            Defect::from(ShapeDefect::PlaceholderExtensionOwner),
            Defect::from(ShapeDefect::PlaceholderTypeExtension),
            Defect::from(ShapeDefect::InvalidShapeId),
            Defect::from(ShapeDefect::InvalidPlaceholderIndex),
            Defect::from(TransitionDefect::SpokeCount),
            Defect::from(TransitionDefect::EffectToken),
            Defect::from(TransitionDefect::AttributeToken),
            Defect::from(TransitionDefect::TimeClamped),
            Defect::from(PresentationDefect::DuplicateSlideId),
            Defect::from(PresentationDefect::DuplicateSlideRelationship),
            Defect::from(PresentationDefect::DuplicateSlideTarget),
            Defect::from(PresentationDefect::SlideListContent),
            Defect::from(MarkupDefect::ProcessingInstruction),
            Defect::from(MarkupDefect::ForeignTextElement),
            Defect::from(ViewDefect::UnknownAttribute),
            Defect::from(ViewDefect::UnknownElement),
            Defect::from(ViewDefect::ElementOrder),
            Defect::from(ViewDefect::DuplicateElement),
            Defect::from(ChartDefect::BarDirection),
            Defect::from(ChartDefect::DuplicateElement),
            Defect::from(ChartDefect::MissingElement),
            Defect::from(CommentDefect::ZeroIndex),
            Defect::from(CommentDefect::DuplicateKey),
            Defect::from(litchi_opc::leniency::Defect::Relationship(
                litchi_opc::leniency::RelationshipDefect::DuplicateId,
            )),
            Defect::from(litchi_ooxml_common::leniency::Defect::CoreProperties(
                litchi_ooxml_common::leniency::CorePropertiesDefect::DuplicateProperty,
            )),
        ];
        for defect in defects {
            assert!(!defect.repair().is_empty());
        }
        assert_eq!(
            Defect::Timing(TimingDefect::OrphanedTarget).repair(),
            TimingDefect::OrphanedTarget.repair()
        );
        assert_eq!(
            Defect::Timing(TimingDefect::DuplicateShapeId).repair(),
            TimingDefect::DuplicateShapeId.repair()
        );
    }
}
