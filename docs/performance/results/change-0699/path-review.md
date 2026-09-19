# Refusal reachability review

performance_claim: none

The early-name-error fixture reaches a real capture refusal, but not the
borrowed-Inherited code changed in the rejected candidate. This distinction
was established by an independent source review and checked against current
source by the coordinator.

The diagnostic inherits the exact fixture construction from 0698. It authors
complex_package, then injects a duplicate name attribute into the first slide.
It does not use the separately created mce_complex_package. The generated
slide writer declares PresentationML, DrawingML and relationship namespaces;
controlled fixture text contains no MCE namespace URI. Presentation capture
loads each SlidePart through from_part_with_name and processed_xml_with_source.
The unchanged MCE processor checks for its namespace URI before constructing
a Reader. Marker-free bytes return Cow::Borrowed, so the start handler is not
called for these slide payloads. The separate NsReader detects the duplicate
attribute while extracting the cSld name. Capture retains and returns that
name error before later semantic work.

Relevant source locations at the recorded baseline:

- probe/src/main.rs: build_cases, authored_package and add_mce_markers;
- crates/litchi-pptx/src/writer/slide.rs: generated namespace declarations;
- crates/litchi-pptx/src/presentation/package.rs: capture_slides;
- crates/litchi-pptx/src/parts/slide.rs: from_part_with_name;
- crates/litchi-ooxml-common/src/mce/codec.rs: process_markup_compatibility
  marker-free return before Reader construction;
- crates/litchi-pptx/src/namespace.rs: read_attrs; and
- crates/litchi-pptx/src/opened/model.rs: first_slide_name_error precedence.

Thus the prior candidate's handler stack increase is a static tradeoff, not
an established cause of this marker-free refusal slowdown. Moving or changing
cold code can affect binary layout, but the source review does not prove that
layout caused the timing difference. Likewise, timing variability does not
prove host noise. Balanced single-case legs, the unchanged full matrix, and
an MCE-marked positive control are needed to narrow those possibilities.
Profiles include fixture setup and diagnostic assertions. Absence of sampled
start frames alone is not proof of unreachability; the source path supplies
the stronger argument.
