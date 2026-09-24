use std::{
    io::{self, Cursor, Read},
    mem::size_of,
    num::{NonZeroU64, NonZeroUsize},
    panic::{AssertUnwindSafe, catch_unwind},
    sync::{
        Arc,
        atomic::{AtomicUsize, Ordering},
    },
};

use litchi_core::{
    Budget, CancellationSource, Error, ExecutionContext, ExecutionLimits, Limits as BudgetLimits,
    Position, Resource,
};
use litchi_odt::constants::{
    ODF_CHART_TEMPLATE, ODF_CONTENT, ODF_DRAWING_TEMPLATE, ODF_FORMULA_TEMPLATE,
    ODF_IMAGE_TEMPLATE, ODF_MASTER_TEMPLATE, ODF_PRESENTATION_TEMPLATE, ODF_SPREADSHEET_TEMPLATE,
    ODF_TEXT, ODF_TEXT_TEMPLATE, ODF_WEB, get_flat_extension_from_mime_type,
    get_mime_type_from_extension, is_odf_extension,
};
use litchi_odt::core::PackageWriter;
use litchi_odt::flat::{Document, Limits};
use litchi_odt::generic::{Family, FlatDocument, Package};
use litchi_odt::variable_declaration::{Body, Declaration, Group, Kind, Part, Scope, ValueType};
use litchi_odt::{RdfaAttributes, RdfaOccurrence, TextMeta, XFormsModel};

const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const TEXT: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const XFORMS: &str = "http://www.w3.org/2002/xforms";

// A small synthetic flat fixture. Its MIME whitespace and prefix spelling are
// intentional: bytes remain authoritative while metadata is detector-canonical.
const TEMPLATE_XML: &str = r#"<?xml version="1.0" encoding="UTF-8"?><o:document xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:h="http://www.w3.org/1999/xhtml" o:version="1.3" o:mimetype=" application/vnd.oasis.opendocument.text-template "><!-- preserve this comment --><o:body><o:text><t:p>Alpha</t:p><t:p>Structured <t:span>content</t:span></t:p></o:text></o:body></o:document>"#;

const MUTATOR_TEMPLATE_XML: &str = r##"<?xml version="1.0"?><o:document xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:h="http://www.w3.org/1999/xhtml" xmlns:xf="http://www.w3.org/2002/xforms" o:mimetype=" application/vnd.oasis.opendocument.text-template " o:version="1.3"><o:body><o:text><o:forms><xf:model id="m"><xf:instance id="i"><d:data xmlns:d="urn:example:data"><d:value>one</d:value></d:data></xf:instance><xf:bind id="b" nodeset="/data/value" type="xsd:string"/></xf:model></o:forms><t:p h:about="#doc" h:property="dc:title"><t:bookmark-start t:name="mark"/><t:meta xml:id="meta" h:property="dc:description">value</t:meta>visible</t:p></o:text></o:body></o:document>"##;

fn normal_xml() -> String {
    TEMPLATE_XML.replace(ODF_TEXT_TEMPLATE, ODF_TEXT)
}

fn flat_xml(mimetype: &str, body: &str) -> Vec<u8> {
    format!(
        r#"<o:document xmlns:o="{OFFICE}" o:mimetype="{mimetype}"><o:body><o:{body}/></o:body></o:document>"#
    )
    .into_bytes()
}

fn many_metadata_nodes(count: usize) -> Vec<u8> {
    let mut xml = format!(
        r#"<o:document xmlns:o="{OFFICE}" xmlns:t="{TEXT}" xmlns:h="http://www.w3.org/1999/xhtml" o:mimetype="{ODF_TEXT}"><o:body><o:text>"#
    );
    for index in 0..count {
        xml.push_str(&format!(
            r##"<t:p h:about="#metadata-{index}" h:property="dc:title">x</t:p>"##
        ));
    }
    xml.push_str("</o:text></o:body></o:document>");
    let bytes = xml.into_bytes();
    let mut exact = Vec::with_capacity(bytes.len());
    exact.extend_from_slice(&bytes);
    exact
}

fn many_xforms_models(count: usize) -> Vec<u8> {
    let mut xml = format!(
        r#"<o:document xmlns:o="{OFFICE}" xmlns:t="{TEXT}" xmlns:xf="{XFORMS}" o:mimetype="{ODF_TEXT}"><o:body><o:text><o:forms>"#
    );
    for index in 0..count {
        xml.push_str(&format!(r#"<xf:model id="model-{index}"/>"#));
    }
    xml.push_str("</o:forms><t:p>content</t:p></o:text></o:body></o:document>");
    let bytes = xml.into_bytes();
    let mut exact = Vec::with_capacity(bytes.len());
    exact.extend_from_slice(&bytes);
    exact
}

fn whitespace_heavy_metadata() -> Vec<u8> {
    let mut xml = format!(
        r#"<o:document xmlns:o="{OFFICE}" xmlns:t="{TEXT}" xmlns:h="http://www.w3.org/1999/xhtml" o:mimetype="{ODF_TEXT}">
  <!-- leading inert metadata comment -->
  <o:body>
    <o:text>
"#
    );
    for _ in 0..96 {
        xml.push_str("      <!-- whitespace and comments are still parser work -->\n");
        xml.push_str("      \n\t  \n");
    }
    xml.push_str(
        r#"      <t:p h:property="dc:title">Alpha</t:p>
    </o:text>
  </o:body>
</o:document>"#,
    );
    let bytes = xml.into_bytes();
    let mut exact = Vec::with_capacity(bytes.len());
    exact.extend_from_slice(&bytes);
    exact
}

fn oversized_reader_fixture() -> Vec<u8> {
    let mut xml = format!(
        r#"<o:document xmlns:o="{OFFICE}" xmlns:t="{TEXT}" o:mimetype="{ODF_TEXT}"><o:body><o:text>"#
    );
    while xml.len() <= 70 * 1024 {
        xml.push_str("<!-- bounded reader growth fixture -->\n");
    }
    xml.push_str("<t:p>Reader growth</t:p></o:text></o:body></o:document>");
    xml.into_bytes()
}

fn assert_no_owned_flat_temporary_files(directory: &std::path::Path) {
    assert!(std::fs::read_dir(directory).unwrap().all(|entry| {
        let name = entry.unwrap().file_name();
        let name = name.to_string_lossy();
        !(name.starts_with(".litchi-") && name.ends_with(".tmp"))
    }));
}

fn template_bytes() -> Vec<u8> {
    TEMPLATE_XML.as_bytes().to_vec()
}

fn execution_context_with_memory(memory: u64) -> (Budget, CancellationSource, ExecutionContext) {
    execution_context_with_memory_and_work(memory, u64::MAX)
}

fn execution_context_with_memory_and_work(
    memory: u64,
    work: u64,
) -> (Budget, CancellationSource, ExecutionContext) {
    let budget = Budget::root(
        "flat-template execution test",
        BudgetLimits::new(memory, u64::MAX, u64::MAX, u64::MAX, u64::MAX, work),
    );
    let (cancellation, token) = CancellationSource::pair();
    let limits = ExecutionLimits::new(
        NonZeroUsize::new(1).unwrap(),
        NonZeroUsize::new(1).unwrap(),
        NonZeroU64::new(u64::MAX).unwrap(),
        0,
    )
    .unwrap();
    let context = ExecutionContext::new(budget.clone(), token, limits);
    (budget, cancellation, context)
}

fn retained_template_bytes(extra_capacity: usize) -> (Vec<u8>, u64) {
    let source = TEMPLATE_XML.as_bytes();
    let mut bytes = Vec::with_capacity(source.len() + extra_capacity);
    bytes.extend_from_slice(source);
    let capacity = u64::try_from(bytes.capacity()).unwrap();
    assert!(capacity > u64::try_from(bytes.len()).unwrap());
    (bytes, capacity)
}

fn assert_memory_limit(error: Error) {
    assert!(matches!(
        error,
        Error::ResourceLimit(limit) if limit.resource == Resource::Memory
    ));
}

fn assert_work_limit(error: Error) {
    assert!(matches!(
        error,
        Error::ResourceLimit(limit) if limit.resource == Resource::Work
    ));
}

fn assert_cancelled(error: Error) {
    assert!(
        error.to_string().contains("cancelled"),
        "unexpected error: {error}"
    );
}

fn mutator_rdfa_value() -> RdfaAttributes {
    RdfaAttributes {
        about: Some("#changed".to_string()),
        property: Some("dc:subject".to_string()),
        ..RdfaAttributes::default()
    }
}

fn apply_paragraph_rdfa(document: &mut FlatDocument) -> litchi_core::Result<()> {
    document.set_paragraph_rdfa(Position::new(0), &mutator_rdfa_value())
}

type FlatMutator = fn(&mut FlatDocument) -> litchi_core::Result<()>;

fn apply_bookmark_rdfa(document: &mut FlatDocument) -> litchi_core::Result<()> {
    document.set_bookmark_rdfa("mark", &mutator_rdfa_value())
}

fn apply_text_meta_rdfa(document: &mut FlatDocument) -> litchi_core::Result<()> {
    document.set_text_meta_rdfa(Position::new(0), &mutator_rdfa_value())
}

fn apply_insert_text_meta(document: &mut FlatDocument) -> litchi_core::Result<()> {
    document.insert_text_meta(Position::new(0), &TextMeta::from_text("inserted")?)
}

fn apply_replace_text_meta(document: &mut FlatDocument) -> litchi_core::Result<()> {
    document.replace_text_meta(Position::new(0), &TextMeta::from_text("replacement-long")?)
}

fn apply_replace_xforms_model(document: &mut FlatDocument) -> litchi_core::Result<()> {
    let mut model = document.xforms_models()?.remove(0);
    model.id = Some("replacement-model-with-more-bytes".to_string());
    document.replace_xforms_model(Position::new(0), &model)
}

fn apply_insert_xforms_model(document: &mut FlatDocument) -> litchi_core::Result<()> {
    let mut model = document.xforms_models()?.remove(0);
    model.id = Some("inserted-model-with-more-bytes".to_string());
    model.children.clear();
    document.insert_xforms_model(&model)
}

fn apply_variable_declaration(document: &mut FlatDocument) -> litchi_core::Result<()> {
    let group = Group {
        kind: Kind::Simple,
        part: Part::Flat,
        scope: Scope::Body(Body::Text),
        declarations: vec![Declaration::Simple {
            name: "counter-with-more-bytes".to_string(),
            value_type: ValueType::String,
        }],
    };
    document.set_variable_declaration_group(&group).map(|_| ())
}

fn assert_growing_mutator_output_limit(name: &str, mutator: FlatMutator) {
    let source = MUTATOR_TEMPLATE_XML.as_bytes().to_vec();
    let source_len = source.len();
    let mut probe = FlatDocument::from_bytes(source.clone()).unwrap();
    mutator(&mut probe).unwrap();
    let exact_output_len = probe.as_bytes().len();
    assert!(
        exact_output_len > source_len,
        "{name} must exercise an output growth boundary"
    );

    let exact_limit = u64::try_from(exact_output_len).unwrap();
    let mut exact = FlatDocument::from_bytes_with_limit(source.clone(), exact_limit).unwrap();
    mutator(&mut exact).unwrap();
    assert_eq!(
        exact.as_bytes().len(),
        exact_output_len,
        "{name} accepted its exact output cap"
    );
    assert_template_metadata(&exact);

    let mut under = FlatDocument::from_bytes_with_limit(source, exact_limit - 1).unwrap();
    let before = under.to_bytes();
    assert!(
        matches!(mutator(&mut under), Err(Error::ResourceLimit(_))),
        "{name} accepted output over its caller cap"
    );
    assert_eq!(
        under.as_bytes(),
        before.as_slice(),
        "{name} mutated on refusal"
    );
    assert_template_metadata(&under);
}

fn assert_template_metadata<F>(document: &F)
where
    F: TemplateMetadata,
{
    assert_eq!(document.family(), Family::Text);
    assert!(document.is_template());
    assert_eq!(document.mimetype(), ODF_TEXT_TEMPLATE);
    assert_eq!(document.extension(), "fott");
}

fn assert_normal_metadata<F>(document: &F)
where
    F: TemplateMetadata,
{
    assert_eq!(document.family(), Family::Text);
    assert!(!document.is_template());
    assert_eq!(document.mimetype(), ODF_TEXT);
    assert_eq!(document.extension(), "fodt");
}

trait TemplateMetadata {
    fn family(&self) -> Family;
    fn is_template(&self) -> bool;
    fn mimetype(&self) -> &str;
    fn extension(&self) -> &str;
}

impl TemplateMetadata for FlatDocument {
    fn family(&self) -> Family {
        FlatDocument::family(self)
    }

    fn is_template(&self) -> bool {
        FlatDocument::is_template(self)
    }

    fn mimetype(&self) -> &str {
        FlatDocument::mimetype(self)
    }

    fn extension(&self) -> &str {
        FlatDocument::extension(self)
    }
}

impl TemplateMetadata for Document {
    fn family(&self) -> Family {
        // The specialized facade has the same family boundary as its generic
        // source snapshot; this delegate is part of the template contract.
        Document::family(self)
    }

    fn is_template(&self) -> bool {
        Document::is_template(self)
    }

    fn mimetype(&self) -> &str {
        Document::mimetype(self)
    }

    fn extension(&self) -> &str {
        Document::extension(self)
    }
}

#[test]
fn common_template_mapping_and_neutral_detection_are_public() {
    assert_eq!(
        get_mime_type_from_extension("fott"),
        Some(ODF_TEXT_TEMPLATE)
    );
    assert_eq!(
        get_flat_extension_from_mime_type(ODF_TEXT_TEMPLATE),
        Some("fott")
    );
    assert!(is_odf_extension("fott"));
    assert_eq!(
        litchi_odt::detect::mime(ODF_TEXT_TEMPLATE.as_bytes()),
        Some(litchi_odt::detect::Format::Odt)
    );
    assert_eq!(
        litchi_odt::detect::flat_mime(TEMPLATE_XML.as_bytes()).as_deref(),
        Some(ODF_TEXT_TEMPLATE)
    );
    assert_eq!(
        litchi_odt::detect::flat(TEMPLATE_XML.as_bytes()),
        Some(litchi_odt::detect::Format::Odt)
    );
}

#[test]
fn generic_and_specialized_template_metadata_preserve_exact_source_bytes() {
    let bytes = template_bytes();
    let generic = FlatDocument::from_bytes(bytes.clone()).unwrap();
    assert_template_metadata(&generic);
    assert_eq!(generic.xml(), TEMPLATE_XML);
    assert_eq!(generic.as_bytes(), bytes.as_slice());
    assert_eq!(generic.to_bytes(), bytes);
    assert_eq!(generic.into_bytes(), TEMPLATE_XML.as_bytes());

    let specialized = Document::from_bytes(template_bytes()).unwrap();
    assert_template_metadata(&specialized);
    assert_eq!(specialized.xml(), TEMPLATE_XML);
    assert_eq!(specialized.as_bytes(), TEMPLATE_XML.as_bytes());
    assert_eq!(specialized.to_bytes(), TEMPLATE_XML.as_bytes());
    assert_eq!(specialized.text().unwrap(), "AlphaStructured content");

    let normal = normal_xml();
    let generic_normal = FlatDocument::from_bytes(normal.as_bytes().to_vec()).unwrap();
    assert_normal_metadata(&generic_normal);
    assert_eq!(generic_normal.as_bytes(), normal.as_bytes());
    let specialized_normal = Document::from_bytes(normal.as_bytes().to_vec()).unwrap();
    assert_normal_metadata(&specialized_normal);
    assert_eq!(specialized_normal.as_bytes(), normal.as_bytes());
}

#[test]
fn execution_context_admission_charges_retained_vec_capacity_and_releases_on_drop() {
    let (bytes, retained_capacity) = retained_template_bytes(4096);
    let maximum = u64::try_from(TEMPLATE_XML.len()).unwrap();
    let (zero_budget, _zero_cancel, zero_context) = execution_context_with_memory(0);
    assert_memory_limit(
        FlatDocument::from_bytes_with_execution_context(bytes, maximum, zero_context)
            .err()
            .unwrap(),
    );
    assert_eq!(zero_budget.used(Resource::Memory), 0);

    let (bytes, retained_capacity_again) = retained_template_bytes(4096);
    assert_eq!(retained_capacity_again, retained_capacity);
    let (under_budget, _under_cancel, under_context) =
        execution_context_with_memory(retained_capacity - 1);
    assert_memory_limit(
        FlatDocument::from_bytes_with_execution_context(bytes, maximum, under_context)
            .err()
            .unwrap(),
    );
    assert_eq!(under_budget.used(Resource::Memory), 0);

    let (bytes, retained_capacity_again) = retained_template_bytes(4096);
    assert_eq!(retained_capacity_again, retained_capacity);
    let (budget, _cancel, context) = execution_context_with_memory(retained_capacity);
    let document =
        FlatDocument::from_bytes_with_execution_context(bytes, maximum, context).unwrap();
    assert_eq!(budget.used(Resource::Memory), retained_capacity);
    drop(document);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn execution_context_transfer_preserves_old_reservation_on_failure_and_moves_on_success() {
    let (bytes, retained_capacity) = retained_template_bytes(4096);
    let maximum = u64::try_from(TEMPLATE_XML.len()).unwrap();
    let (old_budget, _old_cancel, old_context) = execution_context_with_memory(retained_capacity);
    let mut document =
        FlatDocument::from_bytes_with_execution_context(bytes, maximum, old_context).unwrap();
    let before = document.to_bytes();
    assert_eq!(old_budget.used(Resource::Memory), retained_capacity);

    let (failed_budget, _failed_cancel, failed_context) =
        execution_context_with_memory(retained_capacity - 1);
    assert_memory_limit(
        document
            .with_execution_context(failed_context)
            .err()
            .unwrap(),
    );
    assert_eq!(document.as_bytes(), before.as_slice());
    assert_eq!(old_budget.used(Resource::Memory), retained_capacity);
    assert_eq!(failed_budget.used(Resource::Memory), 0);

    let (new_budget, _new_cancel, new_context) = execution_context_with_memory(retained_capacity);
    document.with_execution_context(new_context).unwrap();
    assert_eq!(old_budget.used(Resource::Memory), 0);
    assert_eq!(new_budget.used(Resource::Memory), retained_capacity);
    drop(document);
    assert_eq!(new_budget.used(Resource::Memory), 0);
}

struct PanicIfRead {
    reads: Arc<AtomicUsize>,
}

impl Read for PanicIfRead {
    fn read(&mut self, _buffer: &mut [u8]) -> io::Result<usize> {
        self.reads.fetch_add(1, Ordering::SeqCst);
        panic!("cancelled flat input must not call Read");
    }
}

struct CancelAfterFirstRead {
    cancellation: CancellationSource,
    payload: Vec<u8>,
    reads: Arc<AtomicUsize>,
}

impl Read for CancelAfterFirstRead {
    fn read(&mut self, buffer: &mut [u8]) -> io::Result<usize> {
        let read_number = self.reads.fetch_add(1, Ordering::SeqCst);
        if read_number != 0 {
            panic!("flat input must check cancellation before its next read");
        }
        let amount = self.payload.len().min(buffer.len());
        buffer[..amount].copy_from_slice(&self.payload[..amount]);
        self.payload.drain(..amount);
        self.cancellation.cancel();
        Ok(amount)
    }
}

#[test]
fn already_cancelled_flat_reader_refuses_before_invoking_read() {
    let maximum = u64::try_from(TEMPLATE_XML.len()).unwrap();
    let (budget, cancellation, context) = execution_context_with_memory(u64::MAX);
    cancellation.cancel();
    let reads = Arc::new(AtomicUsize::new(0));
    let reader = PanicIfRead {
        reads: Arc::clone(&reads),
    };
    let result = catch_unwind(AssertUnwindSafe(|| {
        FlatDocument::from_reader_with_execution_context(reader, maximum, context)
    }));
    assert!(result.is_ok(), "cancelled flat input invoked its reader");
    assert_cancelled(result.unwrap().err().unwrap());
    assert_eq!(reads.load(Ordering::SeqCst), 0);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn mid_read_flat_cancellation_refuses_without_retaining_input_memory() {
    let maximum = u64::try_from(TEMPLATE_XML.len()).unwrap();
    let (budget, cancellation, context) = execution_context_with_memory(u64::MAX);
    let reads = Arc::new(AtomicUsize::new(0));
    let reader = CancelAfterFirstRead {
        cancellation,
        payload: template_bytes(),
        reads: Arc::clone(&reads),
    };
    let result = catch_unwind(AssertUnwindSafe(|| {
        FlatDocument::from_reader_with_execution_context(reader, maximum, context)
    }));
    assert!(result.is_ok(), "flat input read past cancellation");
    assert_cancelled(result.unwrap().err().unwrap());
    assert_eq!(reads.load(Ordering::SeqCst), 1);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn oversized_flat_reader_charges_growth_peak_and_admits_with_enough_memory() {
    let source = oversized_reader_fixture();
    assert!(source.len() > 64 * 1024);
    let maximum = u64::try_from(source.len()).unwrap();

    // The retained source is below this limit, while the 64 KiB old
    // allocation and the roughly 72 KiB destination coexist during the first
    // growth. A delta-only implementation therefore admits this input, while
    // peak accounting must refuse it at the exact old-plus-new observed size.
    let peak_limit = 96 * 1024;
    assert!(u64::try_from(source.len()).unwrap() < peak_limit);
    let (refusal_budget, _refusal_cancel, refusal_context) =
        execution_context_with_memory(peak_limit);
    let refusal = FlatDocument::from_reader_with_execution_context(
        Cursor::new(source.clone()),
        maximum,
        refusal_context,
    )
    .err()
    .unwrap();
    match refusal {
        Error::ResourceLimit(limit) => {
            assert_eq!(limit.resource, Resource::Memory);
            assert_eq!(limit.limit, peak_limit);
            assert_eq!(
                limit.observed,
                64 * 1024 + u64::try_from(source.len()).unwrap()
            );
        },
        error => panic!("expected peak memory refusal, got {error}"),
    }
    assert_eq!(refusal_budget.used(Resource::Memory), 0);

    let (admission_budget, _admission_cancel, admission_context) =
        execution_context_with_memory(u64::try_from(source.len() * 8).unwrap());
    let document = FlatDocument::from_reader_with_execution_context(
        Cursor::new(source.clone()),
        maximum,
        admission_context,
    )
    .unwrap();
    assert_eq!(document.as_bytes(), source.as_slice());
    drop(document);
    assert_eq!(admission_budget.used(Resource::Memory), 0);
}

#[test]
fn cancellation_before_flat_mutation_refuses_before_candidate_publication() {
    let source = template_bytes();
    let source_memory = u64::try_from(source.len()).unwrap();
    let (budget, cancellation, context) = execution_context_with_memory(source_memory * 8);
    let mut document = FlatDocument::from_bytes_with_execution_context(
        source.clone(),
        u64::try_from(source.len()).unwrap() + 256,
        context,
    )
    .unwrap();
    let before = document.to_bytes();
    assert_eq!(budget.used(Resource::Memory), source_memory);
    cancellation.cancel();

    let error = document
        .set_paragraph_rdfa(Position::new(0), &mutator_rdfa_value())
        .err()
        .unwrap();
    assert_cancelled(error);
    assert_eq!(document.as_bytes(), before.as_slice());
    assert_eq!(budget.used(Resource::Memory), source_memory);
    drop(document);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn flat_mutation_work_limits_refuse_atomically_and_release_memory_on_drop() {
    let maximum = u64::try_from(TEMPLATE_XML.len() + 1024).unwrap();
    for work in [0, 1] {
        let (source, source_capacity) = retained_template_bytes(4096);
        let (admission_budget, _admission_cancel, admission_context) =
            execution_context_with_memory_and_work(source_capacity * 16, u64::MAX);
        let mut document =
            FlatDocument::from_bytes_with_execution_context(source, maximum, admission_context)
                .unwrap();

        let (mutation_budget, _mutation_cancel, mutation_context) =
            execution_context_with_memory_and_work(source_capacity * 16, work);
        document.with_execution_context(mutation_context).unwrap();
        assert_eq!(admission_budget.used(Resource::Memory), 0);
        assert_eq!(mutation_budget.used(Resource::Memory), source_capacity);
        let before = document.to_bytes();

        let error = document
            .set_paragraph_rdfa(Position::new(0), &mutator_rdfa_value())
            .err()
            .unwrap();
        assert_work_limit(error);
        assert_eq!(document.as_bytes(), before.as_slice());
        assert_eq!(
            FlatDocument::from_bytes(before.clone()).unwrap().as_bytes(),
            before.as_slice()
        );
        assert_eq!(mutation_budget.used(Resource::Memory), source_capacity);
        drop(document);
        assert_eq!(mutation_budget.used(Resource::Memory), 0);
    }
}

#[test]
fn whitespace_heavy_metadata_scan_consumes_work_without_candidate_allocation() {
    let source = whitespace_heavy_metadata();
    let source_capacity = u64::try_from(source.capacity()).unwrap();
    let maximum = u64::try_from(source.len() + 256).unwrap();
    let (admission_budget, _admission_cancel, admission_context) =
        execution_context_with_memory_and_work(source_capacity * 16, u64::MAX);
    let mut document =
        FlatDocument::from_bytes_with_execution_context(source.clone(), maximum, admission_context)
            .unwrap();
    let (mutation_budget, _mutation_cancel, mutation_context) =
        execution_context_with_memory_and_work(source_capacity * 16, 16);
    document.with_execution_context(mutation_context).unwrap();
    assert_eq!(admission_budget.used(Resource::Memory), 0);
    assert_eq!(mutation_budget.used(Resource::Memory), source_capacity);
    let before = document.to_bytes();

    // There is no bookmark to rewrite, so this path only scans comments and
    // whitespace before refusing; it cannot succeed by reserving a candidate.
    let error = document
        .set_bookmark_rdfa("missing-bookmark", &RdfaAttributes::default())
        .err()
        .unwrap();
    assert_work_limit(error);
    assert_eq!(document.as_bytes(), before.as_slice());
    assert_eq!(
        FlatDocument::from_bytes(before.clone()).unwrap().as_bytes(),
        before.as_slice()
    );
    assert_eq!(mutation_budget.used(Resource::Memory), source_capacity);
    drop(document);
    assert_eq!(mutation_budget.used(Resource::Memory), 0);
}

#[test]
fn many_small_rdfa_nodes_cannot_evade_context_memory_accounting() {
    let count = 512;
    let source = many_metadata_nodes(count);
    let source_len = u64::try_from(source.len()).unwrap();
    let logical_nodes =
        u64::try_from(count.checked_mul(size_of::<RdfaOccurrence>()).unwrap()).unwrap();
    let memory_limit = source_len
        .checked_mul(4)
        .and_then(|value| value.checked_add(logical_nodes))
        .and_then(|value| value.checked_sub(1))
        .unwrap();
    let (budget, _cancellation, context) = execution_context_with_memory(memory_limit);
    let mut document =
        FlatDocument::from_bytes_with_execution_context(source.clone(), source_len + 1024, context)
            .unwrap();
    assert_eq!(document.in_content_metadata().unwrap().rdfa.len(), count);
    let before = document.to_bytes();

    let error = document
        .set_paragraph_rdfa(Position::new(0), &mutator_rdfa_value())
        .err()
        .unwrap();
    assert_memory_limit(error);
    assert_eq!(document.as_bytes(), before.as_slice());
    assert_eq!(budget.used(Resource::Memory), source_len);
    drop(document);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn many_small_xforms_models_cannot_evade_context_memory_accounting() {
    let count = 256;
    let source = many_xforms_models(count);
    let source_len = u64::try_from(source.len()).unwrap();
    let logical_models =
        u64::try_from(count.checked_mul(size_of::<XFormsModel>()).unwrap()).unwrap();
    let memory_limit = source_len
        .checked_mul(7)
        .and_then(|value| value.checked_add(logical_models))
        .and_then(|value| value.checked_sub(1))
        .unwrap();
    let (budget, _cancellation, context) = execution_context_with_memory(memory_limit);
    let mut document =
        FlatDocument::from_bytes_with_execution_context(source.clone(), source_len + 1024, context)
            .unwrap();
    assert_eq!(document.xforms_models().unwrap().len(), count);
    let before = document.to_bytes();
    let mut replacement = document.xforms_models().unwrap().remove(0);
    replacement.id = Some("replacement-model".to_string());

    let error = document
        .replace_xforms_model(Position::new(0), &replacement)
        .err()
        .unwrap();
    assert_memory_limit(error);
    assert_eq!(document.as_bytes(), before.as_slice());
    assert_eq!(budget.used(Resource::Memory), source_len);
    drop(document);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn generic_mutators_keep_the_canonical_text_template_classification() {
    let value = mutator_rdfa_value();

    let mut document = FlatDocument::from_bytes(MUTATOR_TEMPLATE_XML.as_bytes().to_vec()).unwrap();
    document
        .set_paragraph_rdfa(Position::new(0), &value)
        .unwrap();
    assert_template_metadata(&document);

    let mut document = FlatDocument::from_bytes(MUTATOR_TEMPLATE_XML.as_bytes().to_vec()).unwrap();
    document.set_bookmark_rdfa("mark", &value).unwrap();
    assert_template_metadata(&document);

    let mut document = FlatDocument::from_bytes(MUTATOR_TEMPLATE_XML.as_bytes().to_vec()).unwrap();
    document
        .set_text_meta_rdfa(Position::new(0), &value)
        .unwrap();
    assert_template_metadata(&document);

    let mut document = FlatDocument::from_bytes(MUTATOR_TEMPLATE_XML.as_bytes().to_vec()).unwrap();
    document
        .insert_text_meta(Position::new(0), &TextMeta::from_text("inserted").unwrap())
        .unwrap();
    assert_template_metadata(&document);

    let mut document = FlatDocument::from_bytes(MUTATOR_TEMPLATE_XML.as_bytes().to_vec()).unwrap();
    document
        .replace_text_meta(
            Position::new(0),
            &TextMeta::from_text("replacement").unwrap(),
        )
        .unwrap();
    assert_template_metadata(&document);

    let mut document = FlatDocument::from_bytes(MUTATOR_TEMPLATE_XML.as_bytes().to_vec()).unwrap();
    document.remove_text_meta(Position::new(0)).unwrap();
    assert_template_metadata(&document);

    let mut document = FlatDocument::from_bytes(MUTATOR_TEMPLATE_XML.as_bytes().to_vec()).unwrap();
    let mut model = document.xforms_models().unwrap().remove(0);
    model.id = Some("changed-model".to_string());
    document
        .replace_xforms_model(Position::new(0), &model)
        .unwrap();
    assert_template_metadata(&document);

    let mut document = FlatDocument::from_bytes(MUTATOR_TEMPLATE_XML.as_bytes().to_vec()).unwrap();
    let mut model = document.xforms_models().unwrap().remove(0);
    model.id = Some("inserted-model".to_string());
    model.children.clear();
    document.insert_xforms_model(&model).unwrap();
    assert_template_metadata(&document);

    let mut document = FlatDocument::from_bytes(MUTATOR_TEMPLATE_XML.as_bytes().to_vec()).unwrap();
    document.remove_xforms_model(Position::new(0)).unwrap();
    assert_template_metadata(&document);

    let group = Group {
        kind: Kind::Simple,
        part: Part::Flat,
        scope: Scope::Body(Body::Text),
        declarations: vec![Declaration::Simple {
            name: "counter".to_string(),
            value_type: ValueType::String,
        }],
    };
    let mut document = FlatDocument::from_bytes(MUTATOR_TEMPLATE_XML.as_bytes().to_vec()).unwrap();
    assert_eq!(
        document.set_variable_declaration_group(&group).unwrap(),
        None
    );
    assert_template_metadata(&document);
    assert_eq!(
        document
            .remove_variable_declaration_group(&group.scope, group.kind)
            .unwrap(),
        Some(group)
    );
    assert_template_metadata(&document);
}

#[test]
fn growing_generic_mutators_honor_exact_caller_output_caps() {
    let mutators: &[(&str, FlatMutator)] = &[
        ("paragraph RDFa", apply_paragraph_rdfa),
        ("bookmark RDFa", apply_bookmark_rdfa),
        ("text:meta RDFa", apply_text_meta_rdfa),
        ("text:meta insertion", apply_insert_text_meta),
        ("text:meta replacement", apply_replace_text_meta),
        ("XForms replacement", apply_replace_xforms_model),
        ("XForms insertion", apply_insert_xforms_model),
        ("variable declaration insertion", apply_variable_declaration),
    ];
    for &(name, mutator) in mutators {
        assert_growing_mutator_output_limit(name, mutator);
    }
}

#[test]
fn malformed_and_non_text_template_flat_inputs_are_refused() {
    for malformed in [
        b"<o:document xmlns:o=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" o:mimetype=\"application/vnd.oasis.opendocument.text-template\"><o:body><o:text/></o:body>".to_vec(),
        b"not XML".to_vec(),
        b"PK\x03\x04mimetype".to_vec(),
    ] {
        assert!(FlatDocument::from_bytes(malformed.clone()).is_err());
        assert!(Document::from_bytes(malformed).is_err());
    }

    for mimetype in [
        ODF_SPREADSHEET_TEMPLATE,
        ODF_PRESENTATION_TEMPLATE,
        ODF_DRAWING_TEMPLATE,
        ODF_CHART_TEMPLATE,
        ODF_FORMULA_TEMPLATE,
        ODF_IMAGE_TEMPLATE,
        ODF_MASTER_TEMPLATE,
        ODF_WEB,
    ] {
        let bytes = flat_xml(mimetype, "text");
        assert!(
            FlatDocument::from_bytes(bytes.clone()).is_err(),
            "accepted {mimetype}"
        );
        assert!(Document::from_bytes(bytes).is_err(), "accepted {mimetype}");
    }

    let wrong_body = flat_xml(ODF_TEXT_TEMPLATE, "spreadsheet");
    assert!(FlatDocument::from_bytes(wrong_body.clone()).is_err());
    assert!(Document::from_bytes(wrong_body).is_err());
}

#[test]
fn flat_template_ingress_accepts_exact_cap_and_rejects_cap_plus_one() {
    let exact = u64::try_from(TEMPLATE_XML.len()).unwrap();
    assert!(FlatDocument::from_bytes_with_limit(template_bytes(), exact).is_ok());
    assert!(FlatDocument::from_reader_with_limit(Cursor::new(template_bytes()), exact).is_ok());

    let mut oversized = template_bytes();
    oversized.push(b'x');
    assert!(matches!(
        FlatDocument::from_bytes_with_limit(oversized.clone(), exact),
        Err(Error::ResourceLimit(_))
    ));
    assert!(matches!(
        FlatDocument::from_reader_with_limit(Cursor::new(oversized), exact),
        Err(Error::ResourceLimit(_))
    ));

    let limits = Limits::new()
        .with_max_document_bytes(TEMPLATE_XML.len())
        .unwrap();
    assert!(Document::from_bytes_with_limits(template_bytes(), limits).is_ok());
    assert!(Document::from_reader_with_limits(Cursor::new(template_bytes()), limits).is_ok());

    let mut oversized = template_bytes();
    oversized.push(b'x');
    assert!(matches!(
        Document::from_bytes_with_limits(oversized.clone(), limits),
        Err(Error::ResourceLimit(_))
    ));
    assert!(matches!(
        Document::from_reader_with_limits(Cursor::new(oversized), limits),
        Err(Error::ResourceLimit(_))
    ));
}

#[test]
fn specialized_template_edit_no_op_change_inverse_and_stale_source_are_atomic() {
    let source = Document::from_bytes(template_bytes()).unwrap();
    let no_op = source.edit().commit().unwrap();
    assert_eq!(no_op.document().as_bytes(), source.as_bytes());
    assert_template_metadata(no_op.document());
    let no_op_applied = no_op.patch().apply(&source).unwrap();
    assert_eq!(no_op_applied.as_bytes(), source.as_bytes());
    assert_template_metadata(&no_op_applied);

    let mut edit = source.edit();
    edit.update_paragraph(0, "Omega").unwrap().unwrap();
    let commit = edit.commit().unwrap();
    assert_template_metadata(commit.document());
    assert!(commit.document().xml().contains("<t:p>Omega</t:p>"));
    assert_eq!(source.as_bytes(), TEMPLATE_XML.as_bytes());

    let reverted = commit.patch().inverse().apply(commit.document()).unwrap();
    assert_eq!(reverted.as_bytes(), source.as_bytes());
    assert_template_metadata(&reverted);

    let normal = Document::from_bytes(normal_xml().into_bytes()).unwrap();
    assert!(!normal.is_template());
    assert!(commit.patch().apply(&normal).is_err());
    assert_eq!(normal.as_bytes(), normal_xml().as_bytes());

    let mut structured = source.edit();
    assert!(structured.update_paragraph(1, "discard markup").is_err());
    assert_eq!(source.as_bytes(), TEMPLATE_XML.as_bytes());
}

#[test]
fn specialized_template_output_limit_accepts_exact_output_and_rejects_one_over() {
    let exact_limits = Limits::new()
        .with_max_document_bytes(TEMPLATE_XML.len())
        .unwrap();
    let source = Document::from_bytes_with_limits(template_bytes(), exact_limits).unwrap();
    let mut exact_edit = source.edit();
    exact_edit.update_paragraph(0, "Omega").unwrap().unwrap();
    let exact_commit = exact_edit.commit().unwrap();
    assert_template_metadata(exact_commit.document());
    assert_eq!(exact_commit.document().as_bytes().len(), TEMPLATE_XML.len());

    let mut over_edit = source.edit();
    over_edit.update_paragraph(0, "Omega!").unwrap().unwrap();
    assert!(matches!(over_edit.commit(), Err(Error::ResourceLimit(_))));
    assert_eq!(source.as_bytes(), TEMPLATE_XML.as_bytes());
}

#[test]
fn formatted_template_no_op_is_exact_but_changed_edit_keeps_compactness_refusal() {
    let formatted = TEMPLATE_XML.replace("<o:body>", "\n<o:body>");
    let source = Document::from_bytes(formatted.as_bytes().to_vec()).unwrap();
    let no_op = source.edit().commit().unwrap();
    assert_eq!(no_op.document().as_bytes(), formatted.as_bytes());
    assert_template_metadata(no_op.document());

    let mut edit = source.edit();
    edit.update_paragraph(0, "Omega").unwrap().unwrap();
    assert!(matches!(edit.commit(), Err(Error::XmlCompactness { .. })));
    assert_eq!(source.as_bytes(), formatted.as_bytes());
}

#[test]
fn generic_and_specialized_template_saves_are_exact_and_suffix_neutral() {
    let directory = tempfile::tempdir().unwrap();
    let generic = FlatDocument::from_bytes(template_bytes()).unwrap();
    let generic_path = directory.path().join("template.fodt");
    generic.save(&generic_path).unwrap();
    assert_eq!(
        std::fs::read(&generic_path).unwrap(),
        TEMPLATE_XML.as_bytes()
    );
    assert_template_metadata(&FlatDocument::open(&generic_path).unwrap());

    let specialized = Document::from_bytes(template_bytes()).unwrap();
    let specialized_path = directory.path().join("template.fott");
    specialized.save(&specialized_path).unwrap();
    assert_eq!(
        std::fs::read(&specialized_path).unwrap(),
        TEMPLATE_XML.as_bytes()
    );
    assert_template_metadata(&Document::open(&specialized_path).unwrap());

    let normal = FlatDocument::from_bytes(normal_xml().into_bytes()).unwrap();
    let normal_path = directory.path().join("normal.fott");
    normal.save(&normal_path).unwrap();
    let reopened = FlatDocument::open(&normal_path).unwrap();
    assert_normal_metadata(&reopened);

    let normal_specialized = Document::from_bytes(normal_xml().into_bytes()).unwrap();
    let normal_specialized_path = directory.path().join("normal-specialized.fott");
    normal_specialized.save(&normal_specialized_path).unwrap();
    assert_normal_metadata(&Document::open(&normal_specialized_path).unwrap());
}

#[test]
fn cancelled_generic_and_specialized_saves_preserve_destinations_and_temps() {
    let directory = tempfile::tempdir().unwrap();
    let foreign_temp = directory.path().join("foreign.tmp");
    std::fs::write(&foreign_temp, b"foreign").unwrap();

    let generic_destination = directory.path().join("cancelled-generic.fott");
    std::fs::write(&generic_destination, b"generic-old").unwrap();
    let generic = FlatDocument::from_bytes(template_bytes()).unwrap();
    let (generic_budget, generic_cancel, generic_context) = execution_context_with_memory(u64::MAX);
    generic_cancel.cancel();
    let error = generic
        .save_with_execution_context(&generic_destination, generic_context)
        .err()
        .unwrap();
    assert_cancelled(error);
    assert_eq!(std::fs::read(&generic_destination).unwrap(), b"generic-old");
    assert_eq!(std::fs::read(&foreign_temp).unwrap(), b"foreign");
    assert_no_owned_flat_temporary_files(directory.path());
    assert_eq!(generic_budget.used(Resource::Memory), 0);

    let specialized_destination = directory.path().join("cancelled-specialized.fott");
    std::fs::write(&specialized_destination, b"specialized-old").unwrap();
    let specialized = Document::from_bytes(template_bytes()).unwrap();
    let (specialized_budget, specialized_cancel, specialized_context) =
        execution_context_with_memory(u64::MAX);
    specialized_cancel.cancel();
    let error = specialized
        .save_with_execution_context(&specialized_destination, specialized_context)
        .err()
        .unwrap();
    assert_cancelled(error);
    assert_eq!(
        std::fs::read(&specialized_destination).unwrap(),
        b"specialized-old"
    );
    assert_eq!(std::fs::read(&foreign_temp).unwrap(), b"foreign");
    assert_no_owned_flat_temporary_files(directory.path());
    assert_eq!(specialized_budget.used(Resource::Memory), 0);
}

#[test]
fn generic_and_specialized_saves_refuse_directory_destinations_without_mutation() {
    let directory = tempfile::tempdir().unwrap();
    let generic_destination = directory.path().join("generic-blocked.fott");
    std::fs::create_dir(&generic_destination).unwrap();
    let generic_marker = generic_destination.join("marker");
    std::fs::write(&generic_marker, b"preserve").unwrap();
    assert!(
        FlatDocument::from_bytes(template_bytes())
            .unwrap()
            .save(&generic_destination)
            .is_err()
    );
    assert_eq!(std::fs::read(&generic_marker).unwrap(), b"preserve");

    let specialized_destination = directory.path().join("specialized-blocked.fott");
    std::fs::create_dir(&specialized_destination).unwrap();
    let specialized_marker = specialized_destination.join("marker");
    std::fs::write(&specialized_marker, b"preserve").unwrap();
    assert!(
        Document::from_bytes(template_bytes())
            .unwrap()
            .save(&specialized_destination)
            .is_err()
    );
    assert_eq!(std::fs::read(&specialized_marker).unwrap(), b"preserve");
}

#[test]
fn packaged_ott_template_behavior_remains_package_owned() {
    let mut writer = PackageWriter::new();
    writer.set_mimetype(ODF_TEXT_TEMPLATE).unwrap();
    writer
        .add_file(
            ODF_CONTENT,
            format!(
                r#"<o:document-content xmlns:o="{OFFICE}" xmlns:t="{TEXT}"><o:body><o:text><t:p>Package template</t:p></o:text></o:body></o:document-content>"#
            )
            .as_bytes(),
        )
        .unwrap();
    let bytes = writer.finish_to_bytes().unwrap();

    let package = Package::from_bytes(bytes.clone()).unwrap();
    assert_eq!(package.family(), Family::Text);
    assert!(package.is_template());
    assert_eq!(package.mimetype(), ODF_TEXT_TEMPLATE);
    assert_eq!(package.to_bytes(), bytes);
    assert!(FlatDocument::from_bytes(bytes.clone()).is_err());
    assert!(Document::from_bytes(bytes).is_err());
}
