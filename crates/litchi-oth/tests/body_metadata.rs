#![allow(
    clippy::unwrap_used,
    reason = "fixture tests use panic-on-failure assertions"
)]

use litchi_core::Position;
use litchi_oth::{Builder, Patch, Template, change, index, table};

const XML_PREFIX: &str = concat!(
    r##"<?xml version="1.0" encoding="UTF-8"?><office:document-content "##,
    r##"xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" "##,
    r##"xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" "##,
    r##"xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" "##,
    r##"xmlns:xlink="http://www.w3.org/1999/xlink" "##,
    r##"xmlns:xhtml="http://www.w3.org/1999/xhtml" "##,
    r##"xmlns:dc="http://purl.org/dc/elements/1.1/" "##,
    r##"xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" "##,
    r##"xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" "##,
    r##"xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" "##,
    r##"xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" "##,
    r##"xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" "##,
    r##"xmlns:foreign="urn:example:foreign" "##,
    r##"xmlns:xml="http://www.w3.org/XML/1998/namespace" office:version="1.4">"##,
    r##"<office:body><office:text>"##,
);
const XML_SUFFIX: &str = r##"</office:text></office:body></office:document-content>"##;

const BODY: &str = concat!(
    r##"<text:tracked-changes text:track-changes="false"><text:changed-region xml:id="change-forward"><text:insertion><office:change-info><dc:creator>Ada</dc:creator><dc:date>2026-09-02T03:04:05Z</dc:date><text:p>insert reason</text:p><text:p>second reason</text:p></office:change-info></text:insertion></text:changed-region><text:changed-region xml:id="change-delete" text:id="change-delete"><text:deletion><office:change-info><dc:creator>Bob</dc:creator><dc:date>2026-09-03T04:05:06Z</dc:date></office:change-info><text:p>deleted body</text:p></text:deletion></text:changed-region><text:changed-region xml:id="change-format"><text:format-change><office:change-info><dc:creator>Cy</dc:creator><dc:date>2026-09-04T05:06:07Z</dc:date></office:change-info></text:format-change></text:changed-region></text:tracked-changes>"##,
    r##"<text:change-start text:change-id="change-forward"/>"##,
    r##"<!-- keep-before --><text:section text:name="EditSection" xml:id="section-id"><!-- section-comment --><?section keep?><text:p>editable</text:p><foreign:opaque foreign:flag="x &amp; y"><![CDATA[opaque &lt;bytes&gt;]]></foreign:opaque></text:section>"##,
    r##"<office:annotation office:name="comment" office:display="false" draw:caption="Comment &amp; caption" draw:style-name="CommentStyle" svg:x="1cm" svg:y="2cm" svg:width="3cm" svg:height="4cm"><dc:creator>Ada &amp; Co</dc:creator><dc:date>2026-09-01T01:02:03Z</dc:date><meta:date-string>September &amp; 1</meta:date-string><meta:creator-initials>A&amp;C</meta:creator-initials><text:list><text:list-item><text:p>rich &amp; list</text:p></text:list-item></text:list></office:annotation>"##,
    r##"<table:table table:name="Data" table:style-name="DataStyle" table:template-name="Template &amp; Name" table:use-first-row-styles="false" table:use-last-row-styles="true" table:use-first-column-styles="false" table:use-last-column-styles="true" table:use-banding-rows-styles="false" table:use-banding-columns-styles="true" table:protected="true" table:protection-key="key &amp; one" table:protection-key-digest-algorithm="urn:example:digest?x=1&amp;y=2" table:print="false" table:print-ranges="$Data.$A$1:$B$2 arbitrary &amp; range" xml:id="table-id" table:is-sub-table="true"><!-- table-comment --><?table keep?><table:table-source table:mode="copy-results-only" table:table-name="External" xlink:type="simple" xlink:href="https://example.test/data?a=1&amp;b=2" xlink:actuate="onRequest" table:filter-name="Filter &amp; Name" table:filter-options="option=&amp;raw" table:refresh-delay="P1Y2M3DT4H5M6.789S"/><table:table-column-group table:display="false"><table:table-column table:number-columns-repeated="2" table:style-name="ColumnStyle" table:visibility="collapse" table:default-cell-style-name="DefaultCell" xml:id="column-collapse"/></table:table-column-group><table:table-column table:style-name="VisibleColumn" table:visibility="visible" xml:id="column-visible"/><table:table-column table:style-name="FilterColumn" table:visibility="filter" xml:id="column-filter"/><table:table-row-group table:display="false"><table:table-row table:number-rows-repeated="2" table:style-name="RowStyle" table:default-cell-style-name="RowDefault" table:visibility="collapse" xml:id="row-collapse"><table:table-cell table:number-columns-repeated="2" table:style-name="FloatCell" table:content-validation-name="AmountValidation" table:formula="of:=SUM([.A1:.B1])" office:value-type="float" office:value="1.2300" table:protect="false" table:protected="true" xml:id="cell-float" table:number-columns-spanned="2" table:number-rows-spanned="3" table:number-matrix-columns-spanned="4" table:number-matrix-rows-spanned="5" xhtml:about="#item" xhtml:property="schema:price dc:title" xhtml:datatype="xsd:decimal" xhtml:content="1.2300"><text:p>one point two three</text:p></table:table-cell><table:table-cell office:value-type="percentage" office:value="25.00"><text:p>25%</text:p></table:table-cell><table:table-cell office:value-type="currency" office:value="-7.500" office:currency="USD"><text:p>USD</text:p></table:table-cell><table:table-cell office:value-type="date" office:date-value="2024-02-03T04:05:06Z"><text:p>date</text:p></table:table-cell><table:table-cell office:value-type="time" office:time-value="P1Y2M3DT4H5M6.789S"><text:p>time</text:p></table:table-cell><table:table-cell office:value-type="boolean" office:boolean-value="false"><text:p>false</text:p></table:table-cell><table:table-cell office:value-type="string" office:string-value="&amp;escaped"><text:p>visible string</text:p></table:table-cell><table:table-cell office:value-type="error" office:string-value="#N/A"><text:p>error</text:p></table:table-cell><table:covered-table-cell><text:p>covered</text:p></table:covered-table-cell></table:table-row></table:table-row-group><table:table-row table:style-name="VisibleRow" table:visibility="visible" xml:id="row-visible"><table:table-cell><text:p>visible row</text:p></table:table-cell></table:table-row><table:table-row table:style-name="FilterRow" table:visibility="filter" xml:id="row-filter"><table:table-cell><text:p>filter row</text:p></table:table-cell></table:table-row></table:table>"##,
    r##"<table:table table:name="Sparse"><table:table-column/><table:table-row><table:table-cell><text:p>sparse</text:p></table:table-cell></table:table-row></table:table>"##,
    r##"<text:table-of-content text:name="Contents" text:style-name="IndexStyle" text:protected="false" text:protection-key="index-key" text:protection-key-digest-algorithm="urn:example:index-digest" xml:id="index-toc"><text:table-of-content-source text:outline-level="3" text:use-outline-level="true" text:use-index-marks="false" text:use-index-source-styles="true" text:index-scope="chapter" text:relative-tab-stop-position="false"><text:index-title-template><text:p>TOC template</text:p></text:index-title-template></text:table-of-content-source><text:index-body><text:p>cached toc</text:p></text:index-body></text:table-of-content>"##,
    r##"<text:illustration-index text:name="Illustrations"><text:illustration-index-source text:use-caption="true" text:caption-sequence-name="Figure" text:caption-sequence-format="category-and-value" text:index-scope="document" text:relative-tab-stop-position="true"/><text:index-body><text:p>cached illustrations</text:p></text:index-body></text:illustration-index>"##,
    r##"<text:table-index text:name="Tables"><text:table-index-source text:use-caption="false" text:caption-sequence-name="Table" text:caption-sequence-format="caption" text:index-scope="chapter" text:relative-tab-stop-position="false"/><text:index-body><text:p>cached tables</text:p></text:index-body></text:table-index>"##,
    r##"<text:object-index text:name="Objects"><text:object-index-source text:use-spreadsheet-objects="true" text:use-math-objects="false" text:use-draw-objects="true" text:use-chart-objects="false" text:use-other-objects="true" text:index-scope="document" text:relative-tab-stop-position="true"/><text:index-body><text:p>cached objects</text:p></text:index-body></text:object-index>"##,
    r##"<text:user-index text:name="UserIndexRoot"><text:user-index-source text:use-index-marks="true" text:use-index-source-styles="false" text:use-graphics="true" text:use-tables="false" text:use-floating-frames="true" text:use-objects="false" text:copy-outline-levels="true" text:index-name="UserMarks" text:index-scope="chapter" text:relative-tab-stop-position="false"/><text:index-body><text:p>cached user</text:p></text:index-body></text:user-index>"##,
    r##"<text:alphabetical-index text:name="Alphabetical"><text:alphabetical-index-source text:ignore-case="true" text:main-entry-style-name="MainEntry" text:alphabetical-separators="false" text:combine-entries="true" text:combine-entries-with-dash="false" text:combine-entries-with-pp="true" text:use-keys-as-entries="false" text:capitalize-entries="true" text:comma-separated="false" fo:language="en" fo:country="US" fo:script="Latn" style:rfc-language-tag="en-US" text:sort-algorithm="unicode" text:index-scope="document" text:relative-tab-stop-position="true"/><text:index-body><text:p>cached alphabetical</text:p></text:index-body></text:alphabetical-index>"##,
    r##"<text:bibliography text:name="Bibliography"><text:bibliography-source/><text:index-body><text:p>cached bibliography</text:p></text:index-body></text:bibliography>"##,
    r##"<text:change-end text:change-id="change-forward"/><text:change text:change-id="change-delete"/>"##,
    r##"<foreign:wrapper><table:table table:name="foreign-table"><table:table-row><table:table-cell><text:p>foreign table</text:p></table:table-cell></table:table-row></table:table><text:table-of-content text:name="foreign-index"><text:table-of-content-source/><text:index-body/></text:table-of-content><text:changed-region xml:id="foreign-change"><text:insertion><office:change-info><dc:creator>foreign</dc:creator><dc:date>2026-01-01T00:00:00Z</dc:date></office:change-info></text:insertion></text:changed-region></foreign:wrapper><?keep-after?>"##,
);

fn content(body: &str) -> String {
    let mut xml = String::with_capacity(XML_PREFIX.len() + body.len() + XML_SUFFIX.len());
    xml.push_str(XML_PREFIX);
    xml.push_str(body);
    xml.push_str(XML_SUFFIX);
    xml
}

fn template(body: &str) -> Template {
    Template::from_bytes(Builder::new().content_xml(content(body)).build().unwrap()).unwrap()
}

fn template_with_escaped_table_alias(body: &str) -> Template {
    let prefix = XML_PREFIX.replace(
        r#"xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" "#,
        r#"xmlns:tbl="urn:oasis:names&#x3A;tc&#x3A;opendocument&#x3A;xmlns&#x3A;table&#x3A;1.0" "#,
    );
    let body = body.replace("table:", "tbl:");
    let mut xml = String::with_capacity(prefix.len() + body.len() + XML_SUFFIX.len());
    xml.push_str(&prefix);
    xml.push_str(&body);
    xml.push_str(XML_SUFFIX);
    Template::from_bytes(Builder::new().content_xml(xml).build().unwrap()).unwrap()
}

fn table_by_name<'a>(tables: &'a [table::Table], name: &str) -> &'a table::Table {
    tables
        .iter()
        .find(|table| table.name() == Some(name))
        .unwrap()
}

fn index_by_name<'a>(indexes: &'a [index::Index], name: &str) -> &'a index::Index {
    indexes
        .iter()
        .find(|index| index.name() == Some(name))
        .unwrap()
}

#[test]
fn table_metadata_covers_scalars_source_layout_and_typed_values() {
    let source = template(BODY);
    let body = source.text_body().unwrap();
    let tables = body.tables().unwrap();
    assert_eq!(tables.len(), 2);

    let data = table_by_name(tables, "Data");
    assert_eq!(data.style_name(), Some("DataStyle"));
    let properties = data.properties();
    assert_eq!(properties.template_name(), Some("Template & Name"));
    assert_eq!(properties.use_first_row_styles(), Some(false));
    assert_eq!(properties.use_last_row_styles(), Some(true));
    assert_eq!(properties.use_first_column_styles(), Some(false));
    assert_eq!(properties.use_last_column_styles(), Some(true));
    assert_eq!(properties.use_banding_rows_styles(), Some(false));
    assert_eq!(properties.use_banding_columns_styles(), Some(true));
    assert_eq!(properties.protected(), Some(true));
    assert_eq!(properties.protection_key(), Some("key & one"));
    assert_eq!(
        properties.protection_key_digest_algorithm(),
        Some("urn:example:digest?x=1&y=2")
    );
    assert_eq!(properties.print(), Some(false));
    assert_eq!(
        properties.print_ranges(),
        Some("$Data.$A$1:$B$2 arbitrary & range")
    );
    assert_eq!(properties.xml_id(), Some("table-id"));
    assert_eq!(properties.is_sub_table(), Some(true));

    let source_metadata = data.source().unwrap();
    assert_eq!(
        source_metadata.mode(),
        Some(table::TableSourceMode::CopyResultsOnly)
    );
    assert_eq!(source_metadata.table_name(), Some("External"));
    assert_eq!(source_metadata.href(), "https://example.test/data?a=1&b=2");
    assert_eq!(
        source_metadata.actuate(),
        Some(table::TableSourceActuate::OnRequest)
    );
    assert_eq!(source_metadata.filter_name(), Some("Filter & Name"));
    assert_eq!(source_metadata.filter_options(), Some("option=&raw"));
    assert_eq!(
        source_metadata.refresh_delay().unwrap().as_str(),
        "P1Y2M3DT4H5M6.789S"
    );

    let copy_all = template(&BODY.replace(
        "table:mode=\"copy-results-only\"",
        "table:mode=\"copy-all\"",
    ));
    assert_eq!(
        table_by_name(copy_all.text_body().unwrap().tables().unwrap(), "Data")
            .source()
            .unwrap()
            .mode(),
        Some(table::TableSourceMode::CopyAll)
    );

    assert_eq!(data.columns().len(), 3);
    assert_eq!(data.declared_column_count(), 4);
    assert_eq!(
        data.columns()[0].visibility(),
        Some(table::Visibility::Collapse)
    );
    assert_eq!(data.columns()[0].xml_id(), Some("column-collapse"));
    assert_eq!(data.columns()[0].declared_repeat_count(), Some(2));
    assert_eq!(data.columns()[0].repeat_count(), 2);
    assert_eq!(
        data.columns()[1].visibility(),
        Some(table::Visibility::Visible)
    );
    assert_eq!(data.columns()[1].xml_id(), Some("column-visible"));
    assert_eq!(data.columns()[1].declared_repeat_count(), None);
    assert_eq!(data.columns()[1].repeat_count(), 1);
    assert_eq!(
        data.columns()[2].visibility(),
        Some(table::Visibility::Filter)
    );
    assert_eq!(data.columns()[2].xml_id(), Some("column-filter"));

    assert_eq!(data.rows().len(), 3);
    let first_row = &data.rows()[0];
    assert_eq!(first_row.default_cell_style_name(), Some("RowDefault"));
    assert_eq!(first_row.visibility(), Some(table::Visibility::Collapse));
    assert_eq!(first_row.xml_id(), Some("row-collapse"));
    assert_eq!(first_row.declared_repeat_count(), Some(2));
    assert_eq!(first_row.repeat_count(), 2);
    assert_eq!(
        data.rows()[1].visibility(),
        Some(table::Visibility::Visible)
    );
    assert_eq!(data.rows()[2].visibility(), Some(table::Visibility::Filter));

    let explicit_one = template(&BODY.replace(
        "table:visibility=\"visible\" xml:id=\"column-visible\"",
        "table:visibility=\"visible\" table:number-columns-repeated=\"1\" xml:id=\"column-visible\"",
    ));
    assert_eq!(
        table_by_name(explicit_one.text_body().unwrap().tables().unwrap(), "Data").columns()[1]
            .declared_repeat_count(),
        Some(1)
    );

    let cells = first_row.cells();
    assert_eq!(cells.len(), 9);
    let float = &cells[0];
    assert_eq!(float.content_validation_name(), Some("AmountValidation"));
    assert_eq!(float.protect(), Some(false));
    assert_eq!(float.protected(), Some(true));
    assert_eq!(float.xml_id(), Some("cell-float"));
    assert_eq!(float.declared_repeat_count(), Some(2));
    assert_eq!(float.repeat_count(), 2);
    assert_eq!(float.declared_columns_spanned(), Some(2));
    assert_eq!(float.declared_rows_spanned(), Some(3));
    assert_eq!(float.matrix_columns_spanned(), Some(4));
    assert_eq!(float.matrix_rows_spanned(), Some(5));
    assert_eq!(float.columns_spanned(), 2);
    assert_eq!(float.rows_spanned(), 3);
    assert_eq!(float.typed_value().unwrap().lexical(), Some("1.2300"));
    assert_eq!(float.typed_value().unwrap().boolean_value(), None);
    let rdfa = float.in_content_meta().unwrap();
    assert_eq!(rdfa.about(), "#item");
    assert_eq!(rdfa.property(), "schema:price dc:title");
    assert_eq!(rdfa.datatype(), Some("xsd:decimal"));
    assert_eq!(rdfa.content(), Some("1.2300"));

    let percentage = cells[1].typed_value().unwrap();
    assert_eq!(percentage.lexical(), Some("25.00"));
    assert_eq!(percentage.currency(), None);
    let currency = cells[2].typed_value().unwrap();
    assert_eq!(currency.lexical(), Some("-7.500"));
    assert_eq!(currency.currency(), Some("USD"));
    let date = cells[3].typed_value().unwrap();
    assert_eq!(date.lexical(), Some("2024-02-03T04:05:06Z"));
    let time = cells[4].typed_value().unwrap();
    assert_eq!(time.duration().unwrap().as_str(), "P1Y2M3DT4H5M6.789S");
    let boolean = cells[5].typed_value().unwrap();
    assert_eq!(boolean.lexical(), Some("false"));
    assert_eq!(boolean.boolean_value(), Some(false));
    let string = cells[6].typed_value().unwrap();
    assert_eq!(string.lexical(), None);
    assert_eq!(string.string_value(), Some("&escaped"));
    let error = cells[7].typed_value().unwrap();
    assert_eq!(error.lexical(), None);
    assert_eq!(error.error_value(), Some("#N/A"));
    assert!(matches!(cells[8].kind(), table::CellKind::Covered));

    let sparse = table_by_name(tables, "Sparse");
    assert!(sparse.source().is_none());
    assert_eq!(sparse.properties().template_name(), None);
    assert_eq!(sparse.properties().protected(), None);
    assert_eq!(sparse.properties().print(), None);
    assert_eq!(sparse.properties().xml_id(), None);
    assert_eq!(sparse.columns()[0].visibility(), None);
    assert_eq!(sparse.columns()[0].xml_id(), None);
    assert_eq!(sparse.columns()[0].declared_repeat_count(), None);
    assert_eq!(sparse.columns()[0].repeat_count(), 1);
    assert_eq!(sparse.rows()[0].visibility(), None);
    assert_eq!(sparse.rows()[0].xml_id(), None);
    assert_eq!(sparse.rows()[0].declared_repeat_count(), None);
    let sparse_cell = &sparse.rows()[0].cells()[0];
    assert_eq!(sparse_cell.declared_repeat_count(), None);
    assert_eq!(sparse_cell.declared_columns_spanned(), None);
    assert_eq!(sparse_cell.declared_rows_spanned(), None);
    assert_eq!(sparse_cell.matrix_columns_spanned(), None);
    assert_eq!(sparse_cell.matrix_rows_spanned(), None);
    assert_eq!(sparse_cell.repeat_count(), 1);
    assert_eq!(sparse_cell.columns_spanned(), 1);
    assert_eq!(sparse_cell.rows_spanned(), 1);
}

#[test]
fn entity_escaped_table_namespace_alias_projects_the_same_table_family() {
    let source = template_with_escaped_table_alias(BODY);
    let body = source.text_body().unwrap();
    let tables = body.tables().unwrap();
    assert_eq!(tables.len(), 2);
    assert_eq!(
        table_by_name(tables, "Data").source().unwrap().table_name(),
        Some("External")
    );
    assert_eq!(
        table_by_name(tables, "Sparse").rows()[0].cells()[0].text(),
        "sparse"
    );
}

#[test]
fn xsd_double_special_lexicals_are_accepted_and_retained() {
    let source = template(
        r#"<table:table table:name="SpecialDoubles"><table:table-row><table:table-cell office:value-type="float" office:value="INF"/><table:table-cell office:value-type="percentage" office:value="-INF"/><table:table-cell office:value-type="currency" office:value="NaN" office:currency="USD"/></table:table-row></table:table>"#,
    );
    let body = source.text_body().unwrap();
    let table = table_by_name(body.tables().unwrap(), "SpecialDoubles");
    let cells = table.rows()[0].cells();
    assert_eq!(cells[0].typed_value().unwrap().lexical(), Some("INF"));
    assert_eq!(cells[1].typed_value().unwrap().lexical(), Some("-INF"));
    assert_eq!(cells[2].typed_value().unwrap().lexical(), Some("NaN"));
    assert_eq!(cells[2].typed_value().unwrap().currency(), Some("USD"));
}

#[test]
fn index_metadata_covers_all_admitted_source_families_and_cached_text() {
    let source = template(BODY);
    let body = source.text_body().unwrap();
    let indexes = body.indexes().unwrap();
    assert_eq!(indexes.len(), 7);

    let toc = index_by_name(indexes, "Contents");
    assert_eq!(toc.style_name(), Some("IndexStyle"));
    assert_eq!(toc.protected(), Some(false));
    assert_eq!(toc.protection_key(), Some("index-key"));
    assert_eq!(
        toc.protection_key_digest_algorithm(),
        Some("urn:example:index-digest")
    );
    assert_eq!(toc.xml_id(), Some("index-toc"));
    assert_eq!(toc.source(), Some("TOC template"));
    assert_eq!(toc.body(), "cached toc");
    assert!(matches!(toc.kind(), index::Kind::TableOfContents));
    let toc_source = toc.source_options().unwrap();
    assert_eq!(toc_source.scope(), Some(index::IndexScope::Chapter));
    assert_eq!(toc_source.relative_tab_stop_position(), Some(false));
    match toc_source.options() {
        index::IndexSourceOptions::TableOfContents {
            outline_level,
            use_outline_level,
            use_index_marks,
            use_index_source_styles,
        } => {
            assert_eq!(*outline_level, Some(3));
            assert_eq!(*use_outline_level, Some(true));
            assert_eq!(*use_index_marks, Some(false));
            assert_eq!(*use_index_source_styles, Some(true));
        },
        other => panic!("unexpected TOC source options: {other:?}"),
    }

    let illustration = index_by_name(indexes, "Illustrations");
    assert!(matches!(illustration.kind(), index::Kind::Illustration));
    assert_eq!(illustration.style_name(), None);
    assert_eq!(illustration.protected(), None);
    assert_eq!(illustration.protection_key(), None);
    assert_eq!(illustration.protection_key_digest_algorithm(), None);
    assert_eq!(illustration.xml_id(), None);
    assert!(!illustration.is_protected());
    assert_eq!(illustration.source(), Some(""));
    assert_eq!(illustration.body(), "cached illustrations");
    let illustration_source = illustration.source_options().unwrap();
    assert_eq!(
        illustration_source.scope(),
        Some(index::IndexScope::Document)
    );
    assert_eq!(illustration_source.relative_tab_stop_position(), Some(true));
    match illustration_source.options() {
        index::IndexSourceOptions::Illustration {
            use_caption,
            caption_sequence_name,
            caption_sequence_format,
        } => {
            assert_eq!(*use_caption, Some(true));
            assert_eq!(caption_sequence_name.as_deref(), Some("Figure"));
            assert_eq!(
                *caption_sequence_format,
                Some(index::CaptionSequenceFormat::CategoryAndValue)
            );
        },
        other => panic!("unexpected illustration source options: {other:?}"),
    }

    let text_format = template(&BODY.replace(
        "text:caption-sequence-format=\"category-and-value\"",
        "text:caption-sequence-format=\"text\"",
    ));
    match index_by_name(
        text_format.text_body().unwrap().indexes().unwrap(),
        "Illustrations",
    )
    .source_options()
    .unwrap()
    .options()
    {
        index::IndexSourceOptions::Illustration {
            caption_sequence_format,
            ..
        } => assert_eq!(
            *caption_sequence_format,
            Some(index::CaptionSequenceFormat::Text)
        ),
        other => panic!("unexpected text-format source options: {other:?}"),
    }

    let table_index = index_by_name(indexes, "Tables");
    assert!(matches!(table_index.kind(), index::Kind::Table));
    match table_index.source_options().unwrap().options() {
        index::IndexSourceOptions::Illustration {
            use_caption,
            caption_sequence_name,
            caption_sequence_format,
        } => {
            assert_eq!(*use_caption, Some(false));
            assert_eq!(caption_sequence_name.as_deref(), Some("Table"));
            assert_eq!(
                *caption_sequence_format,
                Some(index::CaptionSequenceFormat::Caption)
            );
        },
        other => panic!("unexpected table source options: {other:?}"),
    }

    let object = index_by_name(indexes, "Objects");
    assert!(matches!(object.kind(), index::Kind::Object));
    match object.source_options().unwrap().options() {
        index::IndexSourceOptions::Object {
            use_spreadsheet_objects,
            use_math_objects,
            use_draw_objects,
            use_chart_objects,
            use_other_objects,
        } => {
            assert_eq!(*use_spreadsheet_objects, Some(true));
            assert_eq!(*use_math_objects, Some(false));
            assert_eq!(*use_draw_objects, Some(true));
            assert_eq!(*use_chart_objects, Some(false));
            assert_eq!(*use_other_objects, Some(true));
        },
        other => panic!("unexpected object source options: {other:?}"),
    }

    let user = index_by_name(indexes, "UserIndexRoot");
    assert!(matches!(user.kind(), index::Kind::User));
    match user.source_options().unwrap().options() {
        index::IndexSourceOptions::User {
            use_index_marks,
            use_index_source_styles,
            use_graphics,
            use_tables,
            use_floating_frames,
            use_objects,
            copy_outline_levels,
            index_name,
        } => {
            assert_eq!(*use_index_marks, Some(true));
            assert_eq!(*use_index_source_styles, Some(false));
            assert_eq!(*use_graphics, Some(true));
            assert_eq!(*use_tables, Some(false));
            assert_eq!(*use_floating_frames, Some(true));
            assert_eq!(*use_objects, Some(false));
            assert_eq!(*copy_outline_levels, Some(true));
            assert_eq!(index_name, "UserMarks");
        },
        other => panic!("unexpected user source options: {other:?}"),
    }

    let alphabetical = index_by_name(indexes, "Alphabetical");
    assert!(matches!(alphabetical.kind(), index::Kind::Alphabetical));
    match alphabetical.source_options().unwrap().options() {
        index::IndexSourceOptions::Alphabetical {
            ignore_case,
            main_entry_style_name,
            alphabetical_separators,
            combine_entries,
            combine_entries_with_dash,
            combine_entries_with_pp,
            use_keys_as_entries,
            capitalize_entries,
            comma_separated,
            language,
            country,
            script,
            rfc_language_tag,
            sort_algorithm,
        } => {
            assert_eq!(*ignore_case, Some(true));
            assert_eq!(main_entry_style_name.as_deref(), Some("MainEntry"));
            assert_eq!(*alphabetical_separators, Some(false));
            assert_eq!(*combine_entries, Some(true));
            assert_eq!(*combine_entries_with_dash, Some(false));
            assert_eq!(*combine_entries_with_pp, Some(true));
            assert_eq!(*use_keys_as_entries, Some(false));
            assert_eq!(*capitalize_entries, Some(true));
            assert_eq!(*comma_separated, Some(false));
            assert_eq!(language.as_deref(), Some("en"));
            assert_eq!(country.as_deref(), Some("US"));
            assert_eq!(script.as_deref(), Some("Latn"));
            assert_eq!(rfc_language_tag.as_deref(), Some("en-US"));
            assert_eq!(sort_algorithm.as_deref(), Some("unicode"));
        },
        other => panic!("unexpected alphabetical source options: {other:?}"),
    }

    let bibliography = index_by_name(indexes, "Bibliography");
    assert!(matches!(bibliography.kind(), index::Kind::Bibliography));
    assert!(matches!(
        bibliography.source_options().unwrap().options(),
        index::IndexSourceOptions::Bibliography
    ));
    assert_eq!(bibliography.body(), "cached bibliography");
}

#[test]
fn annotation_scalars_and_rich_opaque_content_remain_readable_and_source_owned() {
    let source = template(BODY);
    let body = source.text_body().unwrap();
    let annotations = body.annotations().unwrap();
    assert_eq!(annotations.len(), 1);
    let annotation = &annotations[0];
    assert_eq!(annotation.name(), Some("comment"));
    assert_eq!(annotation.creator(), Some("Ada & Co"));
    assert_eq!(annotation.date(), Some("2026-09-01T01:02:03Z"));
    assert_eq!(annotation.date_string(), Some("September & 1"));
    assert_eq!(annotation.initials(), Some("A&C"));
    assert_eq!(annotation.display(), Some(false));
    assert_eq!(annotation.text(), "rich & list");
    assert!(
        source
            .content_xml()
            .contains("draw:caption=\"Comment &amp; caption\"")
    );
    assert!(source.content_xml().contains("<text:list><text:list-item>"));
    assert!(
        source
            .content_xml()
            .contains("<![CDATA[opaque &lt;bytes&gt;]]>")
    );
}

#[test]
fn tracked_change_metadata_keeps_exact_ids_markers_and_change_info() {
    let source = template(BODY);
    let body = source.text_body().unwrap();
    let tracking = body.change_tracking().unwrap().unwrap();
    assert_eq!(tracking.track_changes(), Some(false));

    let changes = body.changes().unwrap();
    assert_eq!(changes.len(), 6);

    let insertion = changes
        .iter()
        .find(|change| matches!(change.kind(), change::Kind::Insertion))
        .unwrap();
    assert_eq!(insertion.xml_id(), Some("change-forward"));
    assert_eq!(insertion.region_id(), None);
    assert_eq!(insertion.marker_change_id(), None);
    assert_eq!(insertion.id(), Some("change-forward"));
    assert_eq!(insertion.text(), "");
    let insertion_info = insertion.info().unwrap();
    assert_eq!(insertion_info.creator(), "Ada");
    assert_eq!(insertion_info.date(), "2026-09-02T03:04:05Z");
    assert_eq!(
        insertion_info
            .paragraphs()
            .iter()
            .map(|paragraph| paragraph.text())
            .collect::<Vec<_>>(),
        ["insert reason", "second reason"]
    );

    let deletion = changes
        .iter()
        .find(|change| matches!(change.kind(), change::Kind::Deletion))
        .unwrap();
    assert_eq!(deletion.xml_id(), Some("change-delete"));
    assert_eq!(deletion.region_id(), Some("change-delete"));
    assert_eq!(deletion.marker_change_id(), None);
    assert_eq!(deletion.text(), "deleted body");
    assert_eq!(deletion.info().unwrap().creator(), "Bob");

    let format_change = changes
        .iter()
        .find(|change| matches!(change.kind(), change::Kind::Format))
        .unwrap();
    assert_eq!(format_change.xml_id(), Some("change-format"));
    assert_eq!(format_change.region_id(), None);
    assert_eq!(format_change.text(), "");
    assert_eq!(format_change.info().unwrap().creator(), "Cy");

    let start = changes
        .iter()
        .find(|change| matches!(change.kind(), change::Kind::Start))
        .unwrap();
    assert_eq!(start.xml_id(), None);
    assert_eq!(start.region_id(), None);
    assert_eq!(start.marker_change_id(), Some("change-forward"));
    assert_eq!(start.id(), Some("change-forward"));
    assert!(start.info().is_none());

    let end = changes
        .iter()
        .find(|change| matches!(change.kind(), change::Kind::End))
        .unwrap();
    assert_eq!(end.marker_change_id(), Some("change-forward"));
    assert!(end.info().is_none());

    let point = changes
        .iter()
        .find(|change| matches!(change.kind(), change::Kind::Other(value) if value == "change"))
        .unwrap();
    assert_eq!(point.marker_change_id(), Some("change-delete"));
    assert!(point.info().is_none());
}

#[test]
fn tracked_change_declaration_absence_differs_from_present_without_flag() {
    let absent = template(r##"<text:p>plain</text:p>"##);
    let absent_body = absent.text_body().unwrap();
    assert!(absent_body.change_tracking().unwrap().is_none());
    assert!(absent_body.changes().unwrap().is_empty());

    let present_without_flag = template(&BODY.replace(" text:track-changes=\"false\"", ""));
    let present_body = present_without_flag.text_body().unwrap();
    let tracking = present_body.change_tracking().unwrap().unwrap();
    assert_eq!(tracking.track_changes(), None);
    assert_eq!(present_body.changes().unwrap().len(), 6);

    let present_true = template(&BODY.replace(
        "text:track-changes=\"false\"",
        "text:track-changes=\"true\"",
    ));
    assert_eq!(
        present_true
            .text_body()
            .unwrap()
            .change_tracking()
            .unwrap()
            .unwrap()
            .track_changes(),
        Some(true)
    );
}

#[test]
fn malformed_metadata_refuses_typed_projection_without_publishing_partial_values() {
    let invalid_table_values = [
        BODY.replace("table:print=\"false\"", "table:print=\"maybe\""),
        BODY.replace(
            "table:visibility=\"collapse\"",
            "table:visibility=\"hidden\"",
        ),
        BODY.replace(
            "table:number-columns-repeated=\"2\"",
            "table:number-columns-repeated=\"0\"",
        ),
        BODY.replace(
            "office:value-type=\"float\" office:value=\"1.2300\"",
            "office:value-type=\"float\"",
        ),
        BODY.replace(
            "xhtml:property=\"schema:price dc:title\"",
            "xhtml:property=\"bad^property\"",
        ),
        BODY.replace(
            "xhtml:datatype=\"xsd:decimal\"",
            "xhtml:datatype=\"bad^datatype\"",
        ),
        BODY.replace("xlink:type=\"simple\"", "xlink:type=\"extended\""),
        BODY.replace(
            "xlink:href=\"https://example.test/data?a=1&amp;b=2\"",
            "xlink:href=\"https://example.test/bad&#x20;href\"",
        ),
        BODY.replace(" xlink:href=\"https://example.test/data?a=1&amp;b=2\"", ""),
        BODY.replace(
            "table:protection-key-digest-algorithm=\"urn:example:digest?x=1&amp;y=2\"",
            "table:protection-key-digest-algorithm=\"bad&#x20;digest\"",
        ),
        BODY.replace(
            "office:date-value=\"2024-02-03T04:05:06Z\"",
            "office:date-value=\"not-a-date\"",
        ),
        BODY.replace(
            "office:time-value=\"P1Y2M3DT4H5M6.789S\"",
            "office:time-value=\"not-a-duration\"",
        ),
        BODY.replace("xhtml:about=\"#item\"", "xhtml:about=\"bad&#x20;about\""),
        BODY.replace("xml:id=\"column-visible\"", "xml:id=\"table-id\""),
    ];
    for (case, invalid_content) in invalid_table_values.into_iter().enumerate() {
        let source = template(&invalid_content);
        let body = source.text_body().unwrap();
        assert!(body.tables().is_err(), "invalid table case {case}");
        assert!(body.tables().is_err(), "invalid table case {case} retry");
    }

    let invalid_index = template(&BODY.replace(
        "text:use-outline-level=\"true\"",
        "text:use-outline-level=\"maybe\"",
    ));
    let invalid_index_body = invalid_index.text_body().unwrap();
    assert!(invalid_index_body.indexes().is_err());
    assert!(invalid_index_body.indexes().is_err());

    let invalid_index_scope = template(&BODY.replace(
        "text:index-scope=\"chapter\"",
        "text:index-scope=\"section\"",
    ));
    assert!(invalid_index_scope.text_body().unwrap().indexes().is_err());

    let invalid_caption_format = template(&BODY.replace(
        "text:caption-sequence-format=\"category-and-value\"",
        "text:caption-sequence-format=\"bad-format\"",
    ));
    assert!(
        invalid_caption_format
            .text_body()
            .unwrap()
            .indexes()
            .is_err()
    );

    let missing_user_index_name = template(&BODY.replace(" text:index-name=\"UserMarks\"", ""));
    assert!(
        missing_user_index_name
            .text_body()
            .unwrap()
            .indexes()
            .is_err()
    );

    let unresolved_marker = template(&BODY.replace(
        "text:change-id=\"change-forward\"",
        "text:change-id=\"missing-region\"",
    ));
    assert!(unresolved_marker.text_body().unwrap().changes().is_err());

    let missing_xml_id = template(&BODY.replace(" xml:id=\"change-forward\"", ""));
    assert!(missing_xml_id.text_body().unwrap().changes().is_err());

    let text_id_only_region = template(&BODY.replace(
        "xml:id=\"change-delete\" text:id=\"change-delete\"",
        "text:id=\"change-delete\"",
    ));
    assert!(text_id_only_region.text_body().unwrap().changes().is_err());

    let text_id_only_legacy_shape = template(
        r##"<text:tracked-changes><text:changed-region text:id="legacy"><text:insertion><text:p>legacy</text:p></text:insertion></text:changed-region></text:tracked-changes>"##,
    );
    assert!(
        text_id_only_legacy_shape
            .text_body()
            .unwrap()
            .changes()
            .is_err()
    );

    let mismatched_deprecated_id =
        template(&BODY.replace("text:id=\"change-delete\"", "text:id=\"different-region\""));
    assert!(
        mismatched_deprecated_id
            .text_body()
            .unwrap()
            .changes()
            .is_err()
    );

    let invalid_marker_lexical = template(&BODY.replace(
        "text:change-id=\"change-delete\"",
        "text:change-id=\"not an IDREF\"",
    ));
    assert!(
        invalid_marker_lexical
            .text_body()
            .unwrap()
            .changes()
            .is_err()
    );

    let missing_change_info_creator = template(&BODY.replace(
        "<dc:creator>Ada</dc:creator><dc:date>2026-09-02T03:04:05Z</dc:date>",
        "<dc:date>2026-09-02T03:04:05Z</dc:date>",
    ));
    assert!(
        missing_change_info_creator
            .text_body()
            .unwrap()
            .changes()
            .is_err()
    );

    let invalid_change_info_order = template(&BODY.replace(
        "<dc:creator>Ada</dc:creator><dc:date>2026-09-02T03:04:05Z</dc:date><text:p>insert reason</text:p>",
        "<dc:date>2026-09-02T03:04:05Z</dc:date><dc:creator>Ada</dc:creator><text:p>insert reason</text:p>",
    ));
    assert!(
        invalid_change_info_order
            .text_body()
            .unwrap()
            .changes()
            .is_err()
    );
}

#[test]
fn changed_region_shape_rejects_extra_siblings_nonleading_info_and_payload() {
    let cases = [
        (
            "changed-region extra sibling",
            BODY.replace(
                "</text:insertion></text:changed-region>",
                "</text:insertion><text:p>illegal sibling</text:p></text:changed-region>",
            ),
        ),
        (
            "change-info not first",
            BODY.replace(
                r#"<text:deletion><office:change-info><dc:creator>Bob</dc:creator><dc:date>2026-09-03T04:05:06Z</dc:date></office:change-info><text:p>deleted body</text:p></text:deletion>"#,
                r#"<text:deletion><text:p>deleted body</text:p><office:change-info><dc:creator>Bob</dc:creator><dc:date>2026-09-03T04:05:06Z</dc:date></office:change-info></text:deletion>"#,
            ),
        ),
        (
            "insertion payload",
            BODY.replace(
                "</office:change-info></text:insertion></text:changed-region>",
                "</office:change-info><text:p>illegal insertion payload</text:p></text:insertion></text:changed-region>",
            ),
        ),
        (
            "format-change payload",
            BODY.replace(
                "</office:change-info></text:format-change></text:changed-region>",
                "</office:change-info><text:p>illegal format payload</text:p></text:format-change></text:changed-region>",
            ),
        ),
    ];
    for (case, invalid_body) in cases {
        assert!(
            template(&invalid_body)
                .text_body()
                .unwrap()
                .changes()
                .is_err(),
            "invalid changed-region shape: {case}"
        );
    }
}

#[test]
fn cell_value_type_rejects_incompatible_companion_attributes() {
    let cases = [
        (
            "float with date-value",
            BODY.replace(
                r#"office:value-type="float" office:value="1.2300""#,
                r#"office:value-type="float" office:value="1.2300" office:date-value="2026-01-01""#,
            ),
        ),
        (
            "percentage with boolean-value",
            BODY.replace(
                r#"office:value-type="percentage" office:value="25.00""#,
                r#"office:value-type="percentage" office:value="25.00" office:boolean-value="true""#,
            ),
        ),
        (
            "currency with time-value",
            BODY.replace(
                r#"office:value-type="currency" office:value="-7.500" office:currency="USD""#,
                r#"office:value-type="currency" office:value="-7.500" office:currency="USD" office:time-value="PT1S""#,
            ),
        ),
        (
            "date with value",
            BODY.replace(
                r#"office:value-type="date" office:date-value="2024-02-03T04:05:06Z""#,
                r#"office:value-type="date" office:date-value="2024-02-03T04:05:06Z" office:value="1""#,
            ),
        ),
        (
            "boolean with value",
            BODY.replace(
                r#"office:value-type="boolean" office:boolean-value="false""#,
                r#"office:value-type="boolean" office:boolean-value="false" office:value="0""#,
            ),
        ),
        (
            "string with value",
            BODY.replace(
                r#"office:value-type="string" office:string-value="&amp;escaped""#,
                r#"office:value-type="string" office:string-value="&amp;escaped" office:value="1""#,
            ),
        ),
        (
            "error with boolean-value",
            BODY.replace(
                r##"office:value-type="error" office:string-value="#N/A""##,
                r##"office:value-type="error" office:string-value="#N/A" office:boolean-value="true""##,
            ),
        ),
    ];
    for (case, invalid_body) in cases {
        assert!(
            template(&invalid_body)
                .text_body()
                .unwrap()
                .tables()
                .is_err(),
            "invalid cell value companion: {case}"
        );
    }
}

#[test]
fn duplicate_expanded_attributes_refuse_projection() {
    let duplicate = BODY.replace(
        r#"<table:table table:name="Data""#,
        r#"<table:table xmlns:tbl="urn:oasis:names:tc:opendocument:xmlns:table:1.0" table:name="Data" tbl:name="Duplicate""#,
    );
    assert!(template(&duplicate).text_body().unwrap().tables().is_err());
}

#[test]
fn duplicate_table_and_index_metadata_children_refuse_projection() {
    let duplicate_table_source = BODY.replace(
        r#"table:refresh-delay="P1Y2M3DT4H5M6.789S"/><table:table-column-group"#,
        r#"table:refresh-delay="P1Y2M3DT4H5M6.789S"/><table:table-source xlink:type="simple" xlink:href="https://example.test/second"/><table:table-column-group"#,
    );
    assert!(
        template(&duplicate_table_source)
            .text_body()
            .unwrap()
            .tables()
            .is_err()
    );

    let duplicate_index_source = BODY.replace(
        r#"<text:illustration-index-source text:use-caption="true" text:caption-sequence-name="Figure" text:caption-sequence-format="category-and-value" text:index-scope="document" text:relative-tab-stop-position="true"/><text:index-body>"#,
        r#"<text:illustration-index-source text:use-caption="true" text:caption-sequence-name="Figure" text:caption-sequence-format="category-and-value" text:index-scope="document" text:relative-tab-stop-position="true"/><text:illustration-index-source/><text:index-body>"#,
    );
    assert!(
        template(&duplicate_index_source)
            .text_body()
            .unwrap()
            .indexes()
            .is_err()
    );

    let duplicate_index_body = BODY.replace(
        r#"<text:illustration-index-source text:use-caption="true" text:caption-sequence-name="Figure" text:caption-sequence-format="category-and-value" text:index-scope="document" text:relative-tab-stop-position="true"/><text:index-body><text:p>cached illustrations</text:p></text:index-body></text:illustration-index>"#,
        r#"<text:illustration-index-source text:use-caption="true" text:caption-sequence-name="Figure" text:caption-sequence-format="category-and-value" text:index-scope="document" text:relative-tab-stop-position="true"/><text:index-body><text:p>cached illustrations</text:p></text:index-body><text:index-body/></text:illustration-index>"#,
    );
    assert!(
        template(&duplicate_index_body)
            .text_body()
            .unwrap()
            .indexes()
            .is_err()
    );
}

fn date_value_table_body(value: &str) -> String {
    format!(
        r#"<table:table table:name="Dates"><table:table-row><table:table-cell office:value-type="date" office:date-value="{value}"/></table:table-row></table:table>"#
    )
}

#[test]
fn xsd_date_and_datetime_lexicals_follow_the_odf_grammar() {
    for (label, value) in [
        ("date with UTC timezone", "2024-01-01Z"),
        ("date-time at midnight boundary", "2024-01-01T24:00:00Z"),
        (
            "midnight with an all-zero fraction",
            "2024-01-01T24:00:00.00000000000000000000Z",
        ),
        (
            "date-time with arbitrary fractional precision",
            "2024-01-01T00:00:00.12345678901234567890Z",
        ),
    ] {
        let source = template(&date_value_table_body(value));
        let body = source.text_body().unwrap();
        let tables = body.tables().unwrap();
        let cell = &tables[0].rows()[0].cells()[0];
        assert_eq!(
            cell.typed_value().unwrap().lexical(),
            Some(value),
            "valid date lexical rejected or normalized: {label}"
        );
    }

    let change_date = "2024-01-01T00:00:00.12345678901234567890Z";
    let source = template(&BODY.replace("2026-09-02T03:04:05Z", change_date));
    let body = source.text_body().unwrap();
    let insertion = body
        .changes()
        .unwrap()
        .iter()
        .find(|change| matches!(change.kind(), change::Kind::Insertion))
        .unwrap();
    assert_eq!(insertion.info().unwrap().date(), change_date);
}

#[test]
fn xsd_date_and_datetime_lexicals_refuse_year_zero_bad_timezone_and_invalid_fraction() {
    for (label, value) in [
        ("year zero date", "0000-01-01Z"),
        (
            "date timezone beyond XML Schema boundary",
            "2024-01-01+14:01",
        ),
        (
            "nonzero fraction at 24:00:00",
            "2024-01-01T24:00:00.00000000000000000001Z",
        ),
    ] {
        let source = template(&date_value_table_body(value));
        assert!(
            source.text_body().unwrap().tables().is_err(),
            "invalid date lexical was accepted: {label}"
        );
    }

    for (label, value) in [
        ("year zero date-time", "0000-01-01T00:00:00Z"),
        (
            "date-time timezone beyond XML Schema boundary",
            "2024-01-01T00:00:00+14:01",
        ),
        (
            "nonzero fraction at date-time midnight boundary",
            "2024-01-01T24:00:00.1Z",
        ),
    ] {
        let source = template(&BODY.replace("2026-09-02T03:04:05Z", value));
        assert!(
            source.text_body().unwrap().changes().is_err(),
            "invalid change-info date lexical was accepted: {label}"
        );
    }
}

#[test]
fn stray_direct_changed_region_is_rejected_outside_the_unique_tracking_prelude() {
    const STRAY_AND_MARKER: &str = concat!(
        r#"<text:changed-region xml:id="stray-direct"><text:insertion><office:change-info><dc:creator>Stray</dc:creator><dc:date>2024-01-01T00:00:00Z</dc:date></office:change-info></text:insertion></text:changed-region>"#,
        r#"<text:change-start text:change-id="change-forward"/>"#,
    );
    let body = BODY.replace(
        r#"<text:change-start text:change-id="change-forward"/>"#,
        STRAY_AND_MARKER,
    );
    let source = template(&body);
    let body = source.text_body().unwrap();
    assert!(
        body.changes().is_err(),
        "a direct office:text region outside text:tracked-changes was projected"
    );
}

#[test]
fn tracked_change_declarations_require_one_direct_non_nested_owner() {
    let duplicate_direct = BODY.replace(
        r#"</text:tracked-changes><text:change-start"#,
        r#"</text:tracked-changes><text:tracked-changes/><text:change-start"#,
    );
    let nested = BODY.replace(
        r#"<text:tracked-changes text:track-changes="false">"#,
        r#"<text:tracked-changes text:track-changes="false"><text:tracked-changes/>"#,
    );

    for (label, invalid_body) in [
        ("duplicate direct declarations", duplicate_direct),
        ("nested declaration", nested),
    ] {
        let source = template(&invalid_body);
        let body = source.text_body().unwrap();
        assert!(
            body.change_tracking().is_err(),
            "accepted invalid tracked-change ownership: {label}"
        );
        assert!(
            body.changes().is_err(),
            "projected changes from invalid tracked-change ownership: {label}"
        );
    }
}

#[test]
fn office_forms_may_precede_the_tracking_prelude() {
    let with_forms = BODY.replace(
        r#"<text:tracked-changes text:track-changes="false">"#,
        r#"<office:forms/><text:tracked-changes text:track-changes="false">"#,
    );
    let source = template(&with_forms);
    let body = source.text_body().unwrap();
    assert!(body.forms().is_empty());
    assert_eq!(
        body.change_tracking().unwrap().unwrap().track_changes(),
        Some(false)
    );
    assert_eq!(body.changes().unwrap().len(), 6);
}

fn expanded_prelude_body() -> String {
    let with_forms = BODY.replace(
        r#"<text:tracked-changes text:track-changes="false">"#,
        r#"<office:forms/><text:tracked-changes text:track-changes="false">"#,
    );
    with_forms.replace(
        r#"</text:tracked-changes><text:change-start text:change-id="change-forward"/>"#,
        r#"</text:tracked-changes><text:variable-decls/><text:sequence-decls/><text:user-field-decls/><text:dde-connection-decls/><text:alphabetical-index-auto-mark-file xlink:type="simple" xlink:href="concordance.sdi"/><table:calculation-settings/><table:content-validations/><table:label-ranges/><text:change-start text:change-id="change-forward"/>"#,
    )
}

#[test]
fn expanded_office_text_prelude_preserves_all_distinct_declarations() {
    let source = template(&expanded_prelude_body());
    let body = source.text_body().unwrap();
    assert!(body.forms().is_empty());
    assert_eq!(
        body.change_tracking().unwrap().unwrap().track_changes(),
        Some(false)
    );
    assert_eq!(body.changes().unwrap().len(), 6);
    assert_eq!(body.tables().unwrap().len(), 2);
    assert_eq!(body.indexes().unwrap().len(), 7);
}

#[test]
fn expanded_office_text_prelude_rejects_reordering_and_duplicates() {
    let expanded = expanded_prelude_body();
    let cases = [
        (
            "reordered declaration pair",
            expanded.replace(
                "<text:sequence-decls/><text:user-field-decls/>",
                "<text:user-field-decls/><text:sequence-decls/>",
            ),
        ),
        (
            "duplicate declaration",
            expanded.replace(
                "<text:sequence-decls/>",
                "<text:sequence-decls/><text:sequence-decls/>",
            ),
        ),
    ];
    for (label, invalid_body) in cases {
        let source = template(&invalid_body);
        let body = source.text_body().unwrap();
        assert!(
            body.change_tracking().is_err(),
            "accepted invalid expanded prelude: {label}"
        );
    }
}

#[test]
fn late_tracking_declaration_after_body_content_is_refused() {
    let late = BODY.replace(
        r#"<text:tracked-changes text:track-changes="false">"#,
        r#"<text:p>content before tracked changes</text:p><text:tracked-changes text:track-changes="false">"#,
    );
    let source = template(&late);
    let body = source.text_body().unwrap();
    assert!(body.change_tracking().is_err());
    assert!(body.changes().is_err());
}

#[test]
fn no_op_and_inverse_keep_the_exact_metadata_source_and_opaque_markup() {
    let source = template(BODY);
    let source_xml = content(BODY);
    assert_eq!(source.content_xml(), source_xml);

    let no_op = source.edit().commit().unwrap();
    assert!(!no_op.changed());
    assert_eq!(no_op.template().as_bytes(), source.as_bytes());
    assert_eq!(no_op.template().content_xml(), source_xml);

    let mut edit = source.edit();
    edit.set_section_text(Position::new(0), "edited section")
        .unwrap();
    let commit = edit.commit().unwrap();
    assert!(commit.changed());
    assert!(commit.template().content_xml().contains("edited section"));
    assert!(commit.template().content_xml().contains("section-comment"));
    assert!(
        commit
            .template()
            .content_xml()
            .contains("<![CDATA[opaque &lt;bytes&gt;]]>")
    );
    assert!(commit.template().content_xml().contains("foreign:wrapper"));
    assert!(commit.template().content_xml().contains("TOC template"));
    assert!(commit.template().content_xml().contains("rich &amp; list"));
    assert!(
        commit
            .template()
            .content_xml()
            .contains("draw:caption=\"Comment &amp; caption\"")
    );

    let inverse = Patch::from_bytes(&commit.patch().inverse().to_bytes().unwrap()).unwrap();
    assert_eq!(
        inverse.apply(commit.template()).unwrap().as_bytes(),
        source.as_bytes()
    );
}

#[test]
fn foreign_same_named_roots_remain_inert_to_metadata_projection() {
    let source = template(BODY);
    let body = source.text_body().unwrap();
    assert_eq!(body.tables().unwrap().len(), 2);
    assert_eq!(body.indexes().unwrap().len(), 7);
    assert_eq!(body.changes().unwrap().len(), 6);
}

#[test]
fn native_oth_fixture_keeps_its_empty_metadata_boundary() {
    let template =
        Template::from_bytes(include_bytes!("fixtures/libreoffice-desktop-html.oth").to_vec())
            .unwrap();
    let body = template.text_body().unwrap();
    assert!(body.tables().unwrap().is_empty());
    assert!(body.indexes().unwrap().is_empty());
    assert!(body.changes().unwrap().is_empty());
    assert!(body.change_tracking().unwrap().is_none());
}
