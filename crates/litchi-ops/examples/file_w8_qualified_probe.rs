//! Corrected 0362h qualification probe for the unified `File` facade.
//!
//! Every workload route has an equivalent native leaf operation. Commit and
//! publication/save routes are timed separately; output validation is outside
//! the clock and checks semantic edited state, exact no-op/repeated bytes and
//! source identity. Every sample reports raw elapsed totals and operation
//! counts. The companion `docs/unified-ops/results/0362h/abba.sh` driver
//! supplies guarded outer ABBA legs.

#[cfg(all(
    feature = "docx",
    feature = "xlsx",
    feature = "xlsb",
    feature = "pptx",
    feature = "doc",
    feature = "xls",
    feature = "ppt",
    feature = "rtf",
    feature = "odt",
    feature = "ods",
    feature = "odp"
))]
mod probe {
    use std::hint::black_box;
    use std::io::{Cursor, Write};
    use std::time::Instant;

    use litchi_core::{BlobId, FileFormat, Position};
    use litchi_ops::merge::{Options as MergeOptions, Record};
    use litchi_ops::{BatchItem, BatchMode, File, NewNode, Props, open_bytes, props};

    const TARGET_NS: u128 = 500_000_000;
    const MAX_CALLS: usize = 2_000_000;
    const PARAGRAPHS: usize = 1_000;
    const SHAPES: usize = 100;

    fn line(index: usize) -> String {
        format!("Paragraph {index:05} holds a few ordinary words of body text.")
    }

    /// A word-processing file of `count` paragraphs through the unified
    /// `add` (DOCX, ODT); the first paragraph holds the merge placeholder.
    fn unified_words(format: FileFormat, count: usize) -> Vec<u8> {
        let mut file = litchi_ops::create(format).expect("create");
        for index in 0..count {
            let text = if index == 0 {
                "Dear {{name}},".to_owned()
            } else {
                line(index)
            };
            file = file
                .add(
                    "/body",
                    NewNode::new("paragraph").prop("text", text.as_str()),
                )
                .expect("paragraph")
                .file;
        }
        file.to_bytes().expect("bytes")
    }

    fn doc(count: usize) -> Vec<u8> {
        let mut builder = litchi_doc::Builder::new();
        builder.add_paragraph("Dear {{name}},").expect("paragraph");
        for index in 1..count {
            builder.add_paragraph(&line(index)).expect("paragraph");
        }
        builder.build().expect("build").to_bytes().expect("bytes")
    }

    /// A DOCX with a mergeable body and a real cached `NUMCHARS` field after
    /// the 1,000 ordinary paragraphs.  Keeping the field outside the normal
    /// edit paths lets batch/set and merge exercise the same simple paragraph
    /// targets while refresh still performs an observable cache update.
    fn docx_refresh(count: usize) -> Vec<u8> {
        const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        let mut body = String::from(r#"<w:p><w:r><w:t>Dear {{name}},</w:t></w:r></w:p>"#);
        for index in 1..count {
            body.push_str(&format!(
                r#"<w:p><w:r><w:t>{}</w:t></w:r></w:p>"#,
                line(index)
            ));
        }
        body.push_str(
            r#"<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> NUMCHARS </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>0</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>"#,
        );
        let document = format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="{W}"><w:body>{body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:left="1440" w:right="1440" w:top="1440" w:bottom="1440"/></w:sectPr></w:body></w:document>"#
        );
        let content_types = r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#;
        let root = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#;
        let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
        let options = zip::write::SimpleFileOptions::default();
        for (name, contents) in [
            ("[Content_Types].xml", content_types),
            ("_rels/.rels", root),
            ("word/document.xml", document.as_str()),
        ] {
            zip.start_file(name, options).expect("docx field member");
            zip.write_all(contents.as_bytes())
                .expect("docx field member bytes");
        }
        zip.finish().expect("docx field archive").into_inner()
    }

    fn docx_refresh_value(bytes: &[u8]) -> String {
        let package = litchi_docx::Package::from_bytes(bytes).expect("docx refresh fixture");
        let edit = package.edit().expect("docx refresh fixture edit");
        let fields = edit.fields().expect("docx refresh fixture fields");
        let (_, characters) =
            fields
                .paragraphs()
                .iter()
                .fold((0usize, 0usize), |(words, characters), paragraph| {
                    (
                        words + paragraph.text().split_whitespace().count(),
                        characters
                            + paragraph
                                .text()
                                .chars()
                                .filter(|character| !character.is_whitespace())
                                .count(),
                    )
                });
        characters.to_string()
    }

    fn rtf(count: usize) -> Vec<u8> {
        let mut body = vec![r"Dear \{\{name\}\},".to_owned()];
        body.extend((1..count).map(line));
        format!(r"{{\rtf1\ansi {}}}", body.join(r"\par ")).into_bytes()
    }

    /// A workbook whose first sheet holds A1:J100 and the placeholder in
    /// L1, through the unified `set`.
    fn workbook(format: FileFormat) -> Vec<u8> {
        litchi_ops::create(format)
            .expect("create")
            .set("/sheet[1]/A1:J100", props! { "value" => 7 })
            .expect("range")
            .set("/sheet[1]/L1", props! { "value" => "Invoice for {{name}}" })
            .expect("placeholder")
            .to_bytes()
            .expect("bytes")
    }

    fn pptx(shapes: usize) -> Vec<u8> {
        let mut builder = litchi_pptx::builder::Builder::new().expect("builder");
        let slide = builder
            .presentation()
            .expect("presentation")
            .add_slide()
            .expect("slide");
        for index in 0..shapes {
            let text = if index == 0 {
                "Dear {{name}}".to_owned()
            } else {
                format!("text-{index}")
            };
            slide.add_text_box(
                &text,
                i64::try_from(index).unwrap_or(0) * 100,
                0,
                914_400,
                914_400,
            );
        }
        builder.build().expect("build").to_bytes().expect("bytes")
    }

    fn ppt(shapes: usize) -> Vec<u8> {
        let mut builder = litchi_ppt::writer::Builder::new();
        let slide = builder.add_slide().expect("slide");
        for index in 0..shapes {
            let text = if index == 0 {
                "Dear {{name}}"
            } else {
                "before"
            };
            builder
                .add_textbox(slide, 10 + (index % 100) as i32, 10, 240, 40, text)
                .expect("textbox");
        }
        builder.build().expect("build").to_bytes().expect("bytes")
    }

    fn odp(shapes: usize) -> Vec<u8> {
        let mut file = litchi_ops::create(FileFormat::Odp).expect("create");
        // Keep every edited shape in its own scene. ODP's scene splice
        // planner rejects two changes that touch the same serialized scene
        // fragment, so putting all shapes on slide 1 would make a nominal
        // 100-item batch invalid rather than measuring distinct targets.
        for index in 0..shapes {
            if index > 0 {
                file = file.add("/", NewNode::new("slide")).expect("slide").file;
            }
            let text = if index == 0 {
                "Dear {{name}}".to_owned()
            } else {
                format!("text-{index}")
            };
            file = file
                .add(
                    format!("/slide[{}]", index + 1).as_str(),
                    NewNode::new("shape")
                        .prop("text", text.as_str())
                        .prop("width", "2cm")
                        .prop("height", "1cm"),
                )
                .expect("shape")
                .file;
        }
        file.to_bytes().expect("bytes")
    }

    struct Family {
        name: &'static str,
        bytes: Vec<u8>,
        paths: Vec<String>,
        set_props: Props<'static>,
        marker: &'static str,
        refresh_value: Option<String>,
        batch: usize,
        /// CSV rows of the import route; zero for no import.
        import: usize,
    }

    fn family_names() -> &'static [&'static str] {
        &[
            "docx", "odt", "doc", "rtf", "xlsx", "xlsb", "xls", "ods", "pptx", "ppt", "odp",
        ]
    }

    /// Build one fixture only after route selection.  The old probe built all
    /// eleven archives before filtering `LITCHI_PROBE_ONLY`, which made a
    /// focused route pay for unrelated expensive writers and could evict the
    /// source under test from the page cache.
    fn family(name: &str) -> Family {
        let paragraphs: Vec<String> = (1..=PARAGRAPHS)
            .map(|index| format!("/body/p[{index}]"))
            .collect();
        let cells: Vec<String> = (1..=100)
            .flat_map(|row| {
                ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J"]
                    .into_iter()
                    .map(move |column| format!("/sheet[1]/{column}{row}"))
            })
            .collect();
        let shapes: Vec<String> = (1..=SHAPES)
            .map(|index| format!("/slide[1]/shape[{index}]"))
            .collect();
        let text = Props::new().with("text", "Replaced text");
        let value = Props::new().with("value", 42);
        let word = |name, bytes, batch| Family {
            name,
            bytes,
            paths: paragraphs.clone(),
            set_props: text.clone(),
            marker: "/body/p[1]",
            refresh_value: None,
            batch,
            import: 0,
        };
        let sheet = |name, format, import| Family {
            name,
            bytes: workbook(format),
            paths: cells.clone(),
            set_props: value.clone(),
            marker: "/sheet[1]/L1",
            refresh_value: None,
            batch: 1_000,
            import,
        };
        let slides = |name, bytes| Family {
            name,
            bytes,
            paths: shapes.clone(),
            set_props: text.clone(),
            marker: "/slide[1]/shape[1]",
            refresh_value: None,
            batch: 1_000,
            import: 0,
        };
        match name {
            "docx" => {
                let bytes = docx_refresh(PARAGRAPHS);
                let refresh_value = docx_refresh_value(&bytes);
                Family {
                    name: "docx",
                    bytes,
                    paths: paragraphs,
                    set_props: text,
                    marker: "/body/p[1]",
                    refresh_value: Some(refresh_value),
                    batch: 1_000,
                    import: 0,
                }
            },
            "odt" => word("odt", unified_words(FileFormat::Odt, PARAGRAPHS), 1_000),
            "doc" => word("doc", doc(PARAGRAPHS), 256),
            // RTF's native edit profile admits at most 256 staged paragraph
            // operations; keep the facade batch within that same bound.
            "rtf" => word("rtf", rtf(PARAGRAPHS), 256),
            "xlsx" => sheet("xlsx", FileFormat::Xlsx, 1_000),
            // XLSB edits only admit already-indexed cell records, so keep its
            // native import within the 100-row fixture while retaining the
            // same CSV shape and destination cells as the facade route.
            "xlsb" => sheet("xlsb", FileFormat::Xlsb, 100),
            // The XLS import route takes ~13 s for 1,000 rows at base and
            // head alike (record 0362 §5), so it is timed on 100 rows.
            "xls" => sheet("xls", FileFormat::Xls, 100),
            "ods" => sheet("ods", FileFormat::Ods, 1_000),
            "pptx" => slides("pptx", pptx(SHAPES)),
            "ppt" => slides("ppt", ppt(SHAPES)),
            "odp" => {
                // Each shape is on its own slide, so the 100-item batch
                // exercises 100 distinct serialized scene fragments. The
                // earlier one-slide/100-shape attempt is retained as a
                // rejected workload receipt in the qualification record.
                Family {
                    name: "odp",
                    bytes: odp(SHAPES),
                    paths: (1..=SHAPES)
                        .map(|slide| format!("/slide[{slide}]/shape[1]"))
                        .collect(),
                    set_props: text.clone(),
                    marker: "/slide[1]/shape[1]",
                    refresh_value: None,
                    batch: SHAPES,
                    import: 0,
                }
            },
            other => panic!("unknown family {other}"),
        }
    }

    #[allow(
        clippy::large_enum_variant,
        reason = "Keep native owners inline so the probe adds no measured heap indirection."
    )]
    enum LeafPackage {
        Docx(litchi_docx::Package),
        Pptx(litchi_pptx::Package),
        Doc(litchi_doc::Package),
        Xlsx(litchi_xlsx::Package),
        Xlsb(litchi_xlsb::Package),
        Xls(litchi_xls::Package),
        Ppt(litchi_ppt::Package),
        Rtf(litchi_rtf::Package),
        Odt(litchi_odt::Package),
        Ods(litchi_ods::Package),
        Odp(litchi_odp::Package),
    }

    impl LeafPackage {
        fn open(name: &str, bytes: &[u8]) -> Self {
            match name {
                "docx" => {
                    Self::Docx(litchi_docx::Package::from_bytes(bytes).expect("leaf docx open"))
                },
                "pptx" => {
                    Self::Pptx(litchi_pptx::Package::from_bytes(bytes).expect("leaf pptx open"))
                },
                "doc" => Self::Doc(litchi_doc::Package::from_bytes(bytes).expect("leaf doc open")),
                "xlsx" => {
                    Self::Xlsx(litchi_xlsx::Package::from_bytes(bytes).expect("leaf xlsx open"))
                },
                "xlsb" => {
                    Self::Xlsb(litchi_xlsb::Package::from_bytes(bytes).expect("leaf xlsb open"))
                },
                "xls" => Self::Xls(litchi_xls::Package::from_bytes(bytes).expect("leaf xls open")),
                "ppt" => Self::Ppt(litchi_ppt::Package::from_bytes(bytes).expect("leaf ppt open")),
                "rtf" => Self::Rtf(litchi_rtf::Package::from_bytes(bytes).expect("leaf rtf open")),
                "odt" => Self::Odt(litchi_odt::Package::from_bytes(bytes).expect("leaf odt open")),
                "ods" => Self::Ods(litchi_ods::Package::from_bytes(bytes).expect("leaf ods open")),
                "odp" => Self::Odp(litchi_odp::Package::from_bytes(bytes).expect("leaf odp open")),
                other => panic!("unknown family {other}"),
            }
        }

        fn publish(&self) -> Vec<u8> {
            match self {
                Self::Docx(package) => package.to_bytes().expect("leaf docx publish"),
                Self::Pptx(package) => package.to_bytes().expect("leaf pptx publish"),
                Self::Doc(package) => {
                    let mut bytes = Vec::new();
                    package.write_to(&mut bytes).expect("leaf doc publish");
                    bytes
                },
                Self::Xlsx(package) => package.to_bytes().expect("leaf xlsx publish"),
                Self::Xlsb(package) => package.to_bytes().expect("leaf xlsb publish"),
                Self::Xls(package) => package.to_bytes().expect("leaf xls publish"),
                Self::Ppt(package) => package.to_bytes().expect("leaf ppt publish"),
                Self::Rtf(package) => package.to_bytes().expect("leaf rtf publish"),
                Self::Odt(package) => package.to_bytes().expect("leaf odt publish"),
                Self::Ods(package) => package.to_bytes().expect("leaf ods publish"),
                Self::Odp(package) => package.to_bytes().expect("leaf odp publish"),
            }
        }
    }

    #[derive(Clone, Copy, Debug, Eq, PartialEq)]
    enum Route {
        OpenGet,
        Save,
        RepeatedSave,
        SetCommit,
        SetPublication,
        BatchCommit,
        BatchPublication,
        ImportCommit,
        ImportPublication,
        MergeCommit,
        MergePublication,
        RefreshCommit,
        RefreshPublication,
    }

    impl Route {
        const fn suffix(self) -> &'static str {
            match self {
                Self::OpenGet => "open_get",
                Self::Save => "save",
                Self::RepeatedSave => "repeated_save",
                Self::SetCommit => "set_commit",
                Self::SetPublication => "set_publication",
                Self::BatchCommit => "batch_commit",
                Self::BatchPublication => "batch_publication",
                Self::ImportCommit => "import_commit",
                Self::ImportPublication => "import_publication",
                Self::MergeCommit => "merge_commit",
                Self::MergePublication => "merge_publication",
                Self::RefreshCommit => "refresh_commit",
                Self::RefreshPublication => "refresh_publication",
            }
        }

        const fn changed(self) -> bool {
            matches!(
                self,
                Self::SetCommit
                    | Self::SetPublication
                    | Self::BatchCommit
                    | Self::BatchPublication
                    | Self::ImportCommit
                    | Self::ImportPublication
                    | Self::MergeCommit
                    | Self::MergePublication
            )
        }

        const fn is_commit(self) -> bool {
            matches!(
                self,
                Self::SetCommit
                    | Self::BatchCommit
                    | Self::ImportCommit
                    | Self::MergeCommit
                    | Self::RefreshCommit
            )
        }

        const fn is_publication(self) -> bool {
            matches!(
                self,
                Self::Save
                    | Self::RepeatedSave
                    | Self::SetPublication
                    | Self::BatchPublication
                    | Self::ImportPublication
                    | Self::MergePublication
                    | Self::RefreshPublication
            )
        }

        const fn exact_noop(self) -> bool {
            matches!(self, Self::OpenGet | Self::Save | Self::RepeatedSave)
        }
    }

    fn routes_for(name: &str) -> Vec<Route> {
        let mut routes = vec![
            Route::OpenGet,
            Route::Save,
            Route::RepeatedSave,
            Route::SetCommit,
            Route::SetPublication,
            Route::BatchCommit,
            Route::BatchPublication,
            Route::MergeCommit,
            Route::MergePublication,
        ];
        if matches!(name, "xlsx" | "xlsb" | "xls" | "ods") {
            routes.extend([Route::ImportCommit, Route::ImportPublication]);
        }
        if name == "docx" {
            routes.extend([Route::RefreshCommit, Route::RefreshPublication]);
        }
        routes
    }

    fn route_for(name: &str, suffix: &str) -> Option<Route> {
        routes_for(name)
            .into_iter()
            .find(|route| route.suffix() == suffix)
    }

    fn all_route_names() -> Vec<String> {
        family_names()
            .iter()
            .flat_map(|name| {
                routes_for(name)
                    .into_iter()
                    .map(move |route| format!("{name}/{}", route.suffix()))
            })
            .collect()
    }

    #[derive(Clone)]
    struct Sample {
        elapsed_ns: u128,
        calls: usize,
        operation_count: usize,
    }

    fn elapsed_ns(start: Instant) -> u128 {
        start.elapsed().as_nanos()
    }

    fn target_ns() -> u128 {
        std::env::var("LITCHI_FILE_TARGET_NS")
            .ok()
            .and_then(|value| value.parse::<u128>().ok())
            .filter(|value| *value > 0)
            .unwrap_or(TARGET_NS)
    }

    fn headroom() -> f64 {
        std::env::var("LITCHI_FILE_HEADROOM")
            .ok()
            .and_then(|value| value.parse::<f64>().ok())
            .filter(|value| value.is_finite() && *value >= 1.0)
            .unwrap_or(1.25)
    }

    fn max_calls() -> usize {
        std::env::var("LITCHI_FILE_MAX_CALLS")
            .ok()
            .and_then(|value| value.parse::<usize>().ok())
            .filter(|value| *value > 0)
            .unwrap_or(MAX_CALLS)
    }

    fn sample_with_operations<T>(
        calls: usize,
        operations_per_call: usize,
        mut action: impl FnMut() -> T,
    ) -> Sample {
        let started = Instant::now();
        for _ in 0..calls {
            black_box(action());
        }
        Sample {
            elapsed_ns: elapsed_ns(started),
            calls,
            operation_count: calls
                .checked_mul(operations_per_call)
                .expect("operation count overflow"),
        }
    }

    fn sample<T>(calls: usize, action: impl FnMut() -> T) -> Sample {
        sample_with_operations(calls, 1, action)
    }

    fn phase_value(value: Sample) -> serde_json::Value {
        serde_json::json!({
            "samples": [value.elapsed_ns],
            "calls": value.calls,
            "total_ns": value.elapsed_ns,
            "operation_count": value.operation_count,
        })
    }

    #[derive(Clone, Copy)]
    enum Side {
        Leaf,
        Facade,
    }

    fn batch_items(family: &Family, count: usize) -> Vec<BatchItem<'static>> {
        (0..count)
            .map(|index| {
                let mut item = BatchItem::new("set");
                item.path = Some(family.paths[index % family.paths.len()].clone().into());
                item.props = Some(family.set_props.clone());
                item
            })
            .collect()
    }

    fn csv(rows: usize) -> Vec<u8> {
        let mut source = Vec::with_capacity(rows.saturating_mul(20));
        for row in 0..rows {
            source.extend_from_slice(b"42,42,42,42\n");
            if row + 1 == rows {
                break;
            }
        }
        source
    }

    /// Decode the same four-column numeric CSV used by the facade import.
    /// Leaf packages do not expose a format-neutral import operation, so the
    /// equivalent leaf workload parses the input and applies the decoded
    /// cells through its native edit API.  Parsing remains inside the timed
    /// leaf operation, just as it is inside `File::import`.
    fn imported_rows(source: &[u8]) -> usize {
        source
            .split(|byte| *byte == b'\n')
            .filter(|row| !row.is_empty())
            .map(|row| {
                let mut fields = row.split(|byte| *byte == b',');
                for _ in 0..4 {
                    let field = fields.next().expect("four CSV columns");
                    let text = std::str::from_utf8(field).expect("UTF-8 CSV field");
                    let value = text.parse::<f64>().expect("numeric CSV field");
                    black_box(value);
                }
                assert!(fields.next().is_none(), "too many CSV columns");
            })
            .count()
    }

    fn unified_set(family: &Family, source: &File) -> File {
        source
            .set(
                family.paths[family.paths.len() / 2].as_str(),
                family.set_props.clone(),
            )
            .expect("unified set")
    }

    fn unified_batch(family: &Family, source: &File, count: usize) -> File {
        source
            .batch(batch_items(family, count), BatchMode::Atomic)
            .expect("unified batch")
            .into_file()
    }

    fn unified_import(source: &File, input: &[u8]) -> File {
        source
            .import("/sheet[1]/A1", input, &litchi_ops::import::Options::csv())
            .expect("unified import")
            .into_file()
    }

    fn merge_record() -> Record {
        let mut record = Record::new();
        record.insert("name", "Acme");
        record
    }

    fn unified_merge(source: &File, record: &Record) -> File {
        source
            .merge(record, &MergeOptions::new())
            .expect("unified merge")
            .into_file()
    }

    fn unified_refresh(source: &File) -> File {
        source
            .refresh(&litchi_ops::refresh::Options::new())
            .expect("unified refresh")
            .into_file()
    }

    impl LeafPackage {
        fn native_set(&self, _name: &str, paths: &[String], count: usize) -> Self {
            let target = |index: usize| {
                if count == 1 {
                    paths.len() / 2
                } else {
                    index % paths.len()
                }
            };
            match self {
                Self::Docx(package) => {
                    let mut edit = package.edit().expect("docx edit");
                    for index in 0..count {
                        edit.replace_paragraph_text(Position::new(target(index)), "Replaced text")
                            .expect("docx set");
                    }
                    Self::Docx(edit.commit().expect("docx commit").into_package())
                },
                Self::Odt(package) => {
                    let document = package.document().expect("odt document");
                    let nodes: Vec<_> = document
                        .body()
                        .expect("odt body")
                        .nodes()
                        .filter(|node| {
                            matches!(
                                node.kind(),
                                litchi_odt::body::Kind::Paragraph | litchi_odt::body::Kind::Heading
                            )
                        })
                        .map(|node| node.id())
                        .collect();
                    let mut edit = package.edit();
                    for index in 0..count {
                        // `paths` and this filtered node vector use the same
                        // document-order index.  The path itself is rendered
                        // one-based by the facade, so keep the selected index
                        // rather than deriving an unrelated node id.
                        let node = target(index) % nodes.len();
                        edit.set_text(nodes[node], "Replaced text")
                            .expect("odt set");
                    }
                    Self::Odt(edit.commit().expect("odt commit").into_package())
                },
                Self::Doc(package) => {
                    let mut edit = package.edit().expect("doc edit");
                    for index in 0..count {
                        edit.replace_paragraph(Position::new(target(index)), "Replaced text")
                            .expect("doc set");
                    }
                    Self::Doc(edit.commit().expect("doc commit").into_package())
                },
                Self::Rtf(package) => {
                    let mut edit = package.edit();
                    for index in 0..count {
                        edit.replace_paragraph_text(Position::new(target(index)), "Replaced text")
                            .expect("rtf set");
                    }
                    Self::Rtf(edit.commit().expect("rtf commit").into_package())
                },
                Self::Xlsx(package) => {
                    use litchi_xlsx::cell::{Content, Value};
                    let mut edit = package.edit().expect("xlsx edit");
                    for index in 0..count {
                        let target = target(index);
                        let row = target / 10;
                        let column = target % 10;
                        let address = format!("{}{}", char::from(b'A' + column as u8), row + 1);
                        edit.sheet(0)
                            .expect("xlsx sheet")
                            .expect("xlsx worksheet")
                            .set(address.as_str(), Content::Value(Value::from(42)))
                            .expect("xlsx set");
                    }
                    Self::Xlsx(edit.commit().expect("xlsx commit").into_package())
                },
                Self::Xlsb(package) => {
                    use litchi_xlsb::cell_values::{Reference, Value};
                    let mut edit = package
                        .edit_cell_values(0_usize)
                        .expect("xlsb edit")
                        .resolve()
                        .expect("xlsb worksheet resolution")
                        .expect("xlsb worksheet");
                    for index in 0..count {
                        let target = target(index);
                        let row = u32::try_from(target / 10).expect("row");
                        let column = u32::try_from(target % 10).expect("column");
                        edit.set_value(
                            Reference::new(row, column).expect("xlsb reference"),
                            Value::Number(42.0),
                        )
                        .expect("xlsb set");
                    }
                    Self::Xlsb(edit.commit().expect("xlsb commit").into_package())
                },
                Self::Xls(package) => {
                    use litchi_xls::cell_values::{Reference, Selector, Value};
                    let mut edit = package.edit().expect("xls edit");
                    for index in 0..count {
                        let target = target(index);
                        let row = u32::try_from(target / 10).expect("row");
                        let column = u32::try_from(target % 10).expect("column");
                        edit.set_value(
                            Selector::Position(Position::new(0)),
                            Reference::new(row, column).expect("xls reference"),
                            Value::Number(42.0),
                        )
                        .expect("xls set");
                    }
                    Self::Xls(edit.commit().expect("xls commit").into_package())
                },
                Self::Pptx(package) => {
                    let mut edit = package.edit().expect("pptx edit");
                    for index in 0..count {
                        edit.set_shape_text(0, target(index), "Replaced text")
                            .expect("pptx set");
                    }
                    Self::Pptx(edit.commit().expect("pptx commit").into_package())
                },
                Self::Ppt(package) => {
                    let mut edit = package.edit().expect("ppt edit");
                    for index in 0..count {
                        edit.set_shape_text(
                            litchi_ppt::package::edit::Target::new(
                                Position::new(0),
                                Position::new(target(index)),
                            ),
                            "Replaced text",
                        )
                        .expect("ppt set");
                    }
                    Self::Ppt(edit.commit().expect("ppt commit").into_package())
                },
                Self::Ods(package) => {
                    use litchi_ods::worksheet::{CellChange, Selector};
                    use litchi_ods::{Cell, CellValue};
                    let mut edit = package.edit().expect("ods edit");
                    edit.worksheets(|worksheets| {
                        let mut changes = Vec::with_capacity(count);
                        for index in 0..count {
                            let target = target(index);
                            changes.push(CellChange::new(
                                target / 10,
                                target % 10,
                                Cell::new(CellValue::Number(42.0), "42"),
                            ));
                        }
                        worksheets.set_cells(Selector::Position(Position::new(0)), changes)?;
                        Ok(())
                    })
                    .expect("ods set");
                    Self::Ods(edit.commit().expect("ods commit").into_package())
                },
                Self::Odp(package) => {
                    use litchi_odp::scene::{Change, ObjectRef, StoryRef};
                    let mut edit = package.edit().expect("odp edit");
                    let changes: Vec<_> = (0..count)
                        .map(|index| Change::SetParagraphText {
                            story: StoryRef::object(ObjectRef::new(target(index), 0)),
                            paragraph: 0,
                            text: "Replaced text".to_owned(),
                        })
                        .collect();
                    edit.splice_scene(&changes).expect("odp set");
                    Self::Odp(edit.commit().expect("odp commit").into_package())
                },
            }
        }

        fn native_batch(&self, name: &str, paths: &[String], count: usize) -> Self {
            match self {
                Self::Odp(package) => {
                    use litchi_odp::scene::{Change, ObjectRef, StoryRef};
                    let mut edit = package.edit().expect("odp batch edit");
                    let changes: Vec<_> = (0..count)
                        .map(|index| Change::SetParagraphText {
                            story: StoryRef::object(ObjectRef::new(index % paths.len(), 0)),
                            paragraph: 0,
                            text: "Replaced text".to_owned(),
                        })
                        .collect();
                    edit.splice_scene(&changes).expect("odp batch");
                    Self::Odp(edit.commit().expect("odp batch commit").into_package())
                },
                _ => self.native_set(name, paths, count),
            }
        }

        fn native_import(&self, input: &[u8]) -> Self {
            let rows = imported_rows(input);
            let cells = rows.checked_mul(4).expect("CSV cell count");
            match self {
                Self::Xlsx(package) => {
                    use litchi_xlsx::cell::{Content, Value};
                    let mut edit = package.edit().expect("xlsx import edit");
                    {
                        let mut sheet = edit.sheet(0).expect("xlsx sheet").expect("xlsx worksheet");
                        for index in 0..cells {
                            let row = index / 4;
                            let column = index % 4;
                            let address = format!("{}{}", char::from(b'A' + column as u8), row + 1);
                            sheet
                                .set(address.as_str(), Content::Value(Value::from(42)))
                                .expect("xlsx import cell");
                        }
                    }
                    Self::Xlsx(edit.commit().expect("xlsx import commit").into_package())
                },
                Self::Xlsb(package) => {
                    use litchi_xlsb::cell_values::{Reference, Value};
                    let mut edit = package
                        .edit_cell_values(0_usize)
                        .expect("xlsb import edit")
                        .resolve()
                        .expect("xlsb worksheet resolution")
                        .expect("xlsb worksheet");
                    for index in 0..cells {
                        edit.set_value(
                            Reference::new(
                                u32::try_from(index / 4).expect("xlsb import row"),
                                u32::try_from(index % 4).expect("xlsb import column"),
                            )
                            .expect("xlsb import reference"),
                            Value::Number(42.0),
                        )
                        .expect("xlsb import cell");
                    }
                    Self::Xlsb(edit.commit().expect("xlsb import commit").into_package())
                },
                Self::Xls(package) => {
                    use litchi_xls::cell_values::{Reference, Selector, Value};
                    let mut edit = package.edit().expect("xls import edit");
                    for index in 0..cells {
                        edit.set_value(
                            Selector::Position(Position::new(0)),
                            Reference::new(
                                u32::try_from(index / 4).expect("xls import row"),
                                u32::try_from(index % 4).expect("xls import column"),
                            )
                            .expect("xls import reference"),
                            Value::Number(42.0),
                        )
                        .expect("xls import cell");
                    }
                    Self::Xls(edit.commit().expect("xls import commit").into_package())
                },
                Self::Ods(package) => {
                    use litchi_ods::worksheet::{CellChange, Selector};
                    use litchi_ods::{Cell, CellValue};
                    let mut edit = package.edit().expect("ods import edit");
                    edit.worksheets(|worksheets| {
                        let changes = (0..cells)
                            .map(|index| {
                                CellChange::new(
                                    index / 4,
                                    index % 4,
                                    Cell::new(CellValue::Number(42.0), "42"),
                                )
                            })
                            .collect();
                        worksheets.set_cells(Selector::Position(Position::new(0)), changes)?;
                        Ok(())
                    })
                    .expect("ods import cells");
                    Self::Ods(edit.commit().expect("ods import commit").into_package())
                },
                _ => panic!("import route is only supported for workbook families"),
            }
        }

        fn native_merge(&self) -> Self {
            match self {
                Self::Docx(package) => {
                    let mut edit = package.edit().expect("docx edit");
                    edit.replace_paragraph_text(Position::new(0), "Dear Acme,")
                        .expect("docx merge");
                    Self::Docx(edit.commit().expect("docx commit").into_package())
                },
                Self::Odt(package) => {
                    let document = package.document().expect("odt document");
                    let node = document
                        .body()
                        .expect("odt body")
                        .nodes()
                        .find(|node| {
                            matches!(
                                node.kind(),
                                litchi_odt::body::Kind::Paragraph | litchi_odt::body::Kind::Heading
                            )
                        })
                        .expect("odt node");
                    let mut edit = package.edit();
                    edit.set_text(node.id(), "Dear Acme,").expect("odt merge");
                    Self::Odt(edit.commit().expect("odt commit").into_package())
                },
                Self::Doc(package) => {
                    let mut edit = package.edit().expect("doc edit");
                    edit.replace_paragraph(Position::new(0), "Dear Acme,")
                        .expect("doc merge");
                    Self::Doc(edit.commit().expect("doc commit").into_package())
                },
                Self::Rtf(package) => {
                    let mut edit = package.edit();
                    edit.replace_paragraph_text(Position::new(0), "Dear Acme,")
                        .expect("rtf merge");
                    Self::Rtf(edit.commit().expect("rtf commit").into_package())
                },
                Self::Xlsx(package) => {
                    use litchi_xlsx::cell::{Content, Value};
                    let mut edit = package.edit().expect("xlsx edit");
                    edit.sheet(0)
                        .expect("xlsx sheet")
                        .expect("worksheet")
                        .set("L1", Content::Value(Value::text("Invoice for Acme")))
                        .expect("xlsx merge");
                    Self::Xlsx(edit.commit().expect("xlsx commit").into_package())
                },
                Self::Xlsb(package) => {
                    use litchi_xlsb::cell_values::{Reference, Value};
                    let mut edit = package
                        .edit_cell_values(0_usize)
                        .expect("xlsb edit")
                        .resolve()
                        .expect("xlsb worksheet resolution")
                        .expect("worksheet");
                    edit.set_value(
                        Reference::new(0, 11).expect("xlsb ref"),
                        Value::InlineString("Invoice for Acme".to_owned()),
                    )
                    .expect("xlsb merge");
                    Self::Xlsb(edit.commit().expect("xlsb commit").into_package())
                },
                Self::Xls(package) => {
                    use litchi_xls::cell_values::{Reference, Selector, Value};
                    let mut edit = package.edit().expect("xls edit");
                    edit.set_value(
                        Selector::Position(Position::new(0)),
                        Reference::new(0, 11).expect("xls ref"),
                        Value::Text("Invoice for Acme".to_owned()),
                    )
                    .expect("xls merge");
                    Self::Xls(edit.commit().expect("xls commit").into_package())
                },
                Self::Pptx(package) => {
                    let mut edit = package.edit().expect("pptx edit");
                    edit.set_shape_text(0, 0_usize, "Dear Acme")
                        .expect("pptx merge");
                    Self::Pptx(edit.commit().expect("pptx commit").into_package())
                },
                Self::Ppt(package) => {
                    let mut edit = package.edit().expect("ppt edit");
                    edit.set_shape_text(
                        litchi_ppt::package::edit::Target::new(Position::new(0), Position::new(0)),
                        "Dear Acme",
                    )
                    .expect("ppt merge");
                    Self::Ppt(edit.commit().expect("ppt commit").into_package())
                },
                Self::Ods(package) => {
                    use litchi_ods::worksheet::{CellChange, Selector};
                    use litchi_ods::{Cell, CellValue};
                    let mut edit = package.edit().expect("ods edit");
                    edit.worksheets(|worksheets| {
                        worksheets.set_cells(
                            Selector::Position(Position::new(0)),
                            vec![CellChange::new(
                                0,
                                11,
                                Cell::new(
                                    CellValue::Text("Invoice for Acme".to_owned()),
                                    "Invoice for Acme",
                                ),
                            )],
                        )?;
                        Ok(())
                    })
                    .expect("ods merge");
                    Self::Ods(edit.commit().expect("ods commit").into_package())
                },
                Self::Odp(package) => {
                    use litchi_odp::scene::{Change, ObjectRef, StoryRef};
                    let mut edit = package.edit().expect("odp edit");
                    edit.splice_scene(&[Change::SetParagraphText {
                        story: StoryRef::object(ObjectRef::new(0, 0)),
                        paragraph: 0,
                        text: "Dear Acme".to_owned(),
                    }])
                    .expect("odp merge");
                    Self::Odp(edit.commit().expect("odp commit").into_package())
                },
            }
        }

        fn native_refresh(&self, name: &str) -> Self {
            match self {
                Self::Docx(package) if name == "docx" => {
                    let mut edit = package.edit().expect("docx edit");
                    let fields = edit.fields().expect("docx fields");
                    let (_, characters) = fields.paragraphs().iter().fold(
                        (0usize, 0usize),
                        |(words, characters), paragraph| {
                            (
                                words + paragraph.text().split_whitespace().count(),
                                characters
                                    + paragraph
                                        .text()
                                        .chars()
                                        .filter(|character| !character.is_whitespace())
                                        .count(),
                            )
                        },
                    );
                    edit.refresh_fields(
                        &fields,
                        &[litchi_docx::document::fields::Update::Text {
                            field: 0,
                            text: characters.to_string(),
                        }],
                    )
                    .expect("docx refresh");
                    Self::Docx(edit.commit().expect("docx refresh commit").into_package())
                },
                _ => panic!("refresh route is only supported for DOCX"),
            }
        }

        fn read_marker(&self, _name: &str) -> usize {
            match self {
                Self::Docx(package) => package
                    .document()
                    .expect("docx document")
                    .body_paragraph(Position::new(0))
                    .expect("docx paragraph")
                    .map_or(0, |paragraph| {
                        paragraph.text().expect("docx paragraph text").len()
                    }),
                Self::Odt(package) => package
                    .document()
                    .expect("odt document")
                    .body()
                    .expect("odt body")
                    .nodes()
                    .find(|node| {
                        matches!(
                            node.kind(),
                            litchi_odt::body::Kind::Paragraph | litchi_odt::body::Kind::Heading
                        )
                    })
                    .map_or(0, |node| node.text().expect("odt text").len()),
                Self::Doc(package) => package
                    .document()
                    .body_paragraph_text(Position::new(0))
                    .expect("doc paragraph")
                    .map_or(0, str::len),
                Self::Rtf(package) => package
                    .document()
                    .body()
                    .paragraphs()
                    .next()
                    .map_or(0, |paragraph| paragraph.len()),
                Self::Xlsx(package) => package
                    .workbook()
                    .sheet(0)
                    .expect("xlsx sheet")
                    .expect("xlsx worksheet")
                    .cell("L1")
                    .expect("xlsx cell")
                    .stored()
                    .map_or(0, |cell| match cell {
                        litchi_xlsx::cell::Cell::Value(litchi_xlsx::cell::Value::Text(text)) => {
                            text.as_str().len()
                        },
                        litchi_xlsx::cell::Cell::Value(litchi_xlsx::cell::Value::Date(date)) => {
                            date.as_str().len()
                        },
                        _ => 0,
                    }),
                Self::Xlsb(package) => package
                    .workbook()
                    .expect("xlsb workbook")
                    .cell_values(0)
                    .expect("xlsb cell values")
                    .cell(
                        litchi_xlsb::cell_values::Reference::new(0, 11)
                            .expect("xlsb marker reference"),
                    )
                    .expect("xlsb marker cell")
                    .map_or(0, |cell| match cell.value() {
                        litchi_xlsb::cell_values::Value::InlineString(text)
                        | litchi_xlsb::cell_values::Value::FormulaStringCache(text) => text.len(),
                        litchi_xlsb::cell_values::Value::RichString(text) => text.text.len(),
                        _ => 0,
                    }),
                Self::Xls(package) => package
                    .cell_values()
                    .expect("xls cell values")
                    .worksheet(litchi_xls::cell_values::Selector::Position(Position::new(
                        0,
                    )))
                    .expect("xls worksheet")
                    .expect("xls worksheet present")
                    .cell(
                        litchi_xls::cell_values::Reference::new(0, 11)
                            .expect("xls marker reference"),
                    )
                    .expect("xls marker cell")
                    .map_or(0, |cell| match cell.value() {
                        litchi_xls::cell_values::Value::Text(text) => text.len(),
                        _ => 0,
                    }),
                Self::Pptx(package) => package
                    .presentation()
                    .expect("pptx presentation")
                    .slide(0)
                    .expect("pptx slide")
                    .expect("pptx slide present")
                    .shapes()
                    .expect("pptx shapes")
                    .iter()
                    .next()
                    .and_then(|shape| shape.text())
                    .map_or(0, str::len),
                Self::Ppt(package) => package
                    .presentation()
                    .expect("ppt presentation")
                    .slide_at(0)
                    .expect("ppt slide")
                    .expect("ppt slide present")
                    .shapes()
                    .expect("ppt shapes")
                    .first()
                    .map_or(0, |shape| shape.text().expect("ppt shape text").len()),
                Self::Ods(package) => package
                    .workbook()
                    .expect("ods workbook")
                    .cell("Sheet1", 0, 11)
                    .and_then(|view| match view {
                        litchi_ods::worksheet::CellView::Stored(cell) => Some(cell),
                        litchi_ods::worksheet::CellView::Missing => None,
                    })
                    .map_or(0, |cell| cell.text.len()),
                Self::Odp(package) => package
                    .presentation()
                    .expect("odp presentation")
                    .slide(0)
                    .expect("odp slide")
                    .expect("odp slide present")
                    // The fixture intentionally puts exactly one text shape
                    // on each slide. The model exposes that shape's text as
                    // the slide body, so `all_text` is the native equivalent
                    // of `/slide[1]/shape[1]` here.
                    .all_text()
                    .len(),
            }
        }
    }

    fn source_leaf(family: &Family) -> LeafPackage {
        LeafPackage::open(family.name, &family.bytes)
    }

    fn source_facade(family: &Family) -> File {
        open_bytes(family.bytes.clone()).unwrap_or_else(|error| panic!("facade open: {error}"))
    }

    fn leaf_commit(
        family: &Family,
        source: &LeafPackage,
        route: Route,
        import_input: &[u8],
    ) -> LeafPackage {
        match route {
            Route::SetCommit | Route::SetPublication => {
                source.native_set(family.name, &family.paths, 1)
            },
            Route::BatchCommit | Route::BatchPublication => {
                source.native_batch(family.name, &family.paths, family.batch)
            },
            Route::ImportCommit | Route::ImportPublication => source.native_import(import_input),
            Route::MergeCommit | Route::MergePublication => source.native_merge(),
            Route::RefreshCommit | Route::RefreshPublication => source.native_refresh(family.name),
            _ => LeafPackage::open(family.name, &family.bytes),
        }
    }

    fn facade_commit(
        family: &Family,
        source: &File,
        route: Route,
        import_input: &[u8],
        record: &Record,
    ) -> File {
        match route {
            Route::SetCommit | Route::SetPublication => unified_set(family, source),
            Route::BatchCommit | Route::BatchPublication => {
                unified_batch(family, source, family.batch)
            },
            Route::ImportCommit | Route::ImportPublication => unified_import(source, import_input),
            Route::MergeCommit | Route::MergePublication => unified_merge(source, record),
            Route::RefreshCommit | Route::RefreshPublication => unified_refresh(source),
            _ => source.clone(),
        }
    }

    fn phase_once(
        route: Route,
        family: &Family,
        leaf_source: &LeafPackage,
        facade_source: &File,
        side: Side,
        calls: usize,
        import_input: &[u8],
        record: &Record,
    ) -> Sample {
        match (side, route) {
            (Side::Leaf, Route::OpenGet) => sample(calls, || {
                let package = LeafPackage::open(family.name, &family.bytes);
                package.read_marker(family.name)
            }),
            (Side::Facade, Route::OpenGet) => sample(calls, || {
                let file = source_facade(family);
                file.get(family.marker)
                    .expect("facade get")
                    .view()
                    .text()
                    .expect("facade text")
                    .map_or(0, |text| text.len())
            }),
            (Side::Leaf, Route::Save) => sample(calls, || leaf_source.publish().len()),
            (Side::Facade, Route::Save) => sample(calls, || {
                facade_source.to_bytes().expect("facade save").len()
            }),
            (Side::Leaf, Route::RepeatedSave) => sample_with_operations(calls, 2, || {
                let first = leaf_source.publish();
                black_box(first.as_slice());
                leaf_source.publish().len()
            }),
            (Side::Facade, Route::RepeatedSave) => sample_with_operations(calls, 2, || {
                let first = facade_source.to_bytes().expect("facade first save");
                black_box(first.as_slice());
                facade_source
                    .to_bytes()
                    .expect("facade repeated save")
                    .len()
            }),
            (Side::Leaf, route) if route.is_commit() => sample(calls, || {
                leaf_commit(family, leaf_source, route, import_input)
            }),
            (Side::Facade, route) if route.is_commit() => sample(calls, || {
                facade_commit(family, facade_source, route, import_input, record)
            }),
            (Side::Leaf, route) if route.is_publication() => {
                let prepared = leaf_commit(family, leaf_source, route, import_input);
                sample(calls, || prepared.publish().len())
            },
            (Side::Facade, route) if route.is_publication() => {
                let prepared = facade_commit(family, facade_source, route, import_input, record);
                sample(calls, || {
                    prepared.to_bytes().expect("facade publication").len()
                })
            },
            _ => unreachable!("route phase mismatch"),
        }
    }

    fn semantic_value(bytes: &[u8], marker: &str) -> String {
        open_bytes(bytes.to_vec())
            .unwrap_or_else(|error| panic!("semantic reopen: {error}"))
            .get(marker)
            .unwrap_or_else(|error| panic!("semantic marker {marker}: {error}"))
            .view()
            .text()
            .unwrap_or_else(|error| panic!("semantic text {marker}: {error}"))
            .map_or_else(String::new, |value| value.into_owned())
    }

    fn facade_marker_len(file: &File, marker: &str) -> usize {
        file.get(marker)
            .unwrap_or_else(|error| panic!("facade marker {marker}: {error}"))
            .view()
            .text()
            .unwrap_or_else(|error| panic!("facade marker text {marker}: {error}"))
            .map_or(0, |value| value.len())
    }

    fn semantic_marker(family: &Family, route: Route) -> &str {
        match route {
            Route::SetCommit | Route::SetPublication => {
                family.paths[family.paths.len() / 2].as_str()
            },
            Route::BatchCommit | Route::BatchPublication => family.paths[0].as_str(),
            Route::ImportCommit | Route::ImportPublication => "/sheet[1]/A1",
            // `docx_refresh` puts the NUMCHARS field after the 1,000 normal
            // paragraphs so the regular set/batch paths never target a field.
            Route::RefreshCommit | Route::RefreshPublication => "/body/p[1001]",
            _ => family.marker,
        }
    }

    fn expected_semantic(family: &Family, route: Route) -> Option<String> {
        let name = family.name;
        match route {
            Route::SetCommit
            | Route::SetPublication
            | Route::BatchCommit
            | Route::BatchPublication => Some(if matches!(name, "xlsx" | "xlsb" | "xls" | "ods") {
                "42".to_owned()
            } else {
                "Replaced text".to_owned()
            }),
            Route::ImportCommit | Route::ImportPublication => Some("42".to_owned()),
            Route::MergeCommit | Route::MergePublication => {
                Some(if matches!(name, "xlsx" | "xlsb" | "xls" | "ods") {
                    "Invoice for Acme".to_owned()
                } else if matches!(name, "pptx" | "ppt" | "odp") {
                    "Dear Acme".to_owned()
                } else {
                    "Dear Acme,".to_owned()
                })
            },
            Route::RefreshCommit | Route::RefreshPublication => family.refresh_value.clone(),
            _ => None,
        }
    }

    fn output_record(bytes: &[u8], source: &[u8], changed: bool) -> serde_json::Value {
        if bytes.is_empty() {
            panic!("publication returned an empty output");
        }
        if changed && bytes == source {
            panic!("changed route reproduced source bytes");
        }
        serde_json::json!({
            "bytes": bytes.len(),
            "sha256": BlobId::of(bytes).as_hex(),
            "source_sha256": BlobId::of(source).as_hex(),
            "changed_from_source": bytes != source,
        })
    }

    fn validation(
        family: &Family,
        route: Route,
        source_leaf: &LeafPackage,
        source_facade: &File,
        source_leaf_before: &[u8],
        source_facade_before: &[u8],
        leaf_after: &LeafPackage,
        facade_after: &File,
    ) -> serde_json::Value {
        let source_bytes = family.bytes.as_slice();
        // Publish the untouched owners after the edit candidate was built and
        // compare them with bytes captured before staging.  This proves that
        // staging a candidate did not mutate either source owner, including on
        // changed routes where a semantic-marker-only check is vacuous.
        let source_leaf_bytes = source_leaf.publish();
        let source_facade_bytes = source_facade.to_bytes().expect("facade source bytes");
        let source_leaf_exact = source_leaf_bytes == source_leaf_before;
        let source_facade_exact = source_facade_bytes == source_facade_before;
        if !source_leaf_exact {
            panic!("leaf source owner changed while staging the candidate");
        }
        if !source_facade_exact {
            panic!("facade source owner changed while staging the candidate");
        }
        // Keep the untimed open/get comparison on the native read path. The
        // semantic checks below reopen both byte streams through File and
        // therefore cannot prove that LeafPackage::read_marker itself read
        // the same logical target as the facade projection.
        let leaf_read_marker_len = source_leaf.read_marker(family.name);
        let facade_read_marker_len = facade_marker_len(source_facade, family.marker);
        let read_marker_equal = leaf_read_marker_len == facade_read_marker_len;
        if route == Route::OpenGet && !read_marker_equal {
            panic!(
                "native/facade open_get marker lengths differ: {} != {}",
                leaf_read_marker_len, facade_read_marker_len
            );
        }
        let leaf_bytes = if route.is_commit() || route.is_publication() {
            leaf_after.publish()
        } else {
            source_leaf_bytes.clone()
        };
        let facade_bytes = if route.is_commit() || route.is_publication() {
            facade_after.to_bytes().expect("facade validation bytes")
        } else {
            source_facade_bytes.clone()
        };
        if leaf_bytes.is_empty() || facade_bytes.is_empty() {
            panic!("a route produced an empty output");
        }
        let marker = semantic_marker(family, route);
        let source_semantic = semantic_value(source_bytes, marker);
        let source_leaf_semantic = semantic_value(&source_leaf_bytes, marker);
        let source_facade_semantic = semantic_value(&source_facade_bytes, marker);
        let leaf_semantic = semantic_value(&leaf_bytes, marker);
        let facade_semantic = semantic_value(&facade_bytes, marker);
        let changed_expected = route.changed();
        let leaf_changed = leaf_bytes != source_bytes;
        let facade_changed = facade_bytes != source_bytes;
        let expected = expected_semantic(family, route);
        let semantic_state_valid = expected
            .is_none_or(|expected| leaf_semantic == expected && facade_semantic == expected);
        let source_state_preserved =
            source_leaf_semantic == source_semantic && source_facade_semantic == source_semantic;
        let semantic_equal = leaf_semantic == facade_semantic
            && (changed_expected || source_state_preserved)
            && semantic_state_valid;
        let leaf_repeat = if route.is_commit() || route.is_publication() {
            leaf_after.publish()
        } else {
            source_leaf_bytes.clone()
        };
        let facade_repeat = if route.is_commit() || route.is_publication() {
            facade_after.to_bytes().expect("facade repeat bytes")
        } else {
            source_facade_bytes.clone()
        };
        if leaf_repeat != leaf_bytes || facade_repeat != facade_bytes {
            panic!("repeated publication changed output bytes");
        }
        let exact_noop = !route.exact_noop()
            || (leaf_bytes == leaf_repeat
                && facade_bytes == facade_repeat
                && leaf_bytes == source_leaf_bytes
                && facade_bytes == source_facade_bytes);
        if !exact_noop {
            panic!("exact no-op route changed its owner output");
        }
        serde_json::json!({
            "leaf_changed": leaf_changed,
            "facade_changed": facade_changed,
            "leaf_output_changed": leaf_bytes != source_bytes,
            "facade_output_changed": facade_bytes != source_bytes,
            "semantic_state_valid": semantic_state_valid,
            "source_state_preserved": source_state_preserved,
            "source_state_exact": source_leaf_exact && source_facade_exact,
            "leaf_read_marker_len": leaf_read_marker_len,
            "facade_read_marker_len": facade_read_marker_len,
            "read_marker_equal": read_marker_equal,
            "source_leaf_before_sha256": BlobId::of(source_leaf_before).as_hex(),
            "source_leaf_after_sha256": BlobId::of(&source_leaf_bytes).as_hex(),
            "source_facade_before_sha256": BlobId::of(source_facade_before).as_hex(),
            "source_facade_after_sha256": BlobId::of(&source_facade_bytes).as_hex(),
            "semantic_equal": semantic_equal,
            "exact_noop": exact_noop,
            "leaf_repeated_equal": leaf_repeat == leaf_bytes,
            "facade_repeated_equal": facade_repeat == facade_bytes,
            "source_leaf_output": output_record(&source_leaf_bytes, source_bytes, false),
            "source_facade_output": output_record(&source_facade_bytes, source_bytes, false),
            "leaf_output": output_record(&leaf_bytes, source_bytes, changed_expected),
            "facade_output": output_record(&facade_bytes, source_bytes, changed_expected),
            "leaf_repeat": output_record(&leaf_repeat, source_bytes, changed_expected),
            "facade_repeat": output_record(&facade_repeat, source_bytes, changed_expected),
            "leaf_semantic_len": leaf_semantic.len(),
            "facade_semantic_len": facade_semantic.len(),
        })
    }

    fn measure(family: &Family, route: Route, blocks: usize, calls: usize) -> serde_json::Value {
        let leaf_source = source_leaf(family);
        let facade_source = source_facade(family);
        let source_leaf_before = leaf_source.publish();
        let source_facade_before = facade_source.to_bytes().expect("facade source bytes");
        let import_input = csv(family.import);
        let record = merge_record();
        let leaf_after = leaf_commit(family, &leaf_source, route, &import_input);
        let facade_after = facade_commit(family, &facade_source, route, &import_input, &record);
        let validation = validation(
            family,
            route,
            &leaf_source,
            &facade_source,
            &source_leaf_before,
            &source_facade_before,
            &leaf_after,
            &facade_after,
        );
        let mut block_values = Vec::with_capacity(blocks);
        for _ in 0..blocks {
            let leaf_first = phase_once(
                route,
                family,
                &leaf_source,
                &facade_source,
                Side::Leaf,
                calls,
                &import_input,
                &record,
            );
            let facade_first = phase_once(
                route,
                family,
                &leaf_source,
                &facade_source,
                Side::Facade,
                calls,
                &import_input,
                &record,
            );
            let facade_second = phase_once(
                route,
                family,
                &leaf_source,
                &facade_source,
                Side::Facade,
                calls,
                &import_input,
                &record,
            );
            let leaf_second = phase_once(
                route,
                family,
                &leaf_source,
                &facade_source,
                Side::Leaf,
                calls,
                &import_input,
                &record,
            );
            block_values.push(serde_json::json!({
                "order": ["leaf", "facade", "facade", "leaf"],
                "leaf": [phase_value(leaf_first), phase_value(leaf_second)],
                "facade": [phase_value(facade_first), phase_value(facade_second)],
            }));
        }
        serde_json::json!({
            "schema": "litchi_file_qualification.v2",
            "route": format!("{}/{}", family.name, route.suffix()),
            "calls": calls,
            "blocks": block_values,
            "validation": validation,
        })
    }

    fn validate_route(family: &Family, route: Route) -> serde_json::Value {
        let leaf_source = source_leaf(family);
        let facade_source = source_facade(family);
        let source_leaf_before = leaf_source.publish();
        let source_facade_before = facade_source.to_bytes().expect("facade source bytes");
        let import_input = csv(family.import);
        let record = merge_record();
        let leaf_after = leaf_commit(family, &leaf_source, route, &import_input);
        let facade_after = facade_commit(family, &facade_source, route, &import_input, &record);
        serde_json::json!({
            "schema": "litchi_file_validation.v1",
            "route": format!("{}/{}", family.name, route.suffix()),
            "validation": validation(
                family,
                route,
                &leaf_source,
                &facade_source,
                &source_leaf_before,
                &source_facade_before,
                &leaf_after,
                &facade_after,
            ),
        })
    }

    fn calibration(family: &Family, route: Route) -> serde_json::Value {
        let leaf_source = source_leaf(family);
        let facade_source = source_facade(family);
        let import_input = csv(family.import);
        let record = merge_record();
        let mut calls = std::env::var("LITCHI_FILE_CALLS")
            .ok()
            .and_then(|value| value.parse::<usize>().ok())
            .filter(|value| *value > 0)
            .unwrap_or(1);
        let ceiling = max_calls();
        loop {
            let leaf = phase_once(
                route,
                family,
                &leaf_source,
                &facade_source,
                Side::Leaf,
                calls,
                &import_input,
                &record,
            );
            let facade = phase_once(
                route,
                family,
                &leaf_source,
                &facade_source,
                Side::Facade,
                calls,
                &import_input,
                &record,
            );
            let target = (target_ns() as f64 * headroom()).ceil() as u128;
            if leaf.elapsed_ns >= target && facade.elapsed_ns >= target {
                return serde_json::json!({
                    "route": format!("{}/{}", family.name, route.suffix()),
                    "common_calls": calls,
                    "minimum_elapsed_ns": leaf.elapsed_ns.min(facade.elapsed_ns),
                    "leaf_elapsed_ns": leaf.elapsed_ns,
                    "facade_elapsed_ns": facade.elapsed_ns,
                    "target_ns": target,
                });
            }
            if calls >= ceiling {
                panic!("calibration cannot meet target for {}", route.suffix());
            }
            calls = calls.saturating_mul(2).min(ceiling);
        }
    }

    fn args() -> (String, String, usize) {
        let mut mode = String::from("measure");
        let mut route = String::from("*");
        let mut blocks = 5;
        let mut arguments = std::env::args().skip(1);
        while let Some(argument) = arguments.next() {
            match argument.as_str() {
                "--mode" => mode = arguments.next().expect("missing --mode value"),
                "--route" => route = arguments.next().expect("missing --route value"),
                "--blocks" => {
                    blocks = arguments
                        .next()
                        .expect("invalid --blocks")
                        .parse()
                        .expect("invalid blocks")
                },
                _ => panic!("unknown argument {argument}"),
            }
        }
        (mode, route, blocks)
    }

    pub(super) fn run() {
        let (mode, wanted, blocks) = args();
        match mode.as_str() {
            "list" => {
                serde_json::to_writer(
                    std::io::stdout().lock(),
                    &serde_json::json!({
                        "schema": "litchi_file_qualification.v2",
                        "routes": all_route_names(),
                    }),
                )
                .expect("write route list");
                println!();
            },
            "validate" | "calibrate" | "measure" => {
                let (name, suffix) = wanted.split_once('/').expect("route must be family/suffix");
                let route = route_for(name, suffix).expect("unknown route");
                let family = family(name);
                if mode == "validate" {
                    let value = validate_route(&family, route);
                    serde_json::to_writer(std::io::stdout().lock(), &value)
                        .expect("write validation");
                    println!();
                } else if mode == "calibrate" {
                    let value = calibration(&family, route);
                    serde_json::to_writer(std::io::stdout().lock(), &value)
                        .expect("write calibration");
                    println!();
                } else {
                    let calls = std::env::var("LITCHI_FILE_CALLS")
                        .ok()
                        .and_then(|value| value.parse::<usize>().ok())
                        .filter(|value| *value > 0)
                        .expect("measure requires LITCHI_FILE_CALLS");
                    let value = measure(&family, route, blocks, calls);
                    serde_json::to_writer(std::io::stdout().lock(), &value).expect("write measure");
                    println!();
                }
            },
            _ => panic!("unknown mode {mode}"),
        }
    }
}

fn main() {
    #[cfg(all(
        feature = "docx",
        feature = "xlsx",
        feature = "xlsb",
        feature = "pptx",
        feature = "doc",
        feature = "xls",
        feature = "ppt",
        feature = "rtf",
        feature = "odt",
        feature = "ods",
        feature = "odp"
    ))]
    probe::run();
}
