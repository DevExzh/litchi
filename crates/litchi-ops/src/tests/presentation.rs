//! The facade's `presentation` in-source tests (`prs.rs`), ported onto
//! `File` and `PresentationView`.
//!
//! Each test keeps its facade semantics: the route a fixture takes, the
//! owner it ends in, and the result or typed refusal it reports. Facade
//! errors that were string-flattened are asserted through the typed code and
//! leaf instead.

use super::test_data;
use crate::{Code, File};

fn open(relative: &str) -> File {
    crate::open(test_data(relative)).unwrap_or_else(|error| panic!("open {relative}: {error}"))
}

// ---- basic reads (`prs.rs` `tests`) ---------------------------------------

#[cfg(feature = "pptx")]
mod pptx_basics {
    use super::*;

    #[test]
    fn test_presentation_open_and_bytes_pptx() {
        assert!(matches!(open("ooxml/pptx/sample.pptx"), File::Pptx(_)));
        let bytes = std::fs::read(test_data("ooxml/pptx/sample.pptx")).unwrap();
        assert!(matches!(crate::open_bytes(bytes).unwrap(), File::Pptx(_)));
    }

    #[test]
    fn test_presentation_text_slides_and_metadata_pptx() {
        let file = open("ooxml/pptx/sample.pptx");
        file.text().unwrap();
        let view = file.presentation().unwrap();
        assert!(view.slide_count().unwrap() > 0);
        let slides = view.slides().unwrap();
        assert!(!slides.is_empty());
        for slide in slides {
            slide.text().unwrap();
        }
        let _ = file.metadata().unwrap().map(|metadata| metadata.title);
    }

    /// Slide sizes are format-specific: they are read on the PPTX handle.
    #[test]
    fn test_presentation_slide_dimensions_pptx() {
        let File::Pptx(pptx) = open("ooxml/pptx/sample.pptx") else {
            panic!("expected PPTX");
        };
        let (width, height) = pptx.slide_size().unwrap();
        assert!(width > 0 && height > 0);
    }

    #[test]
    fn test_presentation_charts_and_connectors_pptx() {
        for relative in [
            "ooxml/pptx/line-chart.pptx",
            "ooxml/pptx/pie-chart.pptx",
            "ooxml/pptx/radar-chart.pptx",
            "ooxml/pptx/scatter-chart.pptx",
        ] {
            assert!(
                open(relative)
                    .presentation()
                    .unwrap()
                    .slide_count()
                    .unwrap()
                    > 0,
                "{relative}"
            );
        }
        for relative in [
            "ooxml/pptx/connectorConnection.pptx",
            "ooxml/pptx/curvedConnectors.pptx",
            "ooxml/pptx/elbowConnectors.pptx",
        ] {
            open(relative);
        }
    }

    #[test]
    fn test_presentation_extract_text_for_markdown_pptx() {
        let file = open("ooxml/pptx/sample.pptx");
        let slides = file
            .presentation()
            .unwrap()
            .extract_text_for_markdown()
            .unwrap();
        assert!(!slides.is_empty());
        assert_eq!(slides[0].0, 1);
    }

    #[test]
    fn test_presentation_open_nonexistent_file() {
        let error = crate::open(test_data("nonexistent_file.pptx")).unwrap_err();
        assert_eq!(error.code(), Code::FileNotFound);
    }

    #[test]
    fn test_presentation_from_bytes_invalid_data() {
        let error =
            crate::open_bytes(b"This is not a valid presentation file".to_vec()).unwrap_err();
        assert_eq!(error.code(), Code::UnsupportedFormat);
        assert_eq!(error.detected(), None);
    }
}

#[cfg(feature = "ppt")]
mod ppt_basics {
    use super::*;

    #[test]
    fn test_presentation_open_and_bytes_ppt() {
        assert!(matches!(open("ole/ppt/SampleShow.ppt"), File::Ppt(_)));
        let bytes = std::fs::read(test_data("ole/ppt/SampleShow.ppt")).unwrap();
        assert!(matches!(crate::open_bytes(bytes).unwrap(), File::Ppt(_)));
    }

    #[test]
    fn test_presentation_text_slides_and_metadata_ppt() {
        let file = open("ole/ppt/SampleShow.ppt");
        file.text().unwrap();
        let view = file.presentation().unwrap();
        assert!(view.slide_count().unwrap() > 0);
        let slides = view.slides().unwrap();
        assert!(!slides.is_empty());
        for slide in &slides {
            slide.text().unwrap();
            // `shape_count` is reported for legacy PowerPoint only.
            assert!(slide.shape_count().is_some());
            assert!(slide.number().is_some());
        }
        let _ = file.metadata().unwrap().map(|metadata| metadata.title);
    }

    #[test]
    fn test_presentation_open_nonexistent_ppt() {
        let error = crate::open(test_data("nonexistent_file.ppt")).unwrap_err();
        assert_eq!(error.code(), Code::FileNotFound);
    }

    #[test]
    fn test_presentation_extract_text_for_markdown_ppt() {
        let file = open("ole/ppt/SampleShow.ppt");
        let slides = file
            .presentation()
            .unwrap()
            .extract_text_for_markdown()
            .unwrap();
        assert!(!slides.is_empty());
    }

    #[test]
    fn slide_selection_matches_the_slide_list() {
        let file = open("ole/ppt/SampleShow.ppt");
        let view = file.presentation().unwrap();
        let slides = view.slides().unwrap();
        for (position, slide) in slides.iter().enumerate() {
            let selected = view.slide(position).unwrap().expect("slide exists");
            assert_eq!(selected.text().unwrap(), slide.text().unwrap());
        }
        assert!(view.slide(slides.len()).unwrap().is_none());
    }
}

// ---- native PPT path routing (`native_ppt_path_tests`) -------------------

#[cfg(all(feature = "ppt", any(unix, windows)))]
mod native_ppt_path {
    use super::*;

    #[test]
    fn native_path_matches_byte_facade_for_core_queries_and_metadata() {
        let path = test_data("ole/ppt/SampleShow.ppt");
        let from_path = crate::open(&path).unwrap();
        let from_bytes = crate::open_bytes(std::fs::read(&path).unwrap()).unwrap();
        let (path_view, bytes_view) = (
            from_path.presentation().unwrap(),
            from_bytes.presentation().unwrap(),
        );
        assert_eq!(
            path_view.slide_count().unwrap(),
            bytes_view.slide_count().unwrap()
        );
        assert_eq!(from_path.text().unwrap(), from_bytes.text().unwrap());
        assert_eq!(
            from_path
                .metadata()
                .unwrap()
                .map(|metadata| format!("{metadata:?}")),
            from_bytes
                .metadata()
                .unwrap()
                .map(|metadata| format!("{metadata:?}"))
        );
    }

    /// A native PPT is not held to a tiny OPC input limit.
    #[cfg(feature = "pptx")]
    #[test]
    fn native_ppt_source_owner_precedes_tiny_pptx_limit() {
        let limits = litchi_opc::SourceLimits::builder()
            .max_input_bytes(1)
            .unwrap()
            .build()
            .unwrap();
        let file = crate::open_with(
            test_data("ole/ppt/SampleShow.ppt"),
            &crate::OpenOptions::new().ooxml_limits(limits),
        )
        .expect("native PPT opens through its source owner");
        assert!(matches!(file, File::Ppt(_)));
    }

    /// A Word compound file is refused as a presentation identically on the
    /// path and byte routes.
    #[cfg(feature = "doc")]
    #[test]
    fn non_powerpoint_ole_is_refused_identically_on_both_routes() {
        let path = test_data("ole/doc/documentProperties.doc");
        let refusal = |file: crate::Result<File>| {
            file.and_then(|file| file.presentation().map(drop))
                .expect_err("a Word document is not a presentation")
        };
        let from_path = refusal(crate::open(&path));
        let from_bytes = refusal(crate::open_bytes(std::fs::read(&path).unwrap()));
        assert_eq!(from_path.code(), from_bytes.code());
        assert_eq!(from_path.detected(), from_bytes.detected());
        assert_eq!(from_path.to_string(), from_bytes.to_string());
    }

    /// `WordDocument` precedes `PowerPoint Document`: a compound file with
    /// both is a Word document on both routes.
    #[cfg(feature = "doc")]
    #[test]
    fn doc_host_stream_keeps_doc_before_ppt_precedence() {
        use std::io::Cursor;

        use litchi_cfb::{OleFile, OleWriter};

        let read = |relative: &str| {
            OleFile::open(
                Cursor::new(std::fs::read(test_data(relative)).unwrap()),
                litchi_cfb::OleFileLimits::default(),
            )
            .unwrap()
        };
        let mut doc = read("ole/doc/documentProperties.doc");
        let mut ppt = read("ole/ppt/SampleShow.ppt");
        let mut writer = OleWriter::new();
        writer
            .create_stream(
                &["WordDocument"],
                &doc.open_stream(&["WordDocument"]).unwrap(),
            )
            .unwrap();
        for stream in ["PowerPoint Document", "Current User"] {
            writer
                .create_stream(&[stream], &ppt.open_stream(&[stream]).unwrap())
                .unwrap();
        }
        let mut output = Cursor::new(Vec::new());
        writer.write_to(&mut output).unwrap();
        let bytes = output.into_inner();

        let outcome = |file: crate::Result<File>| match file {
            Ok(file) => {
                assert_eq!(file.format(), Some(litchi_core::FileFormat::Doc));
                let error = file.presentation().expect_err("a Word document");
                (error.code(), error.to_string())
            },
            Err(error) => {
                assert_ne!(
                    error.format(),
                    Some(litchi_core::FileFormat::Ppt),
                    "{error:?}"
                );
                (error.code(), error.to_string())
            },
        };
        let from_bytes = outcome(crate::open_bytes(bytes.clone()));
        let temporary = super::super::temporary_with(&bytes, "");
        let from_path = outcome(crate::open(temporary.path()));
        assert_eq!(from_path, from_bytes);
    }
}

// ---- source-backed PPTX routing (`source_pptx_path_tests`) ---------------

#[cfg(all(feature = "pptx", any(unix, windows)))]
mod source_pptx_path {
    use std::io::{Cursor, Write as _};

    use super::*;
    use crate::format::pptx::Pptx;
    use crate::{OpenOptions, SourceMode};

    fn two_slide_package(slide1: &[u8], slide2: &[u8]) -> Vec<u8> {
        let mut writer = zip::ZipWriter::new(Cursor::new(Vec::new()));
        let options = zip::write::SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Stored);
        writer.start_file("[Content_Types].xml", options).unwrap();
        writer
            .write_all(
                br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/slides/slide2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>"#,
            )
            .unwrap();
        writer.start_file("_rels/.rels", options).unwrap();
        writer
            .write_all(
                br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>"#,
            )
            .unwrap();
        writer.start_file("ppt/presentation.xml", options).unwrap();
        writer
            .write_all(
                br#"<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldIdLst><p:sldId id="256" r:id="rId1"/><p:sldId id="257" r:id="rId2"/></p:sldIdLst><p:sldSz cx="9144000" cy="6858000"/></p:presentation>"#,
            )
            .unwrap();
        writer
            .start_file("ppt/_rels/presentation.xml.rels", options)
            .unwrap();
        writer
            .write_all(
                br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide2.xml"/></Relationships>"#,
            )
            .unwrap();
        writer.start_file("ppt/slides/slide1.xml", options).unwrap();
        writer.write_all(slide1).unwrap();
        writer.start_file("ppt/slides/slide2.xml", options).unwrap();
        writer.write_all(slide2).unwrap();
        writer.start_file("docProps/core.xml", options).unwrap();
        writer
            .write_all(
                br#"<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>source-backed</dc:title></cp:coreProperties>"#,
            )
            .unwrap();
        writer.finish().unwrap().into_inner()
    }

    const EMPTY_SLIDE: &[u8] = br#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/></p:spTree></p:cSld></p:sld>"#;

    fn corrupt_unselected_slide_package() -> Vec<u8> {
        two_slide_package(EMPTY_SLIDE, b"not PresentationML")
    }

    fn late_reserved_namespace_slide_package() -> Vec<u8> {
        two_slide_package(
            br#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:extLst xmlns:xml="urn:invalid"><p:ext uri="urn:hostile"/></p:extLst></p:sld>"#,
            EMPTY_SLIDE,
        )
    }

    fn pptx(file: &File) -> &Pptx {
        match file {
            File::Pptx(pptx) => pptx,
            other => panic!("expected PPTX, got {other:?}"),
        }
    }

    fn cold_loads(file: &File) -> u64 {
        pptx(file).cold_loads().expect("the source-backed owner")
    }

    #[test]
    fn source_slide_catalog_is_ordered_metadata_only_and_defers_bad_slide() {
        let temporary = super::super::temporary_with(&corrupt_unselected_slide_package(), ".pptx");
        let file = crate::open(temporary.path()).unwrap();
        let before = cold_loads(&file);
        let catalog = pptx(&file)
            .slide_catalog()
            .expect("catalog reads no slide XML");
        assert_eq!(
            catalog
                .iter()
                .map(|slide| (slide.position(), slide.slide_id()))
                .collect::<Vec<_>>(),
            [(0, 256), (1, 257)]
        );
        assert_eq!(cold_loads(&file), before);
        assert!(file.presentation().unwrap().slide(1).is_err());
    }

    /// A slide relationship retargeted at a non-slide part is refused by the
    /// eager owner's count and catalog, as by the source owner's.
    #[test]
    fn eager_slide_catalog_rejects_wrong_slide_target_like_source_catalog() {
        let input = super::super::temporary_with(&corrupt_unselected_slide_package(), ".pptx");
        let mut package = litchi_pptx::Package::open(input.path()).unwrap();
        let second = package
            .presentation()
            .unwrap()
            .slide_references()
            .unwrap()
            .get(1)
            .unwrap()
            .relationship_id()
            .to_owned();
        {
            let mut opc_edit = package.edit_opc().unwrap();
            opc_edit
                .update(|opc| {
                    let presentation = opc.get_part_mut(
                        &litchi_opc::PackURI::new("/ppt/presentation.xml").unwrap(),
                    )?;
                    presentation
                        .rels_mut()
                        .retarget(&second, "../docProps/core.xml".to_owned())?;
                    Ok(())
                })
                .unwrap();
            package = opc_edit.commit().unwrap().into_package();
        }
        let output = tempfile::Builder::new().suffix(".pptx").tempfile().unwrap();
        package.save(output.path()).unwrap();
        let bytes = std::fs::read(output.path()).unwrap();

        let eager = crate::open_bytes_with(
            bytes.clone(),
            &OpenOptions::new().source_mode(SourceMode::Eager),
        )
        .unwrap();
        assert!(!pptx(&eager).is_source_backed());
        assert!(eager.presentation().unwrap().slide_count().is_err());
        assert!(pptx(&eager).slide_catalog().is_err());

        // The source owner validates the slide graph when it opens.
        let source = crate::open_bytes(bytes);
        assert!(source.is_err() || pptx(&source.unwrap()).slide_catalog().is_err());
    }

    #[test]
    fn source_slide_catalog_reports_typed_stale_source() {
        let temporary = super::super::temporary_with(&corrupt_unselected_slide_package(), ".pptx");
        let file = crate::open(temporary.path()).unwrap();
        super::super::mutate(temporary.path());
        let error = pptx(&file).slide_catalog().unwrap_err();
        assert_eq!(error.code(), Code::IoError);
        assert!(super::super::source_change(&error).is_some(), "{error:?}");
    }

    #[test]
    fn path_source_routes_to_the_source_owner_and_defers_corrupt_slide() {
        let bytes = corrupt_unselected_slide_package();
        let temporary = super::super::temporary_with(&bytes, ".pptx");
        let file = crate::open(temporary.path()).unwrap();
        assert!(pptx(&file).is_source_backed());
        let view = file.presentation().unwrap();
        let before = cold_loads(&file);
        assert_eq!(view.slide_count().unwrap(), 2);
        let after_count = cold_loads(&file);
        assert_eq!(after_count, before);

        let metadata = file.metadata().unwrap().expect("core properties");
        assert_eq!(metadata.title.as_deref(), Some("source-backed"));
        let after_metadata = cold_loads(&file);
        assert!(after_metadata > after_count);
        file.metadata().unwrap();
        assert_eq!(cold_loads(&file), after_metadata);
        assert!(view.slide(2).unwrap().is_none());

        let first = view.slide(0).unwrap().expect("first slide");
        first.text().unwrap();
        assert!(cold_loads(&file) > after_metadata);
        assert!(view.slide(1).is_err());
        assert!(view.slides().is_err());

        let from_bytes = crate::open_bytes(bytes).unwrap();
        assert!(pptx(&from_bytes).is_source_backed());
        let view = from_bytes.presentation().unwrap();
        assert_eq!(view.slide_count().unwrap(), 2);
        assert!(view.slide(2).unwrap().is_none());
        assert!(view.slide(1).is_err());
    }

    #[test]
    fn extensionless_path_source_routes_to_the_source_owner() {
        let temporary = super::super::temporary_with(&corrupt_unselected_slide_package(), "");
        let file = crate::open(temporary.path()).unwrap();
        assert!(pptx(&file).is_source_backed());
    }

    #[test]
    fn bytes_source_routes_to_the_source_owner_and_caches_selected_payloads() {
        let file = crate::open_bytes(corrupt_unselected_slide_package()).unwrap();
        assert!(pptx(&file).is_source_backed());
        let view = file.presentation().unwrap();
        let opening = cold_loads(&file);
        assert_eq!(view.slide_count().unwrap(), 2);
        let after_catalog = cold_loads(&file);
        assert_eq!(after_catalog, opening);

        assert!(file.metadata().unwrap().is_some());
        let after_metadata = cold_loads(&file);
        assert!(after_metadata > after_catalog);
        assert!(file.metadata().unwrap().is_some());
        assert_eq!(cold_loads(&file), after_metadata);

        view.slide(0).unwrap().expect("first slide").text().unwrap();
        let after_first = cold_loads(&file);
        assert!(after_first > after_metadata);
        view.slide(0).unwrap().expect("first slide").text().unwrap();
        assert_eq!(cold_loads(&file), after_first);

        // Opening and catalog reads never touch the corrupt second slide;
        // selecting it is the first operation that reports it.
        assert!(view.slide(1).is_err());
    }

    #[test]
    fn bytes_limits_keep_the_source_probe_bounded() {
        let bytes = corrupt_unselected_slide_package();
        let limits = litchi_opc::SourceLimits::builder()
            .max_input_bytes(bytes.len() as u64 - 1)
            .unwrap()
            .build()
            .unwrap();
        crate::detect::reset_opc_probe_count();
        let error = crate::open_bytes_with(bytes.clone(), &OpenOptions::new().ooxml_limits(limits))
            .unwrap_err();
        assert_eq!(error.code(), Code::LimitExceeded);
        assert!(matches!(
            error.leaf(),
            crate::error::Source::Opc(litchi_opc::OpcError::Limit {
                kind: litchi_opc::SourceLimitKind::InputBytes,
                observed: actual,
                maximum,
            }) if *actual == bytes.len() as u64 && *maximum == bytes.len() as u64 - 1
        ));
        // The bounded probe ran once; the typed error triggers no second one.
        assert_eq!(crate::detect::opc_probe_count(), 1);
    }

    #[test]
    fn source_text_rejects_reserved_namespace_prefix_consistently() {
        let bytes = late_reserved_namespace_slide_package();
        let eager = crate::open_bytes_with(
            bytes.clone(),
            &OpenOptions::new().source_mode(SourceMode::Eager),
        )
        .unwrap();
        let temporary = super::super::temporary_with(&bytes, ".pptx");
        let source = crate::open(temporary.path()).unwrap();
        assert!(pptx(&source).is_source_backed());
        let from_bytes = crate::open_bytes(bytes).unwrap();
        for file in [&source, &from_bytes, &eager] {
            let error = file.text().expect_err("a reserved XML prefix is refused");
            assert_eq!(error.code(), Code::CorruptFile, "{error:?}");
        }
        assert!(source.presentation().unwrap().slide(0).is_err());
        assert!(eager.presentation().unwrap().slide(0).is_err());
    }

    #[test]
    fn cached_source_metadata_rechecks_source_revision() {
        let temporary = super::super::temporary_with(&corrupt_unselected_slide_package(), ".pptx");
        let file = crate::open(temporary.path()).unwrap();
        assert!(file.metadata().unwrap().is_some());
        super::super::mutate(temporary.path());
        let error = file
            .metadata()
            .expect_err("a cached answer must not hide a change");
        assert_eq!(error.code(), Code::IoError);
        assert!(super::super::source_change(&error).is_some(), "{error:?}");
    }

    #[test]
    fn cached_source_metadata_rechecks_cancellation() {
        use std::num::{NonZeroU64, NonZeroUsize};

        let temporary = super::super::temporary_with(&corrupt_unselected_slide_package(), ".pptx");
        let budget = litchi_core::Budget::root(
            "litchi-ops-presentation-metadata-test",
            litchi_core::Limits::new(u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
        );
        let (cancellation_source, cancellation) = litchi_core::CancellationSource::pair();
        let limits = litchi_core::ExecutionLimits::new(
            NonZeroUsize::MIN,
            NonZeroUsize::MIN,
            NonZeroU64::MAX,
            0,
        )
        .unwrap();
        let presentation = litchi_pptx::SourceBackedPresentation::open_with_options(
            temporary.path(),
            litchi_pptx::OpenOptions::default()
                .source_limits(litchi_pptx::SourceLimits::default())
                .execution_context(litchi_core::ExecutionContext::new(
                    budget,
                    cancellation,
                    limits,
                )),
        )
        .unwrap();
        let file = File::Pptx(Pptx::from_source(presentation, &OpenOptions::new()));
        assert!(file.metadata().unwrap().is_some());

        cancellation_source.cancel();
        assert_eq!(
            pptx(&file).slide_catalog().unwrap_err().code(),
            Code::Cancelled
        );
        assert_eq!(file.metadata().unwrap_err().code(), Code::Cancelled);
    }

    #[test]
    fn generic_non_zip_is_not_an_office_file_under_tight_ooxml_limits() {
        let temporary = super::super::temporary_with(&[b'x'; 4096], "");
        let limits = litchi_opc::SourceLimits::builder()
            .max_input_bytes(1)
            .unwrap()
            .build()
            .unwrap();
        let error = crate::open_with(temporary.path(), &OpenOptions::new().ooxml_limits(limits))
            .unwrap_err();
        assert_eq!(error.code(), Code::UnsupportedFormat);
        assert_eq!(error.detected(), None);
    }

    #[test]
    fn oversized_ooxml_suffix_non_zip_keeps_the_opc_limit_precedence() {
        let temporary = super::super::temporary_with(&[b'x'; 4096], ".pptx");
        let limits = litchi_opc::SourceLimits::builder()
            .max_input_bytes(1)
            .unwrap()
            .build()
            .unwrap();
        let error = crate::open_with(temporary.path(), &OpenOptions::new().ooxml_limits(limits))
            .unwrap_err();
        assert_eq!(error.code(), Code::LimitExceeded);
        assert!(matches!(
            error.leaf(),
            crate::error::Source::Opc(litchi_opc::OpcError::Limit {
                kind: litchi_opc::SourceLimitKind::InputBytes,
                ..
            })
        ));
    }

    #[test]
    fn malformed_pptx_path_returns_typed_opc_error_without_fallback() {
        let temporary = super::super::temporary_with(b"PK\x03\x04malformed", ".pptx");
        let error = crate::open(temporary.path()).unwrap_err();
        assert_eq!(error.code(), Code::CorruptFile, "{error:?}");
        assert!(
            matches!(
                error.leaf(),
                crate::error::Source::Opc(litchi_opc::OpcError::ZipError(_))
            ),
            "{error:?}"
        );
    }

    /// An ODP opens with its source owner; an unknown ZIP is not an Office
    /// file but stays bound by the caller's OPC limit.
    #[cfg(feature = "odp")]
    #[test]
    fn odp_source_owner_and_unknown_zip_fallback() {
        let mut writer = litchi_odf_common::core::PackageWriter::new(
            litchi_odf_common::core::StreamingArchiveLimits::default(),
        );
        writer
            .set_mimetype(litchi_odf_common::constants::ODF_PRESENTATION)
            .unwrap();
        writer
            .add_file(
                "content.xml",
                br#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><office:body><office:presentation/></office:body></office:document-content>"#,
            )
            .unwrap();
        let odp = super::super::temporary_with(&writer.finish_to_bytes().unwrap(), ".odp");
        let file = crate::open(odp.path()).unwrap();
        let File::Odp(handle) = &file else {
            panic!("expected ODP, got {file:?}");
        };
        assert!(matches!(
            handle.owner(),
            crate::format::odp::OdpOwner::Source(_)
        ));

        let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
        zip.start_file(
            "plain.txt",
            zip::write::SimpleFileOptions::default()
                .compression_method(zip::CompressionMethod::Stored),
        )
        .unwrap();
        zip.write_all(b"not an Office package").unwrap();
        let unknown_bytes = zip.finish().unwrap().into_inner();
        let unknown = super::super::temporary_with(&unknown_bytes, ".zip");
        let limits = litchi_opc::SourceLimits::builder()
            .max_input_bytes(1)
            .unwrap()
            .build()
            .unwrap();
        let error =
            crate::open_with(unknown.path(), &OpenOptions::new().ooxml_limits(limits)).unwrap_err();
        assert!(matches!(
            error.leaf(),
            crate::error::Source::Opc(litchi_opc::OpcError::Limit {
                kind: litchi_opc::SourceLimitKind::InputBytes,
                observed: actual,
                maximum: 1,
            }) if *actual == unknown_bytes.len() as u64
        ));
        let error = crate::open(unknown.path()).unwrap_err();
        assert_eq!(error.code(), Code::UnsupportedFormat);
        assert_eq!(error.detected(), None);
    }
}
