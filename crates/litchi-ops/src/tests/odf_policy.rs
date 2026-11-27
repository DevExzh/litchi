//! The selected ODF options reach every owned and flat projection.

use std::io::Cursor;

use litchi_core::FileFormat;
use litchi_odf_common::{LimitKind, SemanticLimitKind, SemanticLimits};

use crate::{Code, OpenOptions};

fn content_options(maximum: u64) -> OpenOptions {
    let semantic = SemanticLimits::builder()
        .max_content_bytes(maximum)
        .unwrap()
        .build()
        .unwrap();
    OpenOptions::new().odf_options(litchi_odf_common::OpenOptions::new().semantic_limits(semantic))
}

#[test]
fn prepared_odf_owners_apply_semantic_options_without_changing_source_bytes() {
    let formats = [
        #[cfg(feature = "odt")]
        FileFormat::Odt,
        #[cfg(feature = "ods")]
        FileFormat::Ods,
        #[cfg(feature = "odp")]
        FileFormat::Odp,
        #[cfg(feature = "oth")]
        FileFormat::Oth,
        #[cfg(feature = "odm")]
        FileFormat::Odm,
        #[cfg(feature = "odg")]
        FileFormat::Odg,
        #[cfg(feature = "odc")]
        FileFormat::Odc,
        #[cfg(feature = "odi")]
        FileFormat::Odi,
        #[cfg(feature = "odb")]
        FileFormat::Odb,
        #[cfg(feature = "odf-formula")]
        FileFormat::Odf,
    ];
    for format in formats {
        let bytes = crate::create(format).unwrap().to_bytes().unwrap();
        let length = zip::ZipArchive::new(Cursor::new(&bytes))
            .unwrap()
            .by_name("content.xml")
            .unwrap()
            .size();
        assert!(length > 1);
        let exact = crate::open_bytes_with(bytes.clone(), &content_options(length))
            .unwrap_or_else(|error| panic!("{format:?} exact limit: {error}"));
        assert_eq!(
            exact.to_bytes().unwrap(),
            bytes,
            "{format:?} source fidelity"
        );
        #[cfg(feature = "odt")]
        if format == FileFormat::Odt {
            let prepared = litchi_odf_common::detect::prepared(
                bytes.clone(),
                litchi_odf_common::SourceLimits::default(),
            )
            .unwrap();
            let identity = prepared.prepared_index_identity();
            let owner = litchi_odt::Package::from_prepared_package_with_options(
                prepared,
                content_options(length).odf_open_options(),
            )
            .unwrap();
            assert_eq!(
                owner.document().unwrap().prepared_index_identity(),
                identity
            );
        }
        let error = crate::open_bytes_with(bytes.clone(), &content_options(length - 1))
            .expect_err("a prepared owner must enforce the selected semantic ceiling");
        assert_eq!(error.code(), Code::LimitExceeded, "{format:?}: {error}");
        assert!(
            matches!(error.leaf(),
                crate::error::Source::Odf(litchi_odf_common::Error::Limit {
                    kind: LimitKind::Semantic(SemanticLimitKind::ContentBytes),
                    observed,
                    maximum,
                }) if *observed == length && *maximum == length - 1
            ),
            "{format:?}: {error}"
        );

        #[cfg(any(unix, windows))]
        {
            let source = super::temporary_with(&bytes, ".odf");
            let exact = crate::open_with(source.path(), &content_options(length))
                .unwrap_or_else(|error| panic!("{format:?} exact path limit: {error}"));
            assert_eq!(exact.to_bytes().unwrap(), bytes, "{format:?} path fidelity");
            let error = crate::open_with(source.path(), &content_options(length - 1))
                .expect_err("the path owner must enforce the same semantic ceiling");
            assert_eq!(
                error.code(),
                Code::LimitExceeded,
                "{format:?} path: {error}"
            );
        }
    }
}

#[test]
#[cfg(any(feature = "odt", feature = "ods", feature = "odp", feature = "odg"))]
fn flat_odf_owners_apply_semantic_options() {
    let fixtures = [
        #[cfg(feature = "odt")]
        "odf/odt/font-face-declarations-flat.fodt",
        #[cfg(feature = "ods")]
        "odf/drawing/marker-flat.fods",
        #[cfg(feature = "odp")]
        "odf/drawing/fill-image-linked.fodp",
        #[cfg(feature = "odg")]
        "odf/drawing/fill-image-inline.fodg",
    ];
    for relative in fixtures {
        let bytes = std::fs::read(super::test_data(relative)).unwrap();
        let ordinary = crate::open_bytes(bytes.clone()).unwrap();
        assert_eq!(ordinary.to_bytes().unwrap(), bytes);
        let error = crate::open_bytes_with(bytes.clone(), &content_options(1))
            .expect_err("flat owners must apply the selected semantic ceiling");
        assert_eq!(error.code(), Code::LimitExceeded, "{relative}: {error}");
        assert!(
            matches!(error.leaf(),
                crate::error::Source::Odf(litchi_odf_common::Error::Limit {
                    kind: LimitKind::Semantic(SemanticLimitKind::ContentBytes),
                    observed,
                    maximum: 1,
                }) if *observed > 1
            ),
            "{relative}: {error}"
        );
        let error = crate::open_with(super::test_data(relative), &content_options(1))
            .expect_err("flat path owners must apply the selected semantic ceiling");
        assert_eq!(
            error.code(),
            Code::LimitExceeded,
            "{relative} path: {error}"
        );
    }
}
