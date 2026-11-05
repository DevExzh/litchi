//! The `Locked` and `Signed` states: encrypted inputs classified at open
//! (and unlocked under the `encryption` feature), and signed inputs whose
//! reads go to the file they wrap.

/// Compound-file fixtures and builders for the `Locked` and `Signed` tests.
#[cfg(any(
    feature = "doc",
    feature = "ppt",
    feature = "xls",
    feature = "encryption"
))]
#[cfg(any(
    feature = "doc",
    feature = "docx",
    feature = "ppt",
    feature = "pptx",
    feature = "xls",
    feature = "xlsx",
    feature = "xlsb",
    feature = "rtf",
    feature = "odt",
    feature = "ods",
    feature = "odp"
))]
mod cfb {
    use std::io::Cursor;

    use litchi_cfb::{OleFile, OleWriter};

    /// A compound file holding `streams`, top-level streams only.
    pub(super) fn with_streams(streams: &[(&str, &[u8])]) -> Vec<u8> {
        let mut writer = OleWriter::new();
        for (name, bytes) in streams {
            writer.create_stream(&[*name], bytes).unwrap();
        }
        let mut output = Cursor::new(Vec::new());
        writer.write_to(&mut output).unwrap();
        output.into_inner()
    }

    /// An encrypted OOXML container: `EncryptionInfo` with `version` and an
    /// `EncryptedPackage`, without any real key material.
    pub(super) fn encrypted_ooxml(version: [u8; 4]) -> Vec<u8> {
        let mut info = version.to_vec();
        info.extend_from_slice(&[0x40, 0, 0, 0]);
        with_streams(&[
            ("EncryptionInfo", &info),
            (
                "EncryptedPackage",
                &[8, 0, 0, 0, 0, 0, 0, 0, 1, 2, 3, 4, 5, 6, 7, 8],
            ),
        ])
    }

    /// `bytes` rewritten with an added (empty) `_xmlsignatures` storage
    /// stream, the way a signing producer marks a binary document.
    #[allow(
        dead_code,
        reason = "used by the signed suites of the legacy formats enabled"
    )]
    pub(super) fn signed(bytes: Vec<u8>) -> Vec<u8> {
        let mut source =
            OleFile::open(Cursor::new(bytes), litchi_cfb::OleFileLimits::default()).unwrap();
        let mut writer = OleWriter::new();
        for path in source.list_streams() {
            let path = path.iter().map(String::as_str).collect::<Vec<_>>();
            for depth in 1..path.len() {
                writer.create_storage(&path[..depth]).unwrap();
            }
            let data = source.open_stream(&path).unwrap();
            writer.create_stream(&path, &data).unwrap();
        }
        writer
            .create_stream(&["_signatures"], b"signature")
            .unwrap();
        let mut output = Cursor::new(Vec::new());
        writer.write_to(&mut output).unwrap();
        output.into_inner()
    }
}

/// Encrypted inputs open as `Locked`, classified without a password.
#[cfg(any(
    feature = "doc",
    feature = "ppt",
    feature = "xls",
    feature = "encryption"
))]
#[cfg(any(
    feature = "doc",
    feature = "docx",
    feature = "ppt",
    feature = "pptx",
    feature = "xls",
    feature = "xlsx",
    feature = "xlsb",
    feature = "rtf",
    feature = "odt",
    feature = "ods",
    feature = "odp"
))]
mod locked {
    use litchi_core::FileFormat;

    use super::super::test_data;
    use crate::{Code, EncryptionKind, File, Locked};

    /// An encrypted legacy fixture: path, format, scheme and password.
    pub(super) type Legacy = (&'static str, FileFormat, EncryptionKind, &'static str);

    /// The encrypted legacy fixtures of the enabled formats.
    pub(super) fn legacy() -> Vec<Legacy> {
        let groups: &[&[Legacy]] = &[
            #[cfg(feature = "doc")]
            &[
                (
                    "poi/test-data/document/password_password_cryptoapi.doc",
                    FileFormat::Doc,
                    EncryptionKind::CryptoApi,
                    "password",
                ),
                (
                    "poi/test-data/document/password_tika_binaryrc4.doc",
                    FileFormat::Doc,
                    EncryptionKind::BinaryRc4,
                    "tika",
                ),
            ],
            #[cfg(feature = "xls")]
            &[
                (
                    "poi/test-data/spreadsheet/xor-encryption-abc.xls",
                    FileFormat::Xls,
                    EncryptionKind::XorObfuscation,
                    "abc",
                ),
                (
                    "ole/xls/password.xls",
                    FileFormat::Xls,
                    EncryptionKind::BinaryRc4,
                    "password",
                ),
            ],
            #[cfg(feature = "ppt")]
            &[
                (
                    "poi/test-data/slideshow/Password_Protected-hello.ppt",
                    FileFormat::Ppt,
                    EncryptionKind::CryptoApi,
                    "hello",
                ),
                (
                    "poi/test-data/slideshow/Password_Protected-np-hello.ppt",
                    FileFormat::Ppt,
                    EncryptionKind::CryptoApi,
                    "hello",
                ),
            ],
        ];
        groups.concat()
    }

    #[track_caller]
    pub(super) fn locked(file: File) -> Locked {
        match file {
            File::Locked(locked) => locked,
            other => panic!("expected a locked file, got {other:?}"),
        }
    }

    /// Every read of a locked file is refused with the protection code.
    #[track_caller]
    fn assert_reads_refused(file: &File) {
        for error in [
            file.text().map(drop).unwrap_err(),
            file.metadata().map(drop).unwrap_err(),
            file.tolerance_report().map(drop).unwrap_err(),
        ] {
            assert_eq!(error.code(), Code::DocumentProtected, "{error:?}");
            let suggestion = error
                .suggestion()
                .expect("a locked refusal suggests unlocking");
            assert!(suggestion.contains("Locked::unlock"), "{suggestion}");
            #[cfg(not(feature = "encryption"))]
            assert!(suggestion.contains("`encryption`"), "{suggestion}");
        }
    }

    #[test]
    fn encrypted_legacy_documents_open_locked_with_their_scheme() {
        for (relative, format, kind, _) in legacy() {
            let path = test_data(relative);
            for file in [
                crate::open(&path).unwrap_or_else(|error| panic!("{relative}: {error:?}")),
                crate::open_bytes(std::fs::read(&path).unwrap())
                    .unwrap_or_else(|error| panic!("{relative}: {error:?}")),
            ] {
                assert_eq!(file.format(), Some(format), "{relative}");
                assert_reads_refused(&file);
                let handle = locked(file);
                assert_eq!(handle.format(), Some(format), "{relative}");
                assert_eq!(handle.family(), Some(crate::Family::of(format)));
                assert_eq!(handle.encryption(), kind, "{relative}");
                let debug = format!("{handle:?}");
                assert!(
                    debug.starts_with("Locked {") && debug.ends_with(".. }"),
                    "{debug}"
                );
            }
        }
    }

    /// A family view of a locked file is refused as protected when the file
    /// is, or may be, of that family, and as the wrong format otherwise.
    #[test]
    fn family_views_of_a_locked_file() {
        for (relative, format, _, _) in legacy() {
            let file = crate::open(test_data(relative)).unwrap();
            let views: Vec<(crate::Family, crate::Result<()>)> = vec![
                #[cfg(any(feature = "doc", feature = "docx", feature = "rtf", feature = "odt"))]
                (crate::Family::Document, file.document().map(drop)),
                #[cfg(any(feature = "ppt", feature = "pptx", feature = "odp"))]
                (crate::Family::Presentation, file.presentation().map(drop)),
                #[cfg(any(feature = "xls", feature = "xlsx", feature = "xlsb", feature = "ods"))]
                (crate::Family::Workbook, file.workbook().map(drop)),
            ];
            for (family, result) in views {
                let error = result.expect_err("a locked file has no view");
                if family == crate::Family::of(format) {
                    assert_eq!(error.code(), Code::DocumentProtected, "{relative}");
                } else {
                    assert_eq!(error.code(), Code::UnsupportedFormat, "{relative}");
                    assert_eq!(error.detected(), Some(format), "{relative}");
                }
            }
        }
    }

    /// An encrypted OOXML package hides its format: a file name names it,
    /// and in-memory input reports none. The scheme comes from the
    /// `EncryptionInfo` version.
    #[test]
    fn encrypted_ooxml_containers_are_classified_by_their_encryption_info() {
        for (version, kind) in [
            ([4, 0, 4, 0], EncryptionKind::Agile),
            ([4, 0, 2, 0], EncryptionKind::Standard),
            ([3, 0, 2, 0], EncryptionKind::Standard),
            ([9, 0, 9, 0], EncryptionKind::Unknown),
        ] {
            let bytes = super::cfb::encrypted_ooxml(version);
            let file = crate::open_bytes(bytes.clone()).unwrap();
            assert_eq!(file.format(), None);
            assert_eq!(file.family(), None);
            assert_reads_refused(&file);
            let handle = locked(file);
            assert_eq!(handle.encryption(), kind);
            assert_eq!(handle.format(), None);

            #[cfg(any(unix, windows))]
            for (suffix, named) in [
                (".docx", Some(FileFormat::Docx)),
                (".XLSX", Some(FileFormat::Xlsx)),
                (".pptm", Some(FileFormat::Pptx)),
                (".xlsb", Some(FileFormat::Xlsb)),
                ("", None),
                (".bin", None),
            ] {
                let temporary = super::super::temporary_with(&bytes, suffix);
                let handle = locked(crate::open(temporary.path()).unwrap());
                assert_eq!(handle.format(), named, "{suffix}");
                assert_eq!(handle.encryption(), kind, "{suffix}");
            }
        }
    }

    /// Word 2007 writes a stub `WordDocument` stream next to the encrypted
    /// package (LibreOffice's `Encrypted_MSO2007_abc.docx`); the container
    /// is still an encrypted OOXML package, not a DOC (record 0341).
    #[test]
    fn an_encrypted_ooxml_package_with_a_stub_word_stream_is_locked() {
        let mut info = vec![3, 0, 2, 0];
        info.extend_from_slice(&[0x40, 0, 0, 0]);
        let bytes = super::cfb::with_streams(&[
            ("EncryptionInfo", &info),
            (
                "EncryptedPackage",
                &[8, 0, 0, 0, 0, 0, 0, 0, 1, 2, 3, 4, 5, 6, 7, 8],
            ),
            ("WordDocument", &[0; 64]),
        ]);
        let handle = locked(crate::open_bytes(bytes.clone()).unwrap());
        assert_eq!(handle.encryption(), EncryptionKind::Standard);
        assert_eq!(handle.format(), None);
        #[cfg(any(unix, windows))]
        {
            let temporary = super::super::temporary_with(&bytes, ".docx");
            let handle = locked(crate::open(temporary.path()).unwrap());
            assert_eq!(handle.format(), Some(FileFormat::Docx));
            assert_eq!(handle.encryption(), EncryptionKind::Standard);
            // Detection agrees with the opener: not a DOC.
            assert_eq!(
                crate::detect::detect_file_format(
                    temporary.path(),
                    &crate::detect::SourceLimits::default(),
                ),
                None
            );
        }
        assert!(
            crate::detect::detect_format_smart(bytes, &crate::detect::SourceLimits::default(),)
                .is_none()
        );
    }

    /// The facade's suffix preflight still refuses any other compound file
    /// under an OOXML name (`ooxml_suffix_preflight_precedes_native_ole_routing`).
    #[cfg(all(
        any(unix, windows),
        any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb")
    ))]
    #[test]
    fn an_ooxml_name_on_another_compound_file_is_refused() {
        let mut fixtures = vec![litchi_core::detection::utils::OLE2_SIGNATURE.to_vec()];
        fixtures[0].resize(512, 0);
        for (relative, _, _, _) in legacy() {
            fixtures.push(std::fs::read(test_data(relative)).unwrap());
        }
        #[cfg(feature = "ppt")]
        fixtures.push(std::fs::read(test_data("ole/ppt/SampleShow.ppt")).unwrap());
        for bytes in fixtures {
            for suffix in [".pptx", ".docx", ".xlsx"] {
                let temporary = super::super::temporary_with(&bytes, suffix);
                let error = crate::open(temporary.path()).expect_err("not an OOXML package");
                assert_eq!(error.code(), Code::CorruptFile, "{error:?}");
                assert!(
                    error
                        .message()
                        .contains("OOXML-suffixed input does not have ZIP magic"),
                    "{error:?}"
                );
                let eager = crate::OpenOptions::new().source_mode(crate::SourceMode::Eager);
                let error = crate::open_with(temporary.path(), &eager).expect_err("eager too");
                assert_eq!(error.code(), Code::CorruptFile, "{error:?}");
            }
        }
    }

    #[test]
    fn unified_reads_refuse_locked_files_before_parsing_inputs() {
        let file = crate::open_bytes(super::cfb::encrypted_ooxml([4, 0, 4, 0])).unwrap();
        assert_eq!(file.root().unwrap_err().code(), Code::DocumentProtected);
        assert_eq!(
            file.get("INVALID SECRET").unwrap_err().code(),
            Code::DocumentProtected
        );
        assert_eq!(
            file.find("INVALID SECRET").unwrap_err().code(),
            Code::DocumentProtected
        );
        assert_eq!(
            file.query("INVALID SECRET").unwrap_err().code(),
            Code::DocumentProtected
        );
        assert_eq!(
            file.view(crate::view::Mode::Text).unwrap_err().code(),
            Code::DocumentProtected
        );
    }

    /// Clones of a locked handle share it; unlocking one leaves the others.
    #[test]
    fn locked_handles_are_cheap_clones() {
        let file = crate::open_bytes(super::cfb::encrypted_ooxml([4, 0, 4, 0])).unwrap();
        let clone = file.clone();
        drop(file);
        assert_eq!(locked(clone).encryption(), EncryptionKind::Agile);
    }
}

/// `Locked::unlock`: decrypt the retained source once and open the content.
#[cfg(all(
    feature = "encryption",
    any(
        feature = "doc",
        feature = "docx",
        feature = "ppt",
        feature = "pptx",
        feature = "xls",
        feature = "xlsx",
        feature = "xlsb",
        feature = "rtf",
        feature = "odt",
        feature = "ods",
        feature = "odp"
    )
))]
mod unlock {
    #[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
    use litchi_core::FileFormat;

    use super::super::test_data;
    use super::locked::locked;
    #[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx"))]
    use crate::EncryptionKind;
    #[cfg(feature = "docx")]
    use crate::File;
    use crate::{Code, Password};

    fn password(value: &str) -> Password {
        Password::new(value.to_owned())
    }

    #[test]
    fn legacy_documents_unlock_with_their_password() {
        for (relative, format, _, secret) in super::locked::legacy() {
            let path = test_data(relative);
            for file in [
                crate::open(&path).unwrap(),
                crate::open_bytes(std::fs::read(&path).unwrap()).unwrap(),
            ] {
                let clone = locked(file.clone());
                let unlocked = locked(file)
                    .unlock(password(secret))
                    .unwrap_or_else(|error| panic!("{relative}: {error:?}"));
                assert_eq!(unlocked.format(), Some(format), "{relative}");
                unlocked
                    .text()
                    .unwrap_or_else(|error| panic!("{relative}: {error:?}"));

                // A clone that is still shared unlocks independently, and a
                // wrong password is refused without reading anything.
                let error = clone.unlock(password("not the password")).unwrap_err();
                assert!(
                    matches!(
                        error.code(),
                        Code::InvalidArgument | Code::DocumentProtected
                    ),
                    "{relative}: {error:?}"
                );
            }
        }
    }

    /// LibreOffice's Word 2007 fixture (a stub `WordDocument` stream beside
    /// the encrypted package) opens locked and unlocks with its password
    /// (record 0341).
    #[cfg(feature = "docx")]
    #[test]
    fn a_word_2007_encrypted_docx_unlocks() {
        let path =
            test_data("libreoffice-core/sw/qa/extras/ooxmlexport/data/Encrypted_MSO2007_abc.docx");
        let handle = locked(crate::open(&path).unwrap());
        assert_eq!(handle.format(), Some(FileFormat::Docx));
        assert_eq!(handle.encryption(), EncryptionKind::Standard);
        let unlocked = handle.unlock(password("abc")).unwrap();
        assert_eq!(unlocked.format(), Some(FileFormat::Docx));
        assert!(!unlocked.text().unwrap().trim().is_empty());
    }

    /// Encrypt a plain OOXML fixture, open it locked, and unlock it back to
    /// the same content.
    #[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx"))]
    fn round_trip(
        relative: &str,
        format: FileFormat,
        mode: litchi_crypto::ooxml::Mode,
        kind: EncryptionKind,
    ) {
        let plain = std::fs::read(test_data(relative)).unwrap();
        let expected = crate::open_bytes(plain.clone()).unwrap().text().unwrap();
        let encrypted = litchi_crypto::ooxml::encrypt(
            plain,
            "unlock-test",
            mode,
            &litchi_crypto::ooxml::Limits::default(),
        )
        .unwrap();

        let file = crate::open_bytes(encrypted.clone()).unwrap();
        assert_eq!(file.format(), None);
        let handle = locked(file);
        assert_eq!(handle.encryption(), kind);
        let unlocked = handle.clone().unlock(password("unlock-test")).unwrap();
        assert_eq!(unlocked.format(), Some(format));
        assert_eq!(unlocked.text().unwrap(), expected);

        let error = handle.unlock(password("wrong")).unwrap_err();
        assert_eq!(error.code(), Code::InvalidArgument, "{error:?}");
        assert!(error.suggestion().is_some());

        #[cfg(any(unix, windows))]
        {
            let extension = relative.rsplit('.').next().unwrap();
            let temporary = super::super::temporary_with(&encrypted, &format!(".{extension}"));
            let handle = locked(crate::open(temporary.path()).unwrap());
            assert_eq!(handle.format(), Some(format));
            let unlocked = handle.unlock(password("unlock-test")).unwrap();
            assert_eq!(unlocked.text().unwrap(), expected);
        }
    }

    #[cfg(feature = "docx")]
    #[test]
    fn an_encrypted_docx_unlocks() {
        use litchi_crypto::ooxml::Mode;
        round_trip(
            "ooxml/docx/FancyFoot.docx",
            FileFormat::Docx,
            Mode::Agile,
            EncryptionKind::Agile,
        );
        round_trip(
            "ooxml/docx/comment.docx",
            FileFormat::Docx,
            Mode::Standard,
            EncryptionKind::Standard,
        );
    }

    #[cfg(feature = "xlsx")]
    #[test]
    fn an_encrypted_xlsx_unlocks() {
        use litchi_crypto::ooxml::Mode;
        round_trip(
            "ooxml/xlsx/styles.xlsx",
            FileFormat::Xlsx,
            Mode::AgileSha512,
            EncryptionKind::Agile,
        );
        round_trip(
            "ooxml/xlsx/styles.xlsx",
            FileFormat::Xlsx,
            Mode::Standard,
            EncryptionKind::Standard,
        );
    }

    #[cfg(feature = "pptx")]
    #[test]
    fn an_encrypted_pptx_unlocks() {
        use litchi_crypto::ooxml::Mode;
        round_trip(
            "ooxml/pptx/shapes.pptx",
            FileFormat::Pptx,
            Mode::Agile,
            EncryptionKind::Agile,
        );
    }

    /// The unlocked handle keeps the leaf's eager owner with the encryption
    /// profile, so a later save keeps the package encrypted.
    #[cfg(feature = "docx")]
    #[test]
    fn an_unlocked_docx_keeps_its_encryption_profile() {
        let plain = std::fs::read(test_data("ooxml/docx/FancyFoot.docx")).unwrap();
        let encrypted = litchi_crypto::ooxml::encrypt(
            plain,
            "profile",
            litchi_crypto::ooxml::Mode::Agile,
            &litchi_crypto::ooxml::Limits::default(),
        )
        .unwrap();
        let file = locked(crate::open_bytes(encrypted).unwrap())
            .unlock(password("profile"))
            .unwrap();
        let File::Docx(docx) = &file else {
            panic!("expected DOCX, got {file:?}");
        };
        match docx.owner() {
            crate::format::docx::DocxOwner::Eager(package) => {
                assert!(package.encryption().is_some());
            },
            crate::format::docx::DocxOwner::Source(_) => panic!("expected the eager owner"),
        }
    }

    /// Every OOXML eager owner retains the selected outer encryption profile
    /// across the facade's locked-to-unlocked transition. A deliberately
    /// tight output ceiling must therefore refuse re-encryption for DOCX,
    /// PPTX, and XLSX alike rather than silently reverting to the default.
    #[cfg(all(feature = "docx", feature = "pptx", feature = "xlsx"))]
    #[test]
    fn unlocked_ooxml_consumers_retain_custom_encryption_limits() {
        use crate::{File, OpenOptions, SaveOptions, SaveTo};

        let encryption_limits = litchi_crypto::ooxml::Limits {
            max_output_bytes: 512,
            ..litchi_crypto::ooxml::Limits::default()
        };
        for (relative, suffix) in [
            ("ooxml/docx/FancyFoot.docx", ".docx"),
            ("ooxml/pptx/shapes.pptx", ".pptx"),
            ("ooxml/xlsx/styles.xlsx", ".xlsx"),
        ] {
            let plain = std::fs::read(test_data(relative)).unwrap();
            let encrypted = litchi_crypto::ooxml::encrypt(
                plain,
                "profile",
                litchi_crypto::ooxml::Mode::Agile,
                &litchi_crypto::ooxml::Limits::default(),
            )
            .unwrap();
            let locked = crate::open_bytes_with(
                encrypted,
                &OpenOptions::new().ooxml_encryption_limits(encryption_limits),
            )
            .unwrap();
            let unlocked = match locked {
                File::Locked(locked) => locked.unlock(password("profile")).unwrap(),
                other => panic!("expected locked {relative}, got {other:?}"),
            };
            let directory = tempfile::tempdir().unwrap();
            let destination = directory.path().join(format!("output{suffix}"));
            let error = unlocked
                .save_with(
                    SaveTo::Path(&destination),
                    &SaveOptions::new().reencrypt(password("new")),
                )
                .unwrap_err();
            assert_eq!(error.code(), Code::LimitExceeded, "{relative}: {error:?}");
        }
    }

    #[cfg(feature = "xlsb")]
    #[test]
    fn an_encrypted_xlsb_is_refused_as_an_unsupported_capability() {
        let plain = std::fs::read(test_data("ooxml/xlsb/Simple.xlsb")).unwrap();
        let encrypted = litchi_crypto::ooxml::encrypt(
            plain,
            "xlsb",
            litchi_crypto::ooxml::Mode::Agile,
            &litchi_crypto::ooxml::Limits::default(),
        )
        .unwrap();
        let error = locked(crate::open_bytes(encrypted).unwrap())
            .unlock(password("xlsb"))
            .unwrap_err();
        assert_eq!(error.code(), Code::UnsupportedCapability, "{error:?}");
        assert_eq!(error.format(), Some(FileFormat::Xlsb));
    }

    /// A signed package keeps its signature through decryption.
    #[cfg(feature = "docx")]
    #[test]
    fn a_signed_package_unlocks_as_signed() {
        let plain =
            std::fs::read(test_data("poi/test-data/xmldsign/hello-world-signed.docx")).unwrap();
        let encrypted = litchi_crypto::ooxml::encrypt(
            plain,
            "signed",
            litchi_crypto::ooxml::Mode::Agile,
            &litchi_crypto::ooxml::Limits::default(),
        )
        .unwrap();
        let file = locked(crate::open_bytes(encrypted).unwrap())
            .unlock(password("signed"))
            .unwrap();
        let File::Signed(signed) = &file else {
            panic!("expected a signed file, got {file:?}");
        };
        assert_eq!(signed.format(), FileFormat::Docx);
        assert!(!file.text().unwrap().is_empty());
    }

    /// A positional locked source is pinned to the revision opened.
    #[cfg(all(feature = "xls", any(unix, windows)))]
    #[test]
    fn a_changed_locked_source_is_reported() {
        let temporary = super::super::temporary_copy(
            &test_data("poi/test-data/spreadsheet/xor-encryption-abc.xls"),
            "",
        );
        let handle = locked(crate::open(temporary.path()).unwrap());
        super::super::mutate(temporary.path());
        let error = handle.unlock(password("abc")).unwrap_err();
        assert_eq!(error.code(), Code::IoError, "{error:?}");
    }
}

/// Signed inputs open as `Signed` and read through the file they wrap.
mod signed {
    #[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx"))]
    #[test]
    fn signed_ooxml_packages_open_signed_on_every_route() {
        use litchi_core::FileFormat;

        use crate::{File, OpenOptions, SignatureOrigin, SourceMode};

        let mut fixtures = Vec::new();
        #[cfg(feature = "docx")]
        fixtures.extend([
            (
                "poi/test-data/xmldsign/hello-world-signed.docx",
                FileFormat::Docx,
            ),
            (
                "poi/test-data/xmldsign/hello-world-signed-twice.docx",
                FileFormat::Docx,
            ),
            (
                "poi/test-data/xmldsign/ms-office-2010-signed.docx",
                FileFormat::Docx,
            ),
        ]);
        #[cfg(feature = "pptx")]
        fixtures.extend([
            (
                "poi/test-data/xmldsign/hello-world-signed.pptx",
                FileFormat::Pptx,
            ),
            (
                "poi/test-data/xmldsign/ms-office-2010-signed.pptx",
                FileFormat::Pptx,
            ),
        ]);
        #[cfg(feature = "xlsx")]
        fixtures.extend([
            (
                "poi/test-data/xmldsign/hello-world-signed.xlsx",
                FileFormat::Xlsx,
            ),
            (
                "poi/test-data/xmldsign/ms-office-2010-signed.xlsx",
                FileFormat::Xlsx,
            ),
        ]);
        let eager = OpenOptions::new().source_mode(SourceMode::Eager);
        for (relative, format) in fixtures {
            let path = super::super::test_data(relative);
            let bytes = std::fs::read(&path).unwrap();
            for file in [
                crate::open(&path).unwrap(),
                crate::open_with(&path, &eager).unwrap(),
                crate::open_bytes(bytes.clone()).unwrap(),
                crate::open_bytes_with(bytes.clone(), &eager).unwrap(),
            ] {
                let File::Signed(signed) = &file else {
                    panic!("{relative}: expected a signed file, got {file:?}");
                };
                assert_eq!(signed.origin(), SignatureOrigin::Opc);
                assert_eq!(signed.format(), format);
                assert_eq!(signed.family(), crate::Family::of(format));
                assert_eq!(signed.file().format(), Some(format));
                assert!(!matches!(signed.file(), File::Signed(_) | File::Locked(_)));
                assert_eq!(file.format(), Some(format));
                assert_eq!(file.text().unwrap(), signed.file().text().unwrap());
                assert_eq!(
                    format!("{:?}", file.metadata().unwrap()),
                    format!("{:?}", signed.file().metadata().unwrap())
                );
                file.tolerance_report().unwrap();
                let debug = format!("{signed:?}");
                assert!(debug.contains("Opc") && debug.ends_with(".. }"), "{debug}");
                match format {
                    #[cfg(feature = "docx")]
                    FileFormat::Docx => {
                        assert!(!file.document().unwrap().paragraphs().unwrap().is_empty())
                    },
                    #[cfg(feature = "pptx")]
                    FileFormat::Pptx => {
                        assert!(file.presentation().unwrap().slide_count().unwrap() > 0)
                    },
                    #[cfg(feature = "xlsx")]
                    FileFormat::Xlsx => {
                        assert!(file.workbook().unwrap().worksheet_count().unwrap() > 0)
                    },
                    other => panic!("unexpected {other:?}"),
                }
            }
        }
    }

    #[cfg(feature = "docx")]
    #[test]
    fn unsigned_packages_do_not_open_signed() {
        let file = crate::open(super::super::test_data("ooxml/docx/FancyFoot.docx")).unwrap();
        assert!(matches!(file, crate::File::Docx(_)));
    }

    /// A compound file with a `_signatures` stream opens as `Signed` on the
    /// in-memory and positional routes of each legacy format.
    #[cfg(any(feature = "doc", feature = "ppt", feature = "xls"))]
    #[test]
    fn signed_compound_files_open_signed() {
        use litchi_core::FileFormat;

        use crate::{File, SignatureOrigin};

        let fixtures = [
            #[cfg(feature = "doc")]
            ("ole/doc/FancyFoot.doc", FileFormat::Doc),
            #[cfg(feature = "ppt")]
            ("ole/ppt/SampleShow.ppt", FileFormat::Ppt),
            #[cfg(feature = "xls")]
            ("ole/xls/Simple.xls", FileFormat::Xls),
        ];
        for (relative, format) in fixtures {
            let plain = std::fs::read(super::super::test_data(relative)).unwrap();
            let expected = crate::open_bytes(plain.clone()).unwrap().text().unwrap();
            let bytes = super::cfb::signed(plain);
            let mut files = vec![crate::open_bytes(bytes.clone()).unwrap()];
            #[cfg(any(unix, windows))]
            let temporary = super::super::temporary_with(&bytes, "");
            #[cfg(any(unix, windows))]
            files.push(crate::open(temporary.path()).unwrap());
            for file in files {
                let File::Signed(signed) = &file else {
                    panic!("{relative}: expected a signed file, got {file:?}");
                };
                assert_eq!(signed.origin(), SignatureOrigin::Cfb);
                assert_eq!(signed.format(), format);
                assert_eq!(file.format(), Some(format));
                assert_eq!(file.text().unwrap(), expected, "{relative}");
            }
        }
    }
}
