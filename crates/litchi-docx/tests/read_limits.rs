use std::io::Cursor;

use litchi_docx::{Error, Package, SourceLimits};
use litchi_opc::{OpcError, SourceLimitKind};
use tempfile::NamedTempFile;

fn input_limit(bytes: usize) -> SourceLimits {
    SourceLimits::builder()
        .max_input_bytes(u64::try_from(bytes).expect("package size fits in u64"))
        .expect("valid input limit")
        .build()
        .expect("valid resource policy")
}

fn document_bytes() -> Vec<u8> {
    let package = Package::new().expect("new DOCX package");
    let mut output = Cursor::new(Vec::new());
    package
        .write_to(&mut output)
        .expect("serialize DOCX package");
    output.into_inner()
}

fn assert_input_limit(error: Error, actual: usize, maximum: usize) {
    assert!(matches!(
        error,
        Error::Opc(OpcError::Limit {
            kind: SourceLimitKind::InputBytes,
            observed,
            maximum: bound,
        }) if observed == actual as u64 && bound == maximum as u64
    ));
}

#[test]
fn direct_docx_constructors_apply_exact_and_over_input_limits() {
    let bytes = document_bytes();
    let exact = input_limit(bytes.len());
    let over = input_limit(bytes.len() - 1);

    Package::from_reader(Cursor::new(bytes.clone())).expect("default reader accepts valid DOCX");
    Package::from_reader_with_options(
        Cursor::new(bytes.clone()),
        litchi_docx::OpenOptions::default().source_limits(exact),
    )
    .expect("exact reader limit accepts DOCX");
    assert_input_limit(
        Package::from_reader_with_options(
            Cursor::new(bytes.clone()),
            litchi_docx::OpenOptions::default().source_limits(over),
        )
        .err()
        .expect("reader must propagate OPC input limit"),
        bytes.len(),
        bytes.len() - 1,
    );

    let file = NamedTempFile::new().expect("temporary DOCX path");
    std::fs::write(file.path(), &bytes).expect("write DOCX fixture");
    Package::open(file.path()).expect("default path accepts valid DOCX");
    Package::open_with_options(
        file.path(),
        litchi_docx::OpenOptions::default().source_limits(exact),
    )
    .expect("exact path limit accepts DOCX");
    assert_input_limit(
        Package::open_with_options(
            file.path(),
            litchi_docx::OpenOptions::default().source_limits(over),
        )
        .err()
        .expect("path must propagate OPC input limit"),
        bytes.len(),
        bytes.len() - 1,
    );
}

#[cfg(feature = "encryption")]
#[test]
fn decrypted_docx_uses_the_supplied_opc_limits() {
    use litchi_docx::encryption::{Limits as EncryptionLimits, Mode};

    let bytes = document_bytes();
    let package = Package::from_reader(Cursor::new(bytes.clone())).expect("valid DOCX");
    let encrypted = package
        .to_encrypted("read-limits", Mode::Standard)
        .expect("encrypt DOCX fixture");
    let encryption_limits = EncryptionLimits {
        max_input_bytes: encrypted.len(),
        ..EncryptionLimits::default()
    };
    let exact = input_limit(bytes.len());
    let over = input_limit(bytes.len() - 1);

    Package::from_reader_with_password_and_options(
        Cursor::new(encrypted.clone()),
        "read-limits",
        litchi_docx::OpenOptions::default()
            .source_limits(exact)
            .encryption_limits(encryption_limits),
    )
    .expect("exact post-decryption OPC limit accepts DOCX");
    assert_input_limit(
        Package::from_reader_with_password_and_options(
            Cursor::new(encrypted),
            "read-limits",
            litchi_docx::OpenOptions::default()
                .source_limits(over)
                .encryption_limits(encryption_limits),
        )
        .err()
        .expect("decrypted OPC parsing must propagate caller limit"),
        bytes.len(),
        bytes.len() - 1,
    );
}

#[test]
fn byte_slice_ingress_applies_limits_before_adopting_the_archive() {
    use litchi_docx::OpenOptions;

    let bytes = document_bytes();
    Package::from_bytes(&bytes).expect("default byte-slice ingress");
    Package::from_bytes_with_options(
        &bytes,
        OpenOptions::default().source_limits(input_limit(bytes.len())),
    )
    .expect("exact byte-slice ceiling");
    assert_input_limit(
        Package::from_bytes_with_options(
            &bytes,
            OpenOptions::default().source_limits(input_limit(bytes.len() - 1)),
        )
        .err()
        .expect("one-over byte-slice input"),
        bytes.len(),
        bytes.len() - 1,
    );
    let members = SourceLimits::builder()
        .max_members(1)
        .unwrap()
        .max_parts(1)
        .unwrap()
        .max_relationship_parts(1)
        .unwrap()
        .build()
        .unwrap();
    assert!(matches!(
        Package::from_bytes_with_options(&bytes, OpenOptions::default().source_limits(members)),
        Err(Error::Opc(OpcError::Limit {
            kind: SourceLimitKind::Members,
            maximum: 1,
            ..
        }))
    ));
}

#[cfg(feature = "encryption")]
#[test]
fn encrypted_options_preserve_independent_bounds_and_source_profile_on_every_route() {
    use litchi_docx::OpenOptions;
    use litchi_docx::encryption::{self, Error as CryptoError, Limits, Mode};

    let clear = document_bytes();
    let encrypted = Package::from_bytes(&clear)
        .unwrap()
        .to_encrypted("options", Mode::Standard)
        .unwrap();
    let file = NamedTempFile::new().unwrap();
    std::fs::write(file.path(), &encrypted).unwrap();
    let options = OpenOptions::default()
        .source_limits(input_limit(clear.len()))
        .encryption_limits(Limits {
            max_input_bytes: encrypted.len(),
            ..Limits::default()
        });
    for route in 0..3 {
        let open = |options: OpenOptions| match route {
            0 => Package::open_with_password_and_options(file.path(), "options", options.clone()),
            1 => Package::from_reader_with_password_and_options(
                Cursor::new(&encrypted),
                "options",
                options,
            ),
            _ => Package::from_bytes_with_password_and_options(
                &encrypted,
                "options",
                options.clone(),
            ),
        };
        let package = open(options.clone()).expect("both exact bounds admit source");
        assert_eq!(package.encryption(), Some(Mode::Standard));
        assert!(
            package.write_to(&mut Cursor::new(Vec::new())).is_err(),
            "plain output must remain explicit"
        );
        assert_input_limit(
            open(options.clone().source_limits(input_limit(clear.len() - 1)))
                .err()
                .unwrap(),
            clear.len(),
            clear.len() - 1,
        );
        assert!(matches!(
            open(options.clone().encryption_limits(Limits { max_input_bytes: encrypted.len() - 1, ..Limits::default() })),
            Err(Error::Encryption(CryptoError::Limit { resource: "input", actual, maximum }))
                if actual == encrypted.len() as u64 && maximum == (encrypted.len() - 1) as u64
        ));
    }
    let opened = encryption::open(encrypted.clone(), "options", &Limits::default()).unwrap();
    let package = Package::from_decrypted(opened, options.clone()).unwrap();
    assert_eq!(package.encryption(), Some(Mode::Standard));
    let opened = encryption::open(encrypted.clone(), "options", &Limits::default()).unwrap();
    assert_input_limit(
        Package::from_decrypted(opened, options.source_limits(input_limit(clear.len() - 1)))
            .err()
            .unwrap(),
        clear.len(),
        clear.len() - 1,
    );
    assert_eq!(
        Package::from_bytes_with_password(&encrypted, "options")
            .unwrap()
            .encryption(),
        Some(Mode::Standard)
    );
}

#[cfg(feature = "encryption")]
#[test]
fn encryption_policy_validation_remains_at_password_ingress() {
    use litchi_docx::{OpenOptions, encryption};
    let options = OpenOptions::default().encryption_limits(encryption::Limits {
        max_input_bytes: 0,
        ..encryption::Limits::default()
    });
    assert!(matches!(
        Package::from_bytes_with_password_and_options(&[], "options", options.clone()),
        Err(Error::Encryption(encryption::Error::InvalidLimit(_)))
    ));
    // A plaintext open does not perform an encryption operation.
    Package::from_bytes_with_options(&document_bytes(), options).unwrap();
}
