//! Bounded XMLDSig processing with caller-resolved reference bytes.

use crate::{
    Cert, Coverage, Error, Limits, Policy, Reference, Report, Result, Signer, Status, Weak,
};
use base64::{Engine as _, engine::general_purpose::STANDARD as BASE64};
use p256::ecdsa::{Signature as EcSignature, VerifyingKey as EcVerifyingKey};
use p256::pkcs8::DecodePublicKey as _;
use quick_xml::{
    Reader,
    events::{BytesStart, Event},
};
use rsa::pkcs1v15::{Signature as RsaSignature, VerifyingKey as RsaVerifyingKey};
use rsa::traits::PublicKeyParts as _;
use rsa::{BigUint, RsaPublicKey};
use sha1_legacy::Sha1;
use sha2_legacy::{Digest as _, Sha256, Sha384, Sha512};
use signature::Verifier as _;
use std::borrow::Cow;
use std::collections::{BTreeMap, HashMap, HashSet};
use std::str;
use subtle::ConstantTimeEq as _;
use x509_cert::Certificate as X509Certificate;
use x509_cert::der::{Decode as _, Encode as _};

const DS: &str = "http://www.w3.org/2000/09/xmldsig#";
const DSIG11: &str = "http://www.w3.org/2009/xmldsig11#";
const MDSSI: &str = "http://schemas.openxmlformats.org/package/2006/digital-signature";
const OFFICE: &str = "http://schemas.microsoft.com/office/2006/digsig";
const REL_TRANSFORM: &str = "http://schemas.openxmlformats.org/package/2006/RelationshipTransform";
const P256_CURVE: &str = "urn:oid:1.2.840.10045.3.1.7";
const XML_NS: &str = "http://www.w3.org/XML/1998/namespace";
const XADES: &str = "http://uri.etsi.org/01903/v1.3.2#";
const XADES_SIGNED_PROPERTIES: &str = "http://uri.etsi.org/01903#SignedProperties";
const DS_OBJECT: &str = "http://www.w3.org/2000/09/xmldsig#Object";
const PACKAGE_ID: &str = "idPackageObject";
const OFFICE_ID: &str = "idOfficeObject";
const SIGNED_PROPERTIES_ID: &str = "idSignedProperties";
const SIGNATURE_ID: &str = "idPackageSignature";

const C14N: &str = "http://www.w3.org/TR/2001/REC-xml-c14n-20010315";
const C14N_COMMENTS: &str = "http://www.w3.org/TR/2001/REC-xml-c14n-20010315#WithComments";
const EXCLUSIVE_C14N: &str = "http://www.w3.org/2001/10/xml-exc-c14n#";
const EXCLUSIVE_C14N_COMMENTS: &str = "http://www.w3.org/2001/10/xml-exc-c14n#WithComments";
const SHA1: &str = "http://www.w3.org/2000/09/xmldsig#sha1";
const SHA256: &str = "http://www.w3.org/2001/04/xmlenc#sha256";
const SHA384: &str = "http://www.w3.org/2001/04/xmldsig-more#sha384";
const SHA512: &str = "http://www.w3.org/2001/04/xmlenc#sha512";
const RSA_SHA1: &str = "http://www.w3.org/2000/09/xmldsig#rsa-sha1";
const RSA_SHA256: &str = "http://www.w3.org/2001/04/xmldsig-more#rsa-sha256";
const RSA_SHA384: &str = "http://www.w3.org/2001/04/xmldsig-more#rsa-sha384";
const RSA_SHA512: &str = "http://www.w3.org/2001/04/xmldsig-more#rsa-sha512";
const ECDSA_SHA256: &str = "http://www.w3.org/2001/04/xmldsig-more#ecdsa-sha256";

/// Canonical XML mode, always explicit about comment handling.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Canon {
    Inclusive,
    InclusiveComments,
    Exclusive,
    ExclusiveComments,
}

impl Canon {
    pub fn uri(self) -> &'static str {
        match self {
            Self::Inclusive => C14N,
            Self::InclusiveComments => C14N_COMMENTS,
            Self::Exclusive => EXCLUSIVE_C14N,
            Self::ExclusiveComments => EXCLUSIVE_C14N_COMMENTS,
        }
    }

    fn parse(value: &str) -> Result<Self> {
        match value {
            C14N => Ok(Self::Inclusive),
            C14N_COMMENTS => Ok(Self::InclusiveComments),
            EXCLUSIVE_C14N => Ok(Self::Exclusive),
            EXCLUSIVE_C14N_COMMENTS => Ok(Self::ExclusiveComments),
            _ => Err(Error::Unsupported(value.into())),
        }
    }

    fn comments(self) -> bool {
        matches!(self, Self::InclusiveComments | Self::ExclusiveComments)
    }

    fn exclusive(self) -> bool {
        matches!(self, Self::Exclusive | Self::ExclusiveComments)
    }
}

/// Digest method accepted by the verifier.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Hash {
    Sha1,
    Sha256,
    Sha384,
    Sha512,
}

impl Hash {
    pub fn uri(self) -> &'static str {
        match self {
            Self::Sha1 => SHA1,
            Self::Sha256 => SHA256,
            Self::Sha384 => SHA384,
            Self::Sha512 => SHA512,
        }
    }

    fn parse(value: &str) -> Result<Self> {
        match value {
            SHA1 => Ok(Self::Sha1),
            SHA256 => Ok(Self::Sha256),
            SHA384 => Ok(Self::Sha384),
            SHA512 => Ok(Self::Sha512),
            _ => Err(Error::Unsupported(value.into())),
        }
    }

    fn digest(self, data: &[u8]) -> Vec<u8> {
        match self {
            Self::Sha1 => Sha1::digest(data).to_vec(),
            Self::Sha256 => Sha256::digest(data).to_vec(),
            Self::Sha384 => Sha384::digest(data).to_vec(),
            Self::Sha512 => Sha512::digest(data).to_vec(),
        }
    }
}

/// XMLDSig signature method.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Method {
    RsaSha1,
    RsaSha256,
    RsaSha384,
    RsaSha512,
    EcdsaP256Sha256,
}

impl Method {
    pub fn uri(self) -> &'static str {
        match self {
            Self::RsaSha1 => RSA_SHA1,
            Self::RsaSha256 => RSA_SHA256,
            Self::RsaSha384 => RSA_SHA384,
            Self::RsaSha512 => RSA_SHA512,
            Self::EcdsaP256Sha256 => ECDSA_SHA256,
        }
    }

    fn parse(value: &str) -> Result<Self> {
        match value {
            RSA_SHA1 => Ok(Self::RsaSha1),
            RSA_SHA256 => Ok(Self::RsaSha256),
            RSA_SHA384 => Ok(Self::RsaSha384),
            RSA_SHA512 => Ok(Self::RsaSha512),
            ECDSA_SHA256 => Ok(Self::EcdsaP256Sha256),
            _ => Err(Error::Unsupported(value.into())),
        }
    }

    fn uses_sha1(self) -> bool {
        self == Self::RsaSha1
    }
}

/// Office XMLDSig object profile.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Profile {
    /// OPC-like package signatures with one signed package object.
    Package,
    /// Binary Office signatures with package and Office metadata objects.
    Binary,
}

/// A transform whose output bytes are supplied by the container resolver.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Transform {
    Canon(Canon),
    Relationships(Vec<String>),
}

/// Caller-owned or caller-borrowed bytes covered by a Manifest reference.
///
/// Sorting a `Ref` collection moves descriptors only. Borrowed payloads are
/// never copied; an owned payload is needed only for a format-specific filter
/// that cannot be represented as a borrowed subslice.
#[derive(Debug)]
pub struct Ref<'a> {
    uri: Cow<'a, str>,
    data: Cow<'a, [u8]>,
    transforms: Vec<Transform>,
}

impl<'a> Ref<'a> {
    pub fn new(uri: &'a str, data: &'a [u8]) -> Result<Self> {
        validate_uri(uri)?;
        Ok(Self {
            uri: Cow::Borrowed(uri),
            data: Cow::Borrowed(data),
            transforms: Vec::new(),
        })
    }

    pub fn owned(uri: String, data: Vec<u8>) -> Result<Self> {
        validate_uri(&uri)?;
        Ok(Self {
            uri: Cow::Owned(uri),
            data: Cow::Owned(data),
            transforms: Vec::new(),
        })
    }

    pub fn borrowed_uri(uri: String, data: &'a [u8]) -> Result<Self> {
        validate_uri(&uri)?;
        Ok(Self {
            uri: Cow::Owned(uri),
            data: Cow::Borrowed(data),
            transforms: Vec::new(),
        })
    }

    pub fn transform(mut self, value: Transform) -> Self {
        self.transforms.push(value);
        self
    }

    pub fn uri(&self) -> &str {
        &self.uri
    }

    pub fn data(&self) -> &[u8] {
        &self.data
    }

    pub fn transforms(&self) -> &[Transform] {
        &self.transforms
    }
}

fn validate_uri(uri: &str) -> Result<()> {
    if uri.is_empty() || uri.starts_with('#') || uri.chars().any(is_forbidden_xml_char) {
        return Err(Error::Container(format!(
            "invalid external signature reference URI {uri:?}"
        )));
    }
    Ok(())
}

fn is_forbidden_xml_char(value: char) -> bool {
    matches!(value, '\u{0}'..='\u{8}' | '\u{B}' | '\u{C}' | '\u{E}'..='\u{1F}')
}

/// Resolve the exact external reference set required by a container.
///
/// Implementations may return borrowed bytes or a transformed owned buffer.
/// `expected`, `has`, and `get` jointly define exact Manifest coverage.
pub trait Resolver {
    fn expected(&self) -> usize;
    fn has(&self, uri: &str) -> bool;
    fn get<'a>(&'a self, uri: &str, transforms: &[Transform]) -> Result<(Cow<'a, [u8]>, Coverage)>;
}

impl Resolver for [Ref<'_>] {
    fn expected(&self) -> usize {
        self.len()
    }

    fn has(&self, uri: &str) -> bool {
        self.iter().any(|reference| reference.uri() == uri)
    }

    fn get<'a>(&'a self, uri: &str, transforms: &[Transform]) -> Result<(Cow<'a, [u8]>, Coverage)> {
        let mut found = self.iter().filter(|reference| reference.uri() == uri);
        let reference = found
            .next()
            .ok_or_else(|| Error::Container(format!("Manifest references unexpected URI {uri}")))?;
        if found.next().is_some() {
            return Err(Error::Container(format!(
                "resolver contains duplicate URI {uri}"
            )));
        }
        if reference.transforms() != transforms {
            return Err(Error::Container(format!(
                "transform chain differs for reference {uri}"
            )));
        }
        Ok((Cow::Borrowed(reference.data()), Coverage::Complete))
    }
}

struct XmlBuf {
    value: String,
    max: usize,
}

struct ByteBuf {
    value: Vec<u8>,
    max: usize,
}

impl ByteBuf {
    fn new(max: usize) -> Self {
        Self {
            value: Vec::new(),
            max,
        }
    }

    fn push(&mut self, value: &[u8]) -> Result<()> {
        let next = self
            .value
            .len()
            .checked_add(value.len())
            .ok_or_else(|| Error::Limit("canonical XML length overflow".into()))?;
        if next > self.max {
            return Err(Error::Limit(
                "canonical XML exceeds signature policy".into(),
            ));
        }
        self.value
            .try_reserve(value.len())
            .map_err(|_| Error::Limit("canonical XML allocation failed".into()))?;
        self.value.extend_from_slice(value);
        Ok(())
    }

    fn byte(&mut self, value: u8) -> Result<()> {
        self.push(&[value])
    }

    fn finish(self) -> Vec<u8> {
        self.value
    }
}

impl XmlBuf {
    fn new(max: usize) -> Self {
        Self {
            value: String::new(),
            max,
        }
    }

    fn push(&mut self, value: &str) -> Result<()> {
        let next = self
            .value
            .len()
            .checked_add(value.len())
            .ok_or_else(|| Error::Limit("authored XML length overflow".into()))?;
        if next > self.max {
            return Err(Error::Limit("authored XML exceeds signature policy".into()));
        }
        self.value
            .try_reserve(value.len())
            .map_err(|_| Error::Limit("authored XML allocation failed".into()))?;
        self.value.push_str(value);
        Ok(())
    }

    fn attr(&mut self, value: &str) -> Result<()> {
        for character in value.chars() {
            match character {
                '&' => self.push("&amp;")?,
                '<' => self.push("&lt;")?,
                '"' => self.push("&quot;")?,
                '\t' => self.push("&#x9;")?,
                '\n' => self.push("&#xA;")?,
                '\r' => self.push("&#xD;")?,
                value => {
                    let mut buffer = [0_u8; 4];
                    self.push(value.encode_utf8(&mut buffer))?;
                },
            }
        }
        Ok(())
    }

    fn text(&mut self, value: &str) -> Result<()> {
        let mut previous = ['\0'; 2];
        for character in value.chars() {
            match character {
                '&' => self.push("&amp;")?,
                '<' => self.push("&lt;")?,
                '>' if previous == [']', ']'] => self.push("&gt;")?,
                '\r' => self.push("&#xD;")?,
                value => {
                    let mut buffer = [0_u8; 4];
                    self.push(value.encode_utf8(&mut buffer))?;
                },
            }
            previous = [previous[1], character];
        }
        Ok(())
    }

    fn as_str(&self) -> &str {
        &self.value
    }

    fn into_string(self) -> String {
        self.value
    }
}

/// Author an Office-profile XMLDSig document.
pub fn author(
    profile: Profile,
    signer: &Signer,
    references: &[Ref<'_>],
    limits: &Limits,
) -> Result<Vec<u8>> {
    if references.is_empty() {
        return Err(Error::Sign(
            "a signature must cover at least one reference".into(),
        ));
    }
    if references.len() > limits.max_references() {
        return Err(Error::Limit("too many signature references".into()));
    }
    let certificate_xml_bytes = author_key_limits(signer, limits)?;
    let signing_time = signer
        .signing_time()
        .ok_or_else(|| Error::Sign("an explicit RFC 3339 signing time is required".into()))?;
    let mut ordered = Vec::new();
    ordered
        .try_reserve(references.len())
        .map_err(|_| Error::Limit("reference ordering allocation failed".into()))?;
    ordered.extend(references.iter());
    ordered.sort_by(|left, right| left.uri().cmp(right.uri()));
    for pair in ordered.windows(2) {
        if pair[0].uri() == pair[1].uri() {
            return Err(Error::Sign(format!(
                "duplicate reference URI {}",
                pair[0].uri()
            )));
        }
    }

    let mut manifest = XmlBuf::new(limits.max_signature_bytes());
    for reference in ordered {
        manifest.push("<Reference URI=\"")?;
        manifest.attr(reference.uri())?;
        manifest.push("\">")?;
        render_transforms(&mut manifest, reference.transforms())?;
        manifest.push("<DigestMethod Algorithm=\"")?;
        manifest.push(SHA256)?;
        manifest.push("\"></DigestMethod><DigestValue>")?;
        manifest.push(&BASE64.encode(Sha256::digest(reference.data())))?;
        manifest.push("</DigestValue></Reference>")?;
    }
    let mut package_object = XmlBuf::new(limits.max_signature_bytes());
    package_object.push("<Object xmlns=\"")?;
    package_object.push(DS)?;
    package_object.push("\" xmlns:mdssi=\"")?;
    package_object.push(MDSSI)?;
    package_object.push("\" Id=\"")?;
    package_object.push(PACKAGE_ID)?;
    package_object.push("\"><Manifest>")?;
    package_object.push(manifest.as_str())?;
    package_object.push(
        "</Manifest><SignatureProperties><SignatureProperty Id=\"idSignatureTime\" Target=\"#",
    )?;
    package_object.push(SIGNATURE_ID)?;
    package_object.push(
        "\"><mdssi:SignatureTime><mdssi:Format>YYYY-MM-DDThh:mm:ssTZD</mdssi:Format><mdssi:Value>",
    )?;
    package_object.text(signing_time)?;
    package_object.push(
        "</mdssi:Value></mdssi:SignatureTime></SignatureProperty></SignatureProperties></Object>",
    )?;
    let package_object = package_object.into_string();
    let office_object = match profile {
        Profile::Package => None,
        Profile::Binary => {
            let mut office = XmlBuf::new(limits.max_signature_bytes());
            office.push("<Object xmlns=\"")?;
            office.push(DS)?;
            office.push("\" Id=\"")?;
            office.push(OFFICE_ID)?;
            office.push("\"><SignatureProperties><SignatureProperty Id=\"idOfficeSignatureInfo\" Target=\"#")?;
            office.push(SIGNATURE_ID)?;
            office.push("\"><SignatureInfoV1 xmlns=\"")?;
            office.push(OFFICE)?;
            office.push("\"><SetupID></SetupID><SignatureText></SignatureText><SignatureImage></SignatureImage><SignatureComments></SignatureComments><WindowsVersion>1</WindowsVersion><OfficeVersion>12.0</OfficeVersion><ApplicationVersion>12.0</ApplicationVersion><Monitors>1</Monitors><HorizontalResolution>96</HorizontalResolution><VerticalResolution>96</VerticalResolution><ColorDepth>24</ColorDepth><SignatureProviderId></SignatureProviderId><SignatureProviderUrl></SignatureProviderUrl><SignatureProviderDetails>0</SignatureProviderDetails><SignatureType>1</SignatureType><ManifestHashAlgorithm>")?;
            office.push(SHA256)?;
            office.push("</ManifestHashAlgorithm></SignatureInfoV1></SignatureProperty></SignatureProperties></Object>")?;
            Some(office.into_string())
        },
    };
    let canon = signer.canonicalization();
    let package_digest = Sha256::digest(canonicalize_authored(&package_object, canon, limits)?);
    let mut signed_info = XmlBuf::new(limits.max_signature_bytes());
    signed_info.push("<SignedInfo xmlns=\"")?;
    signed_info.push(DS)?;
    signed_info.push("\"><CanonicalizationMethod Algorithm=\"")?;
    signed_info.push(canon.uri())?;
    signed_info.push("\"></CanonicalizationMethod><SignatureMethod Algorithm=\"")?;
    signed_info.push(signer.method().uri())?;
    signed_info.push("\"></SignatureMethod>")?;
    fragment_reference(&mut signed_info, PACKAGE_ID, canon, &package_digest)?;
    if let Some(office) = &office_object {
        let digest = Sha256::digest(canonicalize_authored(office, canon, limits)?);
        fragment_reference(&mut signed_info, OFFICE_ID, canon, &digest)?;
    }
    signed_info.push("</SignedInfo>")?;
    let signed_info = signed_info.into_string();
    let signed_bytes = canonicalize_authored(&signed_info, canon, limits)?;
    let value = signer.sign(&signed_bytes);
    let key_info = build_key_info(signer, certificate_xml_bytes, limits)?;
    let mut result = XmlBuf::new(limits.max_signature_bytes());
    result.push("<?xml version=\"1.0\" encoding=\"UTF-8\"?><Signature xmlns=\"")?;
    result.push(DS)?;
    result.push("\" Id=\"")?;
    result.push(SIGNATURE_ID)?;
    result.push("\">")?;
    result.push(&signed_info)?;
    result.push("<SignatureValue>")?;
    result.push(&BASE64.encode(value))?;
    result.push("</SignatureValue>")?;
    result.push(&key_info)?;
    result.push(&package_object)?;
    if let Some(office) = &office_object {
        result.push(office)?;
    }
    result.push("</Signature>")?;
    Ok(result.into_string().into_bytes())
}

fn render_transforms(output: &mut XmlBuf, transforms: &[Transform]) -> Result<()> {
    if transforms.is_empty() {
        return Ok(());
    }
    output.push("<Transforms>")?;
    for transform in transforms {
        match transform {
            Transform::Canon(canon) => {
                output.push("<Transform Algorithm=\"")?;
                output.push(canon.uri())?;
                output.push("\"></Transform>")?;
            },
            Transform::Relationships(ids) => {
                output.push("<Transform Algorithm=\"")?;
                output.push(REL_TRANSFORM)?;
                output.push("\" xmlns:mdssi=\"")?;
                output.push(MDSSI)?;
                output.push("\">")?;
                for id in ids {
                    output.push("<mdssi:RelationshipReference SourceId=\"")?;
                    output.attr(id)?;
                    output.push("\"></mdssi:RelationshipReference>")?;
                }
                output.push("</Transform>")?;
            },
        }
    }
    output.push("</Transforms>")?;
    Ok(())
}

fn fragment_reference(output: &mut XmlBuf, id: &str, canon: Canon, digest: &[u8]) -> Result<()> {
    output.push("<Reference URI=\"#")?;
    output.attr(id)?;
    output.push("\"><Transforms><Transform Algorithm=\"")?;
    output.push(canon.uri())?;
    output.push("\"></Transform></Transforms><DigestMethod Algorithm=\"")?;
    output.push(SHA256)?;
    output.push("\"></DigestMethod><DigestValue>")?;
    output.push(&BASE64.encode(digest))?;
    output.push("</DigestValue></Reference>")
}

fn author_key_limits(signer: &Signer, limits: &Limits) -> Result<usize> {
    if signer.certificates().len() > limits.max_certificates() {
        return Err(Error::Limit("too many signing certificates".into()));
    }
    if signer
        .rsa_public()
        .is_some_and(|key| key.n().bits() > limits.max_rsa_bits())
    {
        return Err(Error::Limit("RSA signing modulus exceeds policy".into()));
    }
    let mut total_der = 0_usize;
    let mut xml_bytes = 0_usize;
    for certificate in signer.certificates() {
        if certificate.len() > limits.max_certificate_bytes() {
            return Err(Error::Limit("signing certificate is too large".into()));
        }
        total_der = total_der
            .checked_add(certificate.len())
            .ok_or_else(|| Error::Limit("certificate byte count overflow".into()))?;
        if total_der > limits.max_total_certificate_bytes() {
            return Err(Error::Limit("signing certificates are too large".into()));
        }
        let encoded = certificate
            .len()
            .checked_add(2)
            .map(|length| length / 3)
            .and_then(|length| length.checked_mul(4))
            .ok_or_else(|| Error::Limit("certificate base64 length overflow".into()))?;
        xml_bytes = xml_bytes
            .checked_add("<X509Certificate></X509Certificate>".len())
            .and_then(|length| length.checked_add(encoded))
            .ok_or_else(|| Error::Limit("certificate XML length overflow".into()))?;
        if xml_bytes > limits.max_signature_bytes() {
            return Err(Error::Limit(
                "encoded signing certificates exceed signature policy".into(),
            ));
        }
    }
    Ok(xml_bytes)
}

fn build_key_info(
    signer: &Signer,
    certificate_xml_bytes: usize,
    limits: &Limits,
) -> Result<String> {
    let key_value = if let Some(key) = signer.rsa_public() {
        format!(
            "<KeyValue><RSAKeyValue><Modulus>{}</Modulus><Exponent>{}</Exponent></RSAKeyValue></KeyValue>",
            BASE64.encode(key.n().to_bytes_be()),
            BASE64.encode(key.e().to_bytes_be())
        )
    } else if let Some(key) = signer.p256_public() {
        format!(
            "<KeyValue><dsig11:ECKeyValue xmlns:dsig11=\"{DSIG11}\"><dsig11:NamedCurve URI=\"{P256_CURVE}\"></dsig11:NamedCurve><dsig11:PublicKey>{}</dsig11:PublicKey></dsig11:ECKeyValue></KeyValue>",
            BASE64.encode(key.to_encoded_point(false).as_bytes())
        )
    } else {
        String::new()
    };
    if signer.certificates().is_empty() {
        let result = format!("<KeyInfo>{key_value}</KeyInfo>");
        if result.len() > limits.max_signature_bytes() {
            return Err(Error::Limit("KeyInfo exceeds signature policy".into()));
        }
        return Ok(result);
    }
    let capacity = key_value
        .len()
        .checked_add(certificate_xml_bytes)
        .and_then(|length| length.checked_add("<KeyInfo><X509Data></X509Data></KeyInfo>".len()))
        .ok_or_else(|| Error::Limit("KeyInfo length overflow".into()))?;
    if capacity > limits.max_signature_bytes() {
        return Err(Error::Limit("KeyInfo exceeds signature policy".into()));
    }
    let mut result = String::new();
    result
        .try_reserve(capacity)
        .map_err(|_| Error::Limit("KeyInfo allocation failed".into()))?;
    result.push_str("<KeyInfo>");
    result.push_str(&key_value);
    result.push_str("<X509Data>");
    for certificate in signer.certificates() {
        result.push_str("<X509Certificate>");
        BASE64.encode_string(certificate, &mut result);
        result.push_str("</X509Certificate>");
    }
    result.push_str("</X509Data></KeyInfo>");
    Ok(result)
}

/// Canonicalize exactly one safe XML root using the same bounded parser as the
/// signature verifier.
///
/// DTDs, processing instructions, entity references, duplicate IDs, multiple
/// roots, and input or output exceeding `limits` are rejected.
pub fn canonicalize(xml: &[u8], canon: Canon, limits: &Limits) -> Result<Vec<u8>> {
    let document = Document::parse(xml, limits)?;
    document.canonicalize(document.root, canon, limits)
}

/// Verify one Office-profile XMLDSig document against an exact reference set.
pub fn verify(
    profile: Profile,
    xml: &[u8],
    resolver: &(impl Resolver + ?Sized),
    policy: &Policy,
) -> Result<Report> {
    verify_with_certs(profile, xml, resolver, &[], policy)
}

/// Verify using caller-borrowed external DER certificates in addition to any
/// certificates embedded in the signature.
///
/// External evidence is useful for OPC certificate relationship parts. The
/// merged set is DER-validated and shares one deduplicated count, per-item byte
/// limit, and aggregate byte limit with embedded evidence.
pub fn verify_with_certs(
    profile: Profile,
    xml: &[u8],
    resolver: &(impl Resolver + ?Sized),
    certificates: &[&[u8]],
    policy: &Policy,
) -> Result<Report> {
    if xml.len() > policy.limits().max_signature_bytes() {
        return Err(Error::Limit("signature XML is too large".into()));
    }
    let document = Document::parse(xml, policy.limits())?;
    if !document.is(document.root, DS, "Signature") {
        return Err(Error::Xml("root element must be ds:Signature".into()));
    }
    let signed_info = document.required_child(document.root, DS, "SignedInfo")?;
    let canonicalization = document.required_child(signed_info, DS, "CanonicalizationMethod")?;
    let canon = Canon::parse(document.attr(canonicalization, "Algorithm")?)?;
    let signed_bytes = document.canonicalize(signed_info, canon, policy.limits())?;
    let signature_method = document.required_child(signed_info, DS, "SignatureMethod")?;
    let method = Method::parse(document.attr(signature_method, "Algorithm")?)?;
    enforce_weak(method.uses_sha1(), policy)?;

    let package_object = document.required_bound_object(PACKAGE_ID)?;
    let manifest = document.required_child(package_object, DS, "Manifest")?;
    let manifests = document.descendants(document.root, DS, "Manifest");
    if manifests.as_slice() != [manifest] {
        return Err(Error::Xml(
            "the sole Manifest must be a direct child of the signed package Object".into(),
        ));
    }
    let signed_references = document.children(signed_info, DS, "Reference");
    let maximum = match profile {
        Profile::Package => 3,
        Profile::Binary => 2,
    };
    if signed_references.is_empty() || signed_references.len() > maximum {
        return Err(Error::Xml(format!(
            "SignedInfo contains an invalid number of profile references: {}",
            signed_references.len()
        )));
    }
    let mut reports = Vec::new();
    let mut uses_sha1 = method.uses_sha1();
    let mut signed_ids = HashSet::new();
    for reference in signed_references {
        let uri = document.attr(reference, "URI")?;
        let id = uri
            .strip_prefix('#')
            .ok_or_else(|| Error::Xml(format!("unexpected SignedInfo reference {uri}")))?;
        if !signed_ids.insert(id.to_string()) {
            return Err(Error::Xml(format!("duplicate SignedInfo reference {uri}")));
        }
        let node = match (profile, id) {
            (_, PACKAGE_ID) => {
                require_optional_type(&document, reference, DS_OBJECT, "package Object")?;
                package_object
            },
            (_, OFFICE_ID) => {
                require_optional_type(&document, reference, DS_OBJECT, "Office Object")?;
                document.required_office_object()?
            },
            (Profile::Package, SIGNED_PROPERTIES_ID) => {
                require_type(
                    &document,
                    reference,
                    XADES_SIGNED_PROPERTIES,
                    "XAdES SignedProperties",
                )?;
                document.required_signed_properties()?
            },
            _ => return Err(Error::Xml(format!("unexpected SignedInfo reference {uri}"))),
        };
        let data = dereference_fragment(&document, reference, node, policy.limits())?;
        let (report, weak) =
            verify_digest(&document, reference, &data, Coverage::Complete, policy)?;
        reports.push(report);
        uses_sha1 |= weak;
    }
    if !signed_ids.contains(PACKAGE_ID)
        || profile == Profile::Binary && !signed_ids.contains(OFFICE_ID)
    {
        return Err(Error::Xml(
            "required signed object reference is missing".into(),
        ));
    }
    if profile == Profile::Package {
        for optional in [OFFICE_ID, SIGNED_PROPERTIES_ID] {
            if document.ids.contains_key(optional) && !signed_ids.contains(optional) {
                return Err(Error::Xml(format!(
                    "#{optional} is present but not bound by SignedInfo"
                )));
            }
        }
        require_bound_package_metadata(&document, &signed_ids)?;
    }

    let manifest_references = document.children(manifest, DS, "Reference");
    if manifest_references.len() > resolver.expected() {
        return Err(Error::Container(format!(
            "Manifest covers {} references but the container has only {} eligible references",
            manifest_references.len(),
            resolver.expected()
        )));
    }
    let total_references = manifest_references
        .len()
        .checked_add(signed_ids.len())
        .ok_or_else(|| Error::Limit("reference count overflow".into()))?;
    if total_references > policy.limits().max_references() {
        return Err(Error::Limit("too many signature references".into()));
    }
    let mut seen = HashSet::new();
    let mut coverage = if manifest_references.len() == resolver.expected() {
        Coverage::Complete
    } else {
        Coverage::Partial
    };
    for reference in manifest_references {
        let uri = document.attr(reference, "URI")?;
        if uri.starts_with('#') || !seen.insert(uri.to_string()) {
            return Err(Error::Xml(format!(
                "invalid or duplicate Manifest URI {uri}"
            )));
        }
        if !resolver.has(uri) {
            return Err(Error::Container(format!(
                "Manifest contains unexpected reference {uri}"
            )));
        }
        let transforms = parse_transforms(&document, reference)?;
        let (bytes, reference_coverage) = resolver.get(uri, &transforms)?;
        coverage = coverage.combine(reference_coverage);
        let (report, weak) =
            verify_digest(&document, reference, &bytes, reference_coverage, policy)?;
        reports.push(report);
        uses_sha1 |= weak;
    }
    if coverage == Coverage::Partial && !policy.allows_partial_coverage() {
        return Err(Error::Container(
            "strict policy rejects partial signature coverage".into(),
        ));
    }

    let certificates = extract_certificates(&document, certificates, policy.limits())?;
    let encoded = document.text(document.required_child(document.root, DS, "SignatureValue")?)?;
    let value = decode64(
        &encoded,
        policy.limits().max_signature_bytes(),
        "SignatureValue",
    )?;
    let valid_signature = verify_value(
        &document,
        method,
        &certificates,
        &signed_bytes,
        &value,
        policy.limits(),
    )?;
    let integrity = reports
        .iter()
        .all(|reference| reference.status() == Status::Valid);
    Ok(Report::new(
        status(integrity),
        status(valid_signature),
        coverage,
        reports,
        certificates,
        uses_sha1,
        extract_time(&document)?,
    ))
}

fn enforce_weak(weak: bool, policy: &Policy) -> Result<()> {
    if weak && policy.weak() == Weak::Reject {
        Err(Error::Sha1)
    } else {
        Ok(())
    }
}

fn require_bound_package_metadata(document: &Document, signed_ids: &HashSet<String>) -> Result<()> {
    let office = document.descendants(document.root, OFFICE, "SignatureInfoV1");
    if !office.is_empty() {
        if !signed_ids.contains(OFFICE_ID) || office.len() != 1 {
            return Err(Error::Xml(
                "every Office SignatureInfoV1 must be uniquely bound as #idOfficeObject".into(),
            ));
        }
        document.required_office_object()?;
    }
    let signed_properties = document.descendants(document.root, XADES, "SignedProperties");
    if !signed_properties.is_empty() {
        if !signed_ids.contains(SIGNED_PROPERTIES_ID) || signed_properties.len() != 1 {
            return Err(Error::Xml(
                "every XAdES SignedProperties must be uniquely bound as #idSignedProperties".into(),
            ));
        }
        document.required_signed_properties()?;
    }
    Ok(())
}

fn require_optional_type(
    document: &Document,
    reference: usize,
    expected: &str,
    description: &str,
) -> Result<()> {
    if let Some(actual) = document.optional_attr(reference, "Type")
        && actual != expected
    {
        return Err(Error::Xml(format!(
            "{description} reference has unexpected Type {actual}"
        )));
    }
    Ok(())
}

fn require_type(
    document: &Document,
    reference: usize,
    expected: &str,
    description: &str,
) -> Result<()> {
    match document.optional_attr(reference, "Type") {
        Some(actual) if actual == expected => Ok(()),
        Some(actual) => Err(Error::Xml(format!(
            "{description} reference has unexpected Type {actual}"
        ))),
        None => Err(Error::Xml(format!(
            "{description} reference is missing its Type"
        ))),
    }
}

fn dereference_fragment(
    document: &Document,
    reference: usize,
    node: usize,
    limits: &Limits,
) -> Result<Vec<u8>> {
    let transforms = parse_transforms(document, reference)?;
    let canon = match transforms.as_slice() {
        [] => Canon::Inclusive,
        [Transform::Canon(value)] => *value,
        _ => {
            return Err(Error::Unsupported(
                "invalid same-document transform chain".into(),
            ));
        },
    };
    document.canonicalize(node, canon, limits)
}

fn verify_digest(
    document: &Document,
    reference: usize,
    data: &[u8],
    coverage: Coverage,
    policy: &Policy,
) -> Result<(Reference, bool)> {
    let uri = document.attr(reference, "URI")?.to_string();
    let method = document.required_child(reference, DS, "DigestMethod")?;
    let hash = Hash::parse(document.attr(method, "Algorithm")?)?;
    enforce_weak(hash == Hash::Sha1, policy)?;
    let expected = decode64(
        &document.text(document.required_child(reference, DS, "DigestValue")?)?,
        128,
        "DigestValue",
    )?;
    let actual = hash.digest(data);
    let valid = actual.len() == expected.len() && bool::from(actual.ct_eq(&expected));
    Ok((
        Reference::new(uri, status(valid), coverage),
        hash == Hash::Sha1,
    ))
}

fn parse_transforms(document: &Document, reference: usize) -> Result<Vec<Transform>> {
    let Some(transforms) = document.child(reference, DS, "Transforms") else {
        return Ok(Vec::new());
    };
    let mut result = Vec::new();
    for transform in document.children(transforms, DS, "Transform") {
        let algorithm = document.attr(transform, "Algorithm")?;
        if algorithm == REL_TRANSFORM {
            if !result.is_empty() {
                return Err(Error::Unsupported(
                    "RelationshipTransform must be first".into(),
                ));
            }
            let mut ids = Vec::new();
            for reference in document.children(transform, MDSSI, "RelationshipReference") {
                ids.push(document.attr(reference, "SourceId")?.to_string());
            }
            if ids.is_empty() {
                return Err(Error::Xml("empty RelationshipTransform".into()));
            }
            ids.sort();
            if ids.windows(2).any(|pair| pair[0] == pair[1]) {
                return Err(Error::Xml(
                    "duplicate RelationshipReference SourceId".into(),
                ));
            }
            result.push(Transform::Relationships(ids));
        } else {
            result.push(Transform::Canon(Canon::parse(algorithm)?));
        }
    }
    Ok(result)
}

fn verify_value(
    document: &Document,
    method: Method,
    certificates: &[Cert],
    message: &[u8],
    value: &[u8],
    limits: &Limits,
) -> Result<bool> {
    match method {
        Method::RsaSha1 | Method::RsaSha256 | Method::RsaSha384 | Method::RsaSha512 => {
            let key = rsa_key(document, certificates, limits)?;
            let signature =
                RsaSignature::try_from(value).map_err(|error| Error::Key(error.to_string()))?;
            let valid = match method {
                Method::RsaSha1 => RsaVerifyingKey::<Sha1>::new(key).verify(message, &signature),
                Method::RsaSha256 => {
                    RsaVerifyingKey::<Sha256>::new(key).verify(message, &signature)
                },
                Method::RsaSha384 => {
                    RsaVerifyingKey::<Sha384>::new(key).verify(message, &signature)
                },
                Method::RsaSha512 => {
                    RsaVerifyingKey::<Sha512>::new(key).verify(message, &signature)
                },
                Method::EcdsaP256Sha256 => {
                    return Err(Error::Unsupported("RSA branch received ECDSA".into()));
                },
            };
            Ok(valid.is_ok())
        },
        Method::EcdsaP256Sha256 => {
            let key = ec_key(document, certificates)?;
            let signature =
                EcSignature::from_slice(value).map_err(|error| Error::Key(error.to_string()))?;
            Ok(key.verify(message, &signature).is_ok())
        },
    }
}

fn rsa_key(document: &Document, certificates: &[Cert], limits: &Limits) -> Result<RsaPublicKey> {
    let from_certificate = certificates
        .first()
        .map(cert_spki)
        .transpose()?
        .map(|spki| {
            RsaPublicKey::from_public_key_der(&spki).map_err(|error| Error::Key(error.to_string()))
        })
        .transpose()?;
    let values = document.descendants(document.root, DS, "RSAKeyValue");
    if values.len() > 1 {
        return Err(Error::Key("multiple RSAKeyValue elements".into()));
    }
    let from_xml = values
        .first()
        .map(|value| {
            let modulus = decode64(
                &document.text(document.required_child(*value, DS, "Modulus")?)?,
                limits.max_rsa_bits().div_ceil(8),
                "RSA modulus",
            )?;
            let exponent = decode64(
                &document.text(document.required_child(*value, DS, "Exponent")?)?,
                16,
                "RSA exponent",
            )?;
            RsaPublicKey::new(
                BigUint::from_bytes_be(&modulus),
                BigUint::from_bytes_be(&exponent),
            )
            .map_err(|error| Error::Key(error.to_string()))
        })
        .transpose()?;
    let key = match (from_certificate, from_xml) {
        (Some(certificate), Some(xml)) if certificate != xml => {
            return Err(Error::Key(
                "certificate and RSAKeyValue identify different keys".into(),
            ));
        },
        (Some(key), _) | (_, Some(key)) => key,
        (None, None) => return Err(Error::Key("no RSA verification key".into())),
    };
    if key.n().bits() > limits.max_rsa_bits() {
        return Err(Error::Limit("RSA modulus exceeds policy".into()));
    }
    Ok(key)
}

fn ec_key(document: &Document, certificates: &[Cert]) -> Result<EcVerifyingKey> {
    let from_certificate = certificates
        .first()
        .map(cert_spki)
        .transpose()?
        .map(|spki| {
            EcVerifyingKey::from_public_key_der(&spki)
                .map_err(|error| Error::Key(error.to_string()))
        })
        .transpose()?;
    let values = document.descendants(document.root, DSIG11, "ECKeyValue");
    if values.len() > 1 {
        return Err(Error::Key("multiple ECKeyValue elements".into()));
    }
    let from_xml = values
        .first()
        .map(|value| {
            let curve = document.required_child(*value, DSIG11, "NamedCurve")?;
            if document.attr(curve, "URI")? != P256_CURVE {
                return Err(Error::Unsupported("unsupported EC named curve".into()));
            }
            let point = decode64(
                &document.text(document.required_child(*value, DSIG11, "PublicKey")?)?,
                65,
                "P-256 public key",
            )?;
            EcVerifyingKey::from_sec1_bytes(&point).map_err(|error| Error::Key(error.to_string()))
        })
        .transpose()?;
    match (from_certificate, from_xml) {
        (Some(certificate), Some(xml)) if certificate != xml => Err(Error::Key(
            "certificate and ECKeyValue identify different keys".into(),
        )),
        (Some(key), _) | (_, Some(key)) => Ok(key),
        (None, None) => Err(Error::Key("no P-256 verification key".into())),
    }
}

fn cert_spki(certificate: &Cert) -> Result<Vec<u8>> {
    let certificate = X509Certificate::from_der(certificate.der())
        .map_err(|error| Error::Key(error.to_string()))?;
    certificate
        .tbs_certificate
        .subject_public_key_info
        .to_der()
        .map_err(|error| Error::Key(error.to_string()))
}

fn extract_certificates(
    document: &Document,
    external: &[&[u8]],
    limits: &Limits,
) -> Result<Vec<Cert>> {
    let mut result = Vec::new();
    let mut total = 0_usize;
    for node in document.descendants(document.root, DS, "X509Certificate") {
        let der = decode64(
            &document.text(node)?,
            limits.max_certificate_bytes(),
            "X509Certificate",
        )?;
        push_certificate(&mut result, &mut total, &der, limits)?;
    }
    for &der in external {
        push_certificate(&mut result, &mut total, der, limits)?;
    }
    Ok(result)
}

fn push_certificate(
    certificates: &mut Vec<Cert>,
    total: &mut usize,
    der: &[u8],
    limits: &Limits,
) -> Result<()> {
    if der.len() > limits.max_certificate_bytes() {
        return Err(Error::Limit("certificate is too large".into()));
    }
    if certificates
        .iter()
        .any(|certificate| certificate.der() == der)
    {
        return Ok(());
    }
    if certificates.len() >= limits.max_certificates() {
        return Err(Error::Limit("too many certificates".into()));
    }
    *total = total
        .checked_add(der.len())
        .ok_or_else(|| Error::Limit("certificate byte count overflow".into()))?;
    if *total > limits.max_total_certificate_bytes() {
        return Err(Error::Limit("certificates are too large".into()));
    }
    X509Certificate::from_der(der)
        .map_err(|error| Error::Key(format!("invalid X.509 certificate: {error}")))?;
    certificates.push(Cert::new(der.to_vec()));
    Ok(())
}

fn extract_time(document: &Document) -> Result<Option<String>> {
    let values = document.descendants(document.root, MDSSI, "Value");
    if values.len() > 1 {
        return Err(Error::Xml("multiple Office signing times".into()));
    }
    values.first().map(|node| document.text(*node)).transpose()
}

fn canonicalize_authored(xml: &str, canon: Canon, limits: &Limits) -> Result<Vec<u8>> {
    canonicalize(xml.as_bytes(), canon, limits)
}

fn decode64(value: &str, max: usize, description: &str) -> Result<Vec<u8>> {
    let compact: String = value
        .chars()
        .filter(|character| !character.is_whitespace())
        .collect();
    let encoded_limit = max
        .checked_mul(4)
        .and_then(|value| value.checked_div(3))
        .and_then(|value| value.checked_add(8))
        .unwrap_or(usize::MAX);
    if compact.len() > encoded_limit {
        return Err(Error::Limit(format!("{description} is too large")));
    }
    let decoded = BASE64
        .decode(compact)
        .map_err(|error| Error::Xml(format!("invalid {description} base64: {error}")))?;
    if decoded.len() > max {
        Err(Error::Limit(format!("{description} is too large")))
    } else {
        Ok(decoded)
    }
}

fn status(value: bool) -> Status {
    if value {
        Status::Valid
    } else {
        Status::Invalid
    }
}

#[derive(Debug)]
struct Name {
    qualified: String,
    local: String,
    namespace: String,
}

#[derive(Debug)]
struct Attribute {
    name: Name,
    value: String,
}

#[derive(Debug)]
enum Child {
    Element(usize),
    Text(String),
    Comment(String),
}

#[derive(Debug)]
struct Element {
    name: Name,
    attributes: Vec<Attribute>,
    namespaces: BTreeMap<String, String>,
    children: Vec<Child>,
    parent: Option<usize>,
}

#[derive(Debug)]
struct Document {
    elements: Vec<Element>,
    root: usize,
    ids: HashMap<String, usize>,
}

impl Document {
    fn parse(bytes: &[u8], limits: &Limits) -> Result<Self> {
        if bytes.len() > limits.max_signature_bytes() {
            return Err(Error::Limit("XML is too large".into()));
        }
        let mut reader = Reader::from_reader(bytes);
        reader.config_mut().trim_text(false);
        reader.config_mut().check_end_names = true;
        let mut elements = Vec::new();
        let mut stack = Vec::new();
        let mut ids = HashMap::new();
        let mut root = None;
        let mut buffer = Vec::new();
        loop {
            match reader.read_event_into(&mut buffer) {
                Ok(Event::Start(start)) => Self::start(
                    &start,
                    reader.decoder(),
                    limits,
                    &mut elements,
                    &mut stack,
                    &mut ids,
                    &mut root,
                )?,
                Ok(Event::Empty(start)) => {
                    Self::start(
                        &start,
                        reader.decoder(),
                        limits,
                        &mut elements,
                        &mut stack,
                        &mut ids,
                        &mut root,
                    )?;
                    if stack.pop().is_none() {
                        return Err(Error::Xml("empty-element stack underflow".into()));
                    }
                },
                Ok(Event::End(_)) => {
                    if stack.pop().is_none() {
                        return Err(Error::Xml("unexpected XML end tag".into()));
                    }
                },
                Ok(Event::Text(text)) => {
                    let raw = text.xml10_content().map_err(xml_error)?;
                    let value = quick_xml::escape::unescape(&raw)
                        .map_err(xml_error)?
                        .into_owned();
                    Self::push_text(&mut elements, &stack, value)?;
                },
                Ok(Event::CData(text)) => Self::push_text(
                    &mut elements,
                    &stack,
                    text.xml10_content().map_err(xml_error)?.into_owned(),
                )?,
                Ok(Event::Comment(comment)) => {
                    if let Some(&element) = stack.last() {
                        elements[element].children.push(Child::Comment(
                            comment.xml10_content().map_err(xml_error)?.into_owned(),
                        ));
                    }
                },
                Ok(Event::Decl(_)) if root.is_some() => {
                    return Err(Error::Xml("late XML declaration".into()));
                },
                Ok(Event::Decl(_)) => {},
                Ok(Event::DocType(_) | Event::PI(_) | Event::GeneralRef(_)) => {
                    return Err(Error::Xml(
                        "DTDs, processing instructions, and entity references are prohibited"
                            .into(),
                    ));
                },
                Ok(Event::Eof) => break,
                Err(error) => return Err(Error::Xml(error.to_string())),
            }
            buffer.clear();
        }
        if !stack.is_empty() {
            return Err(Error::Xml("unclosed XML element".into()));
        }
        Ok(Self {
            elements,
            root: root.ok_or_else(|| Error::Xml("signature XML has no root".into()))?,
            ids,
        })
    }

    #[allow(clippy::too_many_arguments)]
    fn start(
        start: &BytesStart<'_>,
        decoder: quick_xml::encoding::Decoder,
        limits: &Limits,
        elements: &mut Vec<Element>,
        stack: &mut Vec<usize>,
        ids: &mut HashMap<String, usize>,
        root: &mut Option<usize>,
    ) -> Result<()> {
        if stack.len() >= limits.max_xml_depth() {
            return Err(Error::Limit("XML depth exceeds policy".into()));
        }
        if elements.len() >= limits.max_xml_elements() {
            return Err(Error::Limit("XML element count exceeds policy".into()));
        }
        let qualified = str::from_utf8(start.name().as_ref())
            .map_err(xml_error)?
            .to_string();
        let mut raw_attributes = Vec::new();
        for attribute in start.attributes().with_checks(true) {
            if raw_attributes.len() >= limits.max_attributes() {
                return Err(Error::Limit("XML attribute count exceeds policy".into()));
            }
            let attribute = attribute.map_err(xml_error)?;
            raw_attributes.push((
                str::from_utf8(attribute.key.as_ref())
                    .map_err(xml_error)?
                    .to_string(),
                attribute
                    .decoded_and_normalized_value(quick_xml::XmlVersion::Explicit1_0, decoder)
                    .map_err(xml_error)?
                    .into_owned(),
            ));
        }
        let mut namespaces = stack.last().map_or_else(
            || {
                let mut namespaces = BTreeMap::new();
                namespaces.insert("xml".into(), XML_NS.into());
                namespaces
            },
            |parent| elements[*parent].namespaces.clone(),
        );
        for (name, value) in &raw_attributes {
            if name == "xmlns" {
                namespaces.insert(String::new(), value.clone());
            } else if let Some(prefix) = name.strip_prefix("xmlns:") {
                if prefix.is_empty()
                    || prefix == "xmlns"
                    || (prefix == "xml" && value != XML_NS)
                    || (prefix != "xml" && value == XML_NS)
                    || value.is_empty()
                {
                    return Err(Error::Xml(format!("invalid namespace declaration {name}")));
                }
                namespaces.insert(prefix.into(), value.clone());
            }
        }
        let name = expanded(&qualified, &namespaces, true)?;
        let mut attributes = Vec::new();
        let mut unique = HashSet::new();
        for (qualified, value) in raw_attributes {
            if qualified == "xmlns" || qualified.starts_with("xmlns:") {
                continue;
            }
            let name = expanded(&qualified, &namespaces, false)?;
            if !unique.insert((name.namespace.clone(), name.local.clone())) {
                return Err(Error::Xml(format!("duplicate attribute {qualified}")));
            }
            attributes.push(Attribute { name, value });
        }
        let index = elements.len();
        let parent = stack.last().copied();
        elements.push(Element {
            name,
            attributes,
            namespaces,
            children: Vec::new(),
            parent,
        });
        if let Some(parent) = parent {
            elements[parent].children.push(Child::Element(index));
        } else if root.replace(index).is_some() {
            return Err(Error::Xml("multiple XML roots".into()));
        }
        for attribute in &elements[index].attributes {
            if attribute.name.namespace.is_empty()
                && attribute.name.local == "Id"
                && (attribute.value.is_empty()
                    || ids.insert(attribute.value.clone(), index).is_some())
            {
                return Err(Error::Xml("empty or duplicate Id attribute".into()));
            }
        }
        stack.push(index);
        Ok(())
    }

    fn push_text(elements: &mut [Element], stack: &[usize], text: String) -> Result<()> {
        if let Some(&element) = stack.last() {
            elements[element].children.push(Child::Text(text));
            Ok(())
        } else if text.trim().is_empty() {
            Ok(())
        } else {
            Err(Error::Xml("text outside XML root".into()))
        }
    }

    fn is(&self, index: usize, namespace: &str, local: &str) -> bool {
        self.elements[index].name.namespace == namespace && self.elements[index].name.local == local
    }

    fn children(&self, index: usize, namespace: &str, local: &str) -> Vec<usize> {
        self.elements[index]
            .children
            .iter()
            .filter_map(|child| match child {
                Child::Element(child) if self.is(*child, namespace, local) => Some(*child),
                _ => None,
            })
            .collect()
    }

    fn child(&self, index: usize, namespace: &str, local: &str) -> Option<usize> {
        self.children(index, namespace, local).first().copied()
    }

    fn required_child(&self, index: usize, namespace: &str, local: &str) -> Result<usize> {
        let children = self.children(index, namespace, local);
        if children.len() == 1 {
            Ok(children[0])
        } else {
            Err(Error::Xml(format!(
                "expected exactly one {{{namespace}}}{local} child"
            )))
        }
    }

    fn descendants(&self, index: usize, namespace: &str, local: &str) -> Vec<usize> {
        let mut output = Vec::new();
        let mut pending = vec![index];
        while let Some(current) = pending.pop() {
            for child in &self.elements[current].children {
                if let Child::Element(child) = child {
                    if self.is(*child, namespace, local) {
                        output.push(*child);
                    }
                    pending.push(*child);
                }
            }
        }
        output
    }

    fn attr(&self, index: usize, local: &str) -> Result<&str> {
        let mut matching = self.elements[index].attributes.iter().filter(|attribute| {
            attribute.name.namespace.is_empty() && attribute.name.local == local
        });
        let attribute = matching
            .next()
            .ok_or_else(|| Error::Xml(format!("missing attribute {local}")))?;
        if matching.next().is_some() {
            return Err(Error::Xml(format!("duplicate attribute {local}")));
        }
        Ok(&attribute.value)
    }

    fn optional_attr(&self, index: usize, local: &str) -> Option<&str> {
        self.elements[index]
            .attributes
            .iter()
            .find(|attribute| attribute.name.namespace.is_empty() && attribute.name.local == local)
            .map(|attribute| attribute.value.as_str())
    }

    fn text(&self, index: usize) -> Result<String> {
        let mut output = String::new();
        for child in &self.elements[index].children {
            match child {
                Child::Text(text) => output.push_str(text),
                Child::Comment(_) => {},
                Child::Element(_) => {
                    return Err(Error::Xml("expected text-only XML element".into()));
                },
            }
        }
        Ok(output)
    }

    fn required_bound_object(&self, id: &str) -> Result<usize> {
        let object = *self
            .ids
            .get(id)
            .ok_or_else(|| Error::Xml(format!("missing signed Object #{id}")))?;
        if !self.is(object, DS, "Object") || self.elements[object].parent != Some(self.root) {
            return Err(Error::Xml(format!(
                "#{id} must identify a direct ds:Object child of Signature"
            )));
        }
        Ok(object)
    }

    fn required_office_object(&self) -> Result<usize> {
        let object = self.required_bound_object(OFFICE_ID)?;
        let properties = self.required_child(object, DS, "SignatureProperties")?;
        let property = self.required_child(properties, DS, "SignatureProperty")?;
        let info = self.required_child(property, OFFICE, "SignatureInfoV1")?;
        let matching = self.descendants(object, OFFICE, "SignatureInfoV1");
        if matching.as_slice() != [info] {
            return Err(Error::Xml(
                "Office Object must contain one directly bound SignatureInfoV1".into(),
            ));
        }
        Ok(object)
    }

    fn required_signed_properties(&self) -> Result<usize> {
        let properties = *self
            .ids
            .get(SIGNED_PROPERTIES_ID)
            .ok_or_else(|| Error::Xml("missing XAdES SignedProperties".into()))?;
        if !self.is(properties, XADES, "SignedProperties") {
            return Err(Error::Xml(
                "#idSignedProperties must identify XAdES SignedProperties".into(),
            ));
        }
        let qualifying = self.elements[properties]
            .parent
            .filter(|parent| self.is(*parent, XADES, "QualifyingProperties"))
            .ok_or_else(|| {
                Error::Xml(
                    "XAdES SignedProperties must be a direct QualifyingProperties child".into(),
                )
            })?;
        if self.optional_attr(qualifying, "Target") != Some("#idPackageSignature") {
            return Err(Error::Xml(
                "XAdES QualifyingProperties must target #idPackageSignature".into(),
            ));
        }
        let object = self.elements[qualifying]
            .parent
            .filter(|parent| self.is(*parent, DS, "Object"))
            .ok_or_else(|| {
                Error::Xml("XAdES QualifyingProperties must be a direct ds:Object child".into())
            })?;
        if self.elements[object].parent != Some(self.root) {
            return Err(Error::Xml(
                "the XAdES ds:Object must be a direct Signature child".into(),
            ));
        }
        let matching = self.descendants(self.root, XADES, "SignedProperties");
        if matching.as_slice() != [properties] {
            return Err(Error::Xml(
                "Signature must contain one unambiguous XAdES SignedProperties element".into(),
            ));
        }
        Ok(properties)
    }

    fn canonicalize(&self, index: usize, canon: Canon, limits: &Limits) -> Result<Vec<u8>> {
        let mut output = ByteBuf::new(limits.max_signature_bytes());
        let mut inherited = BTreeMap::new();
        inherited.insert("xml".into(), XML_NS.into());
        self.canonicalize_element(index, &inherited, canon, &mut output)?;
        Ok(output.finish())
    }

    fn canonicalize_element(
        &self,
        index: usize,
        inherited: &BTreeMap<String, String>,
        canon: Canon,
        output: &mut ByteBuf,
    ) -> Result<()> {
        let element = &self.elements[index];
        output.byte(b'<')?;
        output.push(element.name.qualified.as_bytes())?;
        let visibly_used: HashSet<&str> = if canon.exclusive() {
            std::iter::once(
                element
                    .name
                    .qualified
                    .split_once(':')
                    .map_or("", |(prefix, _)| prefix),
            )
            .chain(element.attributes.iter().filter_map(|attribute| {
                attribute
                    .name
                    .qualified
                    .split_once(':')
                    .map(|value| value.0)
            }))
            .collect()
        } else {
            element.namespaces.keys().map(String::as_str).collect()
        };
        let mut rendered = inherited.clone();
        for (prefix, uri) in &element.namespaces {
            if prefix == "xml" || inherited.get(prefix) == Some(uri) {
                continue;
            }
            if canon.exclusive() && !visibly_used.contains(prefix.as_str()) {
                continue;
            }
            if prefix.is_empty() {
                output.push(b" xmlns=\"")?;
            } else {
                output.push(b" xmlns:")?;
                output.push(prefix.as_bytes())?;
                output.push(b"=\"")?;
            }
            escape_attr_bytes(output, uri)?;
            output.byte(b'\"')?;
            rendered.insert(prefix.clone(), uri.clone());
        }
        let mut attributes = Vec::new();
        attributes
            .try_reserve(element.attributes.len())
            .map_err(|_| Error::Limit("canonical attribute allocation failed".into()))?;
        attributes.extend(element.attributes.iter());
        attributes.sort_by(|left, right| {
            (&left.name.namespace, &left.name.local)
                .cmp(&(&right.name.namespace, &right.name.local))
        });
        for attribute in attributes {
            output.byte(b' ')?;
            output.push(attribute.name.qualified.as_bytes())?;
            output.push(b"=\"")?;
            escape_attr_bytes(output, &attribute.value)?;
            output.byte(b'\"')?;
        }
        output.byte(b'>')?;
        for child in &element.children {
            match child {
                Child::Element(child) => {
                    self.canonicalize_element(*child, &rendered, canon, output)?
                },
                Child::Text(text) => escape_text_bytes(output, text)?,
                Child::Comment(comment) if canon.comments() => {
                    output.push(b"<!--")?;
                    output.push(comment.as_bytes())?;
                    output.push(b"-->")?;
                },
                Child::Comment(_) => {},
            }
        }
        output.push(b"</")?;
        output.push(element.name.qualified.as_bytes())?;
        output.byte(b'>')?;
        Ok(())
    }
}

fn xml_error(error: impl std::fmt::Display) -> Error {
    Error::Xml(error.to_string())
}

fn expanded(qualified: &str, namespaces: &BTreeMap<String, String>, element: bool) -> Result<Name> {
    if qualified.is_empty() || qualified.matches(':').count() > 1 {
        return Err(Error::Xml(format!("invalid XML name {qualified}")));
    }
    let (prefix, local) = qualified.split_once(':').unwrap_or(("", qualified));
    if local.is_empty() || qualified.contains(':') && prefix.is_empty() {
        return Err(Error::Xml(format!("invalid XML name {qualified}")));
    }
    let namespace = if prefix.is_empty() {
        if element {
            namespaces.get("").cloned().unwrap_or_default()
        } else {
            String::new()
        }
    } else {
        namespaces
            .get(prefix)
            .cloned()
            .ok_or_else(|| Error::Xml(format!("unbound XML prefix {prefix}")))?
    };
    Ok(Name {
        qualified: qualified.into(),
        local: local.into(),
        namespace,
    })
}

fn escape_attr_bytes(output: &mut ByteBuf, value: &str) -> Result<()> {
    for character in value.chars() {
        match character {
            '&' => output.push(b"&amp;")?,
            '<' => output.push(b"&lt;")?,
            '"' => output.push(b"&quot;")?,
            '\t' => output.push(b"&#x9;")?,
            '\n' => output.push(b"&#xA;")?,
            '\r' => output.push(b"&#xD;")?,
            value => {
                let mut buffer = [0_u8; 4];
                output.push(value.encode_utf8(&mut buffer).as_bytes())?;
            },
        }
    }
    Ok(())
}

fn escape_text_bytes(output: &mut ByteBuf, value: &str) -> Result<()> {
    let mut previous = ['\0'; 2];
    for character in value.chars() {
        match character {
            '&' => output.push(b"&amp;")?,
            '<' => output.push(b"&lt;")?,
            '>' if previous == [']', ']'] => output.push(b"&gt;")?,
            '\r' => output.push(b"&#xD;")?,
            value => {
                let mut buffer = [0_u8; 4];
                output.push(value.encode_utf8(&mut buffer).as_bytes())?;
            },
        }
        previous = [previous[1], character];
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use p256::ecdsa::SigningKey;
    use zeroize::Zeroizing;

    const EC_KEY: &str = "MIGHAgEAMBMGByqGSM49AgEGCCqGSM49AwEHBG0wawIBAQQgRRzTDESH7IV1uMeZIfZNa21qwnlKbBVI2ytoHQGTf5ahRANCAATScV75bzVtN6Rz/HUh4aYhuWNyBKxl/Ti2+AYOQLiqfNEU7IQhLN68W8LD+6nX6TRbX61AZDycVlNDMlZ+b/B/";
    const EC_CERT: &str = "MIIBjjCCATOgAwIBAgIUK8agRKmBI6A8X42L6J22A/uQB3owCgYIKoZIzj0EAwIwHDEaMBgGA1UEAwwRTGl0Y2hpIEVDRFNBIFRlc3QwHhcNMjYwNzE5MDIwNzEyWhcNMzYwNzE2MDIwNzEyWjAcMRowGAYDVQQDDBFMaXRjaGkgRUNEU0EgVGVzdDBZMBMGByqGSM49AgEGCCqGSM49AwEHA0IABNJxXvlvNW03pHP8dSHhpiG5Y3IErGX9OLb4Bg5AuKp80RTshCEs3rxbwsP7qdfpNFtfrUBkPJxWU0MyVn5v8H+jUzBRMB0GA1UdDgQWBBRXzGhCGAy9zuxds3zh4t1QXPm4MDAfBgNVHSMEGDAWgBRXzGhCGAy9zuxds3zh4t1QXPm4MDAPBgNVHRMBAf8EBTADAQH/MAoGCCqGSM49BAMCA0kAMEYCIQCeGCYXmpANuierr5m/OELnJPUs8hFAo0Odwz8C9M+YjAIhAIZBsbeSpFca7Fn/TB+FadWfXe13bDsidwoo9aIq0IH1";

    fn signer() -> Signer {
        Signer::p256(SigningKey::from_bytes((&[7_u8; 32]).into()).unwrap())
            .time("2026-07-19T12:34:56Z")
            .unwrap()
    }

    fn with_bound_object(
        authored: Vec<u8>,
        signer: &Signer,
        object: &str,
        id: &str,
        kind: &str,
        transform: bool,
    ) -> Vec<u8> {
        let limits = Limits::standard();
        let object_document = Document::parse(object.as_bytes(), &limits).unwrap();
        let target = *object_document.ids.get(id).unwrap();
        let digest = Sha256::digest(
            object_document
                .canonicalize(target, Canon::Inclusive, &limits)
                .unwrap(),
        );
        let transforms = if transform {
            format!(
                "<Transforms><Transform Algorithm=\"{}\"></Transform></Transforms>",
                Canon::Inclusive.uri()
            )
        } else {
            String::new()
        };
        let reference = format!(
            "<Reference URI=\"#{id}\" Type=\"{kind}\">{transforms}<DigestMethod Algorithm=\"{SHA256}\"></DigestMethod><DigestValue>{}</DigestValue></Reference>",
            BASE64.encode(digest)
        );
        let mut xml = String::from_utf8(authored).unwrap();
        let signed_info_end = xml.find("</SignedInfo>").unwrap();
        xml.insert_str(signed_info_end, &reference);
        let signature_end = xml.rfind("</Signature>").unwrap();
        xml.insert_str(signature_end, object);

        let document = Document::parse(xml.as_bytes(), &limits).unwrap();
        let signed_info = document
            .required_child(document.root, DS, "SignedInfo")
            .unwrap();
        let signed = document
            .canonicalize(signed_info, Canon::Inclusive, &limits)
            .unwrap();
        let value = BASE64.encode(signer.sign(&signed));
        let value_start = xml.find("<SignatureValue>").unwrap() + "<SignatureValue>".len();
        let value_end = xml.find("</SignatureValue>").unwrap();
        xml.replace_range(value_start..value_end, &value);
        xml.into_bytes()
    }

    #[test]
    fn binary_profile_round_trips_without_copying_borrowed_payloads() {
        let payload = b"signed bytes";
        let references = [Ref::new("/Payload", payload).unwrap()];
        assert!(matches!(references[0].data, Cow::Borrowed(_)));
        let xml = author(Profile::Binary, &signer(), &references, &Limits::standard()).unwrap();
        let report = verify(Profile::Binary, &xml, &references[..], &Policy::strict()).unwrap();
        assert_eq!(report.integrity(), Status::Valid);
        assert_eq!(report.signature(), Status::Valid);
    }

    #[test]
    fn external_certificate_supplies_an_absent_verification_key() {
        let private = BASE64.decode(EC_KEY).unwrap();
        let certificate = BASE64.decode(EC_CERT).unwrap();
        let signer = Signer::p256_pkcs8(Zeroizing::new(private))
            .unwrap()
            .certs(vec![certificate.clone()])
            .unwrap()
            .time("2026-07-19T12:34:56Z")
            .unwrap();
        let references = [Ref::new("/Payload", b"external certificate").unwrap()];
        let authored = author(Profile::Package, &signer, &references, &Limits::standard()).unwrap();
        let full = String::from_utf8(authored).unwrap();
        let mut xml = full.clone();
        let start = xml.find("<KeyInfo>").unwrap();
        let end = xml.find("</KeyInfo>").unwrap() + "</KeyInfo>".len();
        xml.replace_range(start..end, "");

        assert!(matches!(
            verify(
                Profile::Package,
                xml.as_bytes(),
                &references[..],
                &Policy::strict()
            ),
            Err(Error::Key(_))
        ));

        let evidence = [certificate.as_slice(), certificate.as_slice()];
        let policy = Policy::strict().with_limits(Limits::standard().certificates(1).unwrap());
        let report = verify_with_certs(
            Profile::Package,
            xml.as_bytes(),
            &references[..],
            &evidence,
            &policy,
        )
        .unwrap();
        assert_eq!(report.signature(), Status::Valid);
        assert_eq!(report.certificates().len(), 1);
        assert_eq!(report.certificates()[0].der(), certificate);

        let merged = verify_with_certs(
            Profile::Package,
            full.as_bytes(),
            &references[..],
            &evidence,
            &policy,
        )
        .unwrap();
        assert_eq!(merged.signature(), Status::Valid);
        assert_eq!(merged.certificates().len(), 1);
    }

    #[test]
    fn public_canonicalizer_uses_the_safe_bounded_parser() {
        let limits = Limits::standard();
        let canonical = canonicalize(
            b"<root b=\"2\" a=\"1\"><empty/></root>",
            Canon::Inclusive,
            &limits,
        )
        .unwrap();
        assert_eq!(canonical, b"<root a=\"1\" b=\"2\"><empty></empty></root>");
        assert!(matches!(
            canonicalize(b"<!DOCTYPE root><root/>", Canon::Inclusive, &limits),
            Err(Error::Xml(_))
        ));
        assert!(matches!(
            canonicalize(b"<one/><two/>", Canon::Inclusive, &limits),
            Err(Error::Xml(_))
        ));
        let expansion_limits = Limits::standard().signature_bytes(24).unwrap();
        assert!(matches!(
            canonicalize(b"<r><a/><a/><a/></r>", Canon::Inclusive, &expansion_limits),
            Err(Error::Limit(_))
        ));
    }

    #[test]
    fn author_checks_certificate_and_escaped_output_limits_before_growth() {
        let private = BASE64.decode(EC_KEY).unwrap();
        let certificate = BASE64.decode(EC_CERT).unwrap();
        let cert_signer = Signer::p256_pkcs8(Zeroizing::new(private))
            .unwrap()
            .certs(vec![certificate.clone(), certificate.clone()])
            .unwrap()
            .time("2026-07-19T12:34:56Z")
            .unwrap();
        let references = [Ref::new("/Payload", b"bytes").unwrap()];

        let count = Limits::standard().certificates(1).unwrap();
        assert!(matches!(
            author(Profile::Package, &cert_signer, &references, &count),
            Err(Error::Limit(_))
        ));
        let item = Limits::standard()
            .certificate_bytes(certificate.len() - 1)
            .unwrap();
        assert!(matches!(
            author(Profile::Package, &cert_signer, &references, &item),
            Err(Error::Limit(_))
        ));
        let aggregate = Limits::standard()
            .total_certificate_bytes(certificate.len() * 2 - 1)
            .unwrap();
        assert!(matches!(
            author(Profile::Package, &cert_signer, &references, &aggregate),
            Err(Error::Limit(_))
        ));

        let small = Limits::standard().signature_bytes(1_024).unwrap();
        let plain = signer();
        let huge_uri = format!("/{}", "&".repeat(1_024));
        let huge = [Ref::new(&huge_uri, b"bytes").unwrap()];
        assert!(matches!(
            author(Profile::Package, &plain, &huge, &small),
            Err(Error::Limit(_))
        ));
        let relationship = [Ref::new("/rels", b"bytes")
            .unwrap()
            .transform(Transform::Relationships(vec!["&".repeat(1_024)]))];
        assert!(matches!(
            author(Profile::Package, &plain, &relationship, &small),
            Err(Error::Limit(_))
        ));
    }

    #[test]
    fn exact_manifest_coverage_is_required() {
        let first = Ref::new("/one", b"1").unwrap();
        let authored = author(Profile::Binary, &signer(), &[first], &Limits::standard()).unwrap();
        let required = [
            Ref::new("/one", b"1").unwrap(),
            Ref::new("/two", b"2").unwrap(),
        ];
        assert!(matches!(
            verify(Profile::Binary, &authored, &required[..], &Policy::strict()),
            Err(Error::Container(_))
        ));

        let report = verify(
            Profile::Binary,
            &authored,
            &required[..],
            &Policy::compatible(),
        )
        .unwrap();
        assert_eq!(report.coverage(), Coverage::Partial);
        assert_eq!(report.integrity(), Status::Valid);
        assert_eq!(report.signature(), Status::Valid);
    }

    #[test]
    fn package_profile_binds_poi_and_microsoft_optional_objects() {
        let signer = signer();
        let references = [Ref::new("/Payload", b"signed package").unwrap()];
        let authored = author(Profile::Package, &signer, &references, &Limits::standard()).unwrap();
        let office = format!(
            "<Object xmlns=\"{DS}\" Id=\"{OFFICE_ID}\"><SignatureProperties><SignatureProperty Target=\"#{SIGNATURE_ID}\"><SignatureInfoV1 xmlns=\"{OFFICE}\"><SignatureType>1</SignatureType></SignatureInfoV1></SignatureProperty></SignatureProperties></Object>"
        );
        let poi = with_bound_object(authored, &signer, &office, OFFICE_ID, DS_OBJECT, false);
        let report = verify(Profile::Package, &poi, &references[..], &Policy::strict()).unwrap();
        assert_eq!(report.signature(), Status::Valid);

        let xades = format!(
            "<Object xmlns=\"{DS}\"><xd:QualifyingProperties xmlns:xd=\"{XADES}\" Target=\"#{SIGNATURE_ID}\"><xd:SignedProperties Id=\"{SIGNED_PROPERTIES_ID}\"><xd:SignedSignatureProperties><xd:SigningTime>2026-07-19T12:34:56Z</xd:SigningTime></xd:SignedSignatureProperties></xd:SignedProperties></xd:QualifyingProperties></Object>"
        );
        let microsoft = with_bound_object(
            poi.clone(),
            &signer,
            &xades,
            SIGNED_PROPERTIES_ID,
            XADES_SIGNED_PROPERTIES,
            true,
        );
        let report = verify(
            Profile::Package,
            &microsoft,
            &references[..],
            &Policy::strict(),
        )
        .unwrap();
        assert_eq!(report.signature(), Status::Valid);

        let wrong_type = String::from_utf8(microsoft.clone()).unwrap().replace(
            XADES_SIGNED_PROPERTIES,
            "http://example.invalid/SignedProperties",
        );
        assert!(matches!(
            verify(
                Profile::Package,
                wrong_type.as_bytes(),
                &references[..],
                &Policy::strict()
            ),
            Err(Error::Xml(_))
        ));

        let wrapped = String::from_utf8(microsoft)
            .unwrap()
            .replacen(
                "<xd:SignedProperties Id=\"idSignedProperties\">",
                "<xd:Wrap><xd:SignedProperties Id=\"idSignedProperties\">",
                1,
            )
            .replacen(
                "</xd:SignedProperties>",
                "</xd:SignedProperties></xd:Wrap>",
                1,
            );
        assert!(matches!(
            verify(
                Profile::Package,
                wrapped.as_bytes(),
                &references[..],
                &Policy::strict()
            ),
            Err(Error::Xml(_))
        ));

        let mut unbound = String::from_utf8(
            author(Profile::Package, &signer, &references, &Limits::standard()).unwrap(),
        )
        .unwrap();
        let end = unbound.rfind("</Signature>").unwrap();
        unbound.insert_str(end, &office);
        assert!(matches!(
            verify(
                Profile::Package,
                unbound.as_bytes(),
                &references[..],
                &Policy::strict()
            ),
            Err(Error::Xml(_))
        ));

        let alternate_office = unbound.replace(OFFICE_ID, "alternateOfficeObject");
        assert!(matches!(
            verify(
                Profile::Package,
                alternate_office.as_bytes(),
                &references[..],
                &Policy::strict()
            ),
            Err(Error::Xml(_))
        ));

        let mut alternate_xades = String::from_utf8(
            author(Profile::Package, &signer, &references, &Limits::standard()).unwrap(),
        )
        .unwrap();
        let end = alternate_xades.rfind("</Signature>").unwrap();
        alternate_xades.insert_str(
            end,
            &xades.replace(SIGNED_PROPERTIES_ID, "alternateSignedProperties"),
        );
        assert!(matches!(
            verify(
                Profile::Package,
                alternate_xades.as_bytes(),
                &references[..],
                &Policy::strict()
            ),
            Err(Error::Xml(_))
        ));
    }

    #[test]
    fn rsa_signer_round_trips_without_reparsing_the_private_key() {
        let key = rsa::RsaPrivateKey::new(&mut rsa::rand_core::OsRng, 2_048).unwrap();
        let signer = Signer::rsa(key)
            .unwrap()
            .time("2026-07-19T12:34:56Z")
            .unwrap();
        let references = [Ref::new("/Payload", b"rsa payload").unwrap()];
        let xml = author(Profile::Binary, &signer, &references, &Limits::standard()).unwrap();
        let report = verify(Profile::Binary, &xml, &references[..], &Policy::strict()).unwrap();
        assert_eq!(report.integrity(), Status::Valid);
        assert_eq!(report.signature(), Status::Valid);
    }

    #[test]
    fn manifest_must_be_inside_the_signed_package_object() {
        let references = [Ref::new("/one", b"1").unwrap()];
        let authored =
            author(Profile::Binary, &signer(), &references, &Limits::standard()).unwrap();
        let text = String::from_utf8(authored).unwrap();
        let start = text.find("<Manifest>").unwrap();
        let end = text.find("</Manifest>").unwrap() + "</Manifest>".len();
        let manifest = &text[start..end];
        let moved = format!(
            "{}{}<Object Id=\"unsigned\">{manifest}</Object></Signature>",
            &text[..start],
            &text[end..text.rfind("</Signature>").unwrap()],
        );
        assert!(matches!(
            verify(
                Profile::Binary,
                moved.as_bytes(),
                &references[..],
                &Policy::strict()
            ),
            Err(Error::Xml(_))
        ));
    }

    #[test]
    fn active_xml_and_duplicate_ids_are_rejected() {
        let references = [Ref::new("/one", b"1").unwrap()];
        let xml = author(Profile::Binary, &signer(), &references, &Limits::standard()).unwrap();
        let mut active = b"<!DOCTYPE x>".to_vec();
        active.extend_from_slice(&xml);
        assert!(matches!(
            verify(Profile::Binary, &active, &references[..], &Policy::strict()),
            Err(Error::Xml(_))
        ));
        let duplicate = String::from_utf8(xml)
            .unwrap()
            .replace("Id=\"idOfficeObject\"", "Id=\"idPackageObject\"");
        assert!(matches!(
            verify(
                Profile::Binary,
                duplicate.as_bytes(),
                &references[..],
                &Policy::strict()
            ),
            Err(Error::Xml(_))
        ));
    }
}
