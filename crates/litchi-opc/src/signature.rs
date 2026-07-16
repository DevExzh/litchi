//! Offline, verification-only support for OPC digital signatures.

use crate::{OpcPackage, PackURI, Part, Relationships, TargetMode};
use base64::Engine;
use quick_xml::{events::{BytesStart, Event}, Reader};
use rsa::{BigUint, RsaPublicKey, traits::PublicKeyParts, pkcs1v15::{Signature as RsaSignature,
    VerifyingKey}, pkcs8::DecodePublicKey};
use sha1_legacy::Sha1;
use sha2_legacy::{Digest, Sha256, Sha384, Sha512};
use signature::Verifier;
use std::{collections::{BTreeMap, HashMap, HashSet}, str};
use subtle::ConstantTimeEq;
use thiserror::Error;

const DS: &str = "http://www.w3.org/2000/09/xmldsig#";
const MDSSI: &str = "http://schemas.openxmlformats.org/package/2006/digital-signature";
const REL_NS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const ORIGIN_REL: &str = "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin";
const SIGNATURE_REL: &str = "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature";
const CERTIFICATE_REL: &str = "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/certificate";
const REL_TRANSFORM: &str = "http://schemas.openxmlformats.org/package/2006/RelationshipTransform";
const C14N: &str = "http://www.w3.org/TR/2001/REC-xml-c14n-20010315";
const C14N_COMMENTS: &str = "http://www.w3.org/TR/2001/REC-xml-c14n-20010315#WithComments";
const SHA1: &str = "http://www.w3.org/2000/09/xmldsig#sha1";
const SHA256: &str = "http://www.w3.org/2001/04/xmlenc#sha256";
const SHA384: &str = "http://www.w3.org/2001/04/xmldsig-more#sha384";
const SHA512: &str = "http://www.w3.org/2001/04/xmlenc#sha512";
const RSA_SHA1: &str = "http://www.w3.org/2000/09/xmldsig#rsa-sha1";
const RSA_SHA256: &str = "http://www.w3.org/2001/04/xmldsig-more#rsa-sha256";
const RSA_SHA384: &str = "http://www.w3.org/2001/04/xmldsig-more#rsa-sha384";
const RSA_SHA512: &str = "http://www.w3.org/2001/04/xmldsig-more#rsa-sha512";
const XML_NS: &str = "http://www.w3.org/XML/1998/namespace";
const ORIGIN_CONTENT_TYPE: &str = "application/vnd.openxmlformats-package.digital-signature-origin";
const SIGNATURE_CONTENT_TYPE: &str = "application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml";
const CERTIFICATE_CONTENT_TYPE: &str = "application/vnd.openxmlformats-package.digital-signature-certificate";

pub type Result<T> = std::result::Result<T, DigitalSignatureError>;

#[derive(Debug, Error)]
pub enum DigitalSignatureError {
    #[error("invalid OPC digital-signature graph: {0}")] InvalidGraph(String),
    #[error("invalid or unsafe signature XML: {0}")] InvalidXml(String),
    #[error("digital-signature resource limit exceeded: {0}")] LimitExceeded(String),
    #[error("unsupported digital-signature algorithm or transform: {0}")] UnsupportedAlgorithm(String),
    #[error("SHA-1 is disallowed by the strict verification policy")] Sha1Disallowed,
    #[error("invalid RSA verification key: {0}")] InvalidKey(String),
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Sha1Policy { AllowWithWarning, Reject }

#[derive(Debug, Clone)]
pub struct SignatureVerificationPolicy {
    pub sha1: Sha1Policy,
    pub max_signature_part_bytes: usize,
    pub max_xml_depth: usize,
    pub max_xml_elements: usize,
    pub max_attributes_per_element: usize,
    pub max_references: usize,
    pub max_embedded_certificate_bytes: usize,
    pub max_rsa_modulus_bits: usize,
}
impl SignatureVerificationPolicy {
    pub fn compatibility() -> Self { Self { sha1: Sha1Policy::AllowWithWarning,
        max_signature_part_bytes: 8*1024*1024, max_xml_depth: 128,
        max_xml_elements: 100_000, max_attributes_per_element: 256,
        max_references: 20_000, max_embedded_certificate_bytes: 1024*1024,
        max_rsa_modulus_bits: 16_384 } }
    pub fn strict() -> Self { Self { sha1: Sha1Policy::Reject, ..Self::compatibility() } }
}
impl Default for SignatureVerificationPolicy { fn default() -> Self { Self::compatibility() } }

#[derive(Debug, Clone, Copy, PartialEq, Eq)] pub enum VerificationStatus { Valid, Invalid }
#[derive(Debug, Clone, Copy, PartialEq, Eq)] pub enum CertificateTrust { NotEvaluated }
#[derive(Debug, Clone, PartialEq, Eq)] pub struct EmbeddedCertificate { pub der: Vec<u8> }
#[derive(Debug, Clone, PartialEq, Eq)] pub struct ReferenceVerification {
    pub uri: String, pub status: VerificationStatus,
}
#[derive(Debug, Clone)] pub struct DigitalSignatureVerification {
    pub signature_part: PackURI,
    pub package_integrity: VerificationStatus,
    pub signature_value: VerificationStatus,
    pub certificate_trust: CertificateTrust,
    pub references: Vec<ReferenceVerification>,
    pub certificates: Vec<EmbeddedCertificate>,
    pub uses_sha1: bool,
}

pub fn verify_package(package: &OpcPackage, policy: &SignatureVerificationPolicy)
    -> Result<Vec<DigitalSignatureVerification>> {
    validate_policy(policy)?;
    let origins: Vec<_> = package.rels().iter().filter(|r| r.reltype()==ORIGIN_REL).collect();
    if origins.is_empty() { return Ok(vec![]); }
    if origins.len()!=1 { return Err(DigitalSignatureError::InvalidGraph(
        format!("expected one signature origin relationship, found {}", origins.len()))); }
    internal(origins[0].target_mode(), "signature origin")?;
    let origin_uri=origins[0].target_partname().map_err(graph)?;
    let origin=package.get_part(&origin_uri).map_err(graph)?;
    content_type(origin, ORIGIN_CONTENT_TYPE, "signature origin")?;
    let mut uris=Vec::new(); let mut seen=HashSet::new();
    for rel in origin.rels().iter().filter(|r| r.reltype()==SIGNATURE_REL) {
        internal(rel.target_mode(), "signature")?;
        let uri=rel.target_partname().map_err(graph)?;
        if !seen.insert(uri.clone()) { return Err(DigitalSignatureError::InvalidGraph(
            format!("duplicate signature target {}", uri.as_str()))); }
        let part=package.get_part(&uri).map_err(graph)?;
        content_type(part, SIGNATURE_CONTENT_TYPE, "signature")?;
        for cert_rel in part.rels().iter().filter(|r| r.reltype()==CERTIFICATE_REL) {
            internal(cert_rel.target_mode(), "certificate")?;
            let cert=package.get_part(&cert_rel.target_partname().map_err(graph)?).map_err(graph)?;
            content_type(cert, CERTIFICATE_CONTENT_TYPE, "certificate")?;
        }
        uris.push(uri);
    }
    if uris.is_empty() { return Err(DigitalSignatureError::InvalidGraph(
        "signature origin has no signature relationships".into())); }
    uris.sort_by(|a,b|a.as_str().cmp(b.as_str()));
    uris.into_iter().map(|u| verify_one(package,u,policy)).collect()
}
fn graph<E: std::fmt::Display>(e:E)->DigitalSignatureError {
    DigitalSignatureError::InvalidGraph(e.to_string())
}
fn internal(mode:TargetMode, what:&str)->Result<()> { if mode==TargetMode::External {
    Err(DigitalSignatureError::InvalidGraph(format!("{what} relationship must be internal")))
} else { Ok(()) } }
fn content_type(p:&dyn Part, expected:&str, what:&str)->Result<()> {
    if p.content_type()!=expected { Err(DigitalSignatureError::InvalidGraph(format!(
        "{what} part {} has content type {}, expected {expected}",p.partname().as_str(),p.content_type())))
    } else { Ok(()) }
}
fn validate_policy(p:&SignatureVerificationPolicy)->Result<()> {
    if p.max_signature_part_bytes==0 || p.max_xml_depth==0 || p.max_xml_elements==0 ||
       p.max_attributes_per_element==0 || p.max_references==0 ||
       p.max_embedded_certificate_bytes==0 || p.max_rsa_modulus_bits<512 {
        Err(DigitalSignatureError::LimitExceeded("invalid zero or undersized policy limit".into()))
    } else { Ok(()) }
}

fn verify_one(package:&OpcPackage, uri:PackURI, policy:&SignatureVerificationPolicy)
    ->Result<DigitalSignatureVerification> {
    let part=package.get_part(&uri).map_err(graph)?;
    if part.blob().len()>policy.max_signature_part_bytes { return Err(
        DigitalSignatureError::LimitExceeded(format!("signature part {} is too large",uri.as_str()))); }
    let doc=Document::parse(part.blob(),policy)?;
    if !doc.is(doc.root,DS,"Signature") { return Err(DigitalSignatureError::InvalidXml(
        "root element must be ds:Signature".into())); }
    let signed=doc.required_child(doc.root,DS,"SignedInfo")?;
    let canon=doc.required_child(signed,DS,"CanonicalizationMethod")?;
    let comments=c14n_mode(doc.attr(canon,"Algorithm")?)?;
    let signed_bytes=doc.canonicalize(signed,comments);
    let method=doc.required_child(signed,DS,"SignatureMethod")?;
    let sig_alg=doc.attr(method,"Algorithm")?;
    sha1_policy(sig_alg,policy)?;
    if !matches!(sig_alg,RSA_SHA1|RSA_SHA256|RSA_SHA384|RSA_SHA512) {
        return Err(DigitalSignatureError::UnsupportedAlgorithm(sig_alg.into())); }
    let mut refs=doc.children(signed,DS,"Reference");
    for manifest in doc.descendants(doc.root,DS,"Manifest") {
        refs.extend(doc.children(manifest,DS,"Reference"));
    }
    if refs.is_empty() { return Err(DigitalSignatureError::InvalidXml("no signature references".into())); }
    if refs.len()>policy.max_references { return Err(DigitalSignatureError::LimitExceeded(
        format!("{} references exceed policy",refs.len()))); }
    let mut reports=Vec::with_capacity(refs.len()); let mut uses_sha1=sig_alg==RSA_SHA1;
    for r in refs { let (report,weak)=verify_reference(package,&doc,r,policy)?;
        uses_sha1|=weak; reports.push(report); }
    let certs=extract_certificates(&doc,policy)?;
    let key=extract_key(&doc,&certs,policy)?;
    let value=decode64(&doc.text(doc.required_child(doc.root,DS,"SignatureValue")?)?,
        policy.max_signature_part_bytes,"SignatureValue")?;
    let sig=RsaSignature::try_from(value.as_slice()).map_err(|e|DigitalSignatureError::InvalidKey(e.to_string()))?;
    let sig_ok=match sig_alg { RSA_SHA1=>VerifyingKey::<Sha1>::new(key).verify(&signed_bytes,&sig),
        RSA_SHA256=>VerifyingKey::<Sha256>::new(key).verify(&signed_bytes,&sig),
        RSA_SHA384=>VerifyingKey::<Sha384>::new(key).verify(&signed_bytes,&sig),
        RSA_SHA512=>VerifyingKey::<Sha512>::new(key).verify(&signed_bytes,&sig), _=>unreachable!() }.is_ok();
    let integrity=reports.iter().all(|r|r.status==VerificationStatus::Valid);
    Ok(DigitalSignatureVerification { signature_part:uri, package_integrity:status(integrity),
        signature_value:status(sig_ok), certificate_trust:CertificateTrust::NotEvaluated,
        references:reports, certificates:certs, uses_sha1 })
}
fn status(v:bool)->VerificationStatus { if v {VerificationStatus::Valid}else{VerificationStatus::Invalid} }
fn sha1_policy(a:&str,p:&SignatureVerificationPolicy)->Result<()> { if matches!(a,SHA1|RSA_SHA1)&&
    p.sha1==Sha1Policy::Reject {Err(DigitalSignatureError::Sha1Disallowed)}else{Ok(())} }
fn c14n_mode(a:&str)->Result<bool>{match a{C14N=>Ok(false),C14N_COMMENTS=>Ok(true),
    _=>Err(DigitalSignatureError::UnsupportedAlgorithm(a.into()))}}

#[derive(Debug)] enum Transform { Relationships(Vec<String>), Canonical(bool) }
fn transforms(doc:&Document,r:usize)->Result<Vec<Transform>>{
    let Some(ts)=doc.child(r,DS,"Transforms") else{return Ok(vec![])}; let mut out=Vec::new();
    for t in doc.children(ts,DS,"Transform") { let alg=doc.attr(t,"Algorithm")?;
        match alg { REL_TRANSFORM=>{if !out.is_empty(){return Err(DigitalSignatureError::UnsupportedAlgorithm(
            "RelationshipTransform must be first".into()))} let mut ids=Vec::new();
            for n in doc.children(t,MDSSI,"RelationshipReference"){ids.push(doc.attr(n,"SourceId")?.into())}
            if ids.is_empty(){return Err(DigitalSignatureError::InvalidXml("empty RelationshipTransform".into()))}
            let count=ids.len();ids.sort();ids.dedup();if ids.len()!=count{return Err(
                DigitalSignatureError::InvalidXml("duplicate RelationshipReference".into()))}out.push(Transform::Relationships(ids));},
            C14N|C14N_COMMENTS=>out.push(Transform::Canonical(c14n_mode(alg)?)),
            _=>return Err(DigitalSignatureError::UnsupportedAlgorithm(alg.into())) }
    } Ok(out)
}
fn verify_reference(package:&OpcPackage,doc:&Document,r:usize,p:&SignatureVerificationPolicy)
    ->Result<(ReferenceVerification,bool)>{
    let uri=doc.attr(r,"URI")?.to_string(); let dm=doc.required_child(r,DS,"DigestMethod")?;
    let alg=doc.attr(dm,"Algorithm")?; sha1_policy(alg,p)?;
    let expected=decode64(&doc.text(doc.required_child(r,DS,"DigestValue")?)?,128,"DigestValue")?;
    let data=dereference(package,doc,&uri,&transforms(doc,r)?,p)?;
    let actual=match alg{SHA1=>Sha1::digest(&data).to_vec(),SHA256=>Sha256::digest(&data).to_vec(),
        SHA384=>Sha384::digest(&data).to_vec(),SHA512=>Sha512::digest(&data).to_vec(),
        _=>return Err(DigitalSignatureError::UnsupportedAlgorithm(alg.into()))};
    let valid=actual.len()==expected.len()&&bool::from(actual.ct_eq(&expected));
    Ok((ReferenceVerification{uri,status:status(valid)},alg==SHA1))
}
fn dereference(package:&OpcPackage,doc:&Document,uri:&str,ts:&[Transform],p:&SignatureVerificationPolicy)
    ->Result<Vec<u8>>{
    if let Some(id)=uri.strip_prefix('#'){let n=*doc.ids.get(id).ok_or_else(||
        DigitalSignatureError::InvalidXml(format!("unknown Id {id}")))?;let mut comments=false;
        for t in ts{match t{Transform::Canonical(c)=>comments=*c,Transform::Relationships(_)=>return Err(
            DigitalSignatureError::UnsupportedAlgorithm("relationship transform on fragment".into()))}}
        return Ok(doc.canonicalize(n,comments));}
    if !uri.starts_with('/')||uri.contains('#'){return Err(DigitalSignatureError::InvalidXml(
        format!("invalid package reference {uri}")))}
    let (path,q)=uri.split_once('?').ok_or_else(||DigitalSignatureError::InvalidXml(
        format!("reference lacks ContentType query: {uri}")))?;
    let ct=q.strip_prefix("ContentType=").filter(|s|!s.is_empty()&&!s.contains('&')).ok_or_else(||
        DigitalSignatureError::InvalidXml(format!("invalid ContentType query: {uri}")))?;
    if let Some(ids)=ts.iter().find_map(|t|if let Transform::Relationships(v)=t{Some(v)}else{None}){
        if ct!="application/vnd.openxmlformats-package.relationships+xml"||!path.ends_with(".rels"){
            return Err(DigitalSignatureError::InvalidXml("invalid relationship transform target".into()))}
        return canonical_relationships(relationships(package,path)?,ids); }
    let pu=PackURI::new(path.to_string()).map_err(DigitalSignatureError::InvalidXml)?;
    let part=package.get_part(&pu).map_err(graph)?; if part.content_type()!=ct{return Err(
        DigitalSignatureError::InvalidGraph(format!("reference content type mismatch for {path}")))}
    if ts.is_empty(){return Ok(part.blob().to_vec())}
    if ts.len()!=1{return Err(DigitalSignatureError::UnsupportedAlgorithm("invalid part transforms".into()))}
    let Transform::Canonical(comments)=ts[0] else{unreachable!()};
    let parsed=Document::parse(part.blob(),p)?;Ok(parsed.canonicalize(parsed.root,comments))
}
fn relationships<'a>(p:&'a OpcPackage,path:&str)->Result<&'a Relationships>{
    if path=="/_rels/.rels"{return Ok(p.rels())} let (dir,file)=path.rsplit_once("/_rels/").ok_or_else(||
        DigitalSignatureError::InvalidXml(format!("invalid relationship URI {path}")))?;
    let source=file.strip_suffix(".rels").filter(|s|!s.is_empty()).ok_or_else(||
        DigitalSignatureError::InvalidXml(format!("invalid relationship URI {path}")))?;
    let uri=PackURI::new(format!("{dir}/{source}")).map_err(DigitalSignatureError::InvalidXml)?;
    Ok(p.get_part(&uri).map_err(graph)?.rels())
}
fn canonical_relationships(rels:&Relationships,ids:&[String])->Result<Vec<u8>>{
    let mut chosen=Vec::new();for id in ids{chosen.push(rels.get(id).ok_or_else(||
        DigitalSignatureError::InvalidGraph(format!("relationship {id} not found")))?)}
    chosen.sort_by(|a,b|a.r_id().cmp(b.r_id()));let mut o=Vec::new();o.extend_from_slice(b"<Relationships xmlns=\"");
    attr_escape(&mut o,REL_NS);o.extend_from_slice(b"\">");for r in chosen{o.extend_from_slice(b"<Relationship Id=\"");
    attr_escape(&mut o,r.r_id());o.extend_from_slice(b"\" Target=\"");attr_escape(&mut o,r.target_ref());
    o.extend_from_slice(b"\" TargetMode=\"");o.extend_from_slice(if r.target_mode()==TargetMode::Internal{b"Internal"}else{b"External"});
    o.extend_from_slice(b"\" Type=\"");attr_escape(&mut o,r.reltype());o.extend_from_slice(b"\"></Relationship>");}
    o.extend_from_slice(b"</Relationships>");Ok(o)
}

fn extract_certificates(d:&Document,p:&SignatureVerificationPolicy)->Result<Vec<EmbeddedCertificate>>{
    d.descendants(d.root,DS,"X509Certificate").into_iter().map(|n|Ok(EmbeddedCertificate{
        der:decode64(&d.text(n)?,p.max_embedded_certificate_bytes,"X509Certificate")?})).collect()
}
fn extract_key(d:&Document,certs:&[EmbeddedCertificate],p:&SignatureVerificationPolicy)->Result<RsaPublicKey>{
    if let Some(k)=d.descendants(d.root,DS,"RSAKeyValue").first().copied(){
        let m=decode64(&d.text(d.required_child(k,DS,"Modulus")?)?,p.max_rsa_modulus_bits.div_ceil(8),"modulus")?;
        let e=decode64(&d.text(d.required_child(k,DS,"Exponent")?)?,16,"exponent")?;
        let bits=m.len()*8-m.first().map_or(0,|b|b.leading_zeros() as usize);if !(512..=p.max_rsa_modulus_bits).contains(&bits){
            return Err(DigitalSignatureError::InvalidKey(format!("RSA modulus is {bits} bits")))}
        return RsaPublicKey::new(BigUint::from_bytes_be(&m),BigUint::from_bytes_be(&e)).map_err(|e|
            DigitalSignatureError::InvalidKey(e.to_string()));}
    let cert=certs.first().ok_or_else(||DigitalSignatureError::InvalidKey("no RSA key or certificate".into()))?;
    let key=RsaPublicKey::from_public_key_der(spki(&cert.der)?).map_err(|e|DigitalSignatureError::InvalidKey(e.to_string()))?;
    if key.size()*8>p.max_rsa_modulus_bits{return Err(DigitalSignatureError::InvalidKey("RSA key too large".into()))}Ok(key)
}
fn spki(cert:&[u8])->Result<&[u8]>{let (tag,outer,end)=tlv(cert,0)?;if tag!=0x30||end!=cert.len(){return Err(
    DigitalSignatureError::InvalidKey("invalid DER certificate".into()))}let (tag,tbs,_)=tlv(outer,0)?;if tag!=0x30{return Err(
    DigitalSignatureError::InvalidKey("invalid TBSCertificate".into()))}let mut pos=0;if tbs.first()==Some(&0xa0){pos=tlv(tbs,pos)?.2}
    for _ in 0..5{pos=tlv(tbs,pos)?.2}let start=pos;let (tag,_,end)=tlv(tbs,pos)?;if tag!=0x30{return Err(
        DigitalSignatureError::InvalidKey("invalid SubjectPublicKeyInfo".into()))}Ok(&tbs[start..end])}
fn tlv(d:&[u8],p:usize)->Result<(u8,&[u8],usize)>{let tag=*d.get(p).ok_or_else(||DigitalSignatureError::InvalidKey("truncated DER".into()))?;
    let b=*d.get(p+1).ok_or_else(||DigitalSignatureError::InvalidKey("truncated DER".into()))?;let (len,h)=if b&128==0{(b as usize,2)}else{
    let n=(b&127)as usize;if n==0||n>std::mem::size_of::<usize>(){return Err(DigitalSignatureError::InvalidKey("invalid DER length".into()))}
    let mut l=0usize;for x in d.get(p+2..p+2+n).ok_or_else(||DigitalSignatureError::InvalidKey("truncated DER".into()))?{
        l=l.checked_mul(256).and_then(|v|v.checked_add(*x as usize)).ok_or_else(||DigitalSignatureError::InvalidKey("DER overflow".into()))?}if l<128{return Err(DigitalSignatureError::InvalidKey("non-minimal DER".into()))}(l,2+n)};
    let s=p.checked_add(h).ok_or_else(||DigitalSignatureError::InvalidKey("DER overflow".into()))?;let e=s.checked_add(len).ok_or_else(||DigitalSignatureError::InvalidKey("DER overflow".into()))?;
    Ok((tag,d.get(s..e).ok_or_else(||DigitalSignatureError::InvalidKey("truncated DER".into()))?,e))}
fn decode64(s:&str,max:usize,what:&str)->Result<Vec<u8>>{let compact:String=s.chars().filter(|c|!c.is_whitespace()).collect();
    if compact.len()>max.saturating_mul(4)/3+8{return Err(DigitalSignatureError::LimitExceeded(format!("{what} too large")))}
    let v=base64::engine::general_purpose::STANDARD.decode(compact).map_err(|e|DigitalSignatureError::InvalidXml(e.to_string()))?;
    if v.len()>max{Err(DigitalSignatureError::LimitExceeded(format!("{what} too large")))}else{Ok(v)}}

#[derive(Clone)]struct Name{q:String,local:String,ns:String}
#[derive(Clone)]struct Attribute{name:Name,value:String}
#[derive(Clone)]enum Child{Element(usize),Text(String),Comment(String)}
#[derive(Clone)]struct Element{name:Name,attrs:Vec<Attribute>,namespaces:BTreeMap<String,String>,children:Vec<Child>}
struct Document{elements:Vec<Element>,root:usize,ids:HashMap<String,usize>}
impl Document{
 fn parse(bytes:&[u8],p:&SignatureVerificationPolicy)->Result<Self>{if bytes.len()>p.max_signature_part_bytes{return Err(
    DigitalSignatureError::LimitExceeded("XML too large".into()))}let mut r=Reader::from_reader(bytes);r.config_mut().trim_text(false);
    let(mut elements,mut stack,mut ids,mut root,mut buf)=(Vec::new(),Vec::new(),HashMap::new(),None,Vec::new());loop{
    match r.read_event_into(&mut buf){Ok(Event::Start(s))=>Self::start(&s,r.decoder(),p,&mut elements,&mut stack,&mut ids,&mut root)?,
    Ok(Event::Empty(s))=>{Self::start(&s,r.decoder(),p,&mut elements,&mut stack,&mut ids,&mut root)?;stack.pop();},
    Ok(Event::End(_))=>{if stack.pop().is_none(){return Err(DigitalSignatureError::InvalidXml("unexpected end tag".into()))}},
    Ok(Event::Text(t))=>{let raw=t.xml10_content().map_err(xml)?;let text=quick_xml::escape::unescape(&raw).map_err(xml)?.into_owned();Self::push_text(&mut elements,&stack,text)?},
    Ok(Event::CData(t))=>Self::push_text(&mut elements,&stack,t.xml10_content().map_err(xml)?.into_owned())?,
    Ok(Event::Comment(c))=>if let Some(&n)=stack.last(){elements[n].children.push(Child::Comment(c.xml10_content().map_err(xml)?.into_owned()))},
    Ok(Event::Decl(_))=>if root.is_some(){return Err(DigitalSignatureError::InvalidXml("late XML declaration".into()))},
    Ok(Event::DocType(_)|Event::PI(_)|Event::GeneralRef(_))=>return Err(DigitalSignatureError::InvalidXml("DTD, PI, and entity references are rejected".into())),
    Ok(Event::Eof)=>break,Err(e)=>return Err(DigitalSignatureError::InvalidXml(e.to_string()))};buf.clear()}
    if !stack.is_empty(){return Err(DigitalSignatureError::InvalidXml("unclosed element".into()))}Ok(Self{elements,root:root.ok_or_else(||DigitalSignatureError::InvalidXml("no root".into()))?,ids})}
 #[allow(clippy::too_many_arguments)]fn start(s:&BytesStart<'_>,decoder:quick_xml::encoding::Decoder,p:&SignatureVerificationPolicy,
    es:&mut Vec<Element>,stack:&mut Vec<usize>,ids:&mut HashMap<String,usize>,root:&mut Option<usize>)->Result<()>{
    if stack.len()>=p.max_xml_depth||es.len()>=p.max_xml_elements{return Err(DigitalSignatureError::LimitExceeded("XML structure limit".into()))}
    let q=str::from_utf8(s.name().as_ref()).map_err(xml)?.to_string();let mut raw=Vec::new();for a in s.attributes().with_checks(true){
        let a=a.map_err(xml)?;if raw.len()>=p.max_attributes_per_element{return Err(DigitalSignatureError::LimitExceeded("attribute limit".into()))}
        raw.push((str::from_utf8(a.key.as_ref()).map_err(xml)?.to_string(),a.decoded_and_normalized_value(quick_xml::XmlVersion::Explicit1_0, decoder).map_err(xml)?.into_owned()));}
    let mut ns=stack.last().map_or_else(||{let mut m=BTreeMap::new();m.insert("xml".into(),XML_NS.into());m},|n|es[*n].namespaces.clone());
    for(a,v)in &raw{if a=="xmlns"{ns.insert("".into(),v.clone());}else if let Some(pre)=a.strip_prefix("xmlns:"){
        if pre.is_empty()||pre=="xmlns"||(pre=="xml"&&v!=XML_NS)||(pre!="xml"&&v==XML_NS)||v.is_empty(){return Err(
            DigitalSignatureError::InvalidXml(format!("invalid namespace {a}")))}ns.insert(pre.into(),v.clone());}}
    let name=expanded(&q,&ns,true)?;let mut attrs=Vec::new();let mut unique=HashSet::new();for(q,value)in raw{if q=="xmlns"||q.starts_with("xmlns:"){continue}
        let name=expanded(&q,&ns,false)?;if !unique.insert((name.ns.clone(),name.local.clone())){return Err(DigitalSignatureError::InvalidXml(format!("duplicate attribute {q}")))}attrs.push(Attribute{name,value});}
    let n=es.len();es.push(Element{name,attrs,namespaces:ns,children:Vec::new()});if let Some(&parent)=stack.last(){es[parent].children.push(Child::Element(n))}else if root.replace(n).is_some(){return Err(DigitalSignatureError::InvalidXml("multiple roots".into()))}
    for a in &es[n].attrs{if a.name.ns.is_empty()&&a.name.local=="Id"&&(a.value.is_empty()||ids.insert(a.value.clone(),n).is_some()){
        return Err(DigitalSignatureError::InvalidXml("empty or duplicate Id".into()))}}stack.push(n);Ok(())}
 fn push_text(es:&mut[Element],stack:&[usize],text:String)->Result<()>{if let Some(&n)=stack.last(){es[n].children.push(Child::Text(text));Ok(())}else if text.trim().is_empty(){Ok(())}else{Err(DigitalSignatureError::InvalidXml("text outside root".into()))}}
 fn is(&self,n:usize,ns:&str,local:&str)->bool{self.elements[n].name.ns==ns&&self.elements[n].name.local==local}
 fn children(&self,n:usize,ns:&str,local:&str)->Vec<usize>{self.elements[n].children.iter().filter_map(|c|match c{Child::Element(i)if self.is(*i,ns,local)=>Some(*i),_=>None}).collect()}
 fn child(&self,n:usize,ns:&str,local:&str)->Option<usize>{self.children(n,ns,local).first().copied()}
 fn required_child(&self,n:usize,ns:&str,local:&str)->Result<usize>{let v=self.children(n,ns,local);if v.len()==1{Ok(v[0])}else{Err(DigitalSignatureError::InvalidXml(format!("expected one {{{ns}}}{local}")))}}
 fn descendants(&self,n:usize,ns:&str,local:&str)->Vec<usize>{let mut out=Vec::new();let mut todo=vec![n];while let Some(x)=todo.pop(){for c in &self.elements[x].children{if let Child::Element(i)=c{if self.is(*i,ns,local){out.push(*i)}todo.push(*i)}}}out}
 fn attr(&self,n:usize,local:&str)->Result<&str>{let mut v=self.elements[n].attrs.iter().filter(|a|a.name.ns.is_empty()&&a.name.local==local);let a=v.next().ok_or_else(||DigitalSignatureError::InvalidXml(format!("missing attribute {local}")))?;if v.next().is_some(){return Err(DigitalSignatureError::InvalidXml(format!("duplicate attribute {local}")))}Ok(&a.value)}
 fn text(&self,n:usize)->Result<String>{let mut out=String::new();for c in &self.elements[n].children{match c{Child::Text(s)=>out.push_str(s),Child::Comment(_)=>{},Child::Element(_)=>return Err(DigitalSignatureError::InvalidXml("expected text-only element".into()))}}Ok(out)}
 fn canonicalize(&self,n:usize,comments:bool)->Vec<u8>{let mut out=Vec::new();let mut inherited=BTreeMap::new();inherited.insert("xml".into(),XML_NS.into());self.canon(n,&inherited,comments,&mut out);out}
 fn canon(&self,n:usize,inherited:&BTreeMap<String,String>,comments:bool,out:&mut Vec<u8>){let e=&self.elements[n];out.push(b'<');out.extend_from_slice(e.name.q.as_bytes());
    for(pre,uri)in &e.namespaces{if pre=="xml"||inherited.get(pre)==Some(uri){continue}if pre.is_empty(){out.extend_from_slice(b" xmlns=\"")}else{out.extend_from_slice(b" xmlns:");out.extend_from_slice(pre.as_bytes());out.extend_from_slice(b"=\"")}attr_escape(out,uri);out.push(b'\"')}
    let mut attrs:Vec<_>=e.attrs.iter().collect();attrs.sort_by(|a,b|(&a.name.ns,&a.name.local).cmp(&(&b.name.ns,&b.name.local)));for a in attrs{out.push(b' ');out.extend_from_slice(a.name.q.as_bytes());out.extend_from_slice(b"=\"");attr_escape(out,&a.value);out.push(b'\"')}out.push(b'>');
    for c in &e.children{match c{Child::Element(i)=>self.canon(*i,&e.namespaces,comments,out),Child::Text(s)=>text_escape(out,s),Child::Comment(s)if comments=>{out.extend_from_slice(b"<!--");out.extend_from_slice(s.as_bytes());out.extend_from_slice(b"-->")},_=>{}}}out.extend_from_slice(b"</");out.extend_from_slice(e.name.q.as_bytes());out.push(b'>')}
}
fn xml<E:std::fmt::Display>(e:E)->DigitalSignatureError{DigitalSignatureError::InvalidXml(e.to_string())}
fn expanded(q:&str,ns:&BTreeMap<String,String>,element:bool)->Result<Name>{if q.is_empty()||q.matches(':').count()>1{return Err(xml(format!("invalid name {q}")))}let(pre,local)=q.split_once(':').unwrap_or(("",q));if local.is_empty()||q.contains(':')&&pre.is_empty(){return Err(xml(format!("invalid name {q}")))}let uri=if pre.is_empty(){if element{ns.get("").cloned().unwrap_or_default()}else{String::new()}}else{ns.get(pre).cloned().ok_or_else(||xml(format!("unbound prefix {pre}")))?};Ok(Name{q:q.into(),local:local.into(),ns:uri})}
fn attr_escape(o:&mut Vec<u8>,s:&str){for c in s.chars(){match c{'&'=>o.extend_from_slice(b"&amp;"),'<'=>o.extend_from_slice(b"&lt;"),'"'=>o.extend_from_slice(b"&quot;"),'\t'=>o.extend_from_slice(b"&#x9;"),'\n'=>o.extend_from_slice(b"&#xA;"),'\r'=>o.extend_from_slice(b"&#xD;"),_=>{let mut b=[0;4];o.extend_from_slice(c.encode_utf8(&mut b).as_bytes())}}}}
fn text_escape(o:&mut Vec<u8>,s:&str){let mut prev=['\0';2];for c in s.chars(){match c{'&'=>o.extend_from_slice(b"&amp;"),'<'=>o.extend_from_slice(b"&lt;"),c if c=='>'&&prev==[']',']']=>o.extend_from_slice(b"&gt;"),'\r'=>o.extend_from_slice(b"&#xD;"),_=>{let mut b=[0;4];o.extend_from_slice(c.encode_utf8(&mut b).as_bytes())}}prev=[prev[1],c]}}

#[cfg(test)]mod tests{use super::*;
 const DOCX:&[u8]=include_bytes!("../../../3rdparty/poi/test-data/xmldsign/hello-world-signed.docx");
 const XLSX:&[u8]=include_bytes!("../../../3rdparty/poi/test-data/xmldsign/hello-world-signed.xlsx");
 const PPTX:&[u8]=include_bytes!("../../../3rdparty/poi/test-data/xmldsign/hello-world-signed.pptx");
 const TWICE:&[u8]=include_bytes!("../../../3rdparty/poi/test-data/xmldsign/hello-world-signed-twice.docx");
 fn valid(bytes:&[u8],count:usize){let p=OpcPackage::from_bytes(bytes).unwrap();let reports=p.verify_digital_signatures(&SignatureVerificationPolicy::compatibility()).unwrap();assert_eq!(reports.len(),count);for r in reports{assert_eq!(r.package_integrity,VerificationStatus::Valid);assert_eq!(r.signature_value,VerificationStatus::Valid);assert_eq!(r.certificate_trust,CertificateTrust::NotEvaluated);assert!(r.uses_sha1);assert!(!r.certificates.is_empty())}}
 #[test]fn verifies_real_poi_office_fixtures(){valid(DOCX,1);valid(XLSX,1);valid(PPTX,1)}
 #[test]fn verifies_real_poi_twice_signed_fixture(){valid(TWICE,2)}
 #[test]fn strict_rejects_sha1(){let p=OpcPackage::from_bytes(DOCX).unwrap();assert!(matches!(p.verify_digital_signatures(&SignatureVerificationPolicy::strict()),Err(DigitalSignatureError::Sha1Disallowed)))}
 #[test]fn tamper_is_reported_not_trusted(){let mut p=OpcPackage::from_bytes(DOCX).unwrap();let u=PackURI::new("/word/document.xml").unwrap();let part=p.get_part_mut(&u).unwrap();let mut b=part.blob().to_vec();b.push(b' ');part.set_blob(b);let r=p.verify_digital_signatures(&SignatureVerificationPolicy::compatibility()).unwrap();assert_eq!(r[0].package_integrity,VerificationStatus::Invalid);assert_eq!(r[0].signature_value,VerificationStatus::Valid)}
}
