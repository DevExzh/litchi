//! Bounded ISO/IEC 29500-3:2015 semantic preprocessing.
use quick_xml::{Reader,XmlVersion,encoding::Decoder,events::{BytesStart,Event}};
use std::{borrow::Cow,collections::{BTreeMap,HashSet},str};
use thiserror::Error;
pub const MCE_NAMESPACE:&str="http://schemas.openxmlformats.org/markup-compatibility/2006";
const XML_NS:&str="http://www.w3.org/XML/1998/namespace";
#[derive(Debug,Clone,PartialEq,Eq,Hash)]pub struct ExpandedName{pub namespace:String,pub local_name:String}
#[derive(Debug,Clone)]pub struct MceCapabilities{understood:HashSet<String>,extensions:HashSet<ExpandedName>}
impl MceCapabilities{pub fn new()->Self{Self{understood:HashSet::new(),extensions:HashSet::new()}}pub fn ooxml_baseline()->Self{let mut s=Self::new();for n in ["http://schemas.openxmlformats.org/wordprocessingml/2006/main","http://purl.oclc.org/ooxml/wordprocessingml/main","http://schemas.openxmlformats.org/spreadsheetml/2006/main","http://purl.oclc.org/ooxml/spreadsheetml/main","http://schemas.openxmlformats.org/presentationml/2006/main","http://purl.oclc.org/ooxml/presentationml/main","http://schemas.openxmlformats.org/drawingml/2006/main","http://purl.oclc.org/ooxml/drawingml/main","http://schemas.openxmlformats.org/drawingml/2006/chart","http://purl.oclc.org/ooxml/drawingml/chart","http://schemas.openxmlformats.org/officeDocument/2006/relationships","http://purl.oclc.org/ooxml/officeDocument/relationships","http://schemas.openxmlformats.org/officeDocument/2006/math","http://purl.oclc.org/ooxml/officeDocument/math","urn:schemas-microsoft-com:vml","urn:schemas-microsoft-com:office:office",XML_NS]{s.understood.insert(n.into());}s}pub fn understand_namespace(&mut self,n:impl Into<String>)->&mut Self{self.understood.insert(n.into());self}pub fn preserve_extension_element(&mut self,n:ExpandedName)->&mut Self{self.extensions.insert(n);self}pub fn understands(&self,n:&str)->bool{self.understood.contains(n)}}
impl Default for MceCapabilities{fn default()->Self{Self::ooxml_baseline()}}
#[derive(Debug,Clone)]pub struct MceLimits{pub max_input_bytes:usize,pub max_output_bytes:usize,pub max_depth:usize,pub max_namespace_bindings:usize,pub max_directive_tokens:usize,pub max_choices_per_alternate:usize}
impl Default for MceLimits{fn default()->Self{Self{max_input_bytes:256*1024*1024,max_output_bytes:512*1024*1024,max_depth:256,max_namespace_bindings:4096,max_directive_tokens:4096,max_choices_per_alternate:1024}}}
#[derive(Debug,Clone,Default,PartialEq,Eq)]pub struct MceReport{pub alternate_content_count:usize,pub selected_choices:usize,pub selected_fallbacks:usize,pub ignored_elements:usize,pub ignored_attributes:usize,pub unwrapped_elements:usize}
#[derive(Debug)]pub struct MceOutput<'a>{pub xml:Cow<'a,[u8]>,pub report:MceReport}
#[derive(Debug,Error,Clone,PartialEq,Eq)]pub enum MceError{#[error("non-conformant markup compatibility XML: {0}")]NonConformant(String),#[error("unsupported namespace required by MustUnderstand: {0}")]MustUnderstand(String),#[error("markup compatibility resource limit exceeded: {0}")]LimitExceeded(String),#[error("markup compatibility XML error: {0}")]Xml(String)}
type R<T>=std::result::Result<T,MceError>;
#[derive(Clone)]struct Ctx{ns:BTreeMap<String,String>,ign:HashSet<String>,process:HashSet<ExpandedName>,opaque:bool}
enum Mode{Emit(String),Unwrap,Skip,Alt{choices:usize,selected:bool,fallback:bool},Branch}
struct Frame{ctx:Ctx,mode:Mode,active:bool}
pub fn process_markup_compatibility<'a>(xml:&'a[u8],caps:&MceCapabilities,lim:&MceLimits)->R<MceOutput<'a>>{if !xml.windows(MCE_NAMESPACE.len()).any(|w|w==MCE_NAMESPACE.as_bytes()){return Ok(MceOutput{xml:Cow::Borrowed(xml),report:MceReport::default()})}if xml.len()>lim.max_input_bytes{return Err(limit("input bytes"))}let mut r=Reader::from_reader(xml);r.config_mut().trim_text(false);let(mut stack,mut out,mut rep,mut root,mut buf)=(Vec::new(),Vec::with_capacity(xml.len()),MceReport::default(),false,Vec::new());loop{let d=r.decoder();match r.read_event_into(&mut buf){Ok(Event::Start(e))=>start(&e,d,false,caps,lim,&mut stack,&mut out,&mut rep,&mut root)?,Ok(Event::Empty(e))=>start(&e,d,true,caps,lim,&mut stack,&mut out,&mut rep,&mut root)?,Ok(Event::End(_))=>{let f:Frame=stack.pop().ok_or_else(||bad("unexpected end"))?;match f.mode{Mode::Alt{choices,..}if choices==0=>return Err(bad("AlternateContent requires Choice")),Mode::Emit(q)if f.active=>{out.extend_from_slice(b"</");out.extend_from_slice(q.as_bytes());out.push(b'>')},_=>{}}},Ok(Event::Text(e))=>if visible(&stack){out.extend_from_slice(e.as_ref())},Ok(Event::CData(e))=>if visible(&stack){out.extend_from_slice(b"<![CDATA[");out.extend_from_slice(e.as_ref());out.extend_from_slice(b"]]>" )},Ok(Event::Comment(e))=>if visible(&stack){out.extend_from_slice(b"<!--");out.extend_from_slice(e.as_ref());out.extend_from_slice(b"-->")},Ok(Event::Decl(e))=>if stack.is_empty()&&!root{out.extend_from_slice(b"<?xml ");out.extend_from_slice(e.as_ref());out.extend_from_slice(b"?>")}else{return Err(bad("late XML declaration"))},Ok(Event::GeneralRef(e))=>if visible(&stack){let n=e.decode().map_err(xerr)?;if e.resolve_char_ref().map_err(xerr)?.is_none()&&!matches!(n.as_ref(),"amp"|"lt"|"gt"|"apos"|"quot"){return Err(bad("custom entity"))}out.push(b'&');out.extend_from_slice(e.as_ref());out.push(b';')},Ok(Event::DocType(_)|Event::PI(_))=>return Err(bad("DTD and processing instructions are rejected")),Ok(Event::Eof)=>break,Err(e)=>return Err(MceError::Xml(e.to_string()))}if out.len()>lim.max_output_bytes{return Err(limit("output bytes"))}buf.clear()}if !stack.is_empty(){return Err(bad("unterminated XML"))}Ok(MceOutput{xml:Cow::Owned(out),report:rep})}
fn start(e:&BytesStart<'_>,d:Decoder,empty:bool,caps:&MceCapabilities,lim:&MceLimits,st:&mut Vec<Frame>,out:&mut Vec<u8>,rep:&mut MceReport,root:&mut bool)->R<()>{if st.len()>=lim.max_depth{return Err(limit("depth"))}let q=str::from_utf8(e.name().as_ref()).map_err(xerr)?.to_string();let mut raw=Vec::new();for a in e.attributes().with_checks(true){let a=a.map_err(xerr)?;raw.push((str::from_utf8(a.key.as_ref()).map_err(xerr)?.to_string(),a.decoded_and_normalized_value(XmlVersion::Explicit1_0,d).map_err(xerr)?.into_owned()))}let mut c=st.last().map(|f|f.ctx.clone()).unwrap_or_else(||{let mut ns=BTreeMap::new();ns.insert("xml".into(),XML_NS.into());Ctx{ns,ign:HashSet::new(),process:HashSet::new(),opaque:false}});for(a,v)in &raw{if a=="xmlns"{c.ns.insert("".into(),v.clone());}else if let Some(p)=a.strip_prefix("xmlns:"){if p.is_empty()||v.is_empty(){return Err(bad("invalid namespace"))}c.ns.insert(p.into(),v.clone());}}if c.ns.len()>lim.max_namespace_bindings{return Err(limit("namespace bindings"))}let name=expand(&q,&c.ns,true)?;let parent_active=st.last().is_none_or(|f|f.active);if c.opaque{let f=Frame{ctx:c.clone(),mode:Mode::Emit(q.clone()),active:parent_active};if parent_active{write_start(out,&q,&c.ns,&raw,&c.ign,false,rep)?}return close(st,f,empty,out)}let mut tokens=0;for(a,v)in &raw{if a=="xmlns"||a.starts_with("xmlns:"){continue}let n=expand(a,&c.ns,false)?;if n.namespace!=MCE_NAMESPACE{continue}tokens+=v.split_whitespace().count();match n.local_name.as_str(){"Ignorable"=>for p in v.split_whitespace(){let u=c.ns.get(p).ok_or_else(||bad(format!("unbound Ignorable {p}")))?;if u==MCE_NAMESPACE{return Err(bad("MCE cannot be ignorable"))}c.ign.insert(u.clone());},"ProcessContent"=>for x in v.split_whitespace(){let n=expand(x,&c.ns,true)?;if !c.ign.contains(&n.namespace){return Err(bad("ProcessContent target is not ignorable"))}c.process.insert(n);},"MustUnderstand"=>for p in v.split_whitespace(){let u=c.ns.get(p).ok_or_else(||bad(format!("unbound MustUnderstand {p}")))?;if !caps.understands(u){return Err(MceError::MustUnderstand(u.clone()))}},_=>return Err(bad("unknown MCE attribute"))}}if tokens>lim.max_directive_tokens{return Err(limit("directive tokens"))}c.opaque=caps.extensions.contains(&name);if let Some(parent)=st.last_mut(){if let Mode::Alt{choices,selected,fallback}=&mut parent.mode{if name.namespace!=MCE_NAMESPACE{return Err(bad("non-MCE AlternateContent child"))}let(active,mode)=match name.local_name.as_str(){"Choice"=>{if *fallback{return Err(bad("Choice after Fallback"))}*choices+=1;if *choices>lim.max_choices_per_alternate{return Err(limit("choices"))}let req=attr(&raw,"Requires")?.ok_or_else(||bad("Choice lacks Requires"))?;let mut ok=true;let mut count=0;for p in req.split_whitespace(){count+=1;ok&=caps.understands(c.ns.get(p).ok_or_else(||bad(format!("unbound Requires {p}")))?)}if count==0{return Err(bad("empty Requires"))}let a=parent.active&&!*selected&&ok;if a{*selected=true;rep.selected_choices+=1}(a,Mode::Branch)},"Fallback"=>{if *fallback{return Err(bad("duplicate Fallback"))}*fallback=true;let a=parent.active&&!*selected;if a{*selected=true;rep.selected_fallbacks+=1}(a,Mode::Branch)},_=>return Err(bad("invalid AlternateContent child"))};return close(st,Frame{ctx:c,mode,active},empty,out)}}let mut active=parent_active;let mode=if name.namespace==MCE_NAMESPACE{match name.local_name.as_str(){"AlternateContent"=>{rep.alternate_content_count+=1;Mode::Alt{choices:0,selected:false,fallback:false}},_=>return Err(bad("Choice/Fallback outside AlternateContent"))}}else if c.ign.contains(&name.namespace)&&!caps.understands(&name.namespace){if c.process.contains(&name){for(a,_)in &raw{let n=expand(a,&c.ns,false)?;if n.namespace==XML_NS&&matches!(n.local_name.as_str(),"base"|"lang"|"space"){return Err(bad("xml context attribute on unwrapped element"))}}rep.unwrapped_elements+=1;Mode::Unwrap}else{rep.ignored_elements+=1;active=false;Mode::Skip}}else{Mode::Emit(q.clone())};if matches!(mode,Mode::Emit(_))&&active{write_start(out,&q,&c.ns,&raw,&c.ign,true,rep)?}if st.is_empty(){if *root{return Err(bad("multiple roots"))}*root=true}close(st,Frame{ctx:c,mode,active},empty,out)}
fn close(st:&mut Vec<Frame>,f:Frame,empty:bool,out:&mut Vec<u8>)->R<()>{if empty{match f.mode{Mode::Emit(q)if f.active=>{out.extend_from_slice(b"</");out.extend_from_slice(q.as_bytes());out.push(b'>')},Mode::Alt{..}=>return Err(bad("empty AlternateContent")),_=>{}}}else{st.push(f)}Ok(())}fn visible(s:&[Frame])->bool{s.last().is_some_and(|f|f.active)}fn bad(s:impl Into<String>)->MceError{MceError::NonConformant(s.into())}fn limit(s:&str)->MceError{MceError::LimitExceeded(s.into())}fn xerr(e:impl std::fmt::Display)->MceError{MceError::Xml(e.to_string())}
fn attr<'a>(r:&'a[(String,String)],n:&str)->R<Option<&'a str>>{let mut v=None;for(a,x)in r{if a==n{if v.is_some(){return Err(bad("duplicate attribute"))}v=Some(x.as_str())}}Ok(v)}fn expand(q:&str,ns:&BTreeMap<String,String>,element:bool)->R<ExpandedName>{let(p,l)=q.split_once(':').unwrap_or(("",q));if l.is_empty()||q.matches(':').count()>1{return Err(bad("invalid QName"))}let n=if p.is_empty(){if element{ns.get("").cloned().unwrap_or_default()}else{String::new()}}else{ns.get(p).cloned().ok_or_else(||bad(format!("unbound prefix {p}")))?};Ok(ExpandedName{namespace:n,local_name:l.into()})}
fn write_start(o:&mut Vec<u8>,q:&str,ns:&BTreeMap<String,String>,raw:&[(String,String)],ign:&HashSet<String>,filter:bool,rep:&mut MceReport)->R<()>{o.push(b'<');o.extend_from_slice(q.as_bytes());for(p,u)in ns{if p=="xml"{continue}o.extend_from_slice(if p.is_empty(){b" xmlns"}else{b" xmlns:"});if !p.is_empty(){o.extend_from_slice(p.as_bytes())}o.extend_from_slice(b"=\"");esc(o,u);o.push(b'\"')}for(a,v)in raw{if a=="xmlns"||a.starts_with("xmlns:"){continue}let n=expand(a,ns,false)?;if filter&&(n.namespace==MCE_NAMESPACE||(!n.namespace.is_empty()&&ign.contains(&n.namespace))){rep.ignored_attributes+=1;continue}o.push(b' ');o.extend_from_slice(a.as_bytes());o.extend_from_slice(b"=\"");esc(o,v);o.push(b'\"')}o.push(b'>');Ok(())}fn esc(o:&mut Vec<u8>,s:&str){for c in s.chars(){match c{'&'=>o.extend_from_slice(b"&amp;"),'<'=>o.extend_from_slice(b"&lt;"),'"'=>o.extend_from_slice(b"&quot;"),'\t'=>o.extend_from_slice(b"&#x9;"),'\n'=>o.extend_from_slice(b"&#xA;"),'\r'=>o.extend_from_slice(b"&#xD;"),_=>{let mut b=[0;4];o.extend_from_slice(c.encode_utf8(&mut b).as_bytes())}}}}
pub(crate)fn process_ooxml(x:&[u8])->crate::error::Result<Cow<'_,[u8]>>{process_markup_compatibility(x,&MceCapabilities::default(),&MceLimits::default()).map(|x|x.xml).map_err(crate::error::OoxmlError::MarkupCompatibility)}
pub(crate)fn process_part(part:&dyn litchi_opc::Part)->crate::error::Result<Cow<'_,[u8]>>{process_ooxml(part.blob())}
pub(crate)fn process_part_arc(part:&dyn litchi_opc::Part)->crate::error::Result<std::sync::Arc<Vec<u8>>>{Ok(match process_part(part)?{Cow::Borrowed(_)=>part.blob_arc(),Cow::Owned(v)=>std::sync::Arc::new(v)})}
pub(crate)fn process_str(x:&str)->crate::error::Result<Cow<'_,str>>{match process_ooxml(x.as_bytes())?{Cow::Borrowed(_)=>Ok(Cow::Borrowed(x)),Cow::Owned(v)=>String::from_utf8(v).map(Cow::Owned).map_err(|e|crate::error::OoxmlError::InvalidFormat(format!("MCE output is not UTF-8: {e}")))}}
#[cfg(test)]mod tests{use super::*;fn run(x:&str,c:&MceCapabilities)->R<String>{Ok(String::from_utf8(process_markup_compatibility(x.as_bytes(),c,&MceLimits::default())?.xml.into_owned()).unwrap())}#[test]fn fast_borrowed(){assert!(matches!(process_markup_compatibility(b"<r/>",&MceCapabilities::new(),&MceLimits::default()).unwrap().xml,Cow::Borrowed(_)))}#[test]fn choice_fallback(){let x=r#"<r xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:a="urn:a"><mc:AlternateContent><mc:Choice Requires="a"><yes/></mc:Choice><mc:Fallback><no/></mc:Fallback></mc:AlternateContent></r>"#;let mut c=MceCapabilities::new();assert!(run(x,&c).unwrap().contains("<no"));c.understand_namespace("urn:a");assert!(run(x,&c).unwrap().contains("<yes"))}#[test]fn ignore_and_unwrap(){let x=r#"<r xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:x" mc:Ignorable="x" mc:ProcessContent="x:w"><x:no/><x:w><yes/></x:w></r>"#;let y=run(x,&MceCapabilities::new()).unwrap();assert!(!y.contains("<x:"));assert!(y.contains("<yes"))}#[test]fn security_and_limits(){let x=r#"<r xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:x" mc:MustUnderstand="x"/>"#;assert!(matches!(run(x,&MceCapabilities::new()),Err(MceError::MustUnderstand(_))));let mut l=MceLimits::default();l.max_depth=1;assert!(process_markup_compatibility(b"<r xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\"><x/></r>",&MceCapabilities::new(),&l).is_err())}}

#[cfg(test)]
mod fixture_tests {
    use super::*;
    use litchi_opc::{OpcPackage, PackURI};

    #[test]
    fn poi_styles_select_unsupported_vendor_fallbacks() {
        let package = OpcPackage::from_bytes(include_bytes!("../../../../3rdparty/poi/test-data/spreadsheet/style-alternate-content.xlsx")).unwrap();
        let part = package.get_part(&PackURI::new("/xl/styles.xml").unwrap()).unwrap();
        let output = process_markup_compatibility(part.blob(), &MceCapabilities::default(), &MceLimits::default()).unwrap();
        let xml = std::str::from_utf8(output.xml.as_ref()).unwrap();
        assert!(!xml.contains("mc:AlternateContent"));
        assert!(!xml.contains("hs:extension"));
        assert!(output.report.selected_fallbacks > 10);
    }

    #[test]
    fn libreoffice_pptx_emits_only_fallback_shape() {
        let package = OpcPackage::from_bytes(include_bytes!("../../../../3rdparty/libreoffice-core/oox/qa/unit/data/import-mce.pptx")).unwrap();
        let part = package.get_part(&PackURI::new("/ppt/slides/slide1.xml").unwrap()).unwrap();
        let output = process_markup_compatibility(part.blob(), &MceCapabilities::default(), &MceLimits::default()).unwrap();
        let xml = std::str::from_utf8(output.xml.as_ref()).unwrap();
        assert!(!xml.contains("mc:AlternateContent"));
        assert!(!xml.contains("a14:m"));
        assert!(xml.contains("a:blipFill"));
        assert_eq!(output.report.selected_fallbacks, 1);
    }
}

#[cfg(test)]
mod adapter_tests {
    use crate::docx::enums::WdHeaderFooter;
    use crate::docx::header_footer::HeaderFooter;
    use crate::xlsx::SharedStrings;

    #[test]
    fn docx_header_uses_fallback_without_mutating_raw_xml() {
        let raw = br#"<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:unsupported"><mc:AlternateContent><mc:Choice Requires="x"><w:p><w:r><w:t>choice</w:t></w:r></w:p></mc:Choice><mc:Fallback><w:p><w:r><w:t>fallback</w:t></w:r></w:p></mc:Fallback></mc:AlternateContent></w:hdr>"#;
        let header = HeaderFooter::from_xml_bytes(raw.to_vec(), WdHeaderFooter::Primary);
        assert_eq!(header.xml_bytes(), raw);
        assert_eq!(header.text().unwrap(), "fallback");
        assert_eq!(header.paragraph_count().unwrap(), 1);
    }

    #[test]
    fn xlsx_shared_strings_uses_fallback() {
        let xml = r#"<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:unsupported" count="1" uniqueCount="1"><mc:AlternateContent><mc:Choice Requires="x"><si><t>choice</t></si></mc:Choice><mc:Fallback><si><t>fallback</t></si></mc:Fallback></mc:AlternateContent></sst>"#;
        let strings = SharedStrings::parse(xml).unwrap();
        assert_eq!(strings.get(0), Some("fallback"));
    }

    #[test]
    fn generic_chart_reader_uses_fallback() {
        let xml = br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:unsupported"><mc:AlternateContent><mc:Choice Requires="x"><x:chart/></mc:Choice><mc:Fallback><c:chart/></mc:Fallback></mc:AlternateContent></c:chartSpace>"#;
        crate::charts::reader::parse_chart(xml.as_slice()).unwrap();
    }

    #[test]
    fn alternate_content_picture_selects_fallback() {
        let xml = br#"<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:unsupported"><mc:AlternateContent><mc:Choice Requires="x"><x:picture/></mc:Choice><mc:Fallback><w:pict><w:t>fallback-picture</w:t></w:pict></mc:Fallback></mc:AlternateContent></w:r>"#;
        let output = super::process_ooxml(xml).unwrap();
        let semantic = std::str::from_utf8(output.as_ref()).unwrap();
        assert!(semantic.contains("w:pict"));
        assert!(!semantic.contains("x:picture"));
        assert!(!semantic.contains("AlternateContent"));
    }
}
