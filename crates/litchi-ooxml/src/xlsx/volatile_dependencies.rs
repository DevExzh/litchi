//! Strict, inert SpreadsheetML volatile-dependency records.

use litchi_core::sheet::Result;
use litchi_opc::{OpcPackage, PackURI};
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesEnd, BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use quick_xml::Writer;
use quick_xml::XmlVersion;

const NS: &[u8] = b"http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const STRICT_NS: &[u8] = b"http://purl.oclc.org/ooxml/spreadsheetml/main";
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/volatileDependencies";
const STRICT_REL: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships/volatileDependencies";
const CONTENT_TYPE: &str = "application/vnd.openxmlformats-officedocument.spreadsheetml.volatileDependencies+xml";
const MAX_PART_BYTES: usize = 8 * 1024 * 1024;
const MAX_TYPES: usize = 64;
const MAX_MAINS: usize = 16_384;
const MAX_TOPICS: usize = 65_536;
const MAX_SUBTOPICS: usize = 262_144;
const MAX_REFERENCES: usize = 1_048_576;
const MAX_TEXT_BYTES: usize = 1_048_576;

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum VolatileDependencyType { RealTimeData, OlapFunctions }

#[derive(Clone, Debug, PartialEq)]
pub enum VolatileValue {
    Unspecified(String), Boolean(bool), Number(f64), Error(String), String(String),
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct VolatileReference { pub cell_reference: String, pub sheet_id: u32 }

#[derive(Clone, Debug, PartialEq)]
pub struct VolatileTopic {
    pub value: VolatileValue,
    pub subtopics: Vec<String>,
    pub references: Vec<VolatileReference>,
}

#[derive(Clone, Debug, PartialEq)]
pub struct VolatileMain { pub first: String, pub topics: Vec<VolatileTopic> }

#[derive(Clone, Debug, PartialEq)]
pub struct VolatileType { pub dependency_type: VolatileDependencyType, pub mains: Vec<VolatileMain> }

#[derive(Clone, Debug, PartialEq)]
pub struct VolatileDependencies {
    pub types: Vec<VolatileType>,
    /// Raw `extLst` markup. Its payload is preserved, never interpreted or executed.
    pub extension_list_xml: Option<Vec<u8>>,
}

impl VolatileDependencies {
    pub fn parse(xml: &[u8]) -> Result<Self> {
        if xml.len() > MAX_PART_BYTES { return Err(invalid("volatile-dependencies part exceeds 8 MiB")); }
        let processed = crate::common::mce::process_ooxml(xml)?;
        if processed.len() > MAX_PART_BYTES { return Err(invalid("processed volatile-dependencies part exceeds 8 MiB")); }
        parse_processed(processed.as_ref())
    }

    pub fn to_xml(&self, strict: bool) -> Result<Vec<u8>> {
        validate_document(self)?;
        let ns = if strict { std::str::from_utf8(STRICT_NS).unwrap() } else { std::str::from_utf8(NS).unwrap() };
        let mut out = String::from("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><volTypes xmlns=\"");
        escape_attr(&mut out, ns); out.push_str("\">");
        for ty in &self.types {
            out.push_str("<volType type=\""); out.push_str(match ty.dependency_type { VolatileDependencyType::RealTimeData => "realTimeData", VolatileDependencyType::OlapFunctions => "olapFunctions" }); out.push_str("\">");
            for main in &ty.mains {
                out.push_str("<main first=\""); escape_attr(&mut out, &main.first); out.push_str("\">");
                for topic in &main.topics {
                    out.push_str("<tp");
                    let number = match &topic.value { VolatileValue::Number(v) => Some(v.to_string()), _ => None };
                    let (kind, value) = match &topic.value {
                        VolatileValue::Unspecified(v) => (None, v.as_str()), VolatileValue::Boolean(v) => (Some("b"), if *v { "1" } else { "0" }),
                        VolatileValue::Number(_) => (Some("n"), number.as_deref().unwrap()), VolatileValue::Error(v) => (Some("e"), v.as_str()), VolatileValue::String(v) => (Some("s"), v.as_str()),
                    };
                    if let Some(kind) = kind { out.push_str(" t=\""); out.push_str(kind); out.push('"'); }
                    out.push_str("><v>"); escape_text(&mut out, value); out.push_str("</v>");
                    for value in &topic.subtopics { out.push_str("<stp>"); escape_text(&mut out, value); out.push_str("</stp>"); }
                    for reference in &topic.references { out.push_str("<tr r=\""); escape_attr(&mut out, &reference.cell_reference); out.push_str("\" s=\""); out.push_str(&reference.sheet_id.to_string()); out.push_str("\"/>"); }
                    out.push_str("</tp>");
                }
                out.push_str("</main>");
            }
            out.push_str("</volType>");
        }
        let mut bytes = out.into_bytes();
        if let Some(ext) = &self.extension_list_xml { bytes.extend_from_slice(ext); }
        bytes.extend_from_slice(b"</volTypes>");
        if bytes.len() > MAX_PART_BYTES { return Err(invalid("serialized volatile-dependencies part exceeds 8 MiB")); }
        Ok(bytes)
    }
}

/// Loads the single volatile-dependencies part related to the package workbook.
pub fn load_from_package(package: &OpcPackage) -> Result<Option<VolatileDependencies>> {
    let workbook = package.main_document_part()?;
    let mut matches = workbook.rels().iter().filter(|r| matches!(r.reltype(), REL | STRICT_REL));
    let Some(relationship) = matches.next() else { return Ok(None); };
    if matches.next().is_some() { return Err(invalid("workbook has multiple volatile-dependencies relationships")); }
    if relationship.is_external() { return Err(invalid("volatile-dependencies relationship cannot be external")); }
    let uri: PackURI = relationship.target_partname()?;
    let part = package.get_part(&uri)?;
    if part.content_type() != CONTENT_TYPE { return Err(invalid(format!("volatile-dependencies part '{uri}' has invalid content type '{}'", part.content_type()))); }
    if part.rels().iter().next().is_some() { return Err(invalid("volatile-dependencies part must not have relationships")); }
    if part.blob().len() > MAX_PART_BYTES { return Err(invalid("volatile-dependencies part exceeds 8 MiB")); }
    Ok(Some(VolatileDependencies::parse(part.blob())?))
}

#[derive(Clone, Copy)] enum Context { Root, Type(usize), Main(usize, usize), Topic(usize, usize, usize), Value(usize, usize, usize), Subtopic(usize, usize, usize), Reference }
#[derive(Default)] struct TopicBuilder { kind: Option<u8>, value: Option<String>, subtopics: Vec<String>, references: Vec<VolatileReference> }
struct MainBuilder { first: String, topics: Vec<TopicBuilder> }
struct TypeBuilder { dependency_type: VolatileDependencyType, mains: Vec<MainBuilder> }

fn parse_processed(xml: &[u8]) -> Result<VolatileDependencies> {
    let mut reader = NsReader::from_reader(xml);
    let mut stack = Vec::new(); let mut types: Vec<TypeBuilder> = Vec::new(); let mut extension = None; let mut root_prefixes = Vec::new();
    let mut root_closed = false; let mut capture: Option<(usize, Writer<Vec<u8>>)> = None;
    loop {
        let decoder = reader.decoder(); let event = reader.read_event()?.into_owned();
        if let Some((depth, writer)) = capture.as_mut() {
            match &event { Event::Start(_) => { if *depth >= 256 { return Err(invalid("extension-list depth limit exceeded")); } *depth += 1; }, Event::End(_) => *depth -= 1, Event::DocType(_) | Event::PI(_) => return Err(invalid("DTD and processing instructions are rejected")), _ => {} }
            if *depth == 0 { writer.write_event(Event::End(BytesEnd::new("extLst"))).map_err(xml_error)?; } else { writer.write_event(event.clone()).map_err(xml_error)?; }
            if *depth == 0 { extension = Some(capture.take().unwrap().1.into_inner()); }
            continue;
        }
        let resolver = reader.resolver().clone(); let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(e) if stack.is_empty() => {
                if root_closed || !name(&namespace, &e, b"volTypes") { return Err(invalid("expected one SpreadsheetML volTypes root")); }
                root_prefixes = namespace_attributes(&e, decoder)?; no_attributes(&e)?; stack.push(Context::Root);
            }
            Event::Start(e) if matches!(stack.last(), Some(Context::Root)) && name(&namespace, &e, b"extLst") => {
                if extension.is_some() || types.is_empty() { return Err(invalid("invalid extLst order or duplicate")); }
                no_attributes(&e)?;
                let mut bindings = root_prefixes.clone(); for (key,value) in namespace_attributes(&e, decoder)? { if let Some(binding)=bindings.iter_mut().find(|binding|binding.0==key){binding.1=value;}else{bindings.push((key,value));} } bindings.sort_by(|a,b|a.0.cmp(&b.0));
                let mut wrapper = BytesStart::new("extLst"); for (key,value) in &bindings { wrapper.push_attribute((key.as_str(),value.as_str())); }
                let mut writer=Writer::new(Vec::new()); writer.write_event(Event::Start(wrapper)).map_err(xml_error)?; capture=Some((1,writer));
            }
            Event::Start(e) => start(&mut stack, &mut types, extension.is_some(), &namespace, e, decoder)?,
            Event::Empty(e) => empty(&mut stack, &mut types, &mut extension, &namespace, e, decoder)?,
            Event::Text(t) => push_text(&mut types, stack.last().copied(), &t.decode().map_err(xml_error)?)?,
            Event::CData(t) => push_text(&mut types, stack.last().copied(), &t.decode().map_err(xml_error)?)?,
            Event::GeneralRef(r) => push_text(&mut types, stack.last().copied(), &crate::common::xml::decode_xml_reference(&r)?)?,
            Event::End(_) => { let ended = stack.pop().ok_or_else(|| invalid("closing element outside volTypes"))?; if matches!(ended, Context::Root) { root_closed = true; } }
            Event::DocType(_) | Event::PI(_) => return Err(invalid("DTD and processing instructions are rejected")),
            Event::Eof => break,
            Event::Decl(_) | Event::Comment(_) => {},
        }
        if let Some(Context::Root) = stack.last() {
            if extension.is_none() && reader.buffer_position() > 0 {
                // capture is initiated in `start`; this branch intentionally has no behavior.
            }
        }
    }
    if !root_closed || !stack.is_empty() { return Err(invalid("unterminated volatile-dependencies XML")); }
    let mut result = VolatileDependencies { types: Vec::with_capacity(types.len()), extension_list_xml: extension };
    for ty in types { let mut output = VolatileType { dependency_type: ty.dependency_type, mains: Vec::with_capacity(ty.mains.len()) }; for main in ty.mains { let mut converted = VolatileMain { first: main.first, topics: Vec::with_capacity(main.topics.len()) }; for topic in main.topics { let raw = topic.value.ok_or_else(|| invalid("tp requires one v child"))?; let value = parse_value(topic.kind, raw)?; converted.topics.push(VolatileTopic { value, subtopics: topic.subtopics, references: topic.references }); } output.mains.push(converted); } result.types.push(output); }
    validate_document(&result)?; Ok(result)
}

fn start(stack: &mut Vec<Context>, types: &mut Vec<TypeBuilder>, extension_seen: bool, ns: &ResolveResult, e: BytesStart<'static>, decoder: Decoder) -> Result<()> {
    match stack.last().copied().ok_or_else(|| invalid("element outside volTypes"))? {
        Context::Root if name(ns, &e, b"volType") => { if extension_seen || types.len() >= MAX_TYPES { return Err(invalid("invalid volType order or limit")); } let value = required_attr(&e, decoder, b"type")?; let dependency_type = match value.as_str() { "realTimeData" => VolatileDependencyType::RealTimeData, "olapFunctions" => VolatileDependencyType::OlapFunctions, _ => return Err(invalid("invalid volatile dependency type")) }; only_attrs(&e, &[b"type"])?; types.push(TypeBuilder { dependency_type, mains: Vec::new() }); stack.push(Context::Type(types.len()-1)); }
        Context::Type(t) if name(ns, &e, b"main") => { if total_mains(types) >= MAX_MAINS { return Err(invalid("main-topic limit exceeded")); } let first=required_attr(&e,decoder,b"first")?; bounded(&first)?; only_attrs(&e,&[b"first"])?; types[t].mains.push(MainBuilder{first,topics:Vec::new()}); stack.push(Context::Main(t,types[t].mains.len()-1)); }
        Context::Main(t,m) if name(ns,&e,b"tp") => { if total_topics(types)>=MAX_TOPICS{return Err(invalid("topic limit exceeded"));} let kind=optional_attr(&e,decoder,b"t")?.map(|v|match v.as_str(){"b"=>Ok(b'b'),"n"=>Ok(b'n'),"e"=>Ok(b'e'),"s"=>Ok(b's'),_=>Err(invalid("invalid volatile value type"))}).transpose()?;only_attrs(&e,&[b"t"])?;types[t].mains[m].topics.push(TopicBuilder{kind,..Default::default()});stack.push(Context::Topic(t,m,types[t].mains[m].topics.len()-1)); }
        Context::Topic(t,m,p) if name(ns,&e,b"v") => { let topic=&mut types[t].mains[m].topics[p];if topic.value.is_some()||!topic.subtopics.is_empty()||!topic.references.is_empty(){return Err(invalid("v must be the first and only value child"));}no_attributes(&e)?;topic.value=Some(String::new());stack.push(Context::Value(t,m,p)); }
        Context::Topic(t,m,p) if name(ns,&e,b"stp") => { let limit_reached=total_subtopics(types)>=MAX_SUBTOPICS;let topic=&mut types[t].mains[m].topics[p];if topic.value.is_none()||!topic.references.is_empty()||limit_reached{return Err(invalid("invalid subtopic order or limit"));}no_attributes(&e)?;topic.subtopics.push(String::new());stack.push(Context::Subtopic(t,m,p)); }
        Context::Topic(t,m,p) if name(ns,&e,b"tr") => { add_reference(types,t,m,p,&e,decoder)?;stack.push(Context::Reference); }
        _ => return Err(invalid("unexpected volatile-dependencies element")),
    } Ok(())
}

fn empty(stack:&mut Vec<Context>,types:&mut Vec<TypeBuilder>,extension:&mut Option<Vec<u8>>,ns:&ResolveResult,e:BytesStart<'static>,decoder:Decoder)->Result<()> {
    match stack.last().copied() {
        Some(Context::Topic(t,m,p)) if name(ns,&e,b"v") => { let topic=&mut types[t].mains[m].topics[p];if topic.value.is_some()||!topic.subtopics.is_empty()||!topic.references.is_empty(){return Err(invalid("invalid v order"));}no_attributes(&e)?;topic.value=Some(String::new()); }
        Some(Context::Topic(t,m,p)) if name(ns,&e,b"stp") => { if types[t].mains[m].topics[p].value.is_none()||!types[t].mains[m].topics[p].references.is_empty()||total_subtopics(types)>=MAX_SUBTOPICS{return Err(invalid("invalid subtopic order or limit"));}no_attributes(&e)?;types[t].mains[m].topics[p].subtopics.push(String::new()); }
        Some(Context::Topic(t,m,p)) if name(ns,&e,b"tr") => add_reference(types,t,m,p,&e,decoder)?,
        Some(Context::Root) if name(ns,&e,b"extLst") => { if extension.is_some()||types.is_empty(){return Err(invalid("invalid extLst order or duplicate"));}no_attributes(&e)?;*extension=Some(b"<extLst/>".to_vec()); }
        _ => return Err(invalid("unexpected empty volatile-dependencies element")),
    } Ok(())
}

fn add_reference(types:&mut [TypeBuilder],t:usize,m:usize,p:usize,e:&BytesStart<'_>,decoder:Decoder)->Result<()> { if total_references(types)>=MAX_REFERENCES{return Err(invalid("reference limit exceeded"));}let r=required_attr(e,decoder,b"r")?;bounded(&r)?;if !valid_cell_reference(&r){return Err(invalid("invalid volatile cell reference"));}let s=required_attr(e,decoder,b"s")?.parse::<u32>().map_err(|_|invalid("invalid volatile sheet id"))?;only_attrs(e,&[b"r",b"s"])?;let topic=&mut types[t].mains[m].topics[p];if topic.value.is_none(){return Err(invalid("tr must follow v"));}topic.references.push(VolatileReference{cell_reference:r,sheet_id:s});Ok(()) }
fn push_text(types:&mut[TypeBuilder],ctx:Option<Context>,text:&str)->Result<()> { match ctx { Some(Context::Value(t,m,p))=>append(&mut types[t].mains[m].topics[p].value.as_mut().unwrap(),text),Some(Context::Subtopic(t,m,p))=>append(types[t].mains[m].topics[p].subtopics.last_mut().unwrap(),text),_ if text.trim().is_empty()=>Ok(()),_=>Err(invalid("text outside v or stp")) } }
fn append(out:&mut String,text:&str)->Result<()> { if out.len().saturating_add(text.len())>MAX_TEXT_BYTES{return Err(invalid("volatile text limit exceeded"));}out.push_str(text);Ok(()) }

fn parse_value(kind:Option<u8>,raw:String)->Result<VolatileValue>{Ok(match kind{None=>VolatileValue::Unspecified(raw),Some(b'b')=>VolatileValue::Boolean(match raw.as_str(){"1"|"true"=>true,"0"|"false"=>false,_=>return Err(invalid("invalid volatile boolean"))}),Some(b'n')=>{let v=raw.parse::<f64>().map_err(|_|invalid("invalid volatile number"))?;if !v.is_finite(){return Err(invalid("non-finite volatile number"));}VolatileValue::Number(v)},Some(b'e')=>VolatileValue::Error(raw),Some(b's')=>VolatileValue::String(raw),_=>unreachable!()})}
fn validate_document(d:&VolatileDependencies)->Result<()> { if d.types.is_empty()||d.types.len()>MAX_TYPES{return Err(invalid("volTypes requires 1..64 volType children"));}let mut mains=0;let mut topics=0;let mut subs=0;let mut refs=0;for ty in &d.types{if ty.mains.is_empty(){return Err(invalid("volType requires a main child"));}for main in &ty.mains{bounded(&main.first)?;if main.topics.is_empty(){return Err(invalid("main requires a tp child"));}for topic in &main.topics{match &topic.value{VolatileValue::Unspecified(v)|VolatileValue::Error(v)|VolatileValue::String(v)=>bounded(v)?,VolatileValue::Number(v)if !v.is_finite()=>return Err(invalid("non-finite volatile number")),_=>{}}if topic.references.is_empty(){return Err(invalid("tp requires a tr child"));}for v in &topic.subtopics{bounded(v)?;}for r in &topic.references{bounded(&r.cell_reference)?;if !valid_cell_reference(&r.cell_reference){return Err(invalid("invalid volatile cell reference"));}}subs+=topic.subtopics.len();refs+=topic.references.len();}topics+=main.topics.len();}mains+=ty.mains.len();}if mains>MAX_MAINS||topics>MAX_TOPICS||subs>MAX_SUBTOPICS||refs>MAX_REFERENCES{return Err(invalid("volatile-dependencies resource limit exceeded"));}if let Some(ext)=&d.extension_list_xml{if ext.len()>MAX_PART_BYTES{return Err(invalid("extension list exceeds resource limit"));}}Ok(())}

fn name(ns:&ResolveResult,e:&BytesStart<'_>,local:&[u8])->bool{let namespace_matches=match ns{ResolveResult::Bound(Namespace(v))=>{let bytes:&[u8]=v.as_ref();bytes==NS||bytes==STRICT_NS},_=>false};namespace_matches&&e.local_name().as_ref()==local}
fn required_attr(e:&BytesStart<'_>,d:Decoder,n:&[u8])->Result<String>{optional_attr(e,d,n)?.ok_or_else(||invalid(format!("missing required attribute '{}'",String::from_utf8_lossy(n))))}
fn optional_attr(e:&BytesStart<'_>,d:Decoder,n:&[u8])->Result<Option<String>>{let mut value=None;for a in e.attributes().with_checks(true){let a=a.map_err(xml_error)?;if a.key.as_ref()==n{if value.is_some(){return Err(invalid("duplicate attribute"));}value=Some(a.decoded_and_normalized_value(XmlVersion::Implicit1_0,d).map_err(xml_error)?.into_owned());}}Ok(value)}
fn only_attrs(e:&BytesStart<'_>,allowed:&[&[u8]])->Result<()> {for a in e.attributes().with_checks(true){let a=a.map_err(xml_error)?;let k=a.key.as_ref();if k==b"xmlns"||k.starts_with(b"xmlns:"){continue;}if k.contains(&b':')||!allowed.contains(&k){return Err(invalid(format!("unexpected attribute '{}'",String::from_utf8_lossy(k))));}}Ok(())}
fn namespace_attributes(e:&BytesStart<'_>,d:Decoder)->Result<Vec<(String,String)>>{let mut values=Vec::new();for a in e.attributes().with_checks(true){let a=a.map_err(xml_error)?;let key=std::str::from_utf8(a.key.as_ref()).map_err(xml_error)?;if key.starts_with("xmlns:")&&key!="xmlns:xml"{values.push((key.to_string(),a.decoded_and_normalized_value(XmlVersion::Implicit1_0,d).map_err(xml_error)?.into_owned()));}}Ok(values)}
fn no_attributes(e:&BytesStart<'_>)->Result<()>{only_attrs(e,&[])}
fn bounded(v:&str)->Result<()>{if v.len()>MAX_TEXT_BYTES{Err(invalid("volatile string exceeds 1 MiB"))}else{Ok(())}}
fn total_mains(v:&[TypeBuilder])->usize{v.iter().map(|x|x.mains.len()).sum()}fn total_topics(v:&[TypeBuilder])->usize{v.iter().flat_map(|x|&x.mains).map(|x|x.topics.len()).sum()}fn total_subtopics(v:&[TypeBuilder])->usize{v.iter().flat_map(|x|&x.mains).flat_map(|x|&x.topics).map(|x|x.subtopics.len()).sum()}fn total_references(v:&[TypeBuilder])->usize{v.iter().flat_map(|x|&x.mains).flat_map(|x|&x.topics).map(|x|x.references.len()).sum()}
fn valid_cell_reference(v:&str)->bool{let b=v.as_bytes();let mut i=0;if b.get(i)==Some(&b'$'){i+=1;}let start=i;while b.get(i).is_some_and(u8::is_ascii_alphabetic){i+=1;}if i==start||i-start>3{return false;}if b.get(i)==Some(&b'$'){i+=1;}let row=i;while b.get(i).is_some_and(u8::is_ascii_digit){i+=1;}i==b.len()&&i>row&&b[row]!=b'0'}
fn escape_attr(o:&mut String,v:&str){for c in v.chars(){match c{'&'=>o.push_str("&amp;"),'<'=>o.push_str("&lt;"),'"'=>o.push_str("&quot;"),'\r'=>o.push_str("&#xD;"),'\n'=>o.push_str("&#xA;"),'\t'=>o.push_str("&#x9;"),_=>o.push(c)}}}fn escape_text(o:&mut String,v:&str){for c in v.chars(){match c{'&'=>o.push_str("&amp;"),'<'=>o.push_str("&lt;"),'>'=>o.push_str("&gt;"),_=>o.push(c)}}}
fn invalid(message:impl Into<String>)->Box<dyn std::error::Error+Send+Sync>{std::io::Error::new(std::io::ErrorKind::InvalidData,message.into()).into()}
fn xml_error(e:impl std::fmt::Display)->Box<dyn std::error::Error+Send+Sync>{invalid(e.to_string())}

#[cfg(test)] mod tests {
    use super::*; use litchi_opc::{BlobPart, Part};
    fn sample(ns:&str)->Vec<u8>{format!(r#"<volTypes xmlns="{ns}"><volType type="realTimeData"><main first="server.id"><tp t="s"><v>ready</v><stp>ticker</stp><tr r="$A$1" s="0"/></tp></main></volType><volType type="olapFunctions"><main first="cube"><tp t="n"><v>42.5</v><tr r="B2" s="1"/></tp></main></volType></volTypes>"#).into_bytes()}
    #[test] fn parses_and_writes_transitional_and_strict(){for ns in [std::str::from_utf8(NS).unwrap(),std::str::from_utf8(STRICT_NS).unwrap()]{let parsed=VolatileDependencies::parse(&sample(ns)).unwrap();assert_eq!(parsed.types.len(),2);let strict=parsed.to_xml(true).unwrap();assert_eq!(VolatileDependencies::parse(&strict).unwrap(),parsed);}}
    #[test] fn applies_mce_fallback(){let xml=format!(r#"<volTypes xmlns="{}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:u="urn:future" mc:Ignorable="u"><mc:AlternateContent><mc:Choice Requires="u"><u:item/></mc:Choice><mc:Fallback><volType type="realTimeData"><main first="srv"><tp t="b"><v>true</v><tr r="A1" s="0"/></tp></main></volType></mc:Fallback></mc:AlternateContent></volTypes>"#,std::str::from_utf8(NS).unwrap());assert_eq!(VolatileDependencies::parse(xml.as_bytes()).unwrap().types.len(),1);}
    #[test] fn preserves_non_empty_extension_list_and_inherited_namespaces(){let xml=format!(r#"<volTypes xmlns="{}" xmlns:x14="urn:shadowed" xmlns:x15="urn:inherited"><volType type="realTimeData"><main first="srv"><tp><v>ok</v><tr r="A1" s="0"/></tp></main></volType><extLst xmlns:x14="urn:local"><ext uri="urn:test"><x14:payload value="kept"/><x15:payload/></ext></extLst></volTypes>"#,std::str::from_utf8(NS).unwrap());let parsed=VolatileDependencies::parse(xml.as_bytes()).unwrap();let written=parsed.to_xml(true).unwrap();let text=std::str::from_utf8(&written).unwrap();assert!(text.contains("xmlns:x14=\"urn:local\""));assert!(text.contains("xmlns:x15=\"urn:inherited\""));assert!(!text.contains("urn:shadowed"));assert!(text.contains("x14:payload"));assert_eq!(VolatileDependencies::parse(&written).unwrap(),parsed);}
    #[test] fn rejects_malformed_and_unsafe_input(){for xml in [format!(r#"<volTypes xmlns="{}"/>"#,std::str::from_utf8(NS).unwrap()),format!(r#"<volTypes xmlns="{}"><volType type="bad"><main first="x"><tp><v/><tr r="A1" s="0"/></tp></main></volType></volTypes>"#,std::str::from_utf8(NS).unwrap()),format!(r#"<volTypes xmlns="{}"><volType type="realTimeData"><main first="x"><tp t="b"><v>maybe</v><tr r="A1" s="0"/></tp></main></volType></volTypes>"#,std::str::from_utf8(NS).unwrap()),format!(r#"<!DOCTYPE x [<!ENTITY e "boom">]><volTypes xmlns="{}"><volType type="realTimeData"><main first="x"><tp><v>&e;</v><tr r="A1" s="0"/></tp></main></volType></volTypes>"#,std::str::from_utf8(NS).unwrap())]{assert!(VolatileDependencies::parse(xml.as_bytes()).is_err(),"accepted {xml}");}}
    #[test] fn resolves_package_relationship_and_rejects_outbound_relationships(){let mut package=OpcPackage::new();let workbook_uri=PackURI::new("/custom/book.xml").unwrap();let mut workbook=BlobPart::new(workbook_uri,"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml".into(),Vec::new());workbook.relate_to("deps.xml",REL);package.relate_to("custom/book.xml","http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument");package.add_part(Box::new(workbook));package.add_part(Box::new(BlobPart::new(PackURI::new("/custom/deps.xml").unwrap(),CONTENT_TYPE.into(),sample(std::str::from_utf8(NS).unwrap()))));assert_eq!(load_from_package(&package).unwrap().unwrap().types.len(),2);package.get_part_mut(&PackURI::new("/custom/deps.xml").unwrap()).unwrap().relate_to("other.xml","urn:forbidden");assert!(load_from_package(&package).is_err());}
}
