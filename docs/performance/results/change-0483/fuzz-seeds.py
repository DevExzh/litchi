#!/usr/bin/env python3
"""Generate the legacy and settings/MCE DOCX fuzz seed records.

The original 30 main-story seeds and their manifest/generator record are an
immutable historical input.  ``--extend-settings`` authenticates those bytes,
adds a separate deterministic settings corpus to the same seed directory, and
writes only ``seed-manifest-v2.json`` and ``generator-v2.json``.  It never
rewrites the legacy files.
"""
import argparse
import io
import json
import hashlib
from pathlib import Path
import sys
import zipfile
import zlib

ROOT = Path(__file__).resolve().parent
WORD = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
STRICT = 'http://purl.oclc.org/ooxml/wordprocessingml/main'
REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument'
STRICT_REL = 'http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument'
SETTINGS_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings'
STRICT_SETTINGS_REL = 'http://purl.oclc.org/ooxml/officeDocument/relationships/settings'
MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006'
W14 = 'http://schemas.microsoft.com/office/word/2010/wordml'
W15 = 'http://schemas.microsoft.com/office/word/2012/wordml'
TYPES = ('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
         '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
         '<Default Extension="bin" ContentType="application/octet-stream"/>'
         '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
         '</Types>')


def document(body, namespace=WORD, prefix='w'):
    q = prefix + ':' if prefix else ''
    declaration = 'xmlns:' + prefix if prefix else 'xmlns'
    return f'<{q}document {declaration}="{namespace}"><{q}body>{body}</{q}body></{q}document>'


def package(xml, method, strict=False):
    out = io.BytesIO()
    rel = STRICT_REL if strict else REL
    rels = ('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            f'<Relationship Id="rId1" Type="{rel}" Target="word/document.xml"/>'
            '</Relationships>')
    with zipfile.ZipFile(out, 'w') as archive:
        for name, data in [('[Content_Types].xml', TYPES.encode()),
                           ('_rels/.rels', rels.encode()),
                           ('word/document.xml', xml.encode()),
                           ('word/opaque.bin', bytes(range(256)))]:
            entry = zipfile.ZipInfo(name, (1980, 1, 1, 0, 0, 0))
            entry.compress_type = method
            entry.external_attr = 0o600 << 16
            archive.writestr(entry, data)
    return out.getvalue()


SETTINGS_CASES = (
    'empty-settings',
    'protection-disabled',
    'protection-enabled',
    'duplicate-flags',
    'mce-fallback',
    'must-understand',
    'inherited-namespace',
    'directive-amplification',
)
SETTINGS_DIALECTS = (('word', False), ('strict', True))
SETTINGS_METHODS = ((zipfile.ZIP_STORED, 'stored'), (zipfile.ZIP_DEFLATED, 'deflate'))


def settings_body(case):
    """Return a small bounded settings/MCE case and its root MCE attributes."""

    if case == 'empty-settings':
        return '', ''
    if case == 'protection-disabled':
        return '<w:documentProtection w:edit="readOnly" w:enforcement="0"/>', ''
    if case == 'protection-enabled':
        return ('<w:documentProtection w:edit="readOnly" w:enforcement="1" '
                'w:cryptProviderType="rsaFull" w:cryptAlgorithmClass="hash"/>'), ''
    if case == 'duplicate-flags':
        return ('<w:compat>'
                '<w:compatSetting w:name="doNotExpandShiftReturn" w:uri="urn:settings" w:val="1"/>'
                '<w:compatSetting w:name="doNotExpandShiftReturn" w:uri="urn:settings" w:val="0"/>'
                '<w:compatSetting w:name="useWord2010TableStyleRules" w:uri="urn:settings" w:val="1"/>'
                '</w:compat>'), ''
    if case == 'mce-fallback':
        return ('<mc:AlternateContent>'
                '<mc:Choice Requires="w14"><w14:conflictMode w14:val="choice"/></mc:Choice>'
                '<mc:Fallback><w:compat><w:compatSetting w:name="fallback" w:uri="urn:settings" w:val="1"/></w:compat></mc:Fallback>'
                '</mc:AlternateContent>'), ' mc:Ignorable="w14"'
    if case == 'must-understand':
        return ('<w14:settingsExt w14:val="must-understand"/>'), ' mc:MustUnderstand="w14" mc:Ignorable="w14"'
    if case == 'inherited-namespace':
        return ('<w:compat><w:compatSetting w:name="inherited" w:uri="urn:settings" w:val="1"/>'
                '<w14:settingsExt w14:val="inherited"/></w:compat>'), ' mc:Ignorable="w14"'
    if case == 'directive-amplification':
        return ('<mc:AlternateContent><mc:Choice Requires="w14"><w14:settingsExt w14:val="one"/>'
                '<mc:AlternateContent><mc:Choice Requires="w15"><w15:settingsExt w15:val="two"/>'
                '<mc:AlternateContent><mc:Choice Requires="w14"><w14:settingsExt w14:val="three"/>'
                '</mc:Choice><mc:Fallback><w:compat/></mc:Fallback></mc:AlternateContent>'
                '</mc:Choice><mc:Fallback><w:compat/></mc:Fallback></mc:AlternateContent>'
                '</mc:Choice><mc:Fallback><w:compat/></mc:Fallback></mc:AlternateContent>'), ' mc:Ignorable="w14 w15"'
    raise ValueError(f'unknown settings case: {case}')


def settings_xml(case, strict=False):
    body, attributes = settings_body(case)
    namespace = STRICT if strict else WORD
    return (f'<w:settings xmlns:w="{namespace}" xmlns:mc="{MC}" '
            f'xmlns:w14="{W14}" xmlns:w15="{W15}"{attributes}>{body}</w:settings>').encode()


def settings_main_xml(strict=False):
    namespace = STRICT if strict else WORD
    body = '<w:p><w:r><w:t xml:space="preserve">settings seed</w:t></w:r></w:p>'
    return f'<w:document xmlns:w="{namespace}"><w:body>{body}</w:body></w:document>'.encode()


def settings_package(case, method, strict=False):
    out = io.BytesIO()
    dialect_rel = STRICT_REL if strict else REL
    settings_rel = STRICT_SETTINGS_REL if strict else SETTINGS_REL
    types = ('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
             '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
             '<Default Extension="bin" ContentType="application/octet-stream"/>'
             '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
             '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
             '</Types>').encode()
    rels = ('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            f'<Relationship Id="rId1" Type="{dialect_rel}" Target="word/document.xml"/>'
            f'<Relationship Id="rId2" Type="{settings_rel}" Target="word/settings.xml"/>'
            '</Relationships>').encode()
    entries = (
        ('[Content_Types].xml', types),
        ('_rels/.rels', rels),
        ('word/document.xml', settings_main_xml(strict)),
        ('word/settings.xml', settings_xml(case, strict)),
        ('word/opaque.bin', bytes(range(256))),
    )
    with zipfile.ZipFile(out, 'w') as archive:
        for name, data in entries:
            entry = zipfile.ZipInfo(name, (1980, 1, 1, 0, 0, 0))
            entry.compress_type = method
            entry.external_attr = 0o600 << 16
            archive.writestr(entry, data)
    return out.getvalue()


def seed_record(data):
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        main = archive.read('word/document.xml')
    return {
        'bytes': len(data),
        'sha256': hashlib.sha256(data).hexdigest(),
        'main_xml_sha256': hashlib.sha256(main).hexdigest(),
    }


def verify_seed(path, record, *, settings=False):
    if path.stat().st_size != record['bytes'] or hashlib.sha256(path.read_bytes()).hexdigest() != record['sha256']:
        raise ValueError(f'seed metadata differs: {path.name}')
    with zipfile.ZipFile(path) as archive:
        if archive.testzip() is not None:
            raise ValueError(f'seed CRC check failed: {path.name}')
        if archive.namelist().count('word/document.xml') != 1:
            raise ValueError(f'seed main member differs: {path.name}')
        main = archive.read('word/document.xml')
        if settings and archive.namelist().count('word/settings.xml') != 1:
            raise ValueError(f'settings member missing: {path.name}')
    if hashlib.sha256(main).hexdigest() != record['main_xml_sha256']:
        raise ValueError(f'seed main XML hash differs: {path.name}')


def verify_manifest(manifest, seeds, *, settings_names=(), exact=True):
    if not isinstance(manifest, dict) or not manifest:
        raise ValueError('seed manifest is empty')
    expected = set(manifest)
    actual = {path.name for path in seeds.iterdir() if path.is_file() and not path.is_symlink()}
    if not expected <= actual or (exact and actual != expected):
        raise ValueError(f'seed directory differs: missing={sorted(expected - actual)} extra={sorted(actual - expected)}')
    for name, record in manifest.items():
        if Path(name).name != name or Path(name).suffix != '.docx':
            raise ValueError(f'seed name is malformed: {name}')
        if set(record) != {'bytes', 'sha256', 'main_xml_sha256'}:
            raise ValueError(f'seed record fields differ: {name}')
        verify_seed(seeds / name, record, settings=name in settings_names)


def legacy_inputs():
    fuzz = ROOT / 'fuzz'
    seeds = fuzz / 'seeds'
    manifest_path = fuzz / 'seed-manifest.json'
    generator_path = fuzz / 'generator.json'
    if not seeds.is_dir() or not manifest_path.is_file() or not generator_path.is_file():
        raise ValueError('legacy fuzz seed inputs are incomplete')
    manifest_bytes = manifest_path.read_bytes()
    generator_bytes = generator_path.read_bytes()
    manifest = json.loads(manifest_bytes)
    generator = json.loads(generator_bytes)
    if len(manifest) != 30 or generator.get('seeds') != 30 or generator.get('zip_crc_and_main_hashes_verified') is not True:
        raise ValueError('legacy fuzz seed count/verification record differs')
    legacy_hash = generator.get('generator_sha256')
    if not isinstance(legacy_hash, str) or len(legacy_hash) != 64:
        raise ValueError('legacy generator source digest is missing')
    historical = ROOT / 'driver-history' / legacy_hash / Path(__file__).name
    if not historical.is_file() or hashlib.sha256(historical.read_bytes()).hexdigest() != legacy_hash:
        raise ValueError('legacy generator source history is missing')
    verify_manifest(manifest, seeds, exact=False)
    extras = {
        name for name in (path.name for path in seeds.iterdir() if path.is_file() and not path.is_symlink())
        if name not in manifest
    }
    if any(not name.startswith('settings-') or not name.endswith('.docx') for name in extras):
        raise ValueError(f'legacy seed directory contains unknown extras: {sorted(extras)}')
    return {
        'manifest': manifest,
        'generator': generator,
        'manifest_bytes': manifest_bytes,
        'generator_bytes': generator_bytes,
        'manifest_sha256': hashlib.sha256(manifest_bytes).hexdigest(),
        'generator_sha256': hashlib.sha256(generator_bytes).hexdigest(),
        'seeds': seeds,
    }


def write_bytes_exclusive(path, data):
    with path.open('xb') as stream:
        stream.write(data)


def extend_settings():
    legacy = legacy_inputs()
    fuzz = ROOT / 'fuzz'
    seeds = legacy['seeds']
    manifest_path = fuzz / 'seed-manifest-v2.json'
    generator_path = fuzz / 'generator-v2.json'
    if manifest_path.exists() or generator_path.exists():
        raise ValueError('v2 manifest/generator already exists; refusing replacement')
    new_manifest = {}
    new_data = {}
    for dialect, strict in SETTINGS_DIALECTS:
        for case in SETTINGS_CASES:
            for method, suffix in SETTINGS_METHODS:
                name = f'settings-{case}-{dialect}-{suffix}.docx'
                path = seeds / name
                data = settings_package(case, method, strict)
                if path.exists() and (not path.is_file() or path.is_symlink() or path.read_bytes() != data):
                    raise ValueError(f'existing settings seed differs: {name}')
                new_data[name] = data
                new_manifest[name] = seed_record(data)
    combined = dict(legacy['manifest'])
    combined.update(new_manifest)
    settings_names = set(new_manifest)
    existing_settings = {
        path.name for path in seeds.iterdir()
        if path.is_file() and not path.is_symlink() and path.name not in legacy['manifest']
    }
    if not existing_settings <= settings_names:
        raise ValueError(f'seed directory contains unknown settings extras: {sorted(existing_settings - settings_names)}')
    for name, data in new_data.items():
        path = seeds / name
        if path.exists():
            if not path.is_file() or path.is_symlink() or path.read_bytes() != data:
                raise ValueError(f'existing settings seed differs: {name}')
        else:
            write_bytes_exclusive(path, data)
    manifest_data = (json.dumps(combined, indent=2, sort_keys=True) + '\n').encode()
    current_script_hash = hashlib.sha256(Path(__file__).read_bytes()).hexdigest()
    generator = {
        'schema': 'litchi-docx-fuzz-generator-v2',
        'version': 2,
        'python': sys.version,
        'zlib_build': getattr(zlib, 'ZLIB_VERSION', 'unknown'),
        'zlib_runtime': getattr(zlib, 'ZLIB_RUNTIME_VERSION', getattr(zlib, 'ZLIB_VERSION', 'unknown')),
        'generator_sha256': current_script_hash,
        'seed_count': len(combined),
        'seeds': len(combined),
        'legacy_seed_count': len(legacy['manifest']),
        'settings_seed_count': len(new_manifest),
        'settings_cases': list(SETTINGS_CASES),
        'word_dialects': [name for name, _ in SETTINGS_DIALECTS],
        'compression_methods': [suffix for _, suffix in SETTINGS_METHODS],
        'legacy_manifest_sha256': legacy['manifest_sha256'],
        'legacy_generator_sha256': legacy['generator_sha256'],
        'seed_manifest_sha256': hashlib.sha256(manifest_data).hexdigest(),
        'zip_crc_and_main_hashes_verified': True,
        'settings_member_verified': True,
    }
    write_bytes_exclusive(manifest_path, manifest_data)
    write_bytes_exclusive(generator_path, (json.dumps(generator, indent=2, sort_keys=True) + '\n').encode())
    verify_manifest(combined, seeds, settings_names=settings_names)
    if manifest_path.read_bytes() != manifest_data:
        raise ValueError('manifest custody changed unexpectedly')
    if (fuzz / 'seed-manifest.json').read_bytes() != legacy['manifest_bytes']:
        raise ValueError('legacy seed manifest changed')
    if (fuzz / 'generator.json').read_bytes() != legacy['generator_bytes']:
        raise ValueError('legacy generator record changed')
    print(f'Generated {len(new_manifest)} settings seeds; combined v2 manifest has {len(combined)} seeds.')


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--extend-settings', action='store_true')
    options = parser.parse_args()
    if not options.extend_settings:
        parser.error('explicit --extend-settings is required; legacy records are immutable')
    extend_settings()


if __name__ == '__main__':
    main()
