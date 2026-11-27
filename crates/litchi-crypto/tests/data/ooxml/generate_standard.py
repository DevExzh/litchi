"""Synthetic interop vectors: msoffcrypto-tool==6.0.0, cryptography==50.0.1.

This is a component-built fixture, not output from a native Office application.
Run beside clear.docx. Salt/verifier are deterministic test material only.
Normative envelope: MS-OFFCRYPTO 2.3.4.5; key derivation: 2.3.4.7.

msoffcrypto-tool's CFB writer emits a zero EncryptionBlockSize in its advisory
DataSpaces transform. The generator patches that field to the normative AES
block size (16) while leaving EncryptionInfo and EncryptedPackage untouched.
"""
from pathlib import Path
import hashlib
import io
import json
import struct

import msoffcrypto
from msoffcrypto.method.ecma376_standard import ECMA376Standard
from msoffcrypto.method.container.ecma376_encrypted import (
    DSPos,
    DefaultContent,
    ECMA376Encrypted,
)
from cryptography.hazmat.primitives.ciphers import Cipher, algorithms, modes

root = Path(__file__).resolve().parent
clear = (root / 'clear.docx').read_bytes()
password = 'Litchi synthetic crypto fixture 2026'
results = []
for bits, algid in [(192, 0x660f), (256, 0x6610)]:
    salt = hashlib.sha256(f'synthetic Standard AES{bits} salt'.encode()).digest()[:16]
    verifier = hashlib.sha256(f'synthetic Standard AES{bits} verifier'.encode()).digest()[:16]
    key = ECMA376Standard.makekey_from_password(password, algid, 0x8004, 0x18, bits, 16, salt)

    def enc(data):
        padded = data + b'\0' * ((-len(data)) % 16)
        context = Cipher(algorithms.AES(key), modes.ECB()).encryptor()
        return context.update(padded) + context.finalize()

    header = struct.pack('<8I', 0x24, 0, algid, 0x8004, bits, 0x18, 0, 0)
    header += ('Microsoft Enhanced RSA and AES Cryptographic Provider\0').encode('utf-16le')
    info = struct.pack('<HHII', 3, 2, 0x24, len(header)) + header
    info += struct.pack('<I', 16) + salt + enc(verifier)
    info += struct.pack('<I', 20) + enc(hashlib.sha1(verifier).digest())
    encrypted = struct.pack('<Q', len(clear)) + enc(clear)
    output = io.BytesIO()
    container = ECMA376Encrypted(encrypted, info)
    # The pinned third-party container otherwise injects datetime.now() into
    # every directory entry. Explicit zeros are deterministic, and respect
    # the CFB requirement that stream timestamps be zero.
    for directory in container._dirs:
        directory.CreationTime = 0
        directory.ModificationTime = 0
    primary = bytearray(DefaultContent.Primary)
    primary[188:192] = struct.pack('<I', 16)
    container._dirs[DSPos.iPrimary].Content = bytes(primary)
    container.write_to(output)
    data = output.getvalue()
    filename = f'component-standard-aes{bits}-sha1.docx'
    (root / filename).write_bytes(data)
    opened = msoffcrypto.OfficeFile(io.BytesIO(data))
    opened.load_key(password=password, verify_password=True)
    plaintext = io.BytesIO()
    opened.decrypt(plaintext)
    assert plaintext.getvalue() == clear
    results.append({'file': filename, 'sha256': hashlib.sha256(data).hexdigest(),
                    'bytes': len(data), 'profile': f'Standard AES{bits} SHA1 50000 spins'})
(root / 'standard-manifest.json').write_text(json.dumps({
    'password': password, 'clear_sha256': hashlib.sha256(clear).hexdigest(),
    'provenance': __doc__, 'vectors': results,
}, indent=2) + '\n')
