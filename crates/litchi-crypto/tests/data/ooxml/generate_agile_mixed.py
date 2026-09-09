"""Synthetic Agile interop: AES-128 password wrapping, AES-256 data, SHA-512.

Requires msoffcrypto-tool==6.0.0 and cryptography==50.0.1. This component-built
vector is not native Office output. MS-OFFCRYPTO 2.3.4.10 requires matching
cipherAlgorithm/hashAlgorithm; keyBits is independently specified. Section
2.3.4.13 sizes the wrapped intermediate key using KeyData.keyBits.
All deterministic salts/keys below are test material only.

msoffcrypto-tool's CFB writer emits a zero EncryptionBlockSize in its advisory
DataSpaces transform. The generator patches that field to the normative AES
block size (16) while leaving EncryptionInfo and EncryptedPackage untouched.
"""
import hashlib
import hmac
import io
import json
from pathlib import Path

import msoffcrypto
from msoffcrypto.method.ecma376_agile import (
    ECMA376Agile, ECMA376AgileEncryptionInfo, _encrypt_aes_cbc,
    blkKey_VerifierHashInput, blkKey_encryptedVerifierHashValue,
    blkKey_encryptedKeyValue, blkKey_dataIntegrity1, blkKey_dataIntegrity2,
)
from msoffcrypto.method.container.ecma376_encrypted import (
    DSPos,
    DefaultContent,
    ECMA376Encrypted,
)

root = Path(__file__).resolve().parent
clear = (root / 'clear.docx').read_bytes()
password = 'Litchi synthetic crypto fixture 2026'
info = ECMA376AgileEncryptionInfo()
info.encryptedKey.keyBits = 128
info.encryptedKey.saltValue = hashlib.sha256(b'mixed Agile password salt').digest()[:16]
info.keyData.saltValue = hashlib.sha256(b'mixed Agile package salt').digest()[:16]
secret = hashlib.sha256(b'mixed Agile content key').digest()
verifier = hashlib.sha256(b'mixed Agile password verifier').digest()[:16]
password_hash = ECMA376Agile._derive_iterated_hash_from_password(
    password, info.encryptedKey.saltValue, 'SHA512', info.spinCount,
).digest()

def wrap(block, value):
    key = ECMA376Agile._derive_encryption_key(password_hash, block, 'SHA512', 128)
    return _encrypt_aes_cbc(value, key, info.encryptedKey.saltValue)

info.encryptedVerifierHashInput = wrap(blkKey_VerifierHashInput, verifier)
info.encryptedVerifierHashValue = wrap(blkKey_encryptedVerifierHashValue, hashlib.sha512(verifier).digest())
info.encryptedKeyValue = wrap(blkKey_encryptedKeyValue, secret)
encrypted = ECMA376Agile.encrypt_payload(io.BytesIO(clear), info.keyData, secret, info.keyData.saltValue)
integrity_key = hashlib.sha512(b'mixed Agile integrity key').digest()
integrity_value = hmac.new(integrity_key, encrypted, hashlib.sha512).digest()
for attribute, block, value in [
    ('encryptedHmacKey', blkKey_dataIntegrity1, integrity_key),
    ('encryptedHmacValue', blkKey_dataIntegrity2, integrity_value),
]:
    iv = hashlib.sha512(info.keyData.saltValue + block).digest()[:16]
    setattr(info, attribute, _encrypt_aes_cbc(value, secret, iv))
descriptor = info.getEncryptionDescriptorHeader() + info.toEncryptionDescriptor().encode('utf-8')
container = ECMA376Encrypted(encrypted, descriptor)
for directory in container._dirs:
    directory.CreationTime = 0
    directory.ModificationTime = 0
primary = bytearray(DefaultContent.Primary)
primary[188:192] = (16).to_bytes(4, 'little')
container._dirs[DSPos.iPrimary].Content = bytes(primary)
output = io.BytesIO()
container.write_to(output)
data = output.getvalue()
filename = 'component-agile-aes128-wrap-aes256-data-sha512.docx'
(root / filename).write_bytes(data)
opened = msoffcrypto.OfficeFile(io.BytesIO(data))
opened.load_key(password=password, verify_password=True)
plaintext = io.BytesIO()
opened.decrypt(plaintext, verify_integrity=True)
assert plaintext.getvalue() == clear
(root / 'agile-mixed-manifest.json').write_text(json.dumps({
    'file': filename, 'sha256': hashlib.sha256(data).hexdigest(), 'bytes': len(data),
    'clear_sha256': hashlib.sha256(clear).hexdigest(), 'password': password,
    'provenance': __doc__, 'verification': 'msoffcrypto password+integrity+byte equality passed',
}, indent=2) + '\n')
