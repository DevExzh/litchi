"""Per-callee call counts from one callgrind output (change 0656)."""
import sys
sys.path.insert(0, sys.argv[2])
from callcounts import parse
SYMS = [
    'litchi_pptx::opened::model::package_fingerprint',
    'litchi_pptx::opened::cross_copy_plan::physical_package_fingerprint',
    'litchi_pptx::opened::cross_copy_plan::snapshot_physical_revision',
    'litchi_pptx::opened::cross_copy_plan::bounded_package_bytes',
    'litchi_pptx::opened::cross_copy_plan::build_candidate',
    'litchi_pptx::opened::cross_copy_plan::prepare_cross_slide_copy_for_slides',
    'litchi_pptx::opened::model::capture_internal',
    'litchi_opc::pkgwriter::PackageWriter::write_to_stream',
    'litchi_opc::package::OpcPackage::from_vec_reusing_payloads',
    'zlib_rs::deflate::deflate',
    'sha2::sha256::compress256',
    'soapberry_zip::preserve::PreservationIndex<R>::write_to',
]
calls = parse(sys.argv[1])
print(f'# per-callee call counts summed over the shared fn/cfn name-compression table')
print(f'# {sys.argv[1]}')
for sym in SYMS:
    print(f'{sum(v for k, v in calls.items() if k == sym):>10d}  {sym}')
