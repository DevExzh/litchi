fn main() {
    println!("sha={} avx2={} avx512f={}",
        std::arch::is_x86_feature_detected!("sha"),
        std::arch::is_x86_feature_detected!("avx2"),
        std::arch::is_x86_feature_detected!("avx512f"));
}
