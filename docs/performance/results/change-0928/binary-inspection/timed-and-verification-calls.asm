
/tmp/litchi-performance-0925/measurement-target-v4-final/release/litchi-perf-baseline:     file format elf64-x86-64


Disassembly of section .text:

0000000001057b40 <_ZN20litchi_perf_baseline56run_docx_source_backed_existing_section_layout_edit_save17h1cd3e0d854c7eea3E+0x3060>:
 1057b40:	0e                   	(bad)
 1057b41:	ff                   	(bad)
 1057b42:	ff 15 40 90 85 02    	call   *0x2859040(%rip)        # 38b0b88 <_DYNAMIC+0x2e8>
 1057b48:	48 89 94 24 a8 00 00 	mov    %rdx,0xa8(%rsp)
 1057b4f:	00 
 1057b50:	48 85 c0             	test   %rax,%rax
 1057b53:	0f 85 0b 2c 00 00    	jne    105a764 <_ZN20litchi_perf_baseline56run_docx_source_backed_existing_section_layout_edit_save17h1cd3e0d854c7eea3E+0x5c84>
 1057b59:	ff 15 d9 93 85 02    	call   *0x28593d9(%rip)        # 38b0f38 <_DYNAMIC+0x698>
 1057b5f:	48 89 84 24 70 07 00 	mov    %rax,0x770(%rsp)
 1057b66:	00 
 1057b67:	89 94 24 78 07 00 00 	mov    %edx,0x778(%rsp)
 1057b6e:	48 89 ef             	mov    %rbp,%rdi
 1057b71:	48 8d b4 24 40 16 00 	lea    0x1640(%rsp),%rsi
 1057b78:	00 
 1057b79:	ff 15 31 a8 85 02    	call   *0x285a831(%rip)        # 38b23b0 <_DYNAMIC+0x1b10>
 1057b7f:	48 8b 84 24 00 27 00 	mov    0x2700(%rsp),%rax
 1057b86:	00 
 1057b87:	48 8d 8c 24 10 27 00 	lea    0x2710(%rsp),%rcx
 1057b8e:	00 
 1057b8f:	0f 10 41 f8          	movups -0x8(%rcx),%xmm0
 1057b93:	0f 10 49 08          	movups 0x8(%rcx),%xmm1
 1057b97:	0f 10 51 18          	movups 0x18(%rcx),%xmm2
 1057b9b:	0f 10 59 28          	movups 0x28(%rcx),%xmm3
 1057b9f:	0f 29 84 24 e0 0d 00 	movaps %xmm0,0xde0(%rsp)
 1057ba6:	00 
 1057ba7:	0f 29 8c 24 f0 0d 00 	movaps %xmm1,0xdf0(%rsp)
 1057bae:	00 
 1057baf:	0f 29 94 24 00 0e 00 	movaps %xmm2,0xe00(%rsp)
 1057bb6:	00 
 1057bb7:	0f 29 9c 24 10 0e 00 	movaps %xmm3,0xe10(%rsp)
 1057bbe:	00 
 1057bbf:	48 8b 49 38          	mov    0x38(%rcx),%rcx
 1057bc3:	48 89 8c 24 20 0e 00 	mov    %rcx,0xe20(%rsp)
 1057bca:	00 
 1057bcb:	48 83 f8 02          	cmp    $0x2,%rax
 1057bcf:	0f 84 cc 24 00 00    	je     105a0a1 <_ZN20litchi_perf_baseline56run_docx_source_backed_existing_section_layout_edit_save17h1cd3e0d854c7eea3E+0x55c1>
 1057bd5:	48 8d 8c 24 50 27 00 	lea    0x2750(%rsp),%rcx
 1057bdc:	00 
 1057bdd:	0f 10 41 70          	movups 0x70(%rcx),%xmm0
 1057be1:	48 8d 94 24 78 15 00 	lea    0x1578(%rsp),%rdx
 1057be8:	00 
 1057be9:	0f 11 82 b8 00 00 00 	movups %xmm0,0xb8(%rdx)
 1057bf0:	0f 10 41 60          	movups 0x60(%rcx),%xmm0
 1057bf4:	0f 11 82 a8 00 00 00 	movups %xmm0,0xa8(%rdx)
 1057bfb:	0f 10 41 50          	movups 0x50(%rcx),%xmm0
 1057bff:	0f 11 82 98 00 00 00 	movups %xmm0,0x98(%rdx)
 1057c06:	0f 10 41 40          	movups 0x40(%rcx),%xmm0
 1057c0a:	0f 11 82 88 00 00 00 	movups %xmm0,0x88(%rdx)
 1057c11:	0f 10 01             	movups (%rcx),%xmm0
 1057c14:	0f 10 49 10          	movups 0x10(%rcx),%xmm1
 1057c18:	0f 10 51 20          	movups 0x20(%rcx),%xmm2
 1057c1c:	0f 10 59 30          	movups 0x30(%rcx),%xmm3
 1057c20:	0f 11 5a 78          	movups %xmm3,0x78(%rdx)
 1057c24:	0f 11 52 68          	movups %xmm2,0x68(%rdx)
 1057c28:	0f 11 4a 58          	movups %xmm1,0x58(%rdx)
 1057c2c:	0f 11 42 48          	movups %xmm0,0x48(%rdx)
 1057c30:	48 8b 8c 24 20 0e 00 	mov    0xe20(%rsp),%rcx
 1057c37:	00 
 1057c38:	48 89 4a 40          	mov    %rcx,0x40(%rdx)
 1057c3c:	0f 28 84 24 e0 0d 00 	movaps 0xde0(%rsp),%xmm0
 1057c43:	00 
 1057c44:	0f 28 8c 24 f0 0d 00 	movaps 0xdf0(%rsp),%xmm1
 1057c4b:	00 
 1057c4c:	0f 28 94 24 00 0e 00 	movaps 0xe00(%rsp),%xmm2
 1057c53:	00 
 1057c54:	0f 28 9c 24 10 0e 00 	movaps 0xe10(%rsp),%xmm3
 1057c5b:	00 
 1057c5c:	0f 11 5a 30          	movups %xmm3,0x30(%rdx)
 1057c60:	0f 11 52 20          	movups %xmm2,0x20(%rdx)
 1057c64:	0f 11 4a 10          	movups %xmm1,0x10(%rdx)
 1057c68:	0f 11 02             	movups %xmm0,(%rdx)
 1057c6b:	48 89 84 24 70 15 00 	mov    %rax,0x1570(%rsp)
 1057c72:	00 
 1057c73:	48 8d 84 24 70 15 00 	lea    0x1570(%rsp),%rax
 1057c7a:	00 
 1057c7b:	48 89 84 24 00 27 00 	mov    %rax,0x2700(%rsp)
 1057c82:	00 
 1057c83:	48 8d bc 24 70 07 00 	lea    0x770(%rsp),%rdi
 1057c8a:	00 
 1057c8b:	ff 15 af 92 85 02    	call   *0x28592af(%rip)        # 38b0f40 <_DYNAMIC+0x6a0>
 1057c91:	89 d1                	mov    %edx,%ecx
 1057c93:	ba 00 ca 9a 3b       	mov    $0x3b9aca00,%edx
 1057c98:	48 f7 e2             	mul    %rdx
 1057c9b:	89 c9                	mov    %ecx,%ecx
 1057c9d:	48 01 c1             	add    %rax,%rcx
 1057ca0:	48 89 8c 24 a0 00 00 	mov    %rcx,0xa0(%rsp)
 1057ca7:	00 
 1057ca8:	48 83 d2 00          	adc    $0x0,%rdx
 1057cac:	74 23                	je     1057cd1 <_ZN20litchi_perf_baseline56run_docx_source_backed_existing_section_layout_edit_save17h1cd3e0d854c7eea3E+0x31f1>
 1057cae:	be 25 00 00 00       	mov    $0x25,%esi
 1057cb3:	48 8d 3d 49 a3 0e ff 	lea    -0xf15cb7(%rip),%rdi        # 142003 <anon.41d43701a2c2b5e6feb75b74ff96deac.208.llvm.8476651010627807873>
 1057cba:	ff 15 c8 8e 85 02    	call   *0x2858ec8(%rip)        # 38b0b88 <_DYNAMIC+0x2e8>
 1057cc0:	48 89 94 24 a0 00 00 	mov    %rdx,0xa0(%rsp)
 1057cc7:	00 
 1057cc8:	48 85 c0             	test   %rax,%rax
 1057ccb:	0f 85 b3 2a 00 00    	jne    105a784 <_ZN20litchi_perf_baseline56run_docx_source_backed_existing_section_layout_edit_save17h1cd3e0d854c7eea3E+0x5ca4>
 1057cd1:	ff 15 61 92 85 02    	call   *0x2859261(%rip)        # 38b0f38 <_DYNAMIC+0x698>
 1057cd7:	48 89 84 24 80 07 00 	mov    %rax,0x780(%rsp)
 1057cde:	00 
 1057cdf:	89 94 24 88 07 00 00 	mov    %edx,0x788(%rsp)
 1057ce6:	48 c7 84 24 60 21 00 	movq   $0x0,0x2160(%rsp)
 1057ced:	00 00 00 00 00 
 1057cf2:	48 c7 84 24 68 21 00 	movq   $0x40,0x2168(%rsp)
 1057cf9:	00 40 00 00 00 
 1057cfe:	48 89 ef             	mov    %rbp,%rdi
 1057d01:	48 8d b4 24 70 15 00 	lea    0x1570(%rsp),%rsi
 1057d08:	00 
 1057d09:	48 8d 94 24 60 21 00 	lea    0x2160(%rsp),%rdx
 1057d10:	00 
 1057d11:	e8 ca b4 40 00       	call   14631e0 <_ZN11litchi_docx7section6layout8Snapshot4edit17h7c0fccd1861522f2E>
 1057d16:	48 8b 9c 24 00 27 00 	mov    0x2700(%rsp),%rbx
 1057d1d:	00 
 1057d1e:	48 8d 84 24 10 27 00 	lea    0x2710(%rsp),%rax
 1057d25:	00 
 1057d26:	0f 10 40 f8          	movups -0x8(%rax),%xmm0
 1057d2a:	0f 10 48 08          	movups 0x8(%rax),%xmm1
 1057d2e:	0f 10 50 18          	movups 0x18(%rax),%xmm2
 1057d32:	0f 10 58 28          	movups 0x28(%rax),%xmm3
 1057d36:	0f 29 84 24 30 08 00 	movaps %xmm0,0x830(%rsp)
 1057d3d:	00 
 1057d3e:	0f 29 8c 24 40 08 00 	movaps %xmm1,0x840(%rsp)
 1057d45:	00 
 1057d46:	0f 29 94 24 50 08 00 	movaps %xmm2,0x850(%rsp)
 1057d4d:	00 
 1057d4e:	0f 29 9c 24 60 08 00 	movaps %xmm3,0x860(%rsp)
 1057d55:	00 
 1057d56:	48 8b 40 38          	mov    0x38(%rax),%rax
 1057d5a:	48 89 84 24 70 08 00 	mov    %rax,0x870(%rsp)
 1057d61:	00 
 1057d62:	48 83 fb 02          	cmp    $0x2,%rbx
 1057d66:	0f 84 fa 23 00 00    	je     105a166 <_ZN20litchi_perf_baseline56run_docx_source_backed_existing_section_layout_edit_save17h1cd3e0d854c7eea3E+0x5686>
 1057d6c:	ba c8 04 00 00       	mov    $0x4c8,%edx
 1057d71:	48 8d bc 24 30 0e 00 	lea    0xe30(%rsp),%rdi
 1057d78:	00 
 1057d79:	48 8d b4 24 50 27 00 	lea    0x2750(%rsp),%rsi
 1057d80:	00 
 1057d81:	ff 15 59 8d 85 02    	call   *0x2858d59(%rip)        # 38b0ae0 <memcpy@GLIBC_2.14>
 1057d87:	48 8b 84 24 70 08 00 	mov    0x870(%rsp),%rax
 1057d8e:	00 
 1057d8f:	48 8d 8c 24 e8 0d 00 	lea    0xde8(%rsp),%rcx
 1057d96:	00 
 1057d97:	48 89 41 40          	mov    %rax,0x40(%rcx)
 1057d9b:	0f 28 84 24 30 08 00 	movaps 0x830(%rsp),%xmm0
 1057da2:	00 
 1057da3:	0f 28 8c 24 40 08 00 	movaps 0x840(%rsp),%xmm1
 1057daa:	00 
 1057dab:	0f 28 94 24 50 08 00 	movaps 0x850(%rsp),%xmm2
 1057db2:	00 
 1057db3:	0f 28 9c 24 60 08 00 	movaps 0x860(%rsp),%xmm3
 1057dba:	00 
 1057dbb:	0f 11 59 30          	movups %xmm3,0x30(%rcx)
 1057dbf:	0f 11 51 20          	movups %xmm2,0x20(%rcx)
 1057dc3:	0f 11 49 10          	movups %xmm1,0x10(%rcx)
 1057dc7:	0f 11 01             	movups %xmm0,(%rcx)
 1057dca:	48 89 9c 24 e0 0d 00 	mov    %rbx,0xde0(%rsp)
 1057dd1:	00 
 1057dd2:	48 c7 84 24 00 27 00 	movq   $0x1,0x2700(%rsp)
 1057dd9:	00 01 00 00 00 
 1057dde:	48 c7 84 24 08 27 00 	movq   $0x25350,0x2708(%rsp)
 1057de5:	00 50 53 02 00 
 1057dea:	48 c7 84 24 10 27 00 	movq   $0x0,0x2710(%rsp)
 1057df1:	00 00 00 00 00 
 1057df6:	48 c7 84 24 18 27 00 	movq   $0x8,0x2718(%rsp)
 1057dfd:	00 08 00 00 00 
 1057e02:	48 c7 84 24 20 27 00 	movq   $0x0,0x2720(%rsp)
 1057e09:	00 00 00 00 00 
 1057e0e:	c7 84 24 28 27 00 00 	movl   $0x1010003,0x2728(%rsp)
 1057e15:	03 00 01 01 
 1057e19:	48 8d bc 24 30 08 00 	lea    0x830(%rsp),%rdi
 1057e20:	00 
 1057e21:	48 8d b4 24 e0 0d 00 	lea    0xde0(%rsp),%rsi
 1057e28:	00 
 1057e29:	48 89 ea             	mov    %rbp,%rdx
 1057e2c:	ff 15 9e a7 85 02    	call   *0x285a79e(%rip)        # 38b25d0 <_DYNAMIC+0x1d30>
 1057e32:	48 8b 84 24 30 08 00 	mov    0x830(%rsp),%rax
 1057e39:	00 
 1057e3a:	48 83 f8 23          	cmp    $0x23,%rax
 1057e3e:	0f 85 d0 23 00 00    	jne    105a214 <_ZN20litchi_perf_baseline56run_docx_source_backed_existing_section_layout_edit_save17h1cd3e0d854c7eea3E+0x5734>
 1057e44:	48 8d 84 24 e0 0d 00 	lea    0xde0(%rsp),%rax
 1057e4b:	00 
 1057e4c:	48 89 84 24 00 27 00 	mov    %rax,0x2700(%rsp)
 1057e53:	00 
 1057e54:	48 8d bc 24 80 07 00 	lea    0x780(%rsp),%rdi
 1057e5b:	00 
 1057e5c:	ff 15 de 90 85 02    	call   *0x28590de(%rip)        # 38b0f40 <_DYNAMIC+0x6a0>
 1057e62:	89 d1                	mov    %edx,%ecx
 1057e64:	ba 00 ca 9a 3b       	mov    $0x3b9aca00,%edx
 1057e69:	48 f7 e2             	mul    %rdx
 1057e6c:	89 c9                	mov    %ecx,%ecx
 1057e6e:	48 01 c1             	add    %rax,%rcx
 1057e71:	48 89 8c 24 98 00 00 	mov    %rcx,0x98(%rsp)
 1057e78:	00 
 1057e79:	48 83 d2 00          	adc    $0x0,%rdx
 1057e7d:	74 23                	je     1057ea2 <_ZN20litchi_perf_baseline56run_docx_source_backed_existing_section_layout_edit_save17h1cd3e0d854c7eea3E+0x33c2>
 1057e7f:	be 25 00 00 00       	mov    $0x25,%esi
 1057e84:	48 8d 3d 78 a1 0e ff 	lea    -0xf15e88(%rip),%rdi        # 142003 <anon.41d43701a2c2b5e6feb75b74ff96deac.208.llvm.8476651010627807873>
 1057e8b:	ff 15 f7 8c 85 02    	call   *0x2858cf7(%rip)        # 38b0b88 <_DYNAMIC+0x2e8>
 1057e91:	48 89 94 24 98 00 00 	mov    %rdx,0x98(%rsp)
 1057e98:	00 
 1057e99:	48 85 c0             	test   %rax,%rax
 1057e9c:	0f 85 c3 29 00 00    	jne    105a865 <_ZN20litchi_perf_baseline56run_docx_source_backed_existing_section_layout_edit_save17h1cd3e0d854c7eea3E+0x5d85>
 1057ea2:	ff 15 90 90 85 02    	call   *0x2859090(%rip)        # 38b0f38 <_DYNAMIC+0x698>
 1057ea8:	48 89 84 24 90 07 00 	mov    %rax,0x790(%rsp)
 1057eaf:	00 
 1057eb0:	89 94 24 98 07 00 00 	mov    %edx,0x798(%rsp)
 1057eb7:	ba 18 05 00 00       	mov    $0x518,%edx
 1057ebc:	48 89 ef             	mov    %rbp,%rdi
 1057ebf:	48 8d b4 24 e0 0d 00 	lea    0xde0(%rsp),%rsi
 1057ec6:	00 
 1057ec7:	ff 15 13 8c 85 02    	call   *0x2858c13(%rip)        # 38b0ae0 <memcpy@GLIBC_2.14>
 1057ecd:	b3 01                	mov    $0x1,%bl
 1057ecf:	48 8d bc 24 30 08 00 	lea    0x830(%rsp),%rdi
 1057ed6:	00 
 1057ed7:	48 89 ee             	mov    %rbp,%rsi
 1057eda:	ff 15 e8 a6 85 02    	call   *0x285a6e8(%rip)        # 38b25c8 <_DYNAMIC+0x1d28>
 1057ee0:	48 8b 9c 24 30 08 00 	mov    0x830(%rsp),%rbx
 1057ee7:	00 
 1057ee8:	48 8d 84 24 38 08 00 	lea    0x838(%rsp),%rax
 1057eef:	00 
 1057ef0:	0f 10 00             	movups (%rax),%xmm0
 1057ef3:	0f 10 48 10          	movups 0x10(%rax),%xmm1
 1057ef7:	0f 10 50 20          	movups 0x20(%rax),%xmm2
 1057efb:	0f 10 58 30          	movups 0x30(%rax),%xmm3
 1057eff:	0f 29 84 24 60 21 00 	movaps %xmm0,0x2160(%rsp)
 1057f06:	00 
 1057f07:	0f 29 8c 24 70 21 00 	movaps %xmm1,0x2170(%rsp)
 1057f0e:	00 
 1057f0f:	0f 29 94 24 80 21 00 	movaps %xmm2,0x2180(%rsp)
 1057f16:	00 
 1057f17:	0f 29 9c 24 90 21 00 	movaps %xmm3,0x2190(%rsp)
 1057f1e:	00 
 1057f1f:	48 8b 40 40          	mov    0x40(%rax),%rax
 1057f23:	48 89 84 24 a0 21 00 	mov    %rax,0x21a0(%rsp)
 1057f2a:	00 
 1057f2b:	48 83 fb 02          	cmp    $0x2,%rbx
 1057f2f:	0f 84 95 23 00 00    	je     105a2ca <_ZN20litchi_perf_baseline56run_docx_source_backed_existing_section_layout_edit_save17h1cd3e0d854c7eea3E+0x57ea>
 1057f35:	ba 30 01 00 00       	mov    $0x130,%edx
 1057f3a:	48 8d bc 24 18 1e 00 	lea    0x1e18(%rsp),%rdi
 1057f41:	00 
 1057f42:	48 8d b4 24 80 08 00 	lea    0x880(%rsp),%rsi
 1057f49:	00 
 1057f4a:	ff 15 90 8b 85 02    	call   *0x2858b90(%rip)        # 38b0ae0 <memcpy@GLIBC_2.14>
 1057f50:	48 8b 84 24 a0 21 00 	mov    0x21a0(%rsp),%rax
 1057f57:	00 
 1057f58:	48 8d 8c 24 d0 1d 00 	lea    0x1dd0(%rsp),%rcx
 1057f5f:	00 
 1057f60:	48 89 41 40          	mov    %rax,0x40(%rcx)
 1057f64:	0f 28 84 24 60 21 00 	movaps 0x2160(%rsp),%xmm0
 1057f6b:	00 
 1057f6c:	0f 28 8c 24 70 21 00 	movaps 0x2170(%rsp),%xmm1
 1057f73:	00 
 1057f74:	0f 28 94 24 80 21 00 	movaps 0x2180(%rsp),%xmm2
 1057f7b:	00 
 1057f7c:	0f 28 9c 24 90 21 00 	movaps 0x2190(%rsp),%xmm3
 1057f83:	00 
 1057f84:	0f 11 59 30          	movups %xmm3,0x30(%rcx)
 1057f88:	0f 11 51 20          	movups %xmm2,0x20(%rcx)
 1057f8c:	0f 11 49 10          	movups %xmm1,0x10(%rcx)
 1057f90:	0f 11 01             	movups %xmm0,(%rcx)
 1057f93:	48 89 9c 24 c8 1d 00 	mov    %rbx,0x1dc8(%rsp)
 1057f9a:	00 
 1057f9b:	48 8d 84 24 c8 1d 00 	lea    0x1dc8(%rsp),%rax
 1057fa2:	00 
 1057fa3:	48 89 84 24 00 27 00 	mov    %rax,0x2700(%rsp)
 1057faa:	00 
 1057fab:	b3 01                	mov    $0x1,%bl
 1057fad:	48 8d bc 24 90 07 00 	lea    0x790(%rsp),%rdi
 1057fb4:	00 
 1057fb5:	ff 15 85 8f 85 02    	call   *0x2858f85(%rip)        # 38b0f40 <_DYNAMIC+0x6a0>
 1057fbb:	89 d1                	mov    %edx,%ecx
 1057fbd:	ba 00 ca 9a 3b       	mov    $0x3b9aca00,%edx
 1057fc2:	48 f7 e2             	mul    %rdx
 1057fc5:	89 c9                	mov    %ecx,%ecx
 1057fc7:	48 01 c1             	add    %rax,%rcx
 1057fca:	48 89 8c 24 90 00 00 	mov    %rcx,0x90(%rsp)
 1057fd1:	00 
 1057fd2:	48 83 d2 00          	adc    $0x0,%rdx
 1057fd6:	74 23                	je     1057ffb <_ZN20litchi_perf_baseline56run_docx_source_backed_existing_section_layout_edit_save17h1cd3e0d854c7eea3E+0x351b>
 1057fd8:	be 25 00 00 00       	mov    $0x25,%esi
 1057fdd:	48 8d 3d 1f a0 0e ff 	lea    -0xf15fe1(%rip),%rdi        # 142003 <anon.41d43701a2c2b5e6feb75b74ff96deac.208.llvm.8476651010627807873>
 1057fe4:	ff 15 9e 8b 85 02    	call   *0x2858b9e(%rip)        # 38b0b88 <_DYNAMIC+0x2e8>
 1057fea:	48 89 94 24 90 00 00 	mov    %rdx,0x90(%rsp)
 1057ff1:	00 
 1057ff2:	48 85 c0             	test   %rax,%rax
 1057ff5:	0f 85 a9 29 00 00    	jne    105a9a4 <_ZN20litchi_perf_baseline56run_docx_source_backed_existing_section_layout_edit_save17h1cd3e0d854c7eea3E+0x5ec4>
 1057ffb:	48 8d bc 24 00 26 00 	lea    0x2600(%rsp),%rdi
 1058002:	00 
 1058003:	48 8d b4 24 40 16 00 	lea    0x1640(%rsp),%rsi
 105800a:	00 
 105800b:	ff 15 37 9a 85 02    	call   *0x2859a37(%rip)        # 38b1a48 <_DYNAMIC+0x11a8>
 1058011:	ff 15 21 8f 85 02    	call   *0x2858f21(%rip)        # 38b0f38 <_DYNAMIC+0x698>
 1058017:	48 89 84 24 a0 07 00 	mov    %rax,0x7a0(%rsp)
 105801e:	00 
 105801f:	89 94 24 a8 07 00 00 	mov    %edx,0x7a8(%rsp)
 1058026:	ba 88 07 00 00       	mov    $0x788,%edx
 105802b:	48 89 ef             	mov    %rbp,%rdi
 105802e:	48 8d b4 24 40 16 00 	lea    0x1640(%rsp),%rsi
 1058035:	00 
 1058036:	ff 15 a4 8a 85 02    	call   *0x2858aa4(%rip)        # 38b0ae0 <memcpy@GLIBC_2.14>
 105803c:	31 db                	xor    %ebx,%ebx
 105803e:	48 8d bc 24 30 08 00 	lea    0x830(%rsp),%rdi
 1058045:	00 
 1058046:	48 89 ee             	mov    %rbp,%rsi
 1058049:	48 8d 94 24 00 13 00 	lea    0x1300(%rsp),%rdx
 1058050:	00 
 1058051:	48 8d 8c 24 c8 1d 00 	lea    0x1dc8(%rsp),%rcx
 1058058:	00 
 1058059:	e8 b2 d2 78 00       	call   17e5310 <_ZN11litchi_docx13source_backed7Package39publish_section_layout_commit_to_stream17hae5727a0407a2493E>
 105805e:	48 8b 9c 24 30 08 00 	mov    0x830(%rsp),%rbx
 1058065:	00 
 1058066:	48 8d 84 24 38 08 00 	lea    0x838(%rsp),%rax
 105806d:	00 
 105806e:	0f 10 00             	movups (%rax),%xmm0
 1058071:	0f 10 48 10          	movups 0x10(%rax),%xmm1
 1058075:	0f 10 50 20          	movups 0x20(%rax),%xmm2
 1058079:	0f 10 58 30          	movups 0x30(%rax),%xmm3
 105807d:	0f 29 84 24 40 0c 00 	movaps %xmm0,0xc40(%rsp)
 1058084:	00 
 1058085:	0f 29 8c 24 50 0c 00 	movaps %xmm1,0xc50(%rsp)
 105808c:	00 
 105808d:	0f 29 94 24 60 0c 00 	movaps %xmm2,0xc60(%rsp)
 1058094:	00 
 1058095:	0f 29 9c 24 70 0c 00 	movaps %xmm3,0xc70(%rsp)
 105809c:	00 
 105809d:	48 8b 40 40          	mov    0x40(%rax),%rax
 10580a1:	48 89 84 24 80 0c 00 	mov    %rax,0xc80(%rsp)
 10580a8:	00 
 10580a9:	48 83 fb 02          	cmp    $0x2,%rbx
 10580ad:	0f 84 fc 22 00 00    	je     105a3af <_ZN20litchi_perf_baseline56run_docx_source_backed_existing_section_layout_edit_save17h1cd3e0d854c7eea3E+0x58cf>
 10580b3:	ba 98 02 00 00       	mov    $0x298,%edx
 10580b8:	48 8d bc 24 b0 21 00 	lea    0x21b0(%rsp),%rdi
 10580bf:	00 
 10580c0:	48 8d b4 24 80 08 00 	lea    0x880(%rsp),%rsi
 10580c7:	00 
 10580c8:	ff 15 12 8a 85 02    	call   *0x2858a12(%rip)        # 38b0ae0 <memcpy@GLIBC_2.14>
 10580ce:	48 8b 84 24 80 0c 00 	mov    0xc80(%rsp),%rax
 10580d5:	00 
 10580d6:	48 8d 8c 24 68 21 00 	lea    0x2168(%rsp),%rcx
 10580dd:	00 
 10580de:	48 89 41 40          	mov    %rax,0x40(%rcx)
 10580e2:	0f 28 84 24 40 0c 00 	movaps 0xc40(%rsp),%xmm0
 10580e9:	00 
 10580ea:	0f 28 8c 24 50 0c 00 	movaps 0xc50(%rsp),%xmm1
 10580f1:	00 
 10580f2:	0f 28 94 24 60 0c 00 	movaps 0xc60(%rsp),%xmm2
 10580f9:	00 
 10580fa:	0f 28 9c 24 70 0c 00 	movaps 0xc70(%rsp),%xmm3
 1058101:	00 
 1058102:	0f 11 59 30          	movups %xmm3,0x30(%rcx)
 1058106:	0f 11 51 20          	movups %xmm2,0x20(%rcx)
 105810a:	0f 11 49 10          	movups %xmm1,0x10(%rcx)
 105810e:	0f 11 01             	movups %xmm0,(%rcx)
 1058111:	48 89 9c 24 60 21 00 	mov    %rbx,0x2160(%rsp)
 1058118:	00 
 1058119:	48 8d 84 24 60 21 00 	lea    0x2160(%rsp),%rax
 1058120:	00 
 1058121:	48 89 84 24 00 27 00 	mov    %rax,0x2700(%rsp)
 1058128:	00 
 1058129:	48 8d bc 24 a0 07 00 	lea    0x7a0(%rsp),%rdi
 1058130:	00 
 1058131:	ff 15 09 8e 85 02    	call   *0x2858e09(%rip)        # 38b0f40 <_DYNAMIC+0x6a0>
 1058137:	89 d1                	mov    %edx,%ecx
 1058139:	ba 00 ca 9a 3b       	mov    $0x3b9aca00,%edx
 105813e:	48 f7 e2             	mul    %rdx
 1058141:	89 c9                	mov    %ecx,%ecx
 1058143:	48 01 c1             	add    %rax,%rcx
 1058146:	48 89 8c 24 f8 00 00 	mov    %rcx,0xf8(%rsp)
 105814d:	00 
 105814e:	48 83 d2 00          	adc    $0x0,%rdx
 1058152:	74 23                	je     1058177 <_ZN20litchi_perf_baseline56run_docx_source_backed_existing_section_layout_edit_save17h1cd3e0d854c7eea3E+0x3697>
 1058154:	be 25 00 00 00       	mov    $0x25,%esi
 1058159:	48 8d 3d a3 9e 0e ff 	lea    -0xf1615d(%rip),%rdi        # 142003 <anon.41d43701a2c2b5e6feb75b74ff96deac.208.llvm.8476651010627807873>
 1058160:	ff 15 22 8a 85 02    	call   *0x2858a22(%rip)        # 38b0b88 <_DYNAMIC+0x2e8>
 1058166:	48 89 94 24 f8 00 00 	mov    %rdx,0xf8(%rsp)
 105816d:	00 
 105816e:	48 85 c0             	test   %rax,%rax
 1058171:	0f 85 5c 28 00 00    	jne    105a9d3 <_ZN20litchi_perf_baseline56run_docx_source_backed_existing_section_layout_edit_save17h1cd3e0d854c7eea3E+0x5ef3>
 1058177:	4c 8b ac 24 a8 00 00 	mov    0xa8(%rsp),%r13
 105817e:	00 
 105817f:	4c 03 ac 24 b0 00 00 	add    0xb0(%rsp),%r13
 1058186:	00 
 1058187:	0f 82 1e 1d 00 00    	jb     1059eab <_ZN20litchi_perf_baseline56run_docx_source_backed_existing_section_layout_edit_save17h1cd3e0d854c7eea3E+0x53cb>
 105818d:	4c 03 ac 24 a0 00 00 	add    0xa0(%rsp),%r13
 1058194:	00 
 1058195:	0f 82 10 1d 00 00    	jb     1059eab <_ZN20litchi_perf_baseline56run_docx_source_backed_existing_section_layout_edit_save17h1cd3e0d854c7eea3E+0x53cb>
 105819b:	4c 03 ac 24 98 00 00 	add    0x98(%rsp),%r13
 10581a2:	00 
 10581a3:	0f 82 02 1d 00 00    	jb     1059eab <_ZN20litchi_perf_baseline56run_docx_source_backed_existing_section_layout_edit_save17h1cd3e0d854c7eea3E+0x53cb>
 10581a9:	4c 03 ac 24 90 00 00 	add    0x90(%rsp),%r13
 10581b0:	00 
 10581b1:	0f 82 f4 1c 00 00    	jb     1059eab <_ZN20litchi_perf_baseline56run_docx_source_backed_existing_section_layout_edit_save17h1cd3e0d854c7eea3E+0x53cb>
 10581b7:	4c 03 ac 24 f8 00 00 	add    0xf8(%rsp),%r13
 10581be:	00 
 10581bf:	0f 82 e6 1c 00 00    	jb     1059eab <_ZN20litchi_perf_baseline56run_docx_source_backed_existing_section_layout_edit_save17h1cd3e0d854c7eea3E+0x53cb>
 10581c5:	b8 02 00 00 00       	mov    $0x2,%eax
 10581ca:	48 89 84 24 50 01 00 	mov    %rax,0x150(%rsp)
 10581d1:	00 
 10581d2:	41 80 fc 03          	cmp    $0x3,%r12b
 10581d6:	0f 82 78 02 00 00    	jb     1058454 <_ZN20litchi_perf_baseline56run_docx_source_backed_existing_section_layout_edit_save17h1cd3e0d854c7eea3E+0x3974>
 10581dc:	48 8b 5c 24 40       	mov    0x40(%rsp),%rbx
 10581e1:	48 85 db             	test   %rbx,%rbx
 10581e4:	0f 84 26 32 00 00    	je     105b410 <_ZN20litchi_perf_baseline56run_docx_source_backed_existing_section_layout_edit_save17h1cd3e0d854c7eea3E+0x6930>
 10581ea:	48 8d 73 10          	lea    0x10(%rbx),%rsi
 10581ee:	48 8d bc 24 18 25 00 	lea    0x2518(%rsp),%rdi
 10581f5:	00 
 10581f6:	e8 45 41 db ff       	call   e0c340 <_ZN20litchi_perf_baseline23DocxProviderReadCounter8snapshot17h19c1fa0853483bb2E>
 10581fb:	48 8b 74 24 70       	mov    0x70(%rsp),%rsi
 1058200:	48 85 f6             	test   %rsi,%rsi
 1058203:	0f 84 ff 00 00 00    	je     1058308 <_ZN20litchi_perf_baseline56run_docx_source_backed_existing_section_layout_edit_save17h1cd3e0d854c7eea3E+0x3828>
 1058209:	48 83 c6 10          	add    $0x10,%rsi
 105820d:	48 89 ef             	mov    %rbp,%rdi
 1058210:	ff 15 f2 a5 85 02    	call   *0x285a5f2(%rip)        # 38b2808 <_DYNAMIC+0x1f68>
 1058216:	4c 8b b4 24 00 27 00 	mov    0x2700(%rsp),%r14
 105821d:	00 
 105821e:	48 8b 9c 24 08 27 00 	mov    0x2708(%rsp),%rbx
 1058225:	00 
 1058226:	0f 10 84 24 10 27 00 	movups 0x2710(%rsp),%xmm0
 105822d:	00 
 105822e:	0f 29 84 24 50 01 00 	movaps %xmm0,0x150(%rsp)
 1058235:	00 
 1058236:	0f 10 84 24 20 27 00 	movups 0x2720(%rsp),%xmm0
 105823d:	00 
 105823e:	0f 29 84 24 f0 02 00 	movaps %xmm0,0x2f0(%rsp)
 1058245:	00 
 1058246:	0f 10 84 24 30 27 00 	movups 0x2730(%rsp),%xmm0
 105824d:	00 
 105824e:	0f 29 84 24 60 02 00 	movaps %xmm0,0x260(%rsp)
 1058255:	00 
 1058256:	0f 10 84 24 40 27 00 	movups 0x2740(%rsp),%xmm0
 105825d:	00 
 105825e:	0f 29 84 24 40 01 00 	movaps %xmm0,0x140(%rsp)
 1058265:	00 
 1058266:	4c 8b bc 24 50 27 00 	mov    0x2750(%rsp),%r15
 105826d:	00 
 105826e:	ba 90 00 00 00       	mov    $0x90,%edx
 1058273:	48 8d bc 24 40 0c 00 	lea    0xc40(%rsp),%rdi
 105827a:	00 
 105827b:	48 8d b4 24 58 27 00 	lea    0x2758(%rsp),%rsi
 1058282:	00 
 1058283:	ff 15 57 88 85 02    	call   *0x2858857(%rip)        # 38b0ae0 <memcpy@GLIBC_2.14>
 1058289:	49 83 fe 02          	cmp    $0x2,%r14
 105828d:	0f 84 70 27 00 00    	je     105aa03 <_ZN20litchi_perf_baseline56run_docx_source_backed_existing_section_layout_edit_save17h1cd3e0d854c7eea3E+0x5f23>
 1058293:	ba 90 00 00 00       	mov    $0x90,%edx
 1058298:	48 8d bc 24 88 08 00 	lea    0x888(%rsp),%rdi
 105829f:	00 
 10582a0:	48 8d b4 24 40 0c 00 	lea    0xc40(%rsp),%rsi
 10582a7:	00 
 10582a8:	ff 15 32 88 85 02    	call   *0x2858832(%rip)        # 38b0ae0 <memcpy@GLIBC_2.14>
 10582ae:	4c 89 b4 24 30 08 00 	mov    %r14,0x830(%rsp)
 10582b5:	00 
 10582b6:	48 89 9c 24 38 08 00 	mov    %rbx,0x838(%rsp)
 10582bd:	00 
 10582be:	0f 28 84 24 50 01 00 	movaps 0x150(%rsp),%xmm0
 10582c5:	00 
 10582c6:	0f 11 84 24 40 08 00 	movups %xmm0,0x840(%rsp)
 10582cd:	00 
 10582ce:	0f 28 84 24 f0 02 00 	movaps 0x2f0(%rsp),%xmm0
 10582d5:	00 
 10582d6:	0f 11 84 24 50 08 00 	movups %xmm0,0x850(%rsp)
 10582dd:	00 
 10582de:	0f 28 84 24 60 02 00 	movaps 0x260(%rsp),%xmm0
 10582e5:	00 
 10582e6:	0f 11 84 24 60 08 00 	movups %xmm0,0x860(%rsp)
 10582ed:	00 
 10582ee:	0f 28 84 24 40 01 00 	movaps 0x140(%rsp),%xmm0
 10582f5:	00 
 10582f6:	0f 11 84 24 70 08 00 	movups %xmm0,0x870(%rsp)
 10582fd:	00 
 10582fe:	4c 89 bc 24 80 08 00 	mov    %r15,0x880(%rsp)
 1058305:	00 
 1058306:	eb 10                	jmp    1058318 <_ZN20litchi_perf_baseline56run_docx_source_backed_existing_section_layout_edit_save17h1cd3e0d854c7eea3E+0x3838>
 1058308:	48 8d bc 24 30 08 00 	lea    0x830(%rsp),%rdi
 105830f:	00 
 1058310:	48 89 de             	mov    %rbx,%rsi
 1058313:	e8 58 5d 00 00       	call   105e070 <_ZN20litchi_perf_baseline56run_docx_source_backed_existing_section_layout_edit_save28_$u7b$$u7b$closure$u7d$$u7d$17haca16e54ddf99a44E>
 1058318:	48 8b 84 24 c8 00 00 	mov    0xc8(%rsp),%rax
 105831f:	00 
 1058320:	83 c0 fd             	add    $0xfffffffd,%eax
 1058323:	83 f8 02             	cmp    $0x2,%eax
 1058326:	73 12                	jae    105833a <_ZN20litchi_perf_baseline56run_docx_source_backed_existing_section_layout_edit_save17h1cd3e0d854c7eea3E+0x385a>
 1058328:	48 8b 84 24 80 08 00 	mov    0x880(%rsp),%rax
 105832f:	00 
 1058330:	48 89 84 24 60 02 00 	mov    %rax,0x260(%rsp)
 1058337:	00 
 1058338:	eb 47                	jmp    1058381 <_ZN20litchi_perf_baseline56run_docx_source_backed_existing_section_layout_edit_save17h1cd3e0d854c7eea3E+0x38a1>
 105833a:	83 bc 24 c8 00 00 00 	cmpl   $0x5,0xc8(%rsp)
 1058341:	05 
 1058342:	75 04                	jne    1058348 <_ZN20litchi_perf_baseline56run_docx_source_backed_existing_section_layout_edit_save17h1cd3e0d854c7eea3E+0x3868>
 1058344:	31 c0                	xor    %eax,%eax
 1058346:	eb 05                	jmp    105834d <_ZN20litchi_perf_baseline56run_docx_source_backed_existing_section_layout_edit_save17h1cd3e0d854c7eea3E+0x386d>
 1058348:	b8 40 42 0f 00       	mov    $0xf4240,%eax
 105834d:	48 f7 a4 24 70 08 00 	mulq   0x870(%rsp)
 1058354:	00 
 1058355:	48 89 84 24 60 02 00 	mov    %rax,0x260(%rsp)
 105835c:	00 
 105835d:	0f 80 37 27 00 00    	jo     105aa9a <_ZN20litchi_perf_baseline56run_docx_source_backed_existing_section_layout_edit_save17h1cd3e0d854c7eea3E+0x5fba>
 1058363:	48 8b 84 24 60 02 00 	mov    0x260(%rsp),%rax
 105836a:	00 
 105836b:	48 03 84 24 80 08 00 	add    0x880(%rsp),%rax
 1058372:	00 
 1058373:	48 89 84 24 60 02 00 	mov    %rax,0x260(%rsp)
 105837a:	00 
 105837b:	0f 82 50 30 00 00    	jb     105b3d1 <_ZN20litchi_perf_baseline56run_docx_source_backed_existing_section_layout_edit_save17h1cd3e0d854c7eea3E+0x68f1>
 1058381:	48 8b 84 24 30 08 00 	mov    0x830(%rsp),%rax
 1058388:	00 
 1058389:	48 89 84 24 50 01 00 	mov    %rax,0x150(%rsp)
 1058390:	00 
 1058391:	48 8b 84 24 50 08 00 	mov    0x850(%rsp),%rax
 1058398:	00 
 1058399:	48 89 84 24 78 01 00 	mov    %rax,0x178(%rsp)
 10583a0:	00 
 10583a1:	48 8d 84 24 38 08 00 	lea    0x838(%rsp),%rax
 10583a8:	00 
 10583a9:	0f 10 00             	movups (%rax),%xmm0
 10583ac:	0f 29 84 24 d0 07 00 	movaps %xmm0,0x7d0(%rsp)
 10583b3:	00 
 10583b4:	48 8b 40 10          	mov    0x10(%rax),%rax
 10583b8:	48 89 84 24 e0 07 00 	mov    %rax,0x7e0(%rsp)
 10583bf:	00 
 10583c0:	48 8b 84 24 58 08 00 	mov    0x858(%rsp),%rax
 10583c7:	00 
 10583c8:	48 89 84 24 70 01 00 	mov    %rax,0x170(%rsp)
 10583cf:	00 
 10583d0:	ba b8 00 00 00       	mov    $0xb8,%edx
 10583d5:	48 8d bc 24 88 9e 00 	lea    0x9e88(%rsp),%rdi
 10583dc:	00 
 10583dd:	48 8d b4 24 60 08 00 	lea    0x860(%rsp),%rsi
 10583e4:	00 
 10583e5:	48 8b 1d f4 86 85 02 	mov    0x28586f4(%rip),%rbx        # 38b0ae0 <memcpy@GLIBC_2.14>
 10583ec:	ff d3                	call   *%rbx
 10583ee:	48 8b 84 24 18 25 00 	mov    0x2518(%rsp),%rax
 10583f5:	00 
 10583f6:	48 89 84 24 30 07 00 	mov    %rax,0x730(%rsp)
 10583fd:	00 
 10583fe:	48 8b 84 24 38 25 00 	mov    0x2538(%rsp),%rax
 1058405:	00 
 1058406:	48 89 84 24 68 01 00 	mov    %rax,0x168(%rsp)
 105840d:	00 
 105840e:	48 8d 84 24 20 25 00 	lea    0x2520(%rsp),%rax
 1058415:	00 
 1058416:	0f 10 00             	movups (%rax),%xmm0
 1058419:	0f 29 84 24 b0 07 00 	movaps %xmm0,0x7b0(%rsp)
 1058420:	00 
 1058421:	48 8b 40 10          	mov    0x10(%rax),%rax
 1058425:	48 89 84 24 c0 07 00 	mov    %rax,0x7c0(%rsp)
 105842c:	00 
 105842d:	48 8b 84 24 40 25 00 	mov    0x2540(%rsp),%rax
 1058434:	00 
 1058435:	48 89 84 24 40 03 00 	mov    %rax,0x340(%rsp)
 105843c:	00 
 105843d:	ba b8 00 00 00       	mov    $0xb8,%edx
 1058442:	48 8d bc 24 d0 9d 00 	lea    0x9dd0(%rsp),%rdi
 1058449:	00 
 105844a:	48 8d b4 24 48 25 00 	lea    0x2548(%rsp),%rsi
 1058451:	00 
 1058452:	ff d3                	call   *%rbx
 1058454:	4c 8b 7c 24 40       	mov    0x40(%rsp),%r15
 1058459:	4d 85 ff             	test   %r15,%r15
 105845c:	74 50                	je     10584ae <_ZN20litchi_perf_baseline56run_docx_source_backed_existing_section_layout_edit_save17h1cd3e0d854c7eea3E+0x39ce>
 105845e:	49 8b 47 10          	mov    0x10(%r15),%rax
 1058462:	49 8b 4f 18          	mov    0x18(%r15),%rcx
 1058466:	48 8b 51 10          	mov    0x10(%rcx),%rdx
 105846a:	48 ff ca             	dec    %rdx
 105846d:	48 83 e2 f0          	and    $0xfffffffffffffff0,%rdx
 1058471:	48 8d 34 10          	lea    (%rax,%rdx,1),%rsi
 1058475:	48 83 c6 10          	add    $0x10,%rsi
 1058479:	48 89 ef             	mov    %rbp,%rdi
 105847c:	ff 51 38             	call   *0x38(%rcx)
 105847f:	48 8b 84 24 08 27 00 	mov    0x2708(%rsp),%rax
 1058486:	00 
 1058487:	48 89 84 24 40 01 00 	mov    %rax,0x140(%rsp)
 105848e:	00 
 105848f:	48 83 bc 24 00 27 00 	cmpq   $0x1,0x2700(%rsp)
 1058496:	00 01 
 1058498:	0f 84 c4 24 00 00    	je     105a962 <_ZN20litchi_perf_baseline56run_docx_source_backed_existing_section_layout_edit_save17h1cd3e0d854c7eea3E+0x5e82>
 105849e:	48 8b 84 24 10 27 00 	mov    0x2710(%rsp),%rax
 10584a5:	00 
 10584a6:	48 89 84 24 80 01 00 	mov    %rax,0x180(%rsp)
 10584ad:	00 
 10584ae:	48 8b 44 24 38       	mov    0x38(%rsp),%rax
 10584b3:	48 89 84 24 f0 02 00 	mov    %rax,0x2f0(%rsp)
 10584ba:	00 
 10584bb:	48 85 c0             	test   %rax,%rax
 10584be:	74 2d                	je     10584ed <_ZN20litchi_perf_baseline56run_docx_source_backed_existing_section_layout_edit_save17h1cd3e0d854c7eea3E+0x3a0d>
 10584c0:	48 89 ef             	mov    %rbp,%rdi
 10584c3:	48 8d 74 24 38       	lea    0x38(%rsp),%rsi
 10584c8:	e8 53 58 00 00       	call   105dd20 <_ZN20litchi_perf_baseline56run_docx_source_backed_existing_section_layout_edit_save28_$u7b$$u7b$closure$u7d$$u7d$17h1fb90285dd6bc22cE>
 10584cd:	48 8b 84 24 d0 27 00 	mov    0x27d0(%rsp),%rax
 10584d4:	00 
 10584d5:	48 89 84 24 48 07 00 	mov    %rax,0x748(%rsp)
 10584dc:	00 
 10584dd:	48 8b 84 24 d8 27 00 	mov    0x27d8(%rsp),%rax
 10584e4:	00 
 10584e5:	48 89 84 24 40 07 00 	mov    %rax,0x740(%rsp)
 10584ec:	00 
 10584ed:	48 8d bc 24 70 15 00 	lea    0x1570(%rsp),%rdi
 10584f4:	00 
 10584f5:	48 be 02 00 02 00 02 	movabs $0x200020002,%rsi
 10584fc:	00 00 00 
 10584ff:	e8 5c ca f6 ff       	call   fc4f60 <_ZN20litchi_perf_baseline35verify_docx_section_layout_snapshot17h82d8b060ba051d24E>
 1058504:	48 85 c0             	test   %rax,%rax
 1058507:	0f 85 4e 1b 00 00    	jne    105a05b <_ZN20litchi_perf_baseline56run_docx_source_backed_existing_section_layout_edit_save17h1cd3e0d854c7eea3E+0x557b>
 105850d:	48 8b 84 24 e8 1e 00 	mov    0x1ee8(%rsp),%rax
 1058514:	00 
 1058515:	48 8b 8c 24 f0 1e 00 	mov    0x1ef0(%rsp),%rcx
 105851c:	00 
 105851d:	48 8b 50 20          	mov    0x20(%rax),%rdx
 1058521:	48 3b 51 20          	cmp    0x20(%rcx),%rdx
 1058525:	75 16                	jne    105853d <_ZN20litchi_perf_baseline56run_docx_source_backed_existing_section_layout_edit_save17h1cd3e0d854c7eea3E+0x3a5d>
 1058527:	48 8b 71 18          	mov    0x18(%rcx),%rsi
 105852b:	48 8b 78 18          	mov    0x18(%rax),%rdi
 105852f:	ff 15 4b 86 85 02    	call   *0x285864b(%rip)        # 38b0b80 <bcmp@GLIBC_2.2.5>
 1058535:	85 c0                	test   %eax,%eax
 1058537:	0f 84 0c 1b 00 00    	je     105a049 <_ZN20litchi_perf_baseline56run_docx_source_backed_existing_section_layout_edit_save17h1cd3e0d854c7eea3E+0x5569>
 105853d:	48 83 bc 24 f0 1d 00 	cmpq   $0x0,0x1df0(%rsp)
 1058544:	00 00 
 1058546:	0f 84 c6 33 00 00    	je     105b912 <_ZN20litchi_perf_baseline56run_docx_source_backed_existing_section_layout_edit_save17h1cd3e0d854c7eea3E+0x6e32>
 105854c:	48 8b b4 24 e8 1d 00 	mov    0x1de8(%rsp),%rsi
 1058553:	00 
 1058554:	48 83 c6 10          	add    $0x10,%rsi
 1058558:	48 89 ef             	mov    %rbp,%rdi
 105855b:	ff 15 7f a0 85 02    	call   *0x285a07f(%rip)        # 38b25e0 <_DYNAMIC+0x1d40>
 1058561:	83 bc 24 00 27 00 00 	cmpl   $0x2,0x2700(%rsp)
 1058568:	02 
 1058569:	0f 84 da 1a 00 00    	je     105a049 <_ZN20litchi_perf_baseline56run_docx_source_backed_existing_section_layout_edit_save17h1cd3e0d854c7eea3E+0x5569>
 105856f:	48 8b 84 24 10 27 00 	mov    0x2710(%rsp),%rax
 1058576:	00 
 1058577:	0f b7 9c 24 28 27 00 	movzwl 0x2728(%rsp),%ebx
 105857e:	00 
 105857f:	48 85 c0             	test   %rax,%rax
 1058582:	74 1b                	je     105859f <_ZN20litchi_perf_baseline56run_docx_source_backed_existing_section_layout_edit_save17h1cd3e0d854c7eea3E+0x3abf>
 1058584:	48 8b bc 24 18 27 00 	mov    0x2718(%rsp),%rdi
 105858b:	00 
 105858c:	48 c1 e0 03          	shl    $0x3,%rax
 1058590:	48 8d 34 40          	lea    (%rax,%rax,2),%rsi
 1058594:	ba 08 00 00 00       	mov    $0x8,%edx
 1058599:	ff 15 51 85 85 02    	call   *0x2858551(%rip)        # 38b0af0 <_DYNAMIC+0x250>
 105859f:	66 83 fb 03          	cmp    $0x3,%bx
 10585a3:	0f 85 a0 1a 00 00    	jne    105a049 <_ZN20litchi_perf_baseline56run_docx_source_backed_existing_section_layout_edit_save17h1cd3e0d854c7eea3E+0x5569>
 10585a9:	48 89 ef             	mov    %rbp,%rdi
 10585ac:	48 8d b4 24 98 1e 00 	lea    0x1e98(%rsp),%rsi
 10585b3:	00 
 10585b4:	48 8d 94 24 70 15 00 	lea    0x1570(%rsp),%rdx
 10585bb:	00 
 10585bc:	ff 15 4e a2 85 02    	call   *0x285a24e(%rip)        # 38b2810 <_DYNAMIC+0x1f70>
 10585c2:	48 8b 84 24 00 27 00 	mov    0x2700(%rsp),%rax
 10585c9:	00 
 10585ca:	48 8d 8c 24 10 27 00 	lea    0x2710(%rsp),%rcx
 10585d1:	00 
 10585d2:	0f 10 41 f8          	movups -0x8(%rcx),%xmm0
 10585d6:	0f 10 49 08          	movups 0x8(%rcx),%xmm1
 10585da:	0f 10 51 18          	movups 0x18(%rcx),%xmm2
 10585de:	0f 10 59 28          	movups 0x28(%rcx),%xmm3
 10585e2:	0f 29 84 24 30 08 00 	movaps %xmm0,0x830(%rsp)
 10585e9:	00 
 10585ea:	0f 29 8c 24 40 08 00 	movaps %xmm1,0x840(%rsp)
 10585f1:	00 
 10585f2:	0f 29 94 24 50 08 00 	movaps %xmm2,0x850(%rsp)
 10585f9:	00 
 10585fa:	0f 29 9c 24 60 08 00 	movaps %xmm3,0x860(%rsp)
 1058601:	00 
 1058602:	48 8b 49 38          	mov    0x38(%rcx),%rcx
 1058606:	48 89 8c 24 70 08 00 	mov    %rcx,0x870(%rsp)
 105860d:	00 
 105860e:	48 83 f8 02          	cmp    $0x2,%rax
 1058612:	0f 84 ec 1e 00 00    	je     105a504 <_ZN20litchi_perf_baseline56run_docx_source_backed_existing_section_layout_edit_save17h1cd3e0d854c7eea3E+0x5a24>
 1058618:	48 8d 8c 24 50 27 00 	lea    0x2750(%rsp),%rcx
 105861f:	00 
 1058620:	0f 10 41 70          	movups 0x70(%rcx),%xmm0
 1058624:	48 8d 94 24 50 24 00 	lea    0x2450(%rsp),%rdx
 105862b:	00 
 105862c:	0f 11 82 b8 00 00 00 	movups %xmm0,0xb8(%rdx)
 1058633:	0f 10 41 60          	movups 0x60(%rcx),%xmm0
 1058637:	0f 11 82 a8 00 00 00 	movups %xmm0,0xa8(%rdx)
 105863e:	0f 10 41 50          	movups 0x50(%rcx),%xmm0
 1058642:	0f 11 82 98 00 00 00 	movups %xmm0,0x98(%rdx)
 1058649:	0f 10 41 40          	movups 0x40(%rcx),%xmm0
 105864d:	0f 11 82 88 00 00 00 	movups %xmm0,0x88(%rdx)
 1058654:	0f 10 01             	movups (%rcx),%xmm0
 1058657:	0f 10 49 10          	movups 0x10(%rcx),%xmm1
 105865b:	0f 10 51 20          	movups 0x20(%rcx),%xmm2
 105865f:	0f 10 59 30          	movups 0x30(%rcx),%xmm3
 1058663:	0f 11 5a 78          	movups %xmm3,0x78(%rdx)
 1058667:	0f 11 52 68          	movups %xmm2,0x68(%rdx)
 105866b:	0f 11 4a 58          	movups %xmm1,0x58(%rdx)
 105866f:	0f 11 42 48          	movups %xmm0,0x48(%rdx)
 1058673:	48 8b 8c 24 70 08 00 	mov    0x870(%rsp),%rcx
 105867a:	00 
 105867b:	48 89 4a 40          	mov    %rcx,0x40(%rdx)
 105867f:	0f 28 84 24 30 08 00 	movaps 0x830(%rsp),%xmm0
 1058686:	00 
 1058687:	0f 28 8c 24 40 08 00 	movaps 0x840(%rsp),%xmm1
 105868e:	00 
 105868f:	0f 28 94 24 50 08 00 	movaps 0x850(%rsp),%xmm2
 1058696:	00 
 1058697:	0f 28 9c 24 60 08 00 	movaps 0x860(%rsp),%xmm3
 105869e:	00 
 105869f:	0f 11 5a 30          	movups %xmm3,0x30(%rdx)
 10586a3:	0f 11 52 20          	movups %xmm2,0x20(%rdx)
 10586a7:	0f 11 4a 10          	movups %xmm1,0x10(%rdx)
 10586ab:	0f 11 02             	movups %xmm0,(%rdx)
 10586ae:	48 89 84 24 48 24 00 	mov    %rax,0x2448(%rsp)
 10586b5:	00 
 10586b6:	48 8d 9c 24 30 08 00 	lea    0x830(%rsp),%rbx
 10586bd:	00 
 10586be:	48 89 df             	mov    %rbx,%rdi
 10586c1:	48 8d b4 24 98 1e 00 	lea    0x1e98(%rsp),%rsi
 10586c8:	00 
 10586c9:	e8 62 8c cd ff       	call   d31330 <_ZN11litchi_docx7section6layout5Patch7inverse17hdf9c46b982a6db61E>
 10586ce:	48 89 ef             	mov    %rbp,%rdi
 10586d1:	48 89 de             	mov    %rbx,%rsi
 10586d4:	48 8d 94 24 48 24 00 	lea    0x2448(%rsp),%rdx
 10586db:	00 
 10586dc:	ff 15 2e a1 85 02    	call   *0x285a12e(%rip)        # 38b2810 <_DYNAMIC+0x1f70>
 10586e2:	48 8b 84 24 00 27 00 	mov    0x2700(%rsp),%rax
 10586e9:	00 
 10586ea:	48 8d 8c 24 10 27 00 	lea    0x2710(%rsp),%rcx
 10586f1:	00 
 10586f2:	0f 10 41 f8          	movups -0x8(%rcx),%xmm0
 10586f6:	0f 10 49 08          	movups 0x8(%rcx),%xmm1
 10586fa:	0f 10 51 18          	movups 0x18(%rcx),%xmm2
 10586fe:	0f 10 59 28          	movups 0x28(%rcx),%xmm3
 1058702:	0f 29 84 24 00 02 00 	movaps %xmm0,0x200(%rsp)
 1058709:	00 
 105870a:	0f 29 8c 24 10 02 00 	movaps %xmm1,0x210(%rsp)
 1058711:	00 
 1058712:	0f 29 94 24 20 02 00 	movaps %xmm2,0x220(%rsp)
 1058719:	00 
 105871a:	0f 29 9c 24 30 02 00 	movaps %xmm3,0x230(%rsp)
 1058721:	00 
 1058722:	48 8b 49 38          	mov    0x38(%rcx),%rcx
 1058726:	48 89 8c 24 40 02 00 	mov    %rcx,0x240(%rsp)
 105872d:	00 
 105872e:	48 83 f8 02          	cmp    $0x2,%rax
 1058732:	0f 84 e2 1e 00 00    	je     105a61a <_ZN20litchi_perf_baseline56run_docx_source_backed_existing_section_layout_edit_save17h1cd3e0d854c7eea3E+0x5b3a>
 1058738:	48 8d 8c 24 50 27 00 	lea    0x2750(%rsp),%rcx
 105873f:	00 
 1058740:	0f 10 41 70          	movups 0x70(%rcx),%xmm0
 1058744:	48 8d 94 24 48 0c 00 	lea    0xc48(%rsp),%rdx
 105874b:	00 
 105874c:	0f 11 82 b8 00 00 00 	movups %xmm0,0xb8(%rdx)
 1058753:	0f 10 41 60          	movups 0x60(%rcx),%xmm0
 1058757:	0f 11 82 a8 00 00 00 	movups %xmm0,0xa8(%rdx)
 105875e:	0f 10 41 50          	movups 0x50(%rcx),%xmm0
 1058762:	0f 11 82 98 00 00 00 	movups %xmm0,0x98(%rdx)
 1058769:	0f 10 41 40          	movups 0x40(%rcx),%xmm0
 105876d:	0f 11 82 88 00 00 00 	movups %xmm0,0x88(%rdx)
 1058774:	0f 10 01             	movups (%rcx),%xmm0
 1058777:	0f 10 49 10          	movups 0x10(%rcx),%xmm1
 105877b:	0f 10 51 20          	movups 0x20(%rcx),%xmm2
 105877f:	0f 10 59 30          	movups 0x30(%rcx),%xmm3
 1058783:	0f 11 5a 78          	movups %xmm3,0x78(%rdx)
 1058787:	0f 11 52 68          	movups %xmm2,0x68(%rdx)
 105878b:	0f 11 4a 58          	movups %xmm1,0x58(%rdx)
 105878f:	0f 11 42 48          	movups %xmm0,0x48(%rdx)
 1058793:	48 8b 8c 24 40 02 00 	mov    0x240(%rsp),%rcx
 105879a:	00 
 105879b:	48 89 4a 40          	mov    %rcx,0x40(%rdx)
 105879f:	0f 28 84 24 00 02 00 	movaps 0x200(%rsp),%xmm0
 10587a6:	00 
 10587a7:	0f 28 8c 24 10 02 00 	movaps 0x210(%rsp),%xmm1
 10587ae:	00 
 10587af:	0f 28 94 24 20 02 00 	movaps 0x220(%rsp),%xmm2
 10587b6:	00 
 10587b7:	0f 28 9c 24 30 02 00 	movaps 0x230(%rsp),%xmm3
 10587be:	00 
 10587bf:	0f 11 5a 30          	movups %xmm3,0x30(%rdx)
 10587c3:	0f 11 52 20          	movups %xmm2,0x20(%rdx)
 10587c7:	0f 11 4a 10          	movups %xmm1,0x10(%rdx)
 10587cb:	0f 11 02             	movups %xmm0,(%rdx)
 10587ce:	48 89 84 24 40 0c 00 	mov    %rax,0xc40(%rsp)
 10587d5:	00 
 10587d6:	48 8d bc 24 30 08 00 	lea    0x830(%rsp),%rdi
 10587dd:	00 
 10587de:	e8 bd ee 01 00       	call   10776a0 <_ZN4core3ptr56drop_in_place$LT$litchi_docx..section..layout..Patch$GT$17h79def4891b154840E>
 10587e3:	48 8b 94 24 68 0c 00 	mov    0xc68(%rsp),%rdx
 10587ea:	00 
 10587eb:	48 3b 94 24 98 15 00 	cmp    0x1598(%rsp),%rdx
 10587f2:	00 
 10587f3:	0f 85 20 1f 00 00    	jne    105a719 <_ZN20litchi_perf_baseline56run_docx_source_backed_existing_section_layout_edit_save17h1cd3e0d854c7eea3E+0x5c39>
 10587f9:	48 8b b4 24 90 15 00 	mov    0x1590(%rsp),%rsi
 1058800:	00 
 1058801:	48 8b bc 24 60 0c 00 	mov    0xc60(%rsp),%rdi
 1058808:	00 
 1058809:	48 83 c7 10          	add    $0x10,%rdi
 105880d:	48 83 c6 10          	add    $0x10,%rsi
 1058811:	e8 8a 35 91 00       	call   196bda0 <_ZN63_$LT$A$u20$as$u20$core..slice..cmp..SlicePartialEq$LT$B$GT$$GT$17equal_same_length17h25ddecaa5f48eeb6E>
 1058816:	84 c0                	test   %al,%al
 1058818:	0f 84 fb 1e 00 00    	je     105a719 <_ZN20litchi_perf_baseline56run_docx_source_backed_existing_section_layout_edit_save17h1cd3e0d854c7eea3E+0x5c39>
 105881e:	48 89 ef             	mov    %rbp,%rdi
 1058821:	48 8d b4 24 98 1e 00 	lea    0x1e98(%rsp),%rsi
 1058828:	00 
 1058829:	48 8d 94 24 c8 1d 00 	lea    0x1dc8(%rsp),%rdx
 1058830:	00 
 1058831:	ff 15 d9 9f 85 02    	call   *0x2859fd9(%rip)        # 38b2810 <_DYNAMIC+0x1f70>
 1058837:	83 bc 24 00 27 00 00 	cmpl   $0x2,0x2700(%rsp)
 105883e:	02 
 105883f:	0f 85 c7 1e 00 00    	jne    105a70c <_ZN20litchi_perf_baseline56run_docx_source_backed_existing_section_layout_edit_save17h1cd3e0d854c7eea3E+0x5c2c>
 1058845:	48 89 ef             	mov    %rbp,%rdi
 1058848:	e8 73 e8 00 00       	call   10670c0 <_ZN4core3ptr115drop_in_place$LT$core..result..Result$LT$litchi_docx..section..layout..Snapshot$C$litchi_docx..error..Error$GT$$GT$17h64781b88c065a8aeE>
 105884d:	48 89 ef             	mov    %rbp,%rdi
 1058850:	48 8d b4 24 00 23 00 	lea    0x2300(%rsp),%rsi
 1058857:	00 
 1058858:	48 8d 94 24 60 21 00 	lea    0x2160(%rsp),%rdx
 105885f:	00 
 1058860:	ff 15 aa 9f 85 02    	call   *0x2859faa(%rip)        # 38b2810 <_DYNAMIC+0x1f70>
 1058866:	83 bc 24 00 27 00 00 	cmpl   $0x2,0x2700(%rsp)
 105886d:	02 
 105886e:	0f 84 b9 1e 00 00    	je     105a72d <_ZN20litchi_perf_baseline56run_docx_source_backed_existing_section_layout_edit_save17h1cd3e0d854c7eea3E+0x5c4d>
 1058874:	48 89 ef             	mov    %rbp,%rdi
 1058877:	e8 44 e8 00 00       	call   10670c0 <_ZN4core3ptr115drop_in_place$LT$core..result..Result$LT$litchi_docx..section..layout..Snapshot$C$litchi_docx..error..Error$GT$$GT$17h64781b88c065a8aeE>
 105887c:	48 8b 9c 24 58 14 00 	mov    0x1458(%rsp),%rbx
 1058883:	00 
 1058884:	48 3b 9c 24 d0 01 00 	cmp    0x1d0(%rsp),%rbx
 105888b:	00 
 105888c:	0f 85 bd 18 00 00    	jne    105a14f <_ZN20litchi_perf_baseline56run_docx_source_backed_existing_section_layout_edit_save17h1cd3e0d854c7eea3E+0x566f>
 1058892:	4c 8b b4 24 50 14 00 	mov    0x1450(%rsp),%r14
 1058899:	00 
 105889a:	48 8b b4 24 c8 01 00 	mov    0x1c8(%rsp),%rsi
 10588a1:	00 
 10588a2:	4c 89 f7             	mov    %r14,%rdi
 10588a5:	48 89 da             	mov    %rbx,%rdx
 10588a8:	ff 15 d2 82 85 02    	call   *0x28582d2(%rip)        # 38b0b80 <bcmp@GLIBC_2.2.5>
 10588ae:	85 c0                	test   %eax,%eax
 10588b0:	0f 85 99 18 00 00    	jne    105a14f <_ZN20litchi_perf_baseline56run_docx_source_backed_existing_section_layout_edit_save17h1cd3e0d854c7eea3E+0x566f>
 10588b6:	48 8b bc 24 28 01 00 	mov    0x128(%rsp),%rdi
 10588bd:	00 
 10588be:	4c 89 f6             	mov    %r14,%rsi
 10588c1:	48 89 da             	mov    %rbx,%rdx
 10588c4:	e8 57 4b f3 ff       	call   f8d420 <_ZN20litchi_perf_baseline33verify_docx_section_layout_output17h3d722e8449906e5aE>
 10588c9:	48 85 c0             	test   %rax,%rax
 10588cc:	0f 85 7a 1e 00 00    	jne    105a74c <_ZN20litchi_perf_baseline56run_docx_source_backed_existing_section_layout_edit_save17h1cd3e0d854c7eea3E+0x5c6c>
 10588d2:	48 8d bc 24 00 02 00 	lea    0x200(%rsp),%rdi
 10588d9:	00 
 10588da:	48 8d b4 24 48 14 00 	lea    0x1448(%rsp),%rsi
 10588e1:	00 
 10588e2:	e8 79 b7 91 00       	call   1974060 <_ZN67_$LT$alloc..vec..Vec$LT$T$C$A$GT$$u20$as$u20$core..clone..Clone$GT$5clone17hb7c696c915b55e1cE>
 10588e7:	48 8d bc 24 e8 01 00 	lea    0x1e8(%rsp),%rdi
 10588ee:	00 
 10588ef:	48 8d b4 24 00 02 00 	lea    0x200(%rsp),%rsi
 10588f6:	00 
 10588f7:	ff 15 eb 87 85 02    	call   *0x28587eb(%rip)        # 38b10e8 <_DYNAMIC+0x848>
 10588fd:	48 c7 84 24 00 02 00 	movq   $0x1,0x200(%rsp)
 1058904:	00 01 00 00 00 
 1058909:	48 c7 84 24 08 02 00 	movq   $0x1,0x208(%rsp)
 1058910:	00 01 00 00 00 
 1058915:	48 8b 84 24 f8 01 00 	mov    0x1f8(%rsp),%rax
 105891c:	00 
 105891d:	48 8d 8c 24 10 02 00 	lea    0x210(%rsp),%rcx
 1058924:	00 
 1058925:	48 89 41 10          	mov    %rax,0x10(%rcx)
 1058929:	0f 10 84 24 e8 01 00 	movups 0x1e8(%rsp),%xmm0
 1058930:	00 
 1058931:	0f 11 01             	movups %xmm0,(%rcx)
 1058934:	bf 08 00 00 00       	mov    $0x8,%edi
 1058939:	be 28 00 00 00       	mov    $0x28,%esi
 105893e:	e8 fd 57 cd ff       	call   d2e140 <_RNvNtCslNYArtu3iFV_5alloc5boxed14box_new_uninit>
 1058943:	48 8b 8c 24 20 02 00 	mov    0x220(%rsp),%rcx
 105894a:	00 
 105894b:	48 89 48 20          	mov    %rcx,0x20(%rax)
 105894f:	0f 10 84 24 00 02 00 	movups 0x200(%rsp),%xmm0
 1058956:	00 
 1058957:	0f 10 8c 24 10 02 00 	movups 0x210(%rsp),%xmm1
 105895e:	00 
 105895f:	0f 11 48 10          	movups %xmm1,0x10(%rax)
 1058963:	0f 11 00             	movups %xmm0,(%rax)
 1058966:	48 89 ef             	mov    %rbp,%rdi
 1058969:	48 89 c6             	mov    %rax,%rsi
 105896c:	48 8d 15 35 44 7e 02 	lea    0x27e4435(%rip),%rdx        # 383cda8 <anon.41d43701a2c2b5e6feb75b74ff96deac.936.llvm.8476651010627807873+0x1b8>
 1058973:	ff 15 df 8a 85 02    	call   *0x2858adf(%rip)        # 38b1458 <_DYNAMIC+0xbb8>
 1058979:	48 8b 9c 24 00 27 00 	mov    0x2700(%rsp),%rbx
 1058980:	00 
 1058981:	48 8d 84 24 10 27 00 	lea    0x2710(%rsp),%rax
 1058988:	00 
 1058989:	0f 10 40 f8          	movups -0x8(%rax),%xmm0
 105898d:	0f 10 48 08          	movups 0x8(%rax),%xmm1
 1058991:	0f 10 50 18          	movups 0x18(%rax),%xmm2
 1058995:	0f 10 58 28          	movups 0x28(%rax),%xmm3
 1058999:	0f 29 84 24 30 08 00 	movaps %xmm0,0x830(%rsp)
 10589a0:	00 
 10589a1:	0f 29 8c 24 40 08 00 	movaps %xmm1,0x840(%rsp)
 10589a8:	00 
 10589a9:	0f 29 94 24 50 08 00 	movaps %xmm2,0x850(%rsp)
 10589b0:	00 
 10589b1:	0f 29 9c 24 60 08 00 	movaps %xmm3,0x860(%rsp)
 10589b8:	00 
 10589b9:	48 8b 40 38          	mov    0x38(%rax),%rax
 10589bd:	48 89 84 24 70 08 00 	mov    %rax,0x870(%rsp)
 10589c4:	00 
 10589c5:	48 83 fb 02          	cmp    $0x2,%rbx
 10589c9:	0f 84 e4 1d 00 00    	je     105a7b3 <_ZN20litchi_perf_baseline56run_docx_source_backed_existing_section_layout_edit_save17h1cd3e0d854c7eea3E+0x5cd3>
 10589cf:	ba 38 07 00 00       	mov    $0x738,%edx
 10589d4:	48 8d bc 24 00 28 01 	lea    0x12800(%rsp),%rdi
 10589db:	00 
 10589dc:	48 8d b4 24 50 27 00 	lea    0x2750(%rsp),%rsi
 10589e3:	00 
 10589e4:	ff 15 f6 80 85 02    	call   *0x28580f6(%rip)        # 38b0ae0 <memcpy@GLIBC_2.14>
 10589ea:	48 89 9c 24 b0 27 01 	mov    %rbx,0x127b0(%rsp)
 10589f1:	00 
 10589f2:	48 8b 84 24 70 08 00 	mov    0x870(%rsp),%rax
 10589f9:	00 
 10589fa:	48 8d 8c 24 b8 27 01 	lea    0x127b8(%rsp),%rcx
 1058a01:	00 
 1058a02:	48 89 41 40          	mov    %rax,0x40(%rcx)
 1058a06:	0f 28 84 24 30 08 00 	movaps 0x830(%rsp),%xmm0
 1058a0d:	00 
 1058a0e:	0f 28 8c 24 40 08 00 	movaps 0x840(%rsp),%xmm1
 1058a15:	00 
 1058a16:	0f 28 94 24 50 08 00 	movaps 0x850(%rsp),%xmm2
 1058a1d:	00 
 1058a1e:	0f 28 9c 24 60 08 00 	movaps 0x860(%rsp),%xmm3
 1058a25:	00 
 1058a26:	0f 11 59 30          	movups %xmm3,0x30(%rcx)
 1058a2a:	0f 11 51 20          	movups %xmm2,0x20(%rcx)
 1058a2e:	0f 11 49 10          	movups %xmm1,0x10(%rcx)
 1058a32:	0f 11 01             	movups %xmm0,(%rcx)
 1058a35:	48 c7 84 24 e8 01 00 	movq   $0x0,0x1e8(%rsp)
 1058a3c:	00 00 00 00 00 
 1058a41:	48 c7 84 24 f0 01 00 	movq   $0x1,0x1f0(%rsp)
 1058a48:	00 01 00 00 00 
 1058a4d:	48 c7 84 24 f8 01 00 	movq   $0x0,0x1f8(%rsp)
 1058a54:	00 00 00 00 00 
 1058a59:	48 89 ef             	mov    %rbp,%rdi
 1058a5c:	48 8d b4 24 b0 27 01 	lea    0x127b0(%rsp),%rsi
 1058a63:	00 
 1058a64:	48 8d 94 24 e8 01 00 	lea    0x1e8(%rsp),%rdx
 1058a6b:	00 
 1058a6c:	48 8d 8c 24 60 21 00 	lea    0x2160(%rsp),%rcx
 1058a73:	00 
 1058a74:	e8 97 d5 78 00       	call   17e6010 <_ZN11litchi_docx13source_backed7Package40publish_section_layout_inverse_to_stream17h327cd3bb1a9e76e3E>
 1058a79:	48 8b 84 24 00 27 00 	mov    0x2700(%rsp),%rax
 1058a80:	00 
 1058a81:	48 8d 8c 24 10 27 00 	lea    0x2710(%rsp),%rcx
 1058a88:	00 
 1058a89:	0f 10 41 f8          	movups -0x8(%rcx),%xmm0
 1058a8d:	0f 10 49 08          	movups 0x8(%rcx),%xmm1
 1058a91:	0f 10 51 18          	movups 0x18(%rcx),%xmm2
 1058a95:	0f 10 59 28          	movups 0x28(%rcx),%xmm3
 1058a99:	0f                   	.byte 0xf
 1058a9a:	29                   	.byte 0x29
 1058a9b:	84 24 00             	test   %ah,(%rax,%rax,1)
 1058a9e:	02 00                	add    (%rax),%al
