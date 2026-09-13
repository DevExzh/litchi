/var/tmp/ods-tokenizer-baseline-target-20260913/release/ods-formula-reference-profile:     file format elf64-x86-64


Disassembly of section .text:

0000000000070a50 <litchi_ods::codec::formula::copy_formula_component>:
   70a50:	55                   	push   %rbp
   70a51:	41 57                	push   %r15
   70a53:	41 56                	push   %r14
   70a55:	41 55                	push   %r13
   70a57:	41 54                	push   %r12
   70a59:	53                   	push   %rbx
   70a5a:	48 83 ec 18          	sub    $0x18,%rsp
   70a5e:	4d 89 c4             	mov    %r8,%r12
   70a61:	49 89 cd             	mov    %rcx,%r13
   70a64:	49 89 d6             	mov    %rdx,%r14
   70a67:	49 89 f7             	mov    %rsi,%r15
   70a6a:	48 89 fb             	mov    %rdi,%rbx
   70a6d:	48 c7 04 24 00 00 00 	movq   $0x0,(%rsp)
   70a74:	00
   70a75:	48 c7 44 24 08 01 00 	movq   $0x1,0x8(%rsp)
   70a7c:	00 00
   70a7e:	48 c7 44 24 10 00 00 	movq   $0x0,0x10(%rsp)
   70a85:	00 00
   70a87:	48 89 e7             	mov    %rsp,%rdi
   70a8a:	48 89 d6             	mov    %rdx,%rsi
   70a8d:	ff 15 6d d0 04 00    	call   *0x4d06d(%rip)        # bdb00 <_DYNAMIC+0x3d0>
   70a93:	48 bd 01 00 00 00 00 	movabs $0x8000000000000001,%rbp
   70a9a:	00 00 80
   70a9d:	48 39 e8             	cmp    %rbp,%rax
   70aa0:	75 4c                	jne    70aee <litchi_ods::codec::formula::copy_formula_component+0x9e>
   70aa2:	48 8b 04 24          	mov    (%rsp),%rax
   70aa6:	48 8b 74 24 10       	mov    0x10(%rsp),%rsi
   70aab:	48 29 f0             	sub    %rsi,%rax
   70aae:	49 39 c6             	cmp    %rax,%r14
   70ab1:	77 7a                	ja     70b2d <litchi_ods::codec::formula::copy_formula_component+0xdd>
   70ab3:	4d 85 f6             	test   %r14,%r14
   70ab6:	74 19                	je     70ad1 <litchi_ods::codec::formula::copy_formula_component+0x81>
   70ab8:	48 03 74 24 08       	add    0x8(%rsp),%rsi
   70abd:	48 89 f7             	mov    %rsi,%rdi
   70ac0:	4c 89 fe             	mov    %r15,%rsi
   70ac3:	4c 89 f2             	mov    %r14,%rdx
   70ac6:	ff 15 ec ce 04 00    	call   *0x4ceec(%rip)        # bd9b8 <memcpy@GLIBC_2.14>
   70acc:	48 8b 74 24 10       	mov    0x10(%rsp),%rsi
   70ad1:	4c 01 f6             	add    %r14,%rsi
   70ad4:	48 89 74 24 10       	mov    %rsi,0x10(%rsp)
   70ad9:	48 89 73 18          	mov    %rsi,0x18(%rbx)
   70add:	0f 10 04 24          	movups (%rsp),%xmm0
   70ae1:	0f 11 43 08          	movups %xmm0,0x8(%rbx)
   70ae5:	48 83 c5 10          	add    $0x10,%rbp
   70ae9:	48 89 2b             	mov    %rbp,(%rbx)
   70aec:	eb 30                	jmp    70b1e <litchi_ods::codec::formula::copy_formula_component+0xce>
   70aee:	48 83 c5 0c          	add    $0xc,%rbp
   70af2:	48 89 2b             	mov    %rbp,(%rbx)
   70af5:	48 89 43 08          	mov    %rax,0x8(%rbx)
   70af9:	48 89 53 10          	mov    %rdx,0x10(%rbx)
   70afd:	4c 89 6b 18          	mov    %r13,0x18(%rbx)
   70b01:	4c 89 63 20          	mov    %r12,0x20(%rbx)
   70b05:	48 8b 34 24          	mov    (%rsp),%rsi
   70b09:	48 85 f6             	test   %rsi,%rsi
   70b0c:	74 10                	je     70b1e <litchi_ods::codec::formula::copy_formula_component+0xce>
   70b0e:	48 8b 7c 24 08       	mov    0x8(%rsp),%rdi
   70b13:	ba 01 00 00 00       	mov    $0x1,%edx
   70b18:	ff 15 12 ce 04 00    	call   *0x4ce12(%rip)        # bd930 <_DYNAMIC+0x200>
   70b1e:	48 83 c4 18          	add    $0x18,%rsp
   70b22:	5b                   	pop    %rbx
   70b23:	41 5c                	pop    %r12
   70b25:	41 5d                	pop    %r13
   70b27:	41 5e                	pop    %r14
   70b29:	41 5f                	pop    %r15
   70b2b:	5d                   	pop    %rbp
   70b2c:	c3                   	ret
   70b2d:	48 89 e7             	mov    %rsp,%rdi
   70b30:	b9 01 00 00 00       	mov    $0x1,%ecx
   70b35:	41 b8 01 00 00 00    	mov    $0x1,%r8d
   70b3b:	4c 89 f2             	mov    %r14,%rdx
   70b3e:	e8 8d 15 00 00       	call   720d0 <alloc::raw_vec::RawVecInner<A>::reserve::do_reserve_and_handle>
   70b43:	48 8b 74 24 10       	mov    0x10(%rsp),%rsi
   70b48:	e9 6b ff ff ff       	jmp    70ab8 <litchi_ods::codec::formula::copy_formula_component+0x68>
   70b4d:	48 89 c3             	mov    %rax,%rbx
   70b50:	48 8b 34 24          	mov    (%rsp),%rsi
   70b54:	48 85 f6             	test   %rsi,%rsi
   70b57:	74 10                	je     70b69 <litchi_ods::codec::formula::copy_formula_component+0x119>
   70b59:	48 8b 7c 24 08       	mov    0x8(%rsp),%rdi
   70b5e:	ba 01 00 00 00       	mov    $0x1,%edx
   70b63:	ff 15 c7 cd 04 00    	call   *0x4cdc7(%rip)        # bd930 <_DYNAMIC+0x200>
   70b69:	48 89 df             	mov    %rbx,%rdi
   70b6c:	e8 af 7c 04 00       	call   b8820 <_Unwind_Resume@plt>
