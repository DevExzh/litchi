
/home/zhuhe/code/litchi-target-0682-after/release/probe0682:     file format elf64-x86-64


Disassembly of section .text:

0000000000119c50 <litchi_docx::source_backed::Document::paragraph_count>:
  119c50:	41 57                	push   %r15
  119c52:	41 56                	push   %r14
  119c54:	41 54                	push   %r12
  119c56:	53                   	push   %rbx
  119c57:	48 81 ec 38 02 00 00 	sub    $0x238,%rsp
  119c5e:	49 89 f7             	mov    %rsi,%r15
  119c61:	48 89 fb             	mov    %rdi,%rbx
  119c64:	4c 8d 76 48          	lea    0x48(%rsi),%r14
  119c68:	48 83 7e 48 00       	cmpq   $0x0,0x48(%rsi)
  119c6d:	74 2a                	je     119c99 <litchi_docx::source_backed::Document::paragraph_count+0x49>
  119c6f:	49 bc 1b 00 00 00 00 	movabs $0x800000000000001b,%r12
  119c76:	00 00 80 
  119c79:	48 8d 7c 24 70       	lea    0x70(%rsp),%rdi
  119c7e:	4c 89 f6             	mov    %r14,%rsi
  119c81:	ff 15 f1 6b 18 00    	call   *0x186bf1(%rip)        # 2a0878 <_DYNAMIC+0x8d8>
  119c87:	0f b6 84 24 90 00 00 	movzbl 0x90(%rsp),%eax
  119c8e:	00 
  119c8f:	83 f8 09             	cmp    $0x9,%eax
  119c92:	74 6b                	je     119cff <litchi_docx::source_backed::Document::paragraph_count+0xaf>
  119c94:	83 f8 0d             	cmp    $0xd,%eax
  119c97:	75 31                	jne    119cca <litchi_docx::source_backed::Document::paragraph_count+0x7a>
  119c99:	41 8b 47 28          	mov    0x28(%r15),%eax
  119c9d:	85 c0                	test   %eax,%eax
  119c9f:	0f 85 92 00 00 00    	jne    119d37 <litchi_docx::source_backed::Document::paragraph_count+0xe7>
  119ca5:	49 8b 47 20          	mov    0x20(%r15),%rax
  119ca9:	48 85 c0             	test   %rax,%rax
  119cac:	0f 84 9e 00 00 00    	je     119d50 <litchi_docx::source_backed::Document::paragraph_count+0x100>
  119cb2:	48 8b 40 10          	mov    0x10(%rax),%rax
  119cb6:	48 8b 40 18          	mov    0x18(%rax),%rax
  119cba:	48 89 43 08          	mov    %rax,0x8(%rbx)
  119cbe:	48 c7 03 20 00 00 00 	movq   $0x20,(%rbx)
  119cc5:	e9 82 03 00 00       	jmp    11a04c <litchi_docx::source_backed::Document::paragraph_count+0x3fc>
  119cca:	0f 10 44 24 70       	movups 0x70(%rsp),%xmm0
  119ccf:	0f 10 8c 24 80 00 00 	movups 0x80(%rsp),%xmm1
  119cd6:	00 
  119cd7:	0f 29 8c 24 40 01 00 	movaps %xmm1,0x140(%rsp)
  119cde:	00 
  119cdf:	0f 29 84 24 30 01 00 	movaps %xmm0,0x130(%rsp)
  119ce6:	00 
  119ce7:	8b 8c 24 91 00 00 00 	mov    0x91(%rsp),%ecx
  119cee:	89 0c 24             	mov    %ecx,(%rsp)
  119cf1:	8b 8c 24 94 00 00 00 	mov    0x94(%rsp),%ecx
  119cf8:	89 4c 24 03          	mov    %ecx,0x3(%rsp)
  119cfc:	49 ff c4             	inc    %r12
  119cff:	48 c7 03 00 00 00 00 	movq   $0x0,(%rbx)
  119d06:	4c 89 63 08          	mov    %r12,0x8(%rbx)
  119d0a:	0f 28 84 24 30 01 00 	movaps 0x130(%rsp),%xmm0
  119d11:	00 
  119d12:	0f 28 8c 24 40 01 00 	movaps 0x140(%rsp),%xmm1
  119d19:	00 
  119d1a:	0f 11 43 10          	movups %xmm0,0x10(%rbx)
  119d1e:	0f 11 4b 20          	movups %xmm1,0x20(%rbx)
  119d22:	88 43 30             	mov    %al,0x30(%rbx)
  119d25:	8b 04 24             	mov    (%rsp),%eax
  119d28:	8b 4c 24 03          	mov    0x3(%rsp),%ecx
  119d2c:	89 43 31             	mov    %eax,0x31(%rbx)
  119d2f:	89 4b 34             	mov    %ecx,0x34(%rbx)
  119d32:	e9 15 03 00 00       	jmp    11a04c <litchi_docx::source_backed::Document::paragraph_count+0x3fc>
  119d37:	49 8d 7f 20          	lea    0x20(%r15),%rdi
  119d3b:	4c 89 fe             	mov    %r15,%rsi
  119d3e:	e8 2d 15 01 00       	call   12b270 <std::sync::once_lock::OnceLock<T>::initialize>
  119d43:	49 8b 47 20          	mov    0x20(%r15),%rax
  119d47:	48 85 c0             	test   %rax,%rax
  119d4a:	0f 85 62 ff ff ff    	jne    119cb2 <litchi_docx::source_backed::Document::paragraph_count+0x62>
