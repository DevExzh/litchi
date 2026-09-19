
/home/zhuhe/code/litchi-target-0682-after/release/probe0682:     file format elf64-x86-64


Disassembly of section .text:

00000000001207c0 <litchi_docx::source_backed::Document::paragraph_count>:
  1207c0:	41 57                	push   %r15
  1207c2:	41 56                	push   %r14
  1207c4:	41 54                	push   %r12
  1207c6:	53                   	push   %rbx
  1207c7:	48 81 ec 38 02 00 00 	sub    $0x238,%rsp
  1207ce:	49 89 f7             	mov    %rsi,%r15
  1207d1:	48 89 fb             	mov    %rdi,%rbx
  1207d4:	4c 8d 76 48          	lea    0x48(%rsi),%r14
  1207d8:	48 83 7e 48 00       	cmpq   $0x0,0x48(%rsi)
  1207dd:	74 2a                	je     120809 <litchi_docx::source_backed::Document::paragraph_count+0x49>
  1207df:	49 bc 1b 00 00 00 00 	movabs $0x800000000000001b,%r12
  1207e6:	00 00 80 
  1207e9:	48 8d 7c 24 70       	lea    0x70(%rsp),%rdi
  1207ee:	4c 89 f6             	mov    %r14,%rsi
  1207f1:	ff 15 d9 fb 17 00    	call   *0x17fbd9(%rip)        # 2a03d0 <_DYNAMIC+0x7f0>
  1207f7:	0f b6 84 24 90 00 00 	movzbl 0x90(%rsp),%eax
  1207fe:	00 
  1207ff:	83 f8 09             	cmp    $0x9,%eax
  120802:	74 67                	je     12086b <litchi_docx::source_backed::Document::paragraph_count+0xab>
  120804:	83 f8 0d             	cmp    $0xd,%eax
  120807:	75 2d                	jne    120836 <litchi_docx::source_backed::Document::paragraph_count+0x76>
  120809:	41 8b 47 28          	mov    0x28(%r15),%eax
  12080d:	85 c0                	test   %eax,%eax
  12080f:	0f 85 8e 00 00 00    	jne    1208a3 <litchi_docx::source_backed::Document::paragraph_count+0xe3>
  120815:	49 8b 47 20          	mov    0x20(%r15),%rax
  120819:	48 85 c0             	test   %rax,%rax
  12081c:	0f 84 9a 00 00 00    	je     1208bc <litchi_docx::source_backed::Document::paragraph_count+0xfc>
  120822:	48 8b 40 18          	mov    0x18(%rax),%rax
  120826:	48 89 43 08          	mov    %rax,0x8(%rbx)
  12082a:	48 c7 03 20 00 00 00 	movq   $0x20,(%rbx)
  120831:	e9 82 03 00 00       	jmp    120bb8 <litchi_docx::source_backed::Document::paragraph_count+0x3f8>
  120836:	0f 10 44 24 70       	movups 0x70(%rsp),%xmm0
  12083b:	0f 10 8c 24 80 00 00 	movups 0x80(%rsp),%xmm1
  120842:	00 
  120843:	0f 29 8c 24 40 01 00 	movaps %xmm1,0x140(%rsp)
  12084a:	00 
  12084b:	0f 29 84 24 30 01 00 	movaps %xmm0,0x130(%rsp)
  120852:	00 
  120853:	8b 8c 24 91 00 00 00 	mov    0x91(%rsp),%ecx
  12085a:	89 0c 24             	mov    %ecx,(%rsp)
  12085d:	8b 8c 24 94 00 00 00 	mov    0x94(%rsp),%ecx
  120864:	89 4c 24 03          	mov    %ecx,0x3(%rsp)
  120868:	49 ff c4             	inc    %r12
  12086b:	48 c7 03 00 00 00 00 	movq   $0x0,(%rbx)
  120872:	4c 89 63 08          	mov    %r12,0x8(%rbx)
  120876:	0f 28 84 24 30 01 00 	movaps 0x130(%rsp),%xmm0
  12087d:	00 
  12087e:	0f 28 8c 24 40 01 00 	movaps 0x140(%rsp),%xmm1
  120885:	00 
  120886:	0f 11 43 10          	movups %xmm0,0x10(%rbx)
  12088a:	0f 11 4b 20          	movups %xmm1,0x20(%rbx)
  12088e:	88 43 30             	mov    %al,0x30(%rbx)
  120891:	8b 04 24             	mov    (%rsp),%eax
  120894:	8b 4c 24 03          	mov    0x3(%rsp),%ecx
  120898:	89 43 31             	mov    %eax,0x31(%rbx)
  12089b:	89 4b 34             	mov    %ecx,0x34(%rbx)
  12089e:	e9 15 03 00 00       	jmp    120bb8 <litchi_docx::source_backed::Document::paragraph_count+0x3f8>
  1208a3:	49 8d 7f 20          	lea    0x20(%r15),%rdi
  1208a7:	4c 89 fe             	mov    %r15,%rsi
  1208aa:	e8 1f 97 00 00       	call   129fce <std::sync::once_lock::OnceLock<T>::initialize>
  1208af:	49 8b 47 20          	mov    0x20(%r15),%rax
  1208b3:	48 85 c0             	test   %rax,%rax
  1208b6:	0f 85 66 ff ff ff    	jne    120822 <litchi_docx::source_backed::Document::paragraph_count+0x62>
  1208bc:	49 8b 47 08          	mov    0x8(%r15),%rax
