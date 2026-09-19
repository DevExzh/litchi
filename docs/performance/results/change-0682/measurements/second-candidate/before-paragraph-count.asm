
/home/zhuhe/code/litchi-target-0682-before/release/probe0682:     file format elf64-x86-64


Disassembly of section .text:

00000000001192b0 <litchi_docx::source_backed::Document::paragraph_count>:
  1192b0:	41 57                	push   %r15
  1192b2:	41 56                	push   %r14
  1192b4:	41 54                	push   %r12
  1192b6:	53                   	push   %rbx
  1192b7:	48 81 ec 38 02 00 00 	sub    $0x238,%rsp
  1192be:	49 89 f7             	mov    %rsi,%r15
  1192c1:	48 89 fb             	mov    %rdi,%rbx
  1192c4:	4c 8d 76 40          	lea    0x40(%rsi),%r14
  1192c8:	48 83 7e 40 00       	cmpq   $0x0,0x40(%rsi)
  1192cd:	74 2a                	je     1192f9 <litchi_docx::source_backed::Document::paragraph_count+0x49>
  1192cf:	49 bc 1b 00 00 00 00 	movabs $0x800000000000001b,%r12
  1192d6:	00 00 80 
  1192d9:	48 8d 7c 24 70       	lea    0x70(%rsp),%rdi
  1192de:	4c 89 f6             	mov    %r14,%rsi
  1192e1:	ff 15 91 3b 18 00    	call   *0x183b91(%rip)        # 29ce78 <_DYNAMIC+0x7d8>
  1192e7:	0f b6 84 24 90 00 00 	movzbl 0x90(%rsp),%eax
  1192ee:	00 
  1192ef:	83 f8 09             	cmp    $0x9,%eax
  1192f2:	74 66                	je     11935a <litchi_docx::source_backed::Document::paragraph_count+0xaa>
  1192f4:	83 f8 0d             	cmp    $0xd,%eax
  1192f7:	75 2c                	jne    119325 <litchi_docx::source_backed::Document::paragraph_count+0x75>
  1192f9:	41 8b 47 08          	mov    0x8(%r15),%eax
  1192fd:	85 c0                	test   %eax,%eax
  1192ff:	0f 85 8d 00 00 00    	jne    119392 <litchi_docx::source_backed::Document::paragraph_count+0xe2>
  119305:	49 8b 07             	mov    (%r15),%rax
  119308:	48 85 c0             	test   %rax,%rax
  11930b:	0f 84 98 00 00 00    	je     1193a9 <litchi_docx::source_backed::Document::paragraph_count+0xf9>
  119311:	48 8b 40 18          	mov    0x18(%rax),%rax
  119315:	48 89 43 08          	mov    %rax,0x8(%rbx)
  119319:	48 c7 03 20 00 00 00 	movq   $0x20,(%rbx)
  119320:	e9 71 03 00 00       	jmp    119696 <litchi_docx::source_backed::Document::paragraph_count+0x3e6>
  119325:	0f 10 44 24 70       	movups 0x70(%rsp),%xmm0
  11932a:	0f 10 8c 24 80 00 00 	movups 0x80(%rsp),%xmm1
  119331:	00 
  119332:	0f 29 8c 24 40 01 00 	movaps %xmm1,0x140(%rsp)
  119339:	00 
  11933a:	0f 29 84 24 30 01 00 	movaps %xmm0,0x130(%rsp)
  119341:	00 
  119342:	8b 8c 24 91 00 00 00 	mov    0x91(%rsp),%ecx
  119349:	89 0c 24             	mov    %ecx,(%rsp)
  11934c:	8b 8c 24 94 00 00 00 	mov    0x94(%rsp),%ecx
  119353:	89 4c 24 03          	mov    %ecx,0x3(%rsp)
  119357:	49 ff c4             	inc    %r12
  11935a:	48 c7 03 00 00 00 00 	movq   $0x0,(%rbx)
  119361:	4c 89 63 08          	mov    %r12,0x8(%rbx)
  119365:	0f 28 84 24 30 01 00 	movaps 0x130(%rsp),%xmm0
  11936c:	00 
  11936d:	0f 28 8c 24 40 01 00 	movaps 0x140(%rsp),%xmm1
  119374:	00 
  119375:	0f 11 43 10          	movups %xmm0,0x10(%rbx)
  119379:	0f 11 4b 20          	movups %xmm1,0x20(%rbx)
  11937d:	88 43 30             	mov    %al,0x30(%rbx)
  119380:	8b 04 24             	mov    (%rsp),%eax
  119383:	8b 4c 24 03          	mov    0x3(%rsp),%ecx
  119387:	89 43 31             	mov    %eax,0x31(%rbx)
  11938a:	89 4b 34             	mov    %ecx,0x34(%rbx)
  11938d:	e9 04 03 00 00       	jmp    119696 <litchi_docx::source_backed::Document::paragraph_count+0x3e6>
  119392:	4c 89 ff             	mov    %r15,%rdi
  119395:	4c 89 fe             	mov    %r15,%rsi
  119398:	e8 49 33 00 00       	call   11c6e6 <std::sync::once_lock::OnceLock<T>::initialize>
  11939d:	49 8b 07             	mov    (%r15),%rax
  1193a0:	48 85 c0             	test   %rax,%rax
  1193a3:	0f 85 68 ff ff ff    	jne    119311 <litchi_docx::source_backed::Document::paragraph_count+0x61>
  1193a9:	49 8b 47 20          	mov    0x20(%r15),%rax
  1193ad:	49                   	rex.WB
  1193ae:	8b                   	.byte 0x8b
  1193af:	77                   	.byte 0x77
