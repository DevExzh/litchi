
/home/zhuhe/code/litchi-target-0823/before-real-native:     file format elf64-x86-64


Disassembly of section .text:

0000000000138160 <litchi_pptx::shape::reader::Scene::read_with>:
  138160:	push   %rbp
  138161:	push   %r15
  138163:	push   %r14
  138165:	push   %r13
  138167:	push   %r12
  138169:	push   %rbx
  13816a:	sub    $0x408,%rsp
  138171:	mov    %rdi,%r14
  138174:	mov    (%rcx),%r13
  138177:	cmp    %r13,%rdx
  13817a:	jbe    13819a <litchi_pptx::shape::reader::Scene::read_with+0x3a>
  13817c:	movb   $0x9,0x8(%r14)
  138181:	mov    %r13,0x10(%r14)
  138185:	lea    -0x110f13(%rip),%rax        # 27279 <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13130224050040044120+0x81>
  13818c:	mov    %rax,0x18(%r14)
  138190:	movq   $0x17,0x20(%r14)
  138198:	jmp    13820c <litchi_pptx::shape::reader::Scene::read_with+0xac>
  13819a:	mov    %rcx,0x198(%rsp)
  1381a2:	mov    0x8(%rcx),%r15
  1381a6:	mov    %r15,%rax
  1381a9:	shr    $0x20,%rax
  1381ad:	je     138221 <litchi_pptx::shape::reader::Scene::read_with+0xc1>
  1381af:	mov    $0x36,%edi
  1381b4:	call   *0xf422e(%rip)        # 22c3e8 <malloc@GLIBC_2.2.5>
  1381ba:	test   %rax,%rax
  1381bd:	je     13a64c <litchi_pptx::shape::reader::Scene::read_with+0x24ec>
  1381c3:	movups -0x110f67(%rip),%xmm0        # 27263 <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13130224050040044120+0x6b>
  1381ca:	movups %xmm0,0x20(%rax)
  1381ce:	movups -0x110f82(%rip),%xmm0        # 27253 <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13130224050040044120+0x5b>
  1381d5:	movups %xmm0,0x10(%rax)
  1381d9:	movdqu -0x110f9e(%rip),%xmm0        # 27243 <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13130224050040044120+0x4b>
  1381e1:	movdqu %xmm0,(%rax)
  1381e5:	movabs $0x6e69616d6f64206e,%rcx
  1381ef:	mov    %rcx,0x2e(%rax)
  1381f3:	movb   $0x8,0x8(%r14)
  1381f8:	movq   $0x36,0x10(%r14)
  138200:	mov    %rax,0x18(%r14)
  138204:	movq   $0x36,0x20(%r14)
  13820c:	movabs $0x8000000000000001,%rax
  138216:	dec    %rax
  138219:	mov    %rax,(%r14)
  13821c:	jmp    13a355 <litchi_pptx::shape::reader::Scene::read_with+0x21f5>
  138221:	mov    %rdx,%rbx
  138224:	mov    %rsi,%r12
  138227:	movq   $0x0,0x1c8(%rsp)
  138233:	mov    $0xffffffffffffffd8,%rbp
  13823a:	cmpb   $0x1,%fs:0x10(%rbp)
  13823f:	jne    13a0d7 <litchi_pptx::shape::reader::Scene::read_with+0x1f77>
  138245:	mov    %fs:0x0(%rbp),%rax
  13824a:	mov    %fs:0x8(%rbp),%rdx
  13824f:	lea    0x1(%rax),%rcx
  138253:	mov    %rcx,%fs:0x0(%rbp)
  138258:	movups 0x1c8(%rsp),%xmm0
  138260:	movdqu 0x1d8(%rsp),%xmm1
  138269:	movups 0x1e8(%rsp),%xmm2
  138271:	movaps %xmm0,0x3a0(%rsp)
  138279:	movdqa %xmm1,0x3b0(%rsp)
  138282:	movaps %xmm2,0x3c0(%rsp)
  13828a:	movups 0xeb407(%rip),%xmm0        # 223698 <anon.4cef9af9ef5ad65ef8056305e811a596.14.llvm.8137685535559598875>
  138291:	movaps %xmm0,0x370(%rsp)
  138299:	movdqu 0xeb407(%rip),%xmm0        # 2236a8 <anon.4cef9af9ef5ad65ef8056305e811a596.14.llvm.8137685535559598875+0x10>
  1382a1:	movdqa %xmm0,0x380(%rsp)
  1382aa:	mov    %rax,0x390(%rsp)
  1382b2:	mov    %rdx,0x398(%rsp)
  1382ba:	lea    -0x11379e(%rip),%rsi        # 24b23 <anon.bc3d94ef2394061926bee58178356194.284.llvm.5373085861489949100+0x4b2b>
  1382c1:	lea    0x370(%rsp),%rdi
  1382c9:	mov    $0x38,%edx
  1382ce:	call   15f840 <litchi_ooxml_common::mce::model::Capabilities::understand_namespace>
  1382d3:	lea    -0x1135ae(%rip),%rsi        # 24d2c <anon.bc3d94ef2394061926bee58178356194.284.llvm.5373085861489949100+0x4d34>
  1382da:	lea    0x370(%rsp),%rdi
  1382e2:	mov    $0x38,%edx
  1382e7:	call   15f840 <litchi_ooxml_common::mce::model::Capabilities::understand_namespace>
  1382ec:	lea    -0x112b0d(%rip),%rsi        # 257e6 <anon.bc3d94ef2394061926bee58178356194.284.llvm.5373085861489949100+0x57ee>
  1382f3:	lea    0x370(%rsp),%rdi
  1382fb:	mov    $0x3b,%edx
  138300:	call   15f840 <litchi_ooxml_common::mce::model::Capabilities::understand_namespace>
  138305:	mov    0x198(%rsp),%rax
  13830d:	mov    0x10(%rax),%rax
  138311:	mov    %r13,0x3d0(%rsp)
  138319:	mov    %r15,0x3d8(%rsp)
  138321:	mov    %rax,0x3e0(%rsp)
  138329:	movq   $0x1000,0x3e8(%rsp)
  138335:	movq   $0x1000,0x3f0(%rsp)
  138341:	movq   $0x400,0x3f8(%rsp)
  13834d:	movq   $0x400,0x400(%rsp)
  138359:	lea    0x1c8(%rsp),%rdi
  138361:	lea    0x370(%rsp),%rcx
  138369:	lea    0x3d0(%rsp),%r8
  138371:	mov    %r12,%rsi
  138374:	mov    %rbx,%rdx
  138377:	call   *0xf4583(%rip)        # 22c900 <_DYNAMIC+0x700>
  13837d:	movabs $0x8000000000000001,%rax
  138387:	mov    0x1c8(%rsp),%rcx
  13838f:	mov    0x1d0(%rsp),%rbx
  138397:	mov    0x1d8(%rsp),%r12
  13839f:	cmp    %rax,%rcx
  1383a2:	jne    1383cb <litchi_pptx::shape::reader::Scene::read_with+0x26b>
  1383a4:	movdqu 0x1e0(%rsp),%xmm0
  1383ad:	movb   $0x22,0x8(%r14)
  1383b2:	mov    %rbx,0x10(%r14)
  1383b6:	mov    %r12,0x18(%r14)
  1383ba:	movdqu %xmm0,0x20(%r14)
  1383c0:	dec    %rax
  1383c3:	mov    %rax,(%r14)
  1383c6:	jmp    13a348 <litchi_pptx::shape::reader::Scene::read_with+0x21e8>
  1383cb:	cmp    %r15,%r12
  1383ce:	jbe    13840d <litchi_pptx::shape::reader::Scene::read_with+0x2ad>
  1383d0:	movb   $0x9,0x8(%r14)
  1383d5:	mov    %r15,0x10(%r14)
  1383d9:	lea    -0x1111b8(%rip),%rax        # 27228 <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13130224050040044120+0x30>
  1383e0:	mov    %rax,0x18(%r14)
  1383e4:	movq   $0x1b,0x20(%r14)
  1383ec:	movabs $0x8000000000000001,%rax
  1383f6:	dec    %rax
  1383f9:	mov    %rax,(%r14)
  1383fc:	shl    $1,%rcx
  1383ff:	test   %rcx,%rcx
  138402:	jne    13a3eb <litchi_pptx::shape::reader::Scene::read_with+0x228b>
  138408:	jmp    13a348 <litchi_pptx::shape::reader::Scene::read_with+0x21e8>
  13840d:	mov    %rcx,0x158(%rsp)
  138415:	mov    %rbx,0x248(%rsp)
  13841d:	mov    %r12,0x250(%rsp)
  138425:	mov    0x198(%rsp),%rax
  13842d:	movdqu (%rax),%xmm0
  138431:	movdqu 0x10(%rax),%xmm1
  138436:	movups 0x20(%rax),%xmm2
  13843a:	movdqu %xmm0,0x258(%rsp)
  138443:	movdqu %xmm1,0x268(%rsp)
  13844c:	movups %xmm2,0x278(%rsp)
  138454:	movq   $0x0,0x1e8(%rsp)
  138460:	movq   $0x8,0x1f0(%rsp)
  13846c:	pxor   %xmm0,%xmm0
  138470:	movdqu %xmm0,0x1f8(%rsp)
  138479:	movq   $0x1,0x208(%rsp)
  138485:	movdqu %xmm0,0x210(%rsp)
  13848e:	movq   $0x8,0x220(%rsp)
  13849a:	movdqu %xmm0,0x228(%rsp)
  1384a3:	movdqu %xmm0,0x288(%rsp)
  1384ac:	movq   $0x0,0x298(%rsp)
  1384b8:	movq   $0x1,0x238(%rsp)
  1384c4:	movq   $0x0,0x240(%rsp)
  1384d0:	movq   $0x0,0x1c8(%rsp)
  1384dc:	movq   $0x0,0x1d8(%rsp)
  1384e8:	movb   $0x0,0x2a0(%rsp)
  1384f0:	mov    %rbx,0x80(%rsp)
  1384f8:	mov    %r12,0x88(%rsp)
  138500:	movq   $0x0,0x30(%rsp)
  138509:	movq   $0x1,0x38(%rsp)
  138512:	movdqu %xmm0,0x40(%rsp)
  138518:	movq   $0x8,0x50(%rsp)
  138521:	movq   $0x0,0x58(%rsp)
  13852a:	movl   $0x0,0x5f(%rsp)
  138532:	movw   $0x1,0x63(%rsp)
  138539:	movb   $0x1,0x65(%rsp)
  13853e:	movdqu %xmm0,0x66(%rsp)
  138544:	movl   $0x0,0x75(%rsp)
  13854c:	lea    0xb0(%rsp),%rdi
  138554:	mov    %rbx,0x190(%rsp)
  13855c:	call   *0xf4586(%rip)        # 22cae8 <_DYNAMIC+0x8e8>
  138562:	movups 0x80(%rsp),%xmm0
  13856a:	movaps %xmm0,0x300(%rsp)
  138572:	movups 0x70(%rsp),%xmm0
  138577:	movaps %xmm0,0x2f0(%rsp)
  13857f:	movups 0x30(%rsp),%xmm0
  138584:	movups 0x40(%rsp),%xmm1
  138589:	movups 0x50(%rsp),%xmm2
  13858e:	movups 0x60(%rsp),%xmm3
  138593:	movaps %xmm3,0x2e0(%rsp)
  13859b:	movaps %xmm2,0x2d0(%rsp)
  1385a3:	movaps %xmm1,0x2c0(%rsp)
  1385ab:	movaps %xmm0,0x2b0(%rsp)
  1385b3:	movups 0xb0(%rsp),%xmm0
  1385bb:	movdqu 0xc0(%rsp),%xmm1
  1385c4:	movups 0xd0(%rsp),%xmm2
  1385cc:	movups 0xe0(%rsp),%xmm3
  1385d4:	movaps %xmm0,0x310(%rsp)
  1385dc:	movdqa %xmm1,0x320(%rsp)
  1385e5:	movaps %xmm2,0x330(%rsp)
  1385ed:	movaps %xmm3,0x340(%rsp)
  1385f5:	movb   $0x0,0x350(%rsp)
  1385fd:	cmpq   $0x3,0x250(%rsp)
  138606:	movabs $0x8000000000000001,%r15
  138610:	jb     138639 <litchi_pptx::shape::reader::Scene::read_with+0x4d9>
  138612:	mov    0x248(%rsp),%rax
  13861a:	cmpb   $0xef,(%rax)
  13861d:	jne    138639 <litchi_pptx::shape::reader::Scene::read_with+0x4d9>
  13861f:	cmpb   $0xbb,0x1(%rax)
  138623:	jne    138639 <litchi_pptx::shape::reader::Scene::read_with+0x4d9>
  138625:	xor    %ecx,%ecx
  138627:	cmpb   $0xbf,0x2(%rax)
  13862b:	sete   %cl
  13862e:	lea    (%rcx,%rcx,2),%rax
  138632:	mov    %rax,0x8(%rsp)
  138637:	jmp    138642 <litchi_pptx::shape::reader::Scene::read_with+0x4e2>
  138639:	movq   $0x0,0x8(%rsp)
  138642:	mov    %r12,0x2a8(%rsp)
  13864a:	lea    0x30(%rsp),%rbp
  13864f:	mov    $0x12,%eax
  138654:	movq   %rax,%xmm0
  138659:	movdqa %xmm0,0x20(%rsp)
  13865f:	mov    $0x17,%eax
  138664:	movq   %rax,%xmm0
  138669:	movdqa %xmm0,0x140(%rsp)
  138672:	mov    0x8(%rsp),%rbx
  138677:	nopw   0x0(%rax,%rax,1)
  138680:	mov    0x2e8(%rsp),%r12
  138688:	add    %rbx,%r12
  13868b:	jb     1395ef <litchi_pptx::shape::reader::Scene::read_with+0x148f>
  138691:	cmpb   $0x1,0x350(%rsp)
  138699:	jne    138732 <litchi_pptx::shape::reader::Scene::read_with+0x5d2>
  13869f:	movzwl 0x348(%rsp),%edx
  1386a7:	cmp    $0x1,%dx
  1386ab:	adc    $0xffffffff,%edx
  1386ae:	mov    %dx,0x348(%rsp)
  1386b6:	mov    0x330(%rsp),%rax
  1386be:	mov    0x338(%rsp),%rsi
  1386c6:	mov    %rsi,%rcx
  1386c9:	shl    $0x5,%rcx
  1386cd:	lea    (%rax,%rcx,1),%r8
  1386d1:	xor    %edi,%edi
  1386d3:	data16 data16 data16 cs nopw 0x0(%rax,%rax,1)
  1386e0:	test   %rcx,%rcx
  1386e3:	je     138716 <litchi_pptx::shape::reader::Scene::read_with+0x5b6>
  1386e5:	add    $0xffffffffffffffe0,%rcx
  1386e9:	inc    %rdi
  1386ec:	cmp    %dx,-0x8(%r8)
  1386f1:	lea    -0x20(%r8),%r8
  1386f5:	ja     1386e0 <litchi_pptx::shape::reader::Scene::read_with+0x580>
  1386f7:	mov    %rsi,%rdx
  1386fa:	sub    %rdi,%rdx
  1386fd:	inc    %rdx
  138700:	cmp    %rsi,%rdx
  138703:	jae    13872a <litchi_pptx::shape::reader::Scene::read_with+0x5ca>
  138705:	mov    0x20(%rax,%rcx,1),%rax
  13870a:	cmp    0x320(%rsp),%rax
  138712:	jbe    13871a <litchi_pptx::shape::reader::Scene::read_with+0x5ba>
  138714:	jmp    138722 <litchi_pptx::shape::reader::Scene::read_with+0x5c2>
  138716:	xor    %eax,%eax
  138718:	xor    %edx,%edx
  13871a:	mov    %rax,0x320(%rsp)
  138722:	mov    %rdx,0x338(%rsp)
  13872a:	movb   $0x0,0x350(%rsp)
  138732:	mov    %rbp,%rdi
  138735:	lea    0x2b0(%rsp),%r13
  13873d:	mov    %r13,%rsi
  138740:	call   179230 <quick_xml::reader::Reader<R>::read_event_impl>
  138745:	lea    0xb0(%rsp),%rdi
  13874d:	mov    %r13,%rsi
  138750:	mov    %rbp,%rdx
  138753:	call   179e40 <quick_xml::reader::ns_reader::NsReader<R>::process_event>
  138758:	mov    0xb0(%rsp),%rax
  138760:	lea    0xb8(%rsp),%rdx
  138768:	mov    0x20(%rdx),%rcx
  13876c:	mov    %rcx,0x130(%rsp)
  138774:	lea    0xe(%r15),%rcx
  138778:	movups (%rdx),%xmm0
  13877b:	movups 0x10(%rdx),%xmm1
  13877f:	movaps %xmm0,0x110(%rsp)
  138787:	movaps %xmm1,0x120(%rsp)
  13878f:	cmp    %rcx,%rax
  138792:	jne    139867 <litchi_pptx::shape::reader::Scene::read_with+0x1707>
  138798:	mov    0x130(%rsp),%rax
  1387a0:	mov    %rax,0x1c0(%rsp)
  1387a8:	movdqa 0x110(%rsp),%xmm0
  1387b1:	movdqa 0x120(%rsp),%xmm1
  1387ba:	movdqa %xmm1,0x1b0(%rsp)
  1387c3:	movdqa %xmm0,0x1a0(%rsp)
  1387cc:	mov    0x2e8(%rsp),%r13
  1387d4:	add    %rbx,%r13
  1387d7:	jb     139629 <litchi_pptx::shape::reader::Scene::read_with+0x14c9>
  1387dd:	mov    0x1c0(%rsp),%rax
  1387e5:	mov    %rax,0xd0(%rsp)
  1387ed:	movdqa 0x1a0(%rsp),%xmm0
  1387f6:	movdqa 0x1b0(%rsp),%xmm1
  1387ff:	movdqa %xmm1,0xc0(%rsp)
  138808:	movdqa %xmm0,0xb0(%rsp)
  138811:	mov    %rbp,%rdi
  138814:	lea    0x310(%rsp),%rsi
  13881c:	lea    0xb0(%rsp),%rdx
  138824:	call   *0xf42c6(%rip)        # 22caf0 <_DYNAMIC+0x8f0>
  13882a:	mov    0x40(%rsp),%rax
  13882f:	mov    %rax,0xa0(%rsp)
  138837:	movups 0x30(%rsp),%xmm0
  13883c:	movaps %xmm0,0x90(%rsp)
  138844:	lea    0x48(%rsp),%rcx
  138849:	mov    0x20(%rcx),%rax
  13884d:	mov    %rax,0x180(%rsp)
  138855:	movdqu (%rcx),%xmm0
  138859:	movdqu 0x10(%rcx),%xmm1
  13885e:	movdqa %xmm1,0x170(%rsp)
  138867:	movdqa %xmm0,0x160(%rsp)
  138870:	mov    0x160(%rsp),%rax
  138878:	cmp    $0xa,%rax
  13887c:	ja     138e7c <litchi_pptx::shape::reader::Scene::read_with+0xd1c>
  138882:	lea    -0x1146ad(%rip),%rdx        # 241dc <anon.bc3d94ef2394061926bee58178356194.284.llvm.5373085861489949100+0x41e4>
  138889:	movslq (%rdx,%rax,4),%rcx
  13888d:	add    %rdx,%rcx
  138890:	jmp    *%rcx
  138892:	lea    0x168(%rsp),%rax
  13889a:	movdqu (%rax),%xmm0
  13889e:	movdqu 0x10(%rax),%xmm1
  1388a3:	movdqa %xmm1,0xc0(%rsp)
  1388ac:	movdqa %xmm0,0xb0(%rsp)
  1388b5:	mov    0x270(%rsp),%rbx
  1388bd:	mov    0x298(%rsp),%rbp
  1388c5:	cmp    $0xffffffffffffffff,%rbp
  1388c9:	je     13a157 <litchi_pptx::shape::reader::Scene::read_with+0x1ff7>
  1388cf:	lea    -0x111646(%rip),%rax        # 27290 <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13130224050040044120+0x98>
  1388d6:	mov    %rax,0x40(%rsp)
  1388db:	movq   $0x12,0x48(%rsp)
  1388e4:	mov    %rbx,0x38(%rsp)
  1388e9:	movb   $0x9,0x30(%rsp)
  1388ee:	lea    0x30(%rsp),%rdi
  1388f3:	call   161950 <core::ptr::drop_in_place<litchi_pptx::error::Error>>
  1388f8:	lea    0x1(%rbp),%rax
  1388fc:	mov    %rax,0x298(%rsp)
  138904:	mov    $0x9,%r15b
  138907:	cmp    %rbx,%rbp
  13890a:	jae    13a16a <litchi_pptx::shape::reader::Scene::read_with+0x200a>
  138910:	mov    0x268(%rsp),%rbx
  138918:	mov    0x290(%rsp),%rbp
  138920:	lea    -0x11155d(%rip),%rax        # 273ca <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13130224050040044120+0x1d2>
  138927:	mov    %rax,0x40(%rsp)
  13892c:	movq   $0x17,0x48(%rsp)
  138935:	inc    %rbp
  138938:	je     13a3f9 <litchi_pptx::shape::reader::Scene::read_with+0x2299>
  13893e:	mov    %rbx,0x38(%rsp)
  138943:	movb   $0x9,0x30(%rsp)
  138948:	lea    0x30(%rsp),%rdi
  13894d:	call   161950 <core::ptr::drop_in_place<litchi_pptx::error::Error>>
  138952:	cmp    %rbx,%rbp
  138955:	ja     139992 <litchi_pptx::shape::reader::Scene::read_with+0x1832>
  13895b:	mov    0x240(%rsp),%rax
  138963:	test   %rax,%rax
  138966:	je     138f66 <litchi_pptx::shape::reader::Scene::read_with+0xe06>
  13896c:	mov    0x238(%rsp),%rcx
  138974:	movzbl -0x1(%rcx,%rax,1),%eax
  138979:	jmp    138f68 <litchi_pptx::shape::reader::Scene::read_with+0xe08>
  13897e:	mov    0x168(%rsp),%r12
  138986:	mov    0x170(%rsp),%r13
  13898e:	mov    0x178(%rsp),%rcx
  138996:	mov    %rbp,%rdi
  138999:	lea    0x1c8(%rsp),%rsi
  1389a1:	mov    %r13,%rdx
  1389a4:	xor    %r8d,%r8d
  1389a7:	call   13f320 <litchi_pptx::shape::reader::Scanner::reject_placeholder_character_data>
  1389ac:	movzbl 0x30(%rsp),%r15d
  1389b2:	cmp    $0x27,%r15b
  1389b6:	jne    139be8 <litchi_pptx::shape::reader::Scene::read_with+0x1a88>
  1389bc:	mov    0x228(%rsp),%rax
  1389c4:	test   %rax,%rax
  1389c7:	je     138ac6 <litchi_pptx::shape::reader::Scene::read_with+0x966>
  1389cd:	mov    0x220(%rsp),%rcx
  1389d5:	shl    $0x8,%rax
  1389d9:	cmpq   $0x0,-0xc0(%rcx,%rax,1)
  1389e2:	je     138ac6 <litchi_pptx::shape::reader::Scene::read_with+0x966>
  1389e8:	lea    0xb0(%rsp),%rdi
  1389f0:	lea    0x168(%rsp),%rsi
  1389f8:	call   1a2840 <quick_xml::encoding::Decoder::content>
  1389fd:	mov    0xb0(%rsp),%rbx
  138a05:	movabs $0x8000000000000001,%r15
  138a0f:	cmp    %r15,%rbx
  138a12:	mov    %r12,0x18(%rsp)
  138a17:	je     139d44 <litchi_pptx::shape::reader::Scene::read_with+0x1be4>
  138a1d:	mov    0xb8(%rsp),%rsi
  138a25:	mov    0xc0(%rsp),%rdx
  138a2d:	lea    0x110(%rsp),%rdi
  138a35:	mov    %rsi,0xa8(%rsp)
  138a3d:	call   *0xf4795(%rip)        # 22d1d8 <_DYNAMIC+0xfd8>
  138a43:	lea    0x2(%r15),%rax
  138a47:	cmp    %rax,0x110(%rsp)
  138a4f:	jne    139dbe <litchi_pptx::shape::reader::Scene::read_with+0x1c5e>
  138a55:	mov    0x118(%rsp),%r12
  138a5d:	mov    %rbp,%rdi
  138a60:	mov    0x120(%rsp),%rbp
  138a68:	mov    0x128(%rsp),%rcx
  138a70:	lea    0x1c8(%rsp),%rsi
  138a78:	mov    %rbp,%rdx
  138a7b:	call   13aa10 <litchi_pptx::shape::reader::Scanner::append_text>
  138a80:	movzbl 0x30(%rsp),%r15d
  138a86:	cmp    $0x27,%r15b
  138a8a:	jne    139e54 <litchi_pptx::shape::reader::Scene::read_with+0x1cf4>
  138a90:	shl    $1,%r12
  138a93:	test   %r12,%r12
  138a96:	je     138aa1 <litchi_pptx::shape::reader::Scene::read_with+0x941>
  138a98:	mov    %rbp,%rdi
  138a9b:	call   *0xf3957(%rip)        # 22c3f8 <free@GLIBC_2.2.5>
  138aa1:	shl    $1,%rbx
  138aa4:	test   %rbx,%rbx
  138aa7:	mov    0x8(%rsp),%rbx
  138aac:	lea    0x30(%rsp),%rbp
  138ab1:	mov    0x18(%rsp),%r12
  138ab6:	je     138ac6 <litchi_pptx::shape::reader::Scene::read_with+0x966>
  138ab8:	mov    0xa8(%rsp),%rdi
  138ac0:	call   *0xf3932(%rip)        # 22c3f8 <free@GLIBC_2.2.5>
  138ac6:	test   %r12,%r12
  138ac9:	movabs $0x8000000000000001,%r15
  138ad3:	jle    139573 <litchi_pptx::shape::reader::Scene::read_with+0x1413>
  138ad9:	mov    %r13,%rdi
  138adc:	jmp    13956d <litchi_pptx::shape::reader::Scene::read_with+0x140d>
  138ae1:	lea    0x168(%rsp),%rax
  138ae9:	movdqu (%rax),%xmm0
  138aed:	movdqu 0x10(%rax),%xmm1
  138af2:	movdqa %xmm1,0xc0(%rsp)
  138afb:	movdqa %xmm0,0xb0(%rsp)
  138b04:	mov    0x270(%rsp),%rbx
  138b0c:	mov    0x298(%rsp),%rbp
  138b14:	cmp    $0xffffffffffffffff,%rbp
  138b18:	je     13a157 <litchi_pptx::shape::reader::Scene::read_with+0x1ff7>
  138b1e:	lea    -0x111895(%rip),%rax        # 27290 <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13130224050040044120+0x98>
  138b25:	mov    %rax,0x40(%rsp)
  138b2a:	movq   $0x12,0x48(%rsp)
  138b33:	mov    %rbx,0x38(%rsp)
  138b38:	movb   $0x9,0x30(%rsp)
  138b3d:	lea    0x30(%rsp),%rdi
  138b42:	call   161950 <core::ptr::drop_in_place<litchi_pptx::error::Error>>
  138b47:	lea    0x1(%rbp),%rax
  138b4b:	mov    %rax,0x298(%rsp)
  138b53:	mov    $0x9,%r15b
  138b56:	cmp    %rbx,%rbp
  138b59:	jae    13a16a <litchi_pptx::shape::reader::Scene::read_with+0x200a>
  138b5f:	mov    0x268(%rsp),%rbx
  138b67:	mov    0x290(%rsp),%rbp
  138b6f:	lea    -0x1117ac(%rip),%rax        # 273ca <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13130224050040044120+0x1d2>
  138b76:	mov    %rax,0x40(%rsp)
  138b7b:	movq   $0x17,0x48(%rsp)
  138b84:	inc    %rbp
  138b87:	je     13a3f9 <litchi_pptx::shape::reader::Scene::read_with+0x2299>
  138b8d:	mov    %rbx,0x38(%rsp)
  138b92:	movb   $0x9,0x30(%rsp)
  138b97:	lea    0x30(%rsp),%rdi
  138b9c:	call   161950 <core::ptr::drop_in_place<litchi_pptx::error::Error>>
  138ba1:	cmp    %rbx,%rbp
  138ba4:	ja     139992 <litchi_pptx::shape::reader::Scene::read_with+0x1832>
  138baa:	mov    0x240(%rsp),%rax
  138bb2:	test   %rax,%rax
  138bb5:	je     138eea <litchi_pptx::shape::reader::Scene::read_with+0xd8a>
  138bbb:	mov    0x238(%rsp),%rcx
  138bc3:	movzbl -0x1(%rcx,%rax,1),%eax
  138bc8:	jmp    138eec <litchi_pptx::shape::reader::Scene::read_with+0xd8c>
  138bcd:	mov    0x168(%rsp),%rax
  138bd5:	mov    %rax,0xa8(%rsp)
  138bdd:	mov    0x170(%rsp),%r12
  138be5:	mov    0x290(%rsp),%r15
  138bed:	test   %r15,%r15
  138bf0:	mov    %r12,0x18(%rsp)
  138bf5:	je     139ba0 <litchi_pptx::shape::reader::Scene::read_with+0x1a40>
  138bfb:	mov    0x178(%rsp),%rbx
  138c03:	lea    (%r12,%rbx,1),%rdx
  138c07:	mov    0xf3caa(%rip),%rax        # 22c8b8 <_DYNAMIC+0x6b8>
  138c0e:	mov    (%rax),%rax
  138c11:	mov    $0x3a,%edi
  138c16:	mov    %r12,%rsi
  138c19:	call   *%rax
  138c1b:	cmp    $0x1,%rax
  138c1f:	jne    138ea8 <litchi_pptx::shape::reader::Scene::read_with+0xd48>
  138c25:	mov    %rdx,%rax
  138c28:	sub    %r12,%rax
  138c2b:	not    %rax
  138c2e:	add    %rax,%rbx
  138c31:	inc    %rdx
  138c34:	mov    0x228(%rsp),%rbp
  138c3c:	test   %rbp,%rbp
  138c3f:	jne    138ebc <litchi_pptx::shape::reader::Scene::read_with+0xd5c>
  138c45:	jmp    13940d <litchi_pptx::shape::reader::Scene::read_with+0x12ad>
  138c4a:	mov    0x168(%rsp),%rbx
  138c52:	mov    0x170(%rsp),%r13
  138c5a:	mov    0x228(%rsp),%rcx
  138c62:	test   %rcx,%rcx
  138c65:	je     138d38 <litchi_pptx::shape::reader::Scene::read_with+0xbd8>
  138c6b:	mov    0x220(%rsp),%rax
  138c73:	shl    $0x8,%rcx
  138c77:	cmpb   $0x0,-0x80(%rax,%rcx,1)
  138c7c:	jne    139a13 <litchi_pptx::shape::reader::Scene::read_with+0x18b3>
  138c82:	add    %rcx,%rax
  138c85:	cmpq   $0x0,-0x70(%rax)
  138c8a:	jne    139a13 <litchi_pptx::shape::reader::Scene::read_with+0x18b3>
  138c90:	cmpb   $0x0,-0x60(%rax)
  138c94:	jne    139a13 <litchi_pptx::shape::reader::Scene::read_with+0x18b3>
  138c9a:	cmpb   $0x0,-0x90(%rax)
  138ca1:	je     138cb8 <litchi_pptx::shape::reader::Scene::read_with+0xb58>
  138ca3:	mov    -0x88(%rax),%rcx
  138caa:	cmp    0x290(%rsp),%rcx
  138cb2:	je     13a03b <litchi_pptx::shape::reader::Scene::read_with+0x1edb>
  138cb8:	cmpq   $0x0,-0xc0(%rax)
  138cc0:	je     138d38 <litchi_pptx::shape::reader::Scene::read_with+0xbd8>
  138cc2:	lea    0xb0(%rsp),%rdi
  138cca:	lea    0x168(%rsp),%rsi
  138cd2:	call   1a2840 <quick_xml::encoding::Decoder::content>
  138cd7:	mov    0xb0(%rsp),%r12
  138cdf:	cmp    %r15,%r12
  138ce2:	je     139ee6 <litchi_pptx::shape::reader::Scene::read_with+0x1d86>
  138ce8:	mov    0xb8(%rsp),%rbp
  138cf0:	mov    0xc0(%rsp),%rcx
  138cf8:	lea    0x30(%rsp),%rdi
  138cfd:	lea    0x1c8(%rsp),%rsi
  138d05:	mov    %rbp,%rdx
  138d08:	call   13aa10 <litchi_pptx::shape::reader::Scanner::append_text>
  138d0d:	movzbl 0x30(%rsp),%r15d
  138d13:	cmp    $0x27,%r15b
  138d17:	jne    139fd6 <litchi_pptx::shape::reader::Scene::read_with+0x1e76>
  138d1d:	shl    $1,%r12
  138d20:	test   %r12,%r12
  138d23:	movabs $0x8000000000000001,%r15
  138d2d:	je     138d38 <litchi_pptx::shape::reader::Scene::read_with+0xbd8>
  138d2f:	mov    %rbp,%rdi
  138d32:	call   *0xf36c0(%rip)        # 22c3f8 <free@GLIBC_2.2.5>
  138d38:	test   %rbx,%rbx
  138d3b:	mov    0x8(%rsp),%rbx
  138d40:	lea    0x30(%rsp),%rbp
  138d45:	jle    139573 <litchi_pptx::shape::reader::Scene::read_with+0x1413>
  138d4b:	mov    %r13,%rdi
  138d4e:	jmp    13956d <litchi_pptx::shape::reader::Scene::read_with+0x140d>
  138d53:	lea    0x168(%rsp),%rcx
  138d5b:	mov    0x10(%rcx),%rax
  138d5f:	mov    %rax,0x120(%rsp)
  138d67:	movdqu (%rcx),%xmm0
  138d6b:	movdqa %xmm0,0x110(%rsp)
  138d74:	mov    0x228(%rsp),%rcx
  138d7c:	test   %rcx,%rcx
  138d7f:	je     138e60 <litchi_pptx::shape::reader::Scene::read_with+0xd00>
  138d85:	mov    0x220(%rsp),%rax
  138d8d:	shl    $0x8,%rcx
  138d91:	cmpb   $0x0,-0x80(%rax,%rcx,1)
  138d96:	jne    139a87 <litchi_pptx::shape::reader::Scene::read_with+0x1927>
  138d9c:	add    %rcx,%rax
  138d9f:	cmpq   $0x0,-0x70(%rax)
  138da4:	jne    139a87 <litchi_pptx::shape::reader::Scene::read_with+0x1927>
  138daa:	cmpb   $0x0,-0x60(%rax)
  138dae:	jne    139a87 <litchi_pptx::shape::reader::Scene::read_with+0x1927>
  138db4:	cmpb   $0x0,-0x90(%rax)
  138dbb:	je     138dd2 <litchi_pptx::shape::reader::Scene::read_with+0xc72>
  138dbd:	mov    -0x88(%rax),%rcx
  138dc4:	cmp    0x290(%rsp),%rcx
  138dcc:	je     13a088 <litchi_pptx::shape::reader::Scene::read_with+0x1f28>
  138dd2:	cmpq   $0x0,-0xc0(%rax)
  138dda:	je     138e60 <litchi_pptx::shape::reader::Scene::read_with+0xd00>
  138de0:	lea    0xb0(%rsp),%rdi
  138de8:	lea    0x110(%rsp),%rsi
  138df0:	call   *0xf3c0a(%rip)        # 22ca00 <_DYNAMIC+0x800>
  138df6:	mov    0xb0(%rsp),%r13
  138dfe:	mov    0xb8(%rsp),%rbx
  138e06:	mov    0xc0(%rsp),%r12
  138e0e:	mov    0xc8(%rsp),%rcx
  138e16:	cmp    $0x2,%r13
  138e1a:	jne    139f5d <litchi_pptx::shape::reader::Scene::read_with+0x1dfd>
  138e20:	mov    %rbp,%rdi
  138e23:	lea    0x1c8(%rsp),%rsi
  138e2b:	mov    %r12,%rdx
  138e2e:	call   13aa10 <litchi_pptx::shape::reader::Scanner::append_text>
  138e33:	movzbl 0x30(%rsp),%r15d
  138e39:	cmp    $0x27,%r15b
  138e3d:	jne    139f7c <litchi_pptx::shape::reader::Scene::read_with+0x1e1c>
  138e43:	test   %rbx,%rbx
  138e46:	movabs $0x8000000000000001,%r15
  138e50:	mov    0x8(%rsp),%rbx
  138e55:	je     138e60 <litchi_pptx::shape::reader::Scene::read_with+0xd00>
  138e57:	mov    %r12,%rdi
  138e5a:	call   *0xf3598(%rip)        # 22c3f8 <free@GLIBC_2.2.5>
  138e60:	cmpq   $0x0,0x110(%rsp)
  138e69:	jle    139573 <litchi_pptx::shape::reader::Scene::read_with+0x1413>
  138e6f:	mov    0x118(%rsp),%rdi
  138e77:	jmp    13956d <litchi_pptx::shape::reader::Scene::read_with+0x140d>
  138e7c:	cmp    $0x5,%rax
  138e80:	je     138e8c <litchi_pptx::shape::reader::Scene::read_with+0xd2c>
  138e82:	cmp    $0x6,%rax
  138e86:	jne    139573 <litchi_pptx::shape::reader::Scene::read_with+0x1413>
  138e8c:	cmpq   $0x0,0x168(%rsp)
  138e95:	jle    139573 <litchi_pptx::shape::reader::Scene::read_with+0x1413>
  138e9b:	mov    0x170(%rsp),%rdi
  138ea3:	jmp    13956d <litchi_pptx::shape::reader::Scene::read_with+0x140d>
  138ea8:	mov    %r12,%rdx
  138eab:	mov    0x228(%rsp),%rbp
  138eb3:	test   %rbp,%rbp
  138eb6:	je     13940d <litchi_pptx::shape::reader::Scene::read_with+0x12ad>
  138ebc:	mov    %rbp,%r12
  138ebf:	shl    $0x8,%r12
  138ec3:	add    0x220(%rsp),%r12
  138ecb:	cmp    $0x5,%rbx
  138ecf:	jne    139074 <litchi_pptx::shape::reader::Scene::read_with+0xf14>
  138ed5:	mov    (%rdx),%eax
  138ed7:	mov    $0x656d6163,%ecx
  138edc:	xor    %ecx,%eax
  138ede:	movzbl 0x4(%rdx),%ecx
  138ee2:	xor    $0x6f,%ecx
  138ee5:	jmp    139091 <litchi_pptx::shape::reader::Scene::read_with+0xf31>
  138eea:	xor    %eax,%eax
  138eec:	sub    $0x8,%rsp
  138ef0:	movzbl %al,%eax
  138ef3:	lea    0x38(%rsp),%rdi
  138ef8:	lea    0x1d0(%rsp),%rsi
  138f00:	lea    0x98(%rsp),%rdx
  138f08:	lea    0xb8(%rsp),%rcx
  138f10:	mov    %r12,%r8
  138f13:	mov    %rbp,%r9
  138f16:	push   %r13
  138f18:	push   $0x1
  138f1a:	push   %rax
  138f1b:	call   13b8e0 <litchi_pptx::shape::reader::Scanner::start_element>
  138f20:	add    $0x20,%rsp
  138f24:	movzbl 0x30(%rsp),%r15d
  138f2a:	cmp    $0x27,%r15b
  138f2e:	jne    1399bf <litchi_pptx::shape::reader::Scene::read_with+0x185f>
  138f34:	cmpq   $0x0,0xb0(%rsp)
  138f3d:	jle    138f4d <litchi_pptx::shape::reader::Scene::read_with+0xded>
  138f3f:	mov    0xb8(%rsp),%rdi
  138f47:	call   *0xf34ab(%rip)        # 22c3f8 <free@GLIBC_2.2.5>
  138f4d:	movabs $0x8000000000000001,%r15
  138f57:	mov    0x8(%rsp),%rbx
  138f5c:	lea    0x30(%rsp),%rbp
  138f61:	jmp    139573 <litchi_pptx::shape::reader::Scene::read_with+0x1413>
  138f66:	xor    %eax,%eax
  138f68:	sub    $0x8,%rsp
  138f6c:	movzbl %al,%eax
  138f6f:	lea    0x38(%rsp),%rdi
  138f74:	lea    0x1d0(%rsp),%rsi
  138f7c:	lea    0x98(%rsp),%rdx
  138f84:	lea    0xb8(%rsp),%rcx
  138f8c:	mov    %r12,%r8
  138f8f:	mov    %rbp,%r9
  138f92:	push   %r13
  138f94:	push   $0x0
  138f96:	push   %rax
  138f97:	call   13b8e0 <litchi_pptx::shape::reader::Scanner::start_element>
  138f9c:	add    $0x20,%rsp
  138fa0:	movzbl 0x30(%rsp),%r15d
  138fa6:	cmp    $0x27,%r15b
  138faa:	jne    1399bf <litchi_pptx::shape::reader::Scene::read_with+0x185f>
  138fb0:	mov    0xc0(%rsp),%rdx
  138fb8:	mov    0xc8(%rsp),%r12
  138fc0:	cmp    %rdx,%r12
  138fc3:	ja     13a427 <litchi_pptx::shape::reader::Scene::read_with+0x22c7>
  138fc9:	mov    0xb8(%rsp),%r13
  138fd1:	lea    (%r12,%r13,1),%rdx
  138fd5:	mov    0xf38dc(%rip),%rax        # 22c8b8 <_DYNAMIC+0x6b8>
  138fdc:	mov    (%rax),%rax
  138fdf:	mov    $0x3a,%edi
  138fe4:	mov    %r13,%rsi
  138fe7:	call   *%rax
  138fe9:	cmp    $0x1,%rax
  138fed:	jne    139001 <litchi_pptx::shape::reader::Scene::read_with+0xea1>
  138fef:	mov    %rdx,%rax
  138ff2:	sub    %r13,%rax
  138ff5:	not    %rax
  138ff8:	add    %rax,%r12
  138ffb:	inc    %rdx
  138ffe:	mov    %rdx,%r13
  139001:	movabs $0x8000000000000001,%rax
  13900b:	xor    %ebx,%ebx
  13900d:	cmp    $0x4,%r12
  139011:	jne    139278 <litchi_pptx::shape::reader::Scene::read_with+0x1118>
  139017:	cmpl   $0x7250766e,0x0(%r13)
  13901f:	jne    139278 <litchi_pptx::shape::reader::Scene::read_with+0x1118>
  139025:	cmp    %rax,0x90(%rsp)
  13902d:	jne    139278 <litchi_pptx::shape::reader::Scene::read_with+0x1118>
  139033:	mov    0x98(%rsp),%rdi
  13903b:	mov    0xa0(%rsp),%rax
  139043:	cmp    $0x2e,%rax
  139047:	je     13925c <litchi_pptx::shape::reader::Scene::read_with+0x10fc>
  13904d:	cmp    $0x3a,%rax
  139051:	jne    13906d <litchi_pptx::shape::reader::Scene::read_with+0xf0d>
  139053:	mov    $0x3a,%edx
  139058:	lea    -0x1145a4(%rip),%rsi        # 24abb <anon.bc3d94ef2394061926bee58178356194.284.llvm.5373085861489949100+0x4ac3>
  13905f:	call   *0xf358b(%rip)        # 22c5f0 <bcmp@GLIBC_2.2.5>
  139065:	test   %eax,%eax
  139067:	je     139276 <litchi_pptx::shape::reader::Scene::read_with+0x1116>
  13906d:	xor    %ebx,%ebx
  13906f:	jmp    139278 <litchi_pptx::shape::reader::Scene::read_with+0x1118>
  139074:	cmp    $0x7,%rbx
  139078:	jne    13910c <litchi_pptx::shape::reader::Scene::read_with+0xfac>
  13907e:	mov    (%rdx),%eax
  139080:	mov    $0x6e6b6e75,%ecx
  139085:	xor    %ecx,%eax
  139087:	mov    0x3(%rdx),%ecx
  13908a:	mov    $0x6e776f6e,%edx
  13908f:	xor    %edx,%ecx
  139091:	or     %eax,%ecx
  139093:	mov    0x8(%rsp),%rbx
  139098:	jne    13947a <litchi_pptx::shape::reader::Scene::read_with+0x131a>
  13909e:	movabs $0x8000000000000001,%rax
  1390a8:	cmp    %rax,0x90(%rsp)
  1390b0:	jne    13947a <litchi_pptx::shape::reader::Scene::read_with+0x131a>
  1390b6:	cmpq   $0x3b,0xa0(%rsp)
  1390bf:	jne    13947a <litchi_pptx::shape::reader::Scene::read_with+0x131a>
  1390c5:	mov    0x98(%rsp),%rdi
  1390cd:	mov    $0x3b,%edx
  1390d2:	lea    -0x1138f3(%rip),%rsi        # 257e6 <anon.bc3d94ef2394061926bee58178356194.284.llvm.5373085861489949100+0x57ee>
  1390d9:	call   *0xf3511(%rip)        # 22c5f0 <bcmp@GLIBC_2.2.5>
  1390df:	test   %eax,%eax
  1390e1:	jne    13947a <litchi_pptx::shape::reader::Scene::read_with+0x131a>
  1390e7:	cmpl   $0x1,-0x60(%r12)
  1390ed:	jne    13947a <litchi_pptx::shape::reader::Scene::read_with+0x131a>
  1390f3:	cmp    %r15,-0x58(%r12)
  1390f8:	jne    13947a <litchi_pptx::shape::reader::Scene::read_with+0x131a>
  1390fe:	movq   $0x0,-0x60(%r12)
  139107:	jmp    13947a <litchi_pptx::shape::reader::Scene::read_with+0x131a>
  13910c:	cmp    $0x4,%rbx
  139110:	jne    1391a1 <litchi_pptx::shape::reader::Scene::read_with+0x1041>
  139116:	cmpl   $0x65707974,(%rdx)
  13911c:	mov    0x8(%rsp),%rbx
  139121:	jne    13947a <litchi_pptx::shape::reader::Scene::read_with+0x131a>
  139127:	movabs $0x8000000000000001,%rax
  139131:	cmp    %rax,0x90(%rsp)
  139139:	jne    13947a <litchi_pptx::shape::reader::Scene::read_with+0x131a>
  13913f:	cmpq   $0x3b,0xa0(%rsp)
  139148:	jne    13947a <litchi_pptx::shape::reader::Scene::read_with+0x131a>
  13914e:	mov    0x98(%rsp),%rdi
  139156:	mov    $0x3b,%edx
  13915b:	lea    -0x11397c(%rip),%rsi        # 257e6 <anon.bc3d94ef2394061926bee58178356194.284.llvm.5373085861489949100+0x57ee>
  139162:	call   *0xf3488(%rip)        # 22c5f0 <bcmp@GLIBC_2.2.5>
  139168:	test   %eax,%eax
  13916a:	jne    13947a <litchi_pptx::shape::reader::Scene::read_with+0x131a>
  139170:	cmpb   $0x0,-0x70(%r12)
  139176:	je     13947a <litchi_pptx::shape::reader::Scene::read_with+0x131a>
  13917c:	cmp    %r15,-0x68(%r12)
  139181:	jne    13947a <litchi_pptx::shape::reader::Scene::read_with+0x131a>
  139187:	cmpb   $0x0,-0xb(%r12)
  13918d:	je     13a528 <litchi_pptx::shape::reader::Scene::read_with+0x23c8>
  139193:	movq   $0x0,-0x70(%r12)
  13919c:	jmp    13947a <litchi_pptx::shape::reader::Scene::read_with+0x131a>
  1391a1:	cmp    $0x9,%rbx
  1391a5:	jne    1392cb <litchi_pptx::shape::reader::Scene::read_with+0x116b>
  1391ab:	mov    (%rdx),%rax
  1391ae:	movabs $0x7845657079546870,%rcx
  1391b8:	xor    %rcx,%rax
  1391bb:	movzbl 0x8(%rdx),%ecx
  1391bf:	xor    $0x74,%rcx
  1391c3:	or     %rax,%rcx
  1391c6:	mov    0x8(%rsp),%rbx
  1391cb:	jne    13947a <litchi_pptx::shape::reader::Scene::read_with+0x131a>
  1391d1:	movabs $0x8000000000000001,%rax
  1391db:	cmp    %rax,0x90(%rsp)
  1391e3:	jne    13947a <litchi_pptx::shape::reader::Scene::read_with+0x131a>
  1391e9:	cmpq   $0x3b,0xa0(%rsp)
  1391f2:	jne    13947a <litchi_pptx::shape::reader::Scene::read_with+0x131a>
  1391f8:	mov    0x98(%rsp),%rdi
  139200:	mov    $0x3b,%edx
  139205:	lea    -0x113a26(%rip),%rsi        # 257e6 <anon.bc3d94ef2394061926bee58178356194.284.llvm.5373085861489949100+0x57ee>
  13920c:	call   *0xf33de(%rip)        # 22c5f0 <bcmp@GLIBC_2.2.5>
  139212:	test   %eax,%eax
  139214:	jne    13947a <litchi_pptx::shape::reader::Scene::read_with+0x131a>
  13921a:	cmpb   $0x0,-0x80(%r12)
  139220:	je     13947a <litchi_pptx::shape::reader::Scene::read_with+0x131a>
  139226:	cmp    %r15,-0x78(%r12)
  13922b:	jne    13947a <litchi_pptx::shape::reader::Scene::read_with+0x131a>
  139231:	cmpb   $0x0,-0xc(%r12)
  139237:	je     13a571 <litchi_pptx::shape::reader::Scene::read_with+0x2411>
  13923d:	cmpb   $0x1,-0xb(%r12)
  139243:	jne    13a571 <litchi_pptx::shape::reader::Scene::read_with+0x2411>
  139249:	movq   $0x0,-0x80(%r12)
  139252:	mov    0x8(%rsp),%rbx
  139257:	jmp    13947a <litchi_pptx::shape::reader::Scene::read_with+0x131a>
  13925c:	mov    $0x2e,%edx
  139261:	lea    -0x114773(%rip),%rsi        # 24af5 <anon.bc3d94ef2394061926bee58178356194.284.llvm.5373085861489949100+0x4afd>
  139268:	call   *0xf3382(%rip)        # 22c5f0 <bcmp@GLIBC_2.2.5>
  13926e:	test   %eax,%eax
  139270:	jne    13906d <litchi_pptx::shape::reader::Scene::read_with+0xf0d>
  139276:	mov    $0x1,%bl
  139278:	mov    0x240(%rsp),%r15
  139280:	cmp    0x230(%rsp),%r15
  139288:	jne    139298 <litchi_pptx::shape::reader::Scene::read_with+0x1138>
  13928a:	lea    0x230(%rsp),%rdi
  139292:	call   *0xf3fb8(%rip)        # 22d250 <_DYNAMIC+0x1050>
  139298:	mov    0x238(%rsp),%rax
  1392a0:	mov    %bl,(%rax,%r15,1)
  1392a4:	inc    %r15
  1392a7:	mov    %r15,0x240(%rsp)
  1392af:	mov    %rbp,0x290(%rsp)
  1392b7:	cmpq   $0x0,0xb0(%rsp)
  1392c0:	jg     138f3f <litchi_pptx::shape::reader::Scene::read_with+0xddf>
  1392c6:	jmp    138f4d <litchi_pptx::shape::reader::Scene::read_with+0xded>
  1392cb:	cmp    $0x3,%rbx
  1392cf:	jne    13933b <litchi_pptx::shape::reader::Scene::read_with+0x11db>
  1392d1:	movzwl (%rdx),%eax
  1392d4:	xor    $0x7865,%eax
  1392d9:	movzbl 0x2(%rdx),%ecx
  1392dd:	xor    $0x74,%ecx
  1392e0:	or     %ax,%cx
  1392e3:	mov    0x8(%rsp),%rbx
  1392e8:	jne    13947a <litchi_pptx::shape::reader::Scene::read_with+0x131a>
  1392ee:	movabs $0x8000000000000001,%rax
  1392f8:	cmp    %rax,0x90(%rsp)
  139300:	jne    13947a <litchi_pptx::shape::reader::Scene::read_with+0x131a>
  139306:	mov    0x98(%rsp),%rdi
  13930e:	mov    0xa0(%rsp),%rax
  139316:	cmp    $0x2e,%rax
  13931a:	je     139655 <litchi_pptx::shape::reader::Scene::read_with+0x14f5>
  139320:	cmp    $0x3a,%rax
  139324:	jne    13947a <litchi_pptx::shape::reader::Scene::read_with+0x131a>
  13932a:	mov    $0x3a,%edx
  13932f:	lea    -0x11487b(%rip),%rsi        # 24abb <anon.bc3d94ef2394061926bee58178356194.284.llvm.5373085861489949100+0x4ac3>
  139336:	jmp    139661 <litchi_pptx::shape::reader::Scene::read_with+0x1501>
  13933b:	cmp    $0x6,%rbx
  13933f:	jne    1393ae <litchi_pptx::shape::reader::Scene::read_with+0x124e>
  139341:	mov    (%rdx),%eax
  139343:	mov    $0x4c747865,%ecx
  139348:	xor    %ecx,%eax
  13934a:	movzwl 0x4(%rdx),%ecx
  13934e:	xor    $0x7473,%ecx
  139354:	or     %eax,%ecx
  139356:	mov    0x8(%rsp),%rbx
  13935b:	jne    13947a <litchi_pptx::shape::reader::Scene::read_with+0x131a>
  139361:	movabs $0x8000000000000001,%rax
  13936b:	cmp    %rax,0x90(%rsp)
  139373:	jne    13947a <litchi_pptx::shape::reader::Scene::read_with+0x131a>
  139379:	mov    0x98(%rsp),%rdi
  139381:	mov    0xa0(%rsp),%rax
  139389:	cmp    $0x2e,%rax
  13938d:	je     1396d9 <litchi_pptx::shape::reader::Scene::read_with+0x1579>
  139393:	cmp    $0x3a,%rax
  139397:	jne    13947a <litchi_pptx::shape::reader::Scene::read_with+0x131a>
  13939d:	mov    $0x3a,%edx
  1393a2:	lea    -0x1148ee(%rip),%rsi        # 24abb <anon.bc3d94ef2394061926bee58178356194.284.llvm.5373085861489949100+0x4ac3>
  1393a9:	jmp    1396e5 <litchi_pptx::shape::reader::Scene::read_with+0x1585>
  1393ae:	cmp    $0x2,%rbx
  1393b2:	jne    13940d <litchi_pptx::shape::reader::Scene::read_with+0x12ad>
  1393b4:	cmpw   $0x6870,(%rdx)
  1393b9:	mov    0x8(%rsp),%rbx
  1393be:	jne    13947a <litchi_pptx::shape::reader::Scene::read_with+0x131a>
  1393c4:	movabs $0x8000000000000001,%rax
  1393ce:	cmp    %rax,0x90(%rsp)
  1393d6:	jne    13947a <litchi_pptx::shape::reader::Scene::read_with+0x131a>
  1393dc:	mov    0x98(%rsp),%rdi
  1393e4:	mov    0xa0(%rsp),%rax
  1393ec:	cmp    $0x2e,%rax
  1393f0:	je     139721 <litchi_pptx::shape::reader::Scene::read_with+0x15c1>
  1393f6:	cmp    $0x3a,%rax
  1393fa:	jne    13947a <litchi_pptx::shape::reader::Scene::read_with+0x131a>
  1393fc:	mov    $0x3a,%edx
  139401:	lea    -0x11494d(%rip),%rsi        # 24abb <anon.bc3d94ef2394061926bee58178356194.284.llvm.5373085861489949100+0x4ac3>
  139408:	jmp    13972d <litchi_pptx::shape::reader::Scene::read_with+0x15cd>
  13940d:	cmp    $0x1,%rbx
  139411:	mov    0x8(%rsp),%rbx
  139416:	jne    139475 <litchi_pptx::shape::reader::Scene::read_with+0x1315>
  139418:	cmpb   $0x74,(%rdx)
  13941b:	jne    139475 <litchi_pptx::shape::reader::Scene::read_with+0x1315>
  13941d:	movabs $0x8000000000000001,%rax
  139427:	cmp    %rax,0x90(%rsp)
  13942f:	jne    139475 <litchi_pptx::shape::reader::Scene::read_with+0x1315>
  139431:	mov    0x98(%rsp),%rdi
  139439:	mov    0xa0(%rsp),%rax
  139441:	cmp    $0x29,%rax
  139445:	je     13945b <litchi_pptx::shape::reader::Scene::read_with+0x12fb>
  139447:	cmp    $0x35,%rax
  13944b:	jne    139475 <litchi_pptx::shape::reader::Scene::read_with+0x1315>
  13944d:	mov    $0x35,%edx
  139452:	lea    -0x1149fc(%rip),%rsi        # 24a5d <anon.bc3d94ef2394061926bee58178356194.284.llvm.5373085861489949100+0x4a65>
  139459:	jmp    139467 <litchi_pptx::shape::reader::Scene::read_with+0x1307>
  13945b:	mov    $0x29,%edx
  139460:	lea    -0x1149d5(%rip),%rsi        # 24a92 <anon.bc3d94ef2394061926bee58178356194.284.llvm.5373085861489949100+0x4a9a>
  139467:	call   *0xf3183(%rip)        # 22c5f0 <bcmp@GLIBC_2.2.5>
  13946d:	test   %eax,%eax
  13946f:	je     139595 <litchi_pptx::shape::reader::Scene::read_with+0x1435>
  139475:	test   %rbp,%rbp
  139478:	je     1394d6 <litchi_pptx::shape::reader::Scene::read_with+0x1376>
  13947a:	mov    0x220(%rsp),%rax
  139482:	shl    $0x8,%rbp
  139486:	cmp    %r15,-0x20(%rax,%rbp,1)
  13948b:	mov    0x18(%rsp),%r12
  139490:	jne    1394db <litchi_pptx::shape::reader::Scene::read_with+0x137b>
  139492:	add    %rbp,%rax
  139495:	mov    -0x18(%rax),%edx
  139498:	lea    0x30(%rsp),%rdi
  13949d:	lea    0x1c8(%rsp),%rsi
  1394a5:	mov    %r13,%rcx
  1394a8:	call   13b1c0 <litchi_pptx::shape::reader::Scanner::finish_shape>
  1394ad:	movzbl 0x30(%rsp),%r15d
  1394b3:	cmp    $0x27,%r15b
  1394b7:	jne    139cf9 <litchi_pptx::shape::reader::Scene::read_with+0x1b99>
  1394bd:	mov    0x290(%rsp),%r15
  1394c5:	mov    0x8(%rsp),%rbx
  1394ca:	cmpl   $0x1,0x1d8(%rsp)
  1394d2:	je     1394e5 <litchi_pptx::shape::reader::Scene::read_with+0x1385>
  1394d4:	jmp    1394fb <litchi_pptx::shape::reader::Scene::read_with+0x139b>
  1394d6:	mov    0x18(%rsp),%r12
  1394db:	cmpl   $0x1,0x1d8(%rsp)
  1394e3:	jne    1394fb <litchi_pptx::shape::reader::Scene::read_with+0x139b>
  1394e5:	cmp    %r15,0x1e0(%rsp)
  1394ed:	jne    1394fb <litchi_pptx::shape::reader::Scene::read_with+0x139b>
  1394ef:	movq   $0x0,0x1d8(%rsp)
  1394fb:	cmpl   $0x1,0x1c8(%rsp)
  139503:	lea    0x30(%rsp),%rbp
  139508:	jne    139520 <litchi_pptx::shape::reader::Scene::read_with+0x13c0>
  13950a:	cmp    %r15,0x1d0(%rsp)
  139512:	jne    139520 <litchi_pptx::shape::reader::Scene::read_with+0x13c0>
  139514:	movq   $0x0,0x1c8(%rsp)
  139520:	test   %r15,%r15
  139523:	je     139c0a <litchi_pptx::shape::reader::Scene::read_with+0x1aaa>
  139529:	dec    %r15
  13952c:	mov    %r15,0x290(%rsp)
  139534:	mov    0x240(%rsp),%rax
  13953c:	test   %rax,%rax
  13953f:	je     139c8b <litchi_pptx::shape::reader::Scene::read_with+0x1b2b>
  139545:	dec    %rax
  139548:	mov    %rax,0x240(%rsp)
  139550:	mov    0xa8(%rsp),%rax
  139558:	shl    $1,%rax
  13955b:	test   %rax,%rax
  13955e:	movabs $0x8000000000000001,%r15
  139568:	je     139573 <litchi_pptx::shape::reader::Scene::read_with+0x1413>
  13956a:	mov    %r12,%rdi
  13956d:	call   *0xf2e85(%rip)        # 22c3f8 <free@GLIBC_2.2.5>
  139573:	cmpq   $0x0,0x90(%rsp)
  13957c:	jle    138680 <litchi_pptx::shape::reader::Scene::read_with+0x520>
  139582:	mov    0x98(%rsp),%rdi
  13958a:	call   *0xf2e68(%rip)        # 22c3f8 <free@GLIBC_2.2.5>
  139590:	jmp    138680 <litchi_pptx::shape::reader::Scene::read_with+0x520>
  139595:	test   %rbp,%rbp
  139598:	je     1395d7 <litchi_pptx::shape::reader::Scene::read_with+0x1477>
  13959a:	mov    0x220(%rsp),%rax
  1395a2:	mov    %rbp,%rcx
  1395a5:	shl    $0x8,%rcx
  1395a9:	cmpl   $0x1,-0xc0(%rax,%rcx,1)
  1395b1:	jne    13947a <litchi_pptx::shape::reader::Scene::read_with+0x131a>
  1395b7:	add    %rcx,%rax
  1395ba:	cmp    %r15,-0xb8(%rax)
  1395c1:	jne    13947a <litchi_pptx::shape::reader::Scene::read_with+0x131a>
  1395c7:	movq   $0x0,-0xc0(%rax)
  1395d2:	jmp    13947a <litchi_pptx::shape::reader::Scene::read_with+0x131a>
  1395d7:	mov    0x18(%rsp),%r12
  1395dc:	cmpl   $0x1,0x1d8(%rsp)
  1395e4:	je     1394e5 <litchi_pptx::shape::reader::Scene::read_with+0x1385>
  1395ea:	jmp    1394fb <litchi_pptx::shape::reader::Scene::read_with+0x139b>
  1395ef:	mov    %rbp,%rdi
  1395f2:	call   13f8c0 <litchi_pptx::shape::reader::position::{{closure}}>
  1395f7:	movzbl 0x30(%rsp),%r15d
  1395fd:	mov    0x38(%rsp),%r12
  139602:	cmp    $0x27,%r15b
  139606:	jne    13a43e <litchi_pptx::shape::reader::Scene::read_with+0x22de>
  13960c:	movabs $0x8000000000000001,%r15
  139616:	cmpb   $0x1,0x350(%rsp)
  13961e:	je     13869f <litchi_pptx::shape::reader::Scene::read_with+0x53f>
  139624:	jmp    138732 <litchi_pptx::shape::reader::Scene::read_with+0x5d2>
  139629:	mov    %rbp,%rdi
  13962c:	call   13f8c0 <litchi_pptx::shape::reader::position::{{closure}}>
  139631:	movzbl 0x30(%rsp),%r15d
  139637:	mov    0x38(%rsp),%r13
  13963c:	cmp    $0x27,%r15b
  139640:	jne    13a492 <litchi_pptx::shape::reader::Scene::read_with+0x2332>
  139646:	movabs $0x8000000000000001,%r15
  139650:	jmp    1387dd <litchi_pptx::shape::reader::Scene::read_with+0x67d>
  139655:	mov    $0x2e,%edx
  13965a:	lea    -0x114b6c(%rip),%rsi        # 24af5 <anon.bc3d94ef2394061926bee58178356194.284.llvm.5373085861489949100+0x4afd>
  139661:	call   *0xf2f89(%rip)        # 22c5f0 <bcmp@GLIBC_2.2.5>
  139667:	test   %eax,%eax
  139669:	jne    13947a <litchi_pptx::shape::reader::Scene::read_with+0x131a>
  13966f:	cmpb   $0x0,-0x90(%r12)
  139678:	je     13947a <litchi_pptx::shape::reader::Scene::read_with+0x131a>
  13967e:	cmp    %r15,-0x88(%r12)
  139686:	jne    13947a <litchi_pptx::shape::reader::Scene::read_with+0x131a>
  13968c:	cmpb   $0x1,-0x8(%r12)
  139692:	jne    13a5c7 <litchi_pptx::shape::reader::Scene::read_with+0x2467>
  139698:	mov    -0x38(%r12),%rax
  13969d:	movabs $0x8000000000000001,%rcx
  1396a7:	lea    -0x1(%rcx),%rbx
  1396ab:	cmp    %rbx,%rax
  1396ae:	jne    1397e6 <litchi_pptx::shape::reader::Scene::read_with+0x1686>
  1396b4:	movq   $0x0,-0x90(%r12)
  1396c0:	movb   $0x0,-0x8(%r12)
  1396c6:	movw   $0x0,-0x10(%r12)
  1396ce:	movb   $0x0,-0xe(%r12)
  1396d4:	jmp    139858 <litchi_pptx::shape::reader::Scene::read_with+0x16f8>
  1396d9:	mov    $0x2e,%edx
  1396de:	lea    -0x114bf0(%rip),%rsi        # 24af5 <anon.bc3d94ef2394061926bee58178356194.284.llvm.5373085861489949100+0x4afd>
  1396e5:	call   *0xf2f05(%rip)        # 22c5f0 <bcmp@GLIBC_2.2.5>
  1396eb:	test   %eax,%eax
  1396ed:	jne    13947a <litchi_pptx::shape::reader::Scene::read_with+0x131a>
  1396f3:	cmpb   $0x0,-0xa0(%r12)
  1396fc:	je     13947a <litchi_pptx::shape::reader::Scene::read_with+0x131a>
  139702:	cmp    %r15,-0x98(%r12)
  13970a:	jne    13947a <litchi_pptx::shape::reader::Scene::read_with+0x131a>
  139710:	movq   $0x0,-0xa0(%r12)
  13971c:	jmp    13947a <litchi_pptx::shape::reader::Scene::read_with+0x131a>
  139721:	mov    $0x2e,%edx
  139726:	lea    -0x114c38(%rip),%rsi        # 24af5 <anon.bc3d94ef2394061926bee58178356194.284.llvm.5373085861489949100+0x4afd>
  13972d:	call   *0xf2ebd(%rip)        # 22c5f0 <bcmp@GLIBC_2.2.5>
  139733:	test   %eax,%eax
  139735:	jne    13947a <litchi_pptx::shape::reader::Scene::read_with+0x131a>
  13973b:	cmpl   $0x1,-0xb0(%r12)
  139744:	jne    13947a <litchi_pptx::shape::reader::Scene::read_with+0x131a>
  13974a:	cmp    %r15,-0xa8(%r12)
  139752:	jne    13947a <litchi_pptx::shape::reader::Scene::read_with+0x131a>
  139758:	cmpb   $0x0,-0xd(%r12)
  13975e:	je     139784 <litchi_pptx::shape::reader::Scene::read_with+0x1624>
  139760:	cmpb   $0x0,-0xc(%r12)
  139766:	je     13a60e <litchi_pptx::shape::reader::Scene::read_with+0x24ae>
  13976c:	cmpb   $0x1,-0xb(%r12)
  139772:	jne    13a60e <litchi_pptx::shape::reader::Scene::read_with+0x24ae>
  139778:	cmpb   $0x2,-0xa(%r12)
  13977e:	je     13a60e <litchi_pptx::shape::reader::Scene::read_with+0x24ae>
  139784:	movq   $0x0,-0xb0(%r12)
  139790:	movq   $0x0,-0xa0(%r12)
  13979c:	movq   $0x0,-0x90(%r12)
  1397a8:	movb   $0x0,-0x8(%r12)
  1397ae:	movl   $0x0,-0x11(%r12)
  1397b7:	cmpq   $0x0,-0x38(%r12)
  1397bd:	jle    1397ca <litchi_pptx::shape::reader::Scene::read_with+0x166a>
  1397bf:	mov    -0x30(%r12),%rdi
  1397c4:	call   *0xf2c2e(%rip)        # 22c3f8 <free@GLIBC_2.2.5>
  1397ca:	movabs $0x8000000000000001,%rax
  1397d4:	dec    %rax
  1397d7:	mov    %rax,-0x38(%r12)
  1397dc:	mov    0x8(%rsp),%rbx
  1397e1:	jmp    13947a <litchi_pptx::shape::reader::Scene::read_with+0x131a>
  1397e6:	cmpq   $0x1e,-0x28(%r12)
  1397ec:	jne    139828 <litchi_pptx::shape::reader::Scene::read_with+0x16c8>
  1397ee:	mov    -0x30(%r12),%rcx
  1397f3:	movdqu (%rcx),%xmm0
  1397f7:	pcmpeqb -0x1236bf(%rip),%xmm0        # 16140 <anon.58baeeae7d5fe476157d296c9f08f803.346.llvm.5844774155465566772+0x1c0>
  1397ff:	movdqu 0xe(%rcx),%xmm1
  139804:	pcmpeqb -0x1230ec(%rip),%xmm1        # 16720 <anon.58baeeae7d5fe476157d296c9f08f803.346.llvm.5844774155465566772+0x7a0>
  13980c:	pand   %xmm0,%xmm1
  139810:	pmovmskb %xmm1,%ecx
  139814:	cmp    $0xffff,%ecx
  13981a:	jne    139828 <litchi_pptx::shape::reader::Scene::read_with+0x16c8>
  13981c:	cmpb   $0x0,-0xe(%r12)
  139822:	je     13a65c <litchi_pptx::shape::reader::Scene::read_with+0x24fc>
  139828:	movq   $0x0,-0x90(%r12)
  139834:	movb   $0x0,-0x8(%r12)
  13983a:	movw   $0x0,-0x10(%r12)
  139842:	movb   $0x0,-0xe(%r12)
  139848:	test   %rax,%rax
  13984b:	je     139858 <litchi_pptx::shape::reader::Scene::read_with+0x16f8>
  13984d:	mov    -0x30(%r12),%rdi
  139852:	call   *0xf2ba0(%rip)        # 22c3f8 <free@GLIBC_2.2.5>
  139858:	mov    %rbx,-0x38(%r12)
  13985d:	mov    0x8(%rsp),%rbx
  139862:	jmp    13947a <litchi_pptx::shape::reader::Scene::read_with+0x131a>
  139867:	mov    0x130(%rsp),%rcx
  13986f:	mov    %rcx,0xd8(%rsp)
  139877:	movdqa 0x110(%rsp),%xmm0
  139880:	movdqa 0x120(%rsp),%xmm1
  139889:	movdqu %xmm1,0xc8(%rsp)
  139892:	movdqu %xmm0,0xb8(%rsp)
  13989b:	mov    %rax,0xb0(%rsp)
  1398a3:	lea    0x30(%rsp),%rdi
  1398a8:	lea    0xb0(%rsp),%rsi
  1398b0:	call   *0xf396a(%rip)        # 22d220 <_DYNAMIC+0x1020>
  1398b6:	movzbl 0x30(%rsp),%r15d
  1398bc:	mov    0x31(%rsp),%eax
  1398c0:	mov    %eax,0x10(%rsp)
  1398c4:	mov    0x34(%rsp),%eax
  1398c8:	mov    %eax,0x13(%rsp)
  1398cc:	mov    0x38(%rsp),%r13
  1398d1:	mov    0x40(%rsp),%rax
  1398d6:	mov    %rax,0x8(%rsp)
  1398db:	movups 0x48(%rsp),%xmm0
  1398e0:	movaps %xmm0,0x20(%rsp)
  1398e5:	movups 0x58(%rsp),%xmm0
  1398ea:	movaps %xmm0,0xf0(%rsp)
  1398f2:	movdqu 0x68(%rsp),%xmm0
  1398f8:	movdqa %xmm0,0x100(%rsp)
  139901:	mov    %r13,%r12
  139904:	cmpq   $0x0,0x2b0(%rsp)
  13990d:	jne    13a1b9 <litchi_pptx::shape::reader::Scene::read_with+0x2059>
  139913:	jmp    13a1c7 <litchi_pptx::shape::reader::Scene::read_with+0x2067>
  139918:	mov    $0x3e,%r13d
  13991e:	mov    $0x3e,%edi
  139923:	call   *0xf2abf(%rip)        # 22c3e8 <malloc@GLIBC_2.2.5>
  139929:	test   %rax,%rax
  13992c:	je     13a6da <litchi_pptx::shape::reader::Scene::read_with+0x257a>
  139932:	mov    %rax,%r12
  139935:	movups -0x111f73(%rip),%xmm0        # 279c9 <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13130224050040044120+0x7d1>
  13993c:	movups %xmm0,0x2e(%rax)
  139940:	movups -0x111f8c(%rip),%xmm0        # 279bb <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13130224050040044120+0x7c3>
  139947:	movups %xmm0,0x20(%rax)
  13994b:	movups -0x111fa7(%rip),%xmm0        # 279ab <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13130224050040044120+0x7b3>
  139952:	movups %xmm0,0x10(%rax)
  139956:	movups -0x111fc2(%rip),%xmm0        # 2799b <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13130224050040044120+0x7a3>
  13995d:	movups %xmm0,(%rax)
  139960:	mov    $0x8,%r15b
  139963:	cmpq   $0x0,0x168(%rsp)
  13996c:	jle    139982 <litchi_pptx::shape::reader::Scene::read_with+0x1822>
  13996e:	mov    0x170(%rsp),%rdi
  139976:	call   *0xf2a7c(%rip)        # 22c3f8 <free@GLIBC_2.2.5>
  13997c:	mov    $0x3e,%r13d
  139982:	movq   %r13,%xmm0
  139987:	movdqa %xmm0,0x20(%rsp)
  13998d:	jmp    13a18d <litchi_pptx::shape::reader::Scene::read_with+0x202d>
  139992:	lea    -0x1125cf(%rip),%r12        # 273ca <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13130224050040044120+0x1d2>
  139999:	mov    %rbx,%r13
  13999c:	movdqa 0x140(%rsp),%xmm0
  1399a5:	movdqa %xmm0,0x20(%rsp)
  1399ab:	cmpq   $0x0,0xb0(%rsp)
  1399b4:	jg     13a17f <litchi_pptx::shape::reader::Scene::read_with+0x201f>
  1399ba:	jmp    13a18d <litchi_pptx::shape::reader::Scene::read_with+0x202d>
  1399bf:	mov    0x31(%rsp),%eax
  1399c3:	mov    0x34(%rsp),%ecx
  1399c7:	mov    %ecx,0x13(%rsp)
  1399cb:	mov    %eax,0x10(%rsp)
  1399cf:	mov    0x38(%rsp),%r13
  1399d4:	mov    0x40(%rsp),%r12
  1399d9:	movups 0x48(%rsp),%xmm0
  1399de:	movaps %xmm0,0x20(%rsp)
  1399e3:	movups 0x58(%rsp),%xmm0
  1399e8:	movaps %xmm0,0xf0(%rsp)
  1399f0:	movdqu 0x68(%rsp),%xmm0
  1399f6:	movdqa %xmm0,0x100(%rsp)
  1399ff:	cmpq   $0x0,0xb0(%rsp)
  139a08:	jg     13a17f <litchi_pptx::shape::reader::Scene::read_with+0x201f>
  139a0e:	jmp    13a18d <litchi_pptx::shape::reader::Scene::read_with+0x202d>
  139a13:	mov    %r13,0x18(%rsp)
  139a18:	mov    $0x3e,%r15d
  139a1e:	mov    $0x3e,%edi
  139a23:	call   *0xf29bf(%rip)        # 22c3e8 <malloc@GLIBC_2.2.5>
  139a29:	test   %rax,%rax
  139a2c:	je     13a6ba <litchi_pptx::shape::reader::Scene::read_with+0x255a>
  139a32:	mov    %rax,%r12
  139a35:	movups -0x1120e1(%rip),%xmm0        # 2795b <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13130224050040044120+0x763>
  139a3c:	movups %xmm0,0x2e(%rax)
  139a40:	movups -0x1120fa(%rip),%xmm0        # 2794d <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13130224050040044120+0x755>
  139a47:	movups %xmm0,0x20(%rax)
  139a4b:	movups -0x112115(%rip),%xmm0        # 2793d <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13130224050040044120+0x745>
  139a52:	movups %xmm0,0x10(%rax)
  139a56:	movups -0x112130(%rip),%xmm0        # 2792d <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13130224050040044120+0x735>
  139a5d:	movups %xmm0,(%rax)
  139a60:	mov    $0x3e,%r13d
  139a66:	movd   %r13d,%xmm0
  139a6b:	movdqa %xmm0,0x20(%rsp)
  139a71:	mov    $0x8,%r15b
  139a74:	test   %rbx,%rbx
  139a77:	jle    13a18d <litchi_pptx::shape::reader::Scene::read_with+0x202d>
  139a7d:	mov    0x18(%rsp),%rdi
  139a82:	jmp    13a187 <litchi_pptx::shape::reader::Scene::read_with+0x2027>
  139a87:	mov    $0x3e,%ebx
  139a8c:	mov    $0x3e,%edi
  139a91:	call   *0xf2951(%rip)        # 22c3e8 <malloc@GLIBC_2.2.5>
  139a97:	test   %rax,%rax
  139a9a:	je     13a6ca <litchi_pptx::shape::reader::Scene::read_with+0x256a>
  139aa0:	mov    %rax,%r12
  139aa3:	movups -0x11214f(%rip),%xmm0        # 2795b <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13130224050040044120+0x763>
  139aaa:	movups %xmm0,0x2e(%rax)
  139aae:	movups -0x112168(%rip),%xmm0        # 2794d <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13130224050040044120+0x755>
  139ab5:	movups %xmm0,0x20(%rax)
  139ab9:	movups -0x112183(%rip),%xmm0        # 2793d <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13130224050040044120+0x745>
  139ac0:	movups %xmm0,0x10(%rax)
  139ac4:	movups -0x11219e(%rip),%xmm0        # 2792d <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13130224050040044120+0x735>
  139acb:	movups %xmm0,(%rax)
  139ace:	mov    $0x3e,%r13d
  139ad4:	movd   %r13d,%xmm0
  139ad9:	movdqa %xmm0,0x20(%rsp)
  139adf:	mov    $0x8,%r15b
  139ae2:	cmpq   $0x0,0x110(%rsp)
  139aeb:	jle    13a18d <litchi_pptx::shape::reader::Scene::read_with+0x202d>
  139af1:	mov    0x118(%rsp),%rdi
  139af9:	jmp    13a187 <litchi_pptx::shape::reader::Scene::read_with+0x2027>
  139afe:	cmpq   $0x0,0x90(%rsp)
  139b07:	jle    139b17 <litchi_pptx::shape::reader::Scene::read_with+0x19b7>
  139b09:	mov    0x98(%rsp),%rdi
  139b11:	call   *0xf28e1(%rip)        # 22c3f8 <free@GLIBC_2.2.5>
  139b17:	cmpq   $0x0,0x290(%rsp)
  139b20:	jne    139b3c <litchi_pptx::shape::reader::Scene::read_with+0x19dc>
  139b22:	cmpq   $0x0,0x228(%rsp)
  139b2b:	jne    139b3c <litchi_pptx::shape::reader::Scene::read_with+0x19dc>
  139b2d:	cmpq   $0x0,0x240(%rsp)
  139b36:	je     13a0ec <litchi_pptx::shape::reader::Scene::read_with+0x1f8c>
  139b3c:	mov    $0x26,%r12d
  139b42:	mov    $0x26,%edi
  139b47:	call   *0xf289b(%rip)        # 22c3e8 <malloc@GLIBC_2.2.5>
  139b4d:	test   %rax,%rax
  139b50:	je     13a6ec <litchi_pptx::shape::reader::Scene::read_with+0x258c>
  139b56:	movups -0x112174(%rip),%xmm0        # 279e9 <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13130224050040044120+0x7f1>
  139b5d:	movups %xmm0,0x10(%rax)
  139b61:	movups -0x11218f(%rip),%xmm0        # 279d9 <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13130224050040044120+0x7e1>
  139b68:	movups %xmm0,(%rax)
  139b6b:	movabs $0x73746e656d656c65,%rcx
  139b75:	mov    %rax,0x8(%rsp)
  139b7a:	mov    %rcx,0x1e(%rax)
  139b7e:	movq   %r12,%xmm0
  139b83:	movdqa %xmm0,0x20(%rsp)
  139b89:	mov    $0x8,%r15b
  139b8c:	cmpq   $0x0,0x2b0(%rsp)
  139b95:	jne    13a1b9 <litchi_pptx::shape::reader::Scene::read_with+0x2059>
  139b9b:	jmp    13a1c7 <litchi_pptx::shape::reader::Scene::read_with+0x2067>
  139ba0:	mov    $0x27,%ebx
  139ba5:	mov    $0x27,%edi
  139baa:	call   *0xf2838(%rip)        # 22c3e8 <malloc@GLIBC_2.2.5>
  139bb0:	test   %rax,%rax
  139bb3:	je     13a6aa <litchi_pptx::shape::reader::Scene::read_with+0x254a>
  139bb9:	mov    %rax,%r12
  139bbc:	movups -0x112848(%rip),%xmm0        # 2737b <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13130224050040044120+0x183>
  139bc3:	movups %xmm0,0x10(%rax)
  139bc7:	movups -0x112863(%rip),%xmm0        # 2736b <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13130224050040044120+0x173>
  139bce:	movups %xmm0,(%rax)
  139bd1:	movabs $0x67617420646e6520,%rax
  139bdb:	mov    %rax,0x1f(%r12)
  139be0:	mov    $0x27,%r13d
  139be6:	jmp    139c41 <litchi_pptx::shape::reader::Scene::read_with+0x1ae1>
  139be8:	mov    %r12,0x18(%rsp)
  139bed:	mov    %r13,0x140(%rsp)
  139bf5:	mov    0x31(%rsp),%eax
  139bf9:	mov    0x34(%rsp),%ecx
  139bfd:	mov    %ecx,0x13(%rsp)
  139c01:	mov    %eax,0x10(%rsp)
  139c05:	jmp    139d89 <litchi_pptx::shape::reader::Scene::read_with+0x1c29>
  139c0a:	mov    $0x19,%ebx
  139c0f:	mov    $0x19,%edi
  139c14:	call   *0xf27ce(%rip)        # 22c3e8 <malloc@GLIBC_2.2.5>
  139c1a:	test   %rax,%rax
  139c1d:	je     13a6aa <litchi_pptx::shape::reader::Scene::read_with+0x254a>
  139c23:	mov    %rax,%r12
  139c26:	movups -0x112873(%rip),%xmm0        # 273ba <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13130224050040044120+0x1c2>
  139c2d:	movups %xmm0,0x9(%rax)
  139c31:	movups -0x112887(%rip),%xmm0        # 273b1 <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13130224050040044120+0x1b9>
  139c38:	movups %xmm0,(%rax)
  139c3b:	mov    $0x19,%r13d
  139c41:	movq   %r13,%xmm0
  139c46:	movdqa %xmm0,0x20(%rsp)
  139c4c:	mov    $0x8,%r15b
  139c4f:	mov    0x110(%rsp),%eax
  139c56:	mov    0x113(%rsp),%ecx
  139c5d:	mov    %ecx,0x13(%rsp)
  139c61:	mov    %eax,0x10(%rsp)
  139c65:	movdqa 0xb0(%rsp),%xmm0
  139c6e:	movdqa 0xc0(%rsp),%xmm1
  139c77:	movdqa %xmm0,0xf0(%rsp)
  139c80:	movdqa %xmm1,0x100(%rsp)
  139c89:	jmp    139cdb <litchi_pptx::shape::reader::Scene::read_with+0x1b7b>
  139c8b:	mov    $0x2b,%ebx
  139c90:	mov    $0x2b,%edi
  139c95:	call   *0xf274d(%rip)        # 22c3e8 <malloc@GLIBC_2.2.5>
  139c9b:	test   %rax,%rax
  139c9e:	je     13a6aa <litchi_pptx::shape::reader::Scene::read_with+0x254a>
  139ca4:	mov    %rax,%r12
  139ca7:	movups -0x112294(%rip),%xmm0        # 27a1a <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13130224050040044120+0x822>
  139cae:	movups %xmm0,0x1b(%rax)
  139cb2:	movups -0x1122aa(%rip),%xmm0        # 27a0f <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13130224050040044120+0x817>
  139cb9:	movups %xmm0,0x10(%rax)
  139cbd:	movups -0x1122c5(%rip),%xmm0        # 279ff <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13130224050040044120+0x807>
  139cc4:	movups %xmm0,(%rax)
  139cc7:	mov    $0x2b,%r13d
  139ccd:	movq   %r13,%xmm0
  139cd2:	movdqa %xmm0,0x20(%rsp)
  139cd8:	mov    $0x8,%r15b
  139cdb:	mov    0xa8(%rsp),%rax
  139ce3:	shl    $1,%rax
  139ce6:	test   %rax,%rax
  139ce9:	je     13a18d <litchi_pptx::shape::reader::Scene::read_with+0x202d>
  139cef:	mov    0x18(%rsp),%rdi
  139cf4:	jmp    13a187 <litchi_pptx::shape::reader::Scene::read_with+0x2027>
  139cf9:	mov    0x31(%rsp),%eax
  139cfd:	mov    0x34(%rsp),%ecx
  139d01:	mov    %ecx,0x113(%rsp)
  139d08:	mov    %eax,0x110(%rsp)
  139d0f:	mov    0x38(%rsp),%r13
  139d14:	mov    0x40(%rsp),%r12
  139d19:	movups 0x48(%rsp),%xmm0
  139d1e:	movaps %xmm0,0x20(%rsp)
  139d23:	movups 0x58(%rsp),%xmm0
  139d28:	movaps %xmm0,0xb0(%rsp)
  139d30:	movdqu 0x68(%rsp),%xmm0
  139d36:	movdqa %xmm0,0xc0(%rsp)
  139d3f:	jmp    139c4f <litchi_pptx::shape::reader::Scene::read_with+0x1aef>
  139d44:	mov    %r13,0x140(%rsp)
  139d4c:	lea    0xb8(%rsp),%rax
  139d54:	movdqu (%rax),%xmm0
  139d58:	movdqa %xmm0,0x110(%rsp)
  139d61:	lea    0x30(%rsp),%rdi
  139d66:	lea    0x110(%rsp),%rsi
  139d6e:	call   13f4d0 <litchi_pptx::shape::reader::Scanner::scan::{{closure}}>
  139d73:	movzbl 0x30(%rsp),%r15d
  139d79:	mov    0x31(%rsp),%eax
  139d7d:	mov    %eax,0x10(%rsp)
  139d81:	mov    0x34(%rsp),%eax
  139d85:	mov    %eax,0x13(%rsp)
  139d89:	mov    0x38(%rsp),%r13
  139d8e:	mov    0x40(%rsp),%r12
  139d93:	movups 0x48(%rsp),%xmm0
  139d98:	movaps %xmm0,0x20(%rsp)
  139d9d:	movups 0x58(%rsp),%xmm0
  139da2:	movaps %xmm0,0xf0(%rsp)
  139daa:	movdqu 0x68(%rsp),%xmm0
  139db0:	movdqa %xmm0,0x100(%rsp)
  139db9:	jmp    139ecd <litchi_pptx::shape::reader::Scene::read_with+0x1d6d>
  139dbe:	mov    %r13,0x140(%rsp)
  139dc6:	mov    0x130(%rsp),%rax
  139dce:	mov    %rax,0xd0(%rsp)
  139dd6:	movdqu 0x110(%rsp),%xmm0
  139ddf:	movdqu 0x120(%rsp),%xmm1
  139de8:	movdqa %xmm1,0xc0(%rsp)
  139df1:	movdqa %xmm0,0xb0(%rsp)
  139dfa:	lea    0x30(%rsp),%rdi
  139dff:	lea    0xb0(%rsp),%rsi
  139e07:	call   13f590 <litchi_pptx::shape::reader::Scanner::scan::{{closure}}>
  139e0c:	movzbl 0x30(%rsp),%r15d
  139e12:	mov    0x31(%rsp),%eax
  139e16:	mov    %eax,0x10(%rsp)
  139e1a:	mov    0x34(%rsp),%eax
  139e1e:	mov    %eax,0x13(%rsp)
  139e22:	mov    0x38(%rsp),%r13
  139e27:	mov    0x40(%rsp),%r12
  139e2c:	movups 0x48(%rsp),%xmm0
  139e31:	movaps %xmm0,0x20(%rsp)
  139e36:	movups 0x58(%rsp),%xmm0
  139e3b:	movaps %xmm0,0xf0(%rsp)
  139e43:	movdqu 0x68(%rsp),%xmm0
  139e49:	movdqa %xmm0,0x100(%rsp)
  139e52:	jmp    139eb7 <litchi_pptx::shape::reader::Scene::read_with+0x1d57>
  139e54:	mov    %r13,0x140(%rsp)
  139e5c:	mov    0x31(%rsp),%eax
  139e60:	mov    0x34(%rsp),%ecx
  139e64:	mov    %ecx,0x13(%rsp)
  139e68:	mov    %eax,0x10(%rsp)
  139e6c:	mov    0x38(%rsp),%r13
  139e71:	mov    0x40(%rsp),%rax
  139e76:	mov    %rax,0x8(%rsp)
  139e7b:	movups 0x48(%rsp),%xmm0
  139e80:	movaps %xmm0,0x20(%rsp)
  139e85:	movups 0x58(%rsp),%xmm0
  139e8a:	movaps %xmm0,0xf0(%rsp)
  139e92:	movdqu 0x68(%rsp),%xmm0
  139e98:	movdqa %xmm0,0x100(%rsp)
  139ea1:	shl    $1,%r12
  139ea4:	test   %r12,%r12
  139ea7:	je     139eb2 <litchi_pptx::shape::reader::Scene::read_with+0x1d52>
  139ea9:	mov    %rbp,%rdi
  139eac:	call   *0xf2546(%rip)        # 22c3f8 <free@GLIBC_2.2.5>
  139eb2:	mov    0x8(%rsp),%r12
  139eb7:	shl    $1,%rbx
  139eba:	test   %rbx,%rbx
  139ebd:	je     139ecd <litchi_pptx::shape::reader::Scene::read_with+0x1d6d>
  139ebf:	mov    0xa8(%rsp),%rdi
  139ec7:	call   *0xf252b(%rip)        # 22c3f8 <free@GLIBC_2.2.5>
  139ecd:	cmpq   $0x0,0x18(%rsp)
  139ed3:	jle    13a18d <litchi_pptx::shape::reader::Scene::read_with+0x202d>
  139ed9:	mov    0x140(%rsp),%rdi
  139ee1:	jmp    13a187 <litchi_pptx::shape::reader::Scene::read_with+0x2027>
  139ee6:	mov    %r13,0x18(%rsp)
  139eeb:	lea    0xb8(%rsp),%rax
  139ef3:	movdqu (%rax),%xmm0
  139ef7:	movdqa %xmm0,0x110(%rsp)
  139f00:	lea    0x30(%rsp),%rdi
  139f05:	lea    0x110(%rsp),%rsi
  139f0d:	call   13f4d0 <litchi_pptx::shape::reader::Scanner::scan::{{closure}}>
  139f12:	movzbl 0x30(%rsp),%r15d
  139f18:	mov    0x31(%rsp),%eax
  139f1c:	mov    %eax,0x10(%rsp)
  139f20:	mov    0x34(%rsp),%eax
  139f24:	mov    %eax,0x13(%rsp)
  139f28:	mov    0x38(%rsp),%r13
  139f2d:	mov    0x40(%rsp),%r12
  139f32:	movups 0x48(%rsp),%xmm0
  139f37:	movaps %xmm0,0x20(%rsp)
  139f3c:	movups 0x58(%rsp),%xmm0
  139f41:	movaps %xmm0,0xf0(%rsp)
  139f49:	movdqu 0x68(%rsp),%xmm0
  139f4f:	movdqa %xmm0,0x100(%rsp)
  139f58:	jmp    139a74 <litchi_pptx::shape::reader::Scene::read_with+0x1914>
  139f5d:	movq   %rcx,%xmm0
  139f62:	movq   %r12,%xmm1
  139f67:	punpcklqdq %xmm0,%xmm1
  139f6b:	movdqa %xmm1,0x20(%rsp)
  139f71:	mov    $0x23,%r15b
  139f74:	mov    %rbx,%r12
  139f77:	jmp    139ae2 <litchi_pptx::shape::reader::Scene::read_with+0x1982>
  139f7c:	mov    0x31(%rsp),%eax
  139f80:	mov    0x34(%rsp),%ecx
  139f84:	mov    %ecx,0x13(%rsp)
  139f88:	mov    %eax,0x10(%rsp)
  139f8c:	mov    0x38(%rsp),%r13
  139f91:	mov    0x40(%rsp),%rbp
  139f96:	movups 0x48(%rsp),%xmm0
  139f9b:	movaps %xmm0,0x20(%rsp)
  139fa0:	movups 0x58(%rsp),%xmm0
  139fa5:	movaps %xmm0,0xf0(%rsp)
  139fad:	movdqu 0x68(%rsp),%xmm0
  139fb3:	movdqa %xmm0,0x100(%rsp)
  139fbc:	test   %rbx,%rbx
  139fbf:	je     13a0cf <litchi_pptx::shape::reader::Scene::read_with+0x1f6f>
  139fc5:	mov    %r12,%rdi
  139fc8:	call   *0xf242a(%rip)        # 22c3f8 <free@GLIBC_2.2.5>
  139fce:	mov    %rbp,%r12
  139fd1:	jmp    139ae2 <litchi_pptx::shape::reader::Scene::read_with+0x1982>
  139fd6:	mov    %r13,0x18(%rsp)
  139fdb:	mov    0x31(%rsp),%eax
  139fdf:	mov    0x34(%rsp),%ecx
  139fe3:	mov    %ecx,0x13(%rsp)
  139fe7:	mov    %eax,0x10(%rsp)
  139feb:	mov    0x38(%rsp),%r13
  139ff0:	mov    0x40(%rsp),%rax
  139ff5:	mov    %rax,0x8(%rsp)
  139ffa:	movups 0x48(%rsp),%xmm0
  139fff:	movaps %xmm0,0x20(%rsp)
  13a004:	movups 0x58(%rsp),%xmm0
  13a009:	movaps %xmm0,0xf0(%rsp)
  13a011:	movdqu 0x68(%rsp),%xmm0
  13a017:	movdqa %xmm0,0x100(%rsp)
  13a020:	shl    $1,%r12
  13a023:	test   %r12,%r12
  13a026:	je     13a031 <litchi_pptx::shape::reader::Scene::read_with+0x1ed1>
  13a028:	mov    %rbp,%rdi
  13a02b:	call   *0xf23c7(%rip)        # 22c3f8 <free@GLIBC_2.2.5>
  13a031:	mov    0x8(%rsp),%r12
  13a036:	jmp    139a74 <litchi_pptx::shape::reader::Scene::read_with+0x1914>
  13a03b:	mov    %r13,0x18(%rsp)
  13a040:	mov    $0x30,%r15d
  13a046:	mov    $0x30,%edi
  13a04b:	call   *0xf2397(%rip)        # 22c3e8 <malloc@GLIBC_2.2.5>
  13a051:	test   %rax,%rax
  13a054:	je     13a6ba <litchi_pptx::shape::reader::Scene::read_with+0x255a>
  13a05a:	mov    %rax,%r12
  13a05d:	movups -0x1126d9(%rip),%xmm0        # 2798b <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13130224050040044120+0x793>
  13a064:	movups %xmm0,0x20(%rax)
  13a068:	movups -0x1126f4(%rip),%xmm0        # 2797b <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13130224050040044120+0x783>
  13a06f:	movups %xmm0,0x10(%rax)
  13a073:	movups -0x11270f(%rip),%xmm0        # 2796b <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13130224050040044120+0x773>
  13a07a:	movups %xmm0,(%rax)
  13a07d:	mov    $0x30,%r13d
  13a083:	jmp    139a66 <litchi_pptx::shape::reader::Scene::read_with+0x1906>
  13a088:	mov    $0x30,%ebx
  13a08d:	mov    $0x30,%edi
  13a092:	call   *0xf2350(%rip)        # 22c3e8 <malloc@GLIBC_2.2.5>
  13a098:	test   %rax,%rax
  13a09b:	je     13a6ca <litchi_pptx::shape::reader::Scene::read_with+0x256a>
  13a0a1:	mov    %rax,%r12
  13a0a4:	movups -0x112720(%rip),%xmm0        # 2798b <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13130224050040044120+0x793>
  13a0ab:	movups %xmm0,0x20(%rax)
  13a0af:	movups -0x11273b(%rip),%xmm0        # 2797b <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13130224050040044120+0x783>
  13a0b6:	movups %xmm0,0x10(%rax)
  13a0ba:	movups -0x112756(%rip),%xmm0        # 2796b <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13130224050040044120+0x773>
  13a0c1:	movups %xmm0,(%rax)
  13a0c4:	mov    $0x30,%r13d
  13a0ca:	jmp    139ad4 <litchi_pptx::shape::reader::Scene::read_with+0x1974>
  13a0cf:	mov    %rbp,%r12
  13a0d2:	jmp    139ae2 <litchi_pptx::shape::reader::Scene::read_with+0x1982>
  13a0d7:	call   *0xf282b(%rip)        # 22c908 <_DYNAMIC+0x708>
  13a0dd:	mov    %rdx,%fs:0x8(%rbp)
  13a0e2:	movb   $0x1,%fs:0x10(%rbp)
  13a0e7:	jmp    13824f <litchi_pptx::shape::reader::Scene::read_with+0xef>
  13a0ec:	lea    0x218(%rsp),%rbx
  13a0f4:	mov    0x1e8(%rsp),%r12
  13a0fc:	mov    0x1f0(%rsp),%rax
  13a104:	mov    %rax,0x8(%rsp)
  13a109:	movups 0x1f8(%rsp),%xmm0
  13a111:	movaps %xmm0,0x20(%rsp)
  13a116:	lea    0x208(%rsp),%rax
  13a11e:	movdqu (%rax),%xmm0
  13a122:	movdqa %xmm0,0xf0(%rsp)
  13a12b:	lea    0x2b0(%rsp),%rdi
  13a133:	call   164800 <core::ptr::drop_in_place<quick_xml::reader::ns_reader::NsReader<&[u8]>>>
  13a138:	mov    %rbx,%rdi
  13a13b:	call   1645d0 <core::ptr::drop_in_place<alloc::vec::Vec<litchi_pptx::shape::reader::Active>>>
  13a140:	mov    $0x27,%r15b
  13a143:	cmpq   $0x0,0x230(%rsp)
  13a14c:	jne    13a2bf <litchi_pptx::shape::reader::Scene::read_with+0x215f>
  13a152:	jmp    13a2d7 <litchi_pptx::shape::reader::Scene::read_with+0x2177>
  13a157:	mov    $0x12,%eax
  13a15c:	movq   %rax,%xmm0
  13a161:	movdqa %xmm0,0x20(%rsp)
  13a167:	mov    $0x9,%r15b
  13a16a:	mov    %rbx,%r13
  13a16d:	lea    -0x112ee4(%rip),%r12        # 27290 <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13130224050040044120+0x98>
  13a174:	cmpq   $0x0,0xb0(%rsp)
  13a17d:	jle    13a18d <litchi_pptx::shape::reader::Scene::read_with+0x202d>
  13a17f:	mov    0xb8(%rsp),%rdi
  13a187:	call   *0xf226b(%rip)        # 22c3f8 <free@GLIBC_2.2.5>
  13a18d:	cmpq   $0x0,0x90(%rsp)
  13a196:	jle    13a1a6 <litchi_pptx::shape::reader::Scene::read_with+0x2046>
  13a198:	mov    0x98(%rsp),%rdi
  13a1a0:	call   *0xf2252(%rip)        # 22c3f8 <free@GLIBC_2.2.5>
  13a1a6:	mov    %r12,0x8(%rsp)
  13a1ab:	mov    %r13,%r12
  13a1ae:	cmpq   $0x0,0x2b0(%rsp)
  13a1b7:	je     13a1c7 <litchi_pptx::shape::reader::Scene::read_with+0x2067>
  13a1b9:	mov    0x2b8(%rsp),%rdi
  13a1c1:	call   *0xf2231(%rip)        # 22c3f8 <free@GLIBC_2.2.5>
  13a1c7:	cmpq   $0x0,0x2c8(%rsp)
  13a1d0:	je     13a1e0 <litchi_pptx::shape::reader::Scene::read_with+0x2080>
  13a1d2:	mov    0x2d0(%rsp),%rdi
  13a1da:	call   *0xf2218(%rip)        # 22c3f8 <free@GLIBC_2.2.5>
  13a1e0:	cmpq   $0x0,0x310(%rsp)
  13a1e9:	je     13a1f9 <litchi_pptx::shape::reader::Scene::read_with+0x2099>
  13a1eb:	mov    0x318(%rsp),%rdi
  13a1f3:	call   *0xf21ff(%rip)        # 22c3f8 <free@GLIBC_2.2.5>
  13a1f9:	cmpq   $0x0,0x328(%rsp)
  13a202:	je     13a212 <litchi_pptx::shape::reader::Scene::read_with+0x20b2>
  13a204:	mov    0x330(%rsp),%rdi
  13a20c:	call   *0xf21e6(%rip)        # 22c3f8 <free@GLIBC_2.2.5>
  13a212:	cmpq   $0x0,0x1e8(%rsp)
  13a21b:	je     13a22b <litchi_pptx::shape::reader::Scene::read_with+0x20cb>
  13a21d:	mov    0x1f0(%rsp),%rdi
  13a225:	call   *0xf21cd(%rip)        # 22c3f8 <free@GLIBC_2.2.5>
  13a22b:	cmpq   $0x0,0x200(%rsp)
  13a234:	je     13a244 <litchi_pptx::shape::reader::Scene::read_with+0x20e4>
  13a236:	mov    0x208(%rsp),%rdi
  13a23e:	call   *0xf21b4(%rip)        # 22c3f8 <free@GLIBC_2.2.5>
  13a244:	mov    0x220(%rsp),%rbx
  13a24c:	mov    0x228(%rsp),%r13
  13a254:	test   %r13,%r13
  13a257:	je     13a2a0 <litchi_pptx::shape::reader::Scene::read_with+0x2140>
  13a259:	lea    0xd0(%rbx),%rbp
  13a260:	jmp    13a27c <litchi_pptx::shape::reader::Scene::read_with+0x211c>
  13a262:	data16 data16 data16 data16 cs nopw 0x0(%rax,%rax,1)
  13a270:	add    $0x100,%rbp
  13a277:	dec    %r13
  13a27a:	je     13a2a0 <litchi_pptx::shape::reader::Scene::read_with+0x2140>
  13a27c:	cmpq   $0x0,-0x20(%rbp)
  13a281:	jle    13a28d <litchi_pptx::shape::reader::Scene::read_with+0x212d>
  13a283:	mov    -0x18(%rbp),%rdi
  13a287:	call   *0xf216b(%rip)        # 22c3f8 <free@GLIBC_2.2.5>
  13a28d:	cmpq   $0x0,-0x8(%rbp)
  13a292:	jle    13a270 <litchi_pptx::shape::reader::Scene::read_with+0x2110>
  13a294:	mov    0x0(%rbp),%rdi
  13a298:	call   *0xf215a(%rip)        # 22c3f8 <free@GLIBC_2.2.5>
  13a29e:	jmp    13a270 <litchi_pptx::shape::reader::Scene::read_with+0x2110>
  13a2a0:	cmpq   $0x0,0x218(%rsp)
  13a2a9:	je     13a2b4 <litchi_pptx::shape::reader::Scene::read_with+0x2154>
  13a2ab:	mov    %rbx,%rdi
  13a2ae:	call   *0xf2144(%rip)        # 22c3f8 <free@GLIBC_2.2.5>
  13a2b4:	cmpq   $0x0,0x230(%rsp)
  13a2bd:	je     13a2cd <litchi_pptx::shape::reader::Scene::read_with+0x216d>
  13a2bf:	mov    0x238(%rsp),%rdi
  13a2c7:	call   *0xf212b(%rip)        # 22c3f8 <free@GLIBC_2.2.5>
  13a2cd:	cmp    $0x27,%r15b
  13a2d1:	jne    13a36a <litchi_pptx::shape::reader::Scene::read_with+0x220a>
  13a2d7:	mov    0x158(%rsp),%rax
  13a2df:	movaps 0xf0(%rsp),%xmm0
  13a2e7:	movaps %xmm0,0x360(%rsp)
  13a2ef:	movups %xmm0,0x20(%r14)
  13a2f4:	mov    %r12,(%r14)
  13a2f7:	mov    0x8(%rsp),%rcx
  13a2fc:	mov    %rcx,0x8(%r14)
  13a300:	movaps 0x20(%rsp),%xmm0
  13a305:	movups %xmm0,0x10(%r14)
  13a30a:	mov    %rax,0x30(%r14)
  13a30e:	mov    0x190(%rsp),%rax
  13a316:	mov    %rax,0x38(%r14)
  13a31a:	mov    0x2a8(%rsp),%rax
  13a322:	mov    %rax,0x40(%r14)
  13a326:	mov    0x198(%rsp),%rax
  13a32e:	movups (%rax),%xmm0
  13a331:	movups 0x10(%rax),%xmm1
  13a335:	movups 0x20(%rax),%xmm2
  13a339:	movups %xmm0,0x48(%r14)
  13a33e:	movups %xmm1,0x58(%r14)
  13a343:	movups %xmm2,0x68(%r14)
  13a348:	lea    0x370(%rsp),%rdi
  13a350:	call   163da0 <core::ptr::drop_in_place<litchi_ooxml_common::mce::model::Capabilities>>
  13a355:	mov    %r14,%rax
  13a358:	add    $0x408,%rsp
  13a35f:	pop    %rbx
  13a360:	pop    %r12
  13a362:	pop    %r13
  13a364:	pop    %r14
  13a366:	pop    %r15
  13a368:	pop    %rbp
  13a369:	ret
  13a36a:	mov    0x10(%rsp),%eax
  13a36e:	mov    0x13(%rsp),%ecx
  13a372:	mov    %ecx,0xc(%r14)
  13a376:	mov    %eax,0x9(%r14)
  13a37a:	movaps 0xf0(%rsp),%xmm0
  13a382:	movaps 0x100(%rsp),%xmm1
  13a38a:	movaps %xmm0,0x360(%rsp)
  13a392:	movups %xmm1,0x40(%r14)
  13a397:	movaps 0x360(%rsp),%xmm0
  13a39f:	movups %xmm0,0x30(%r14)
  13a3a4:	mov    %r15b,0x8(%r14)
  13a3a8:	mov    %r12,0x10(%r14)
  13a3ac:	mov    0x8(%rsp),%rax
  13a3b1:	mov    %rax,0x18(%r14)
  13a3b5:	movaps 0x20(%rsp),%xmm0
  13a3ba:	movups %xmm0,0x20(%r14)
  13a3bf:	movabs $0x8000000000000001,%rax
  13a3c9:	dec    %rax
  13a3cc:	mov    %rax,(%r14)
  13a3cf:	mov    0x190(%rsp),%rbx
  13a3d7:	mov    0x158(%rsp),%rcx
  13a3df:	shl    $1,%rcx
  13a3e2:	test   %rcx,%rcx
  13a3e5:	je     138408 <litchi_pptx::shape::reader::Scene::read_with+0x2a8>
  13a3eb:	mov    %rbx,%rdi
  13a3ee:	call   *0xf2004(%rip)        # 22c3f8 <free@GLIBC_2.2.5>
  13a3f4:	jmp    138408 <litchi_pptx::shape::reader::Scene::read_with+0x2a8>
  13a3f9:	mov    $0x17,%eax
  13a3fe:	movq   %rax,%xmm0
  13a403:	movdqa %xmm0,0x20(%rsp)
  13a409:	lea    -0x113046(%rip),%r12        # 273ca <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13130224050040044120+0x1d2>
  13a410:	mov    %rbx,%r13
  13a413:	cmpq   $0x0,0xb0(%rsp)
  13a41c:	jg     13a17f <litchi_pptx::shape::reader::Scene::read_with+0x201f>
  13a422:	jmp    13a18d <litchi_pptx::shape::reader::Scene::read_with+0x202d>
  13a427:	lea    0xeb65a(%rip),%rcx        # 225a88 <anon.bc3d94ef2394061926bee58178356194.1849.llvm.5373085861489949100+0x580>
  13a42e:	xor    %edi,%edi
  13a430:	mov    %r12,%rsi
  13a433:	call   *0xf2147(%rip)        # 22c580 <_DYNAMIC+0x380>
  13a439:	jmp    13a6fc <litchi_pptx::shape::reader::Scene::read_with+0x259c>
  13a43e:	mov    0x31(%rsp),%eax
  13a442:	mov    0x34(%rsp),%ecx
  13a446:	mov    %ecx,0x13(%rsp)
  13a44a:	mov    %eax,0x10(%rsp)
  13a44e:	mov    0x40(%rsp),%rax
  13a453:	mov    %rax,0x8(%rsp)
  13a458:	movups 0x48(%rsp),%xmm0
  13a45d:	movaps %xmm0,0x20(%rsp)
  13a462:	movups 0x58(%rsp),%xmm0
  13a467:	movaps %xmm0,0xf0(%rsp)
  13a46f:	movdqu 0x68(%rsp),%xmm0
  13a475:	movdqa %xmm0,0x100(%rsp)
  13a47e:	cmpq   $0x0,0x2b0(%rsp)
  13a487:	jne    13a1b9 <litchi_pptx::shape::reader::Scene::read_with+0x2059>
  13a48d:	jmp    13a1c7 <litchi_pptx::shape::reader::Scene::read_with+0x2067>
  13a492:	mov    0x31(%rsp),%eax
  13a496:	mov    0x34(%rsp),%ecx
  13a49a:	mov    %ecx,0x13(%rsp)
  13a49e:	mov    %eax,0x10(%rsp)
  13a4a2:	mov    0x40(%rsp),%rax
  13a4a7:	mov    %rax,0x8(%rsp)
  13a4ac:	movups 0x48(%rsp),%xmm0
  13a4b1:	movaps %xmm0,0x20(%rsp)
  13a4b6:	movups 0x58(%rsp),%xmm0
  13a4bb:	movaps %xmm0,0xf0(%rsp)
  13a4c3:	movdqu 0x68(%rsp),%xmm0
  13a4c9:	movdqa %xmm0,0x100(%rsp)
  13a4d2:	mov    0x1a0(%rsp),%rax
  13a4da:	cmp    $0x9,%rax
  13a4de:	ja     13a1ab <litchi_pptx::shape::reader::Scene::read_with+0x204b>
  13a4e4:	lea    -0x116337(%rip),%rcx        # 241b4 <anon.bc3d94ef2394061926bee58178356194.284.llvm.5373085861489949100+0x41bc>
  13a4eb:	movslq (%rcx,%rax,4),%rax
  13a4ef:	add    %rcx,%rax
  13a4f2:	jmp    *%rax
  13a4f4:	cmpq   $0x0,0x1a8(%rsp)
  13a4fd:	jle    13a1ab <litchi_pptx::shape::reader::Scene::read_with+0x204b>
  13a503:	mov    0x1b0(%rsp),%rdi
  13a50b:	call   *0xf1ee7(%rip)        # 22c3f8 <free@GLIBC_2.2.5>
  13a511:	mov    %r13,%r12
  13a514:	cmpq   $0x0,0x2b0(%rsp)
  13a51d:	jne    13a1b9 <litchi_pptx::shape::reader::Scene::read_with+0x2059>
  13a523:	jmp    13a1c7 <litchi_pptx::shape::reader::Scene::read_with+0x2067>
  13a528:	mov    $0x2c,%ebx
  13a52d:	mov    $0x2c,%edi
  13a532:	call   *0xf1eb0(%rip)        # 22c3e8 <malloc@GLIBC_2.2.5>
  13a538:	test   %rax,%rax
  13a53b:	je     13a6aa <litchi_pptx::shape::reader::Scene::read_with+0x254a>
  13a541:	mov    %rax,%r12
  13a544:	movups -0x112da2(%rip),%xmm0        # 277a9 <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13130224050040044120+0x5b1>
  13a54b:	movups %xmm0,0x1c(%rax)
  13a54f:	movups -0x112db9(%rip),%xmm0        # 2779d <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13130224050040044120+0x5a5>
  13a556:	movups %xmm0,0x10(%rax)
  13a55a:	movups -0x112dd4(%rip),%xmm0        # 2778d <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13130224050040044120+0x595>
  13a561:	movups %xmm0,(%r12)
  13a566:	mov    $0x2c,%r13d
  13a56c:	jmp    139c41 <litchi_pptx::shape::reader::Scene::read_with+0x1ae1>
  13a571:	mov    $0x37,%ebx
  13a576:	mov    $0x37,%edi
  13a57b:	call   *0xf1e67(%rip)        # 22c3e8 <malloc@GLIBC_2.2.5>
  13a581:	test   %rax,%rax
  13a584:	je     13a6aa <litchi_pptx::shape::reader::Scene::read_with+0x254a>
  13a58a:	mov    %rax,%r12
  13a58d:	movups -0x112c7e(%rip),%xmm0        # 27916 <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13130224050040044120+0x71e>
  13a594:	movups %xmm0,0x20(%rax)
  13a598:	movups -0x112c99(%rip),%xmm0        # 27906 <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13130224050040044120+0x70e>
  13a59f:	movups %xmm0,0x10(%rax)
  13a5a3:	movups -0x112cb4(%rip),%xmm0        # 278f6 <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13130224050040044120+0x6fe>
  13a5aa:	movups %xmm0,(%rax)
  13a5ad:	movabs $0x617461646174656d,%rax
  13a5b7:	mov    %rax,0x2f(%r12)
  13a5bc:	mov    $0x37,%r13d
  13a5c2:	jmp    139c41 <litchi_pptx::shape::reader::Scene::read_with+0x1ae1>
  13a5c7:	mov    $0x2e,%ebx
  13a5cc:	mov    $0x2e,%edi
  13a5d1:	call   *0xf1e11(%rip)        # 22c3e8 <malloc@GLIBC_2.2.5>
  13a5d7:	test   %rax,%rax
  13a5da:	je     13a6aa <litchi_pptx::shape::reader::Scene::read_with+0x254a>
  13a5e0:	mov    %rax,%r12
  13a5e3:	movups -0x112dc0(%rip),%xmm0        # 2782a <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13130224050040044120+0x632>
  13a5ea:	movups %xmm0,0x1e(%rax)
  13a5ee:	movups -0x112dd9(%rip),%xmm0        # 2781c <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13130224050040044120+0x624>
  13a5f5:	movups %xmm0,0x10(%rax)
  13a5f9:	movups -0x112df4(%rip),%xmm0        # 2780c <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13130224050040044120+0x614>
  13a600:	movups %xmm0,(%rax)
  13a603:	mov    $0x2e,%r13d
  13a609:	jmp    139c41 <litchi_pptx::shape::reader::Scene::read_with+0x1ae1>
  13a60e:	mov    $0x2c,%ebx
  13a613:	mov    $0x2c,%edi
  13a618:	call   *0xf1dca(%rip)        # 22c3e8 <malloc@GLIBC_2.2.5>
  13a61e:	test   %rax,%rax
  13a621:	je     13a6aa <litchi_pptx::shape::reader::Scene::read_with+0x254a>
  13a627:	mov    %rax,%r12
  13a62a:	movups -0x112d84(%rip),%xmm0        # 278ad <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13130224050040044120+0x6b5>
  13a631:	movups %xmm0,0x1c(%rax)
  13a635:	movups -0x112d9b(%rip),%xmm0        # 278a1 <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13130224050040044120+0x6a9>
  13a63c:	movups %xmm0,0x10(%rax)
  13a640:	movups -0x112db6(%rip),%xmm0        # 27891 <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13130224050040044120+0x699>
  13a647:	jmp    13a561 <litchi_pptx::shape::reader::Scene::read_with+0x2401>
  13a64c:	mov    $0x1,%edi
  13a651:	mov    $0x36,%esi
  13a656:	call   *0xf1dd4(%rip)        # 22c430 <_DYNAMIC+0x230>
  13a65c:	mov    $0x39,%ebx
  13a661:	mov    $0x39,%edi
  13a666:	call   *0xf1d7c(%rip)        # 22c3e8 <malloc@GLIBC_2.2.5>
  13a66c:	test   %rax,%rax
  13a66f:	je     13a6aa <litchi_pptx::shape::reader::Scene::read_with+0x254a>
  13a671:	mov    %rax,%r12
  13a674:	movups -0x112d95(%rip),%xmm0        # 278e6 <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13130224050040044120+0x6ee>
  13a67b:	movups %xmm0,0x29(%rax)
  13a67f:	movups -0x112da9(%rip),%xmm0        # 278dd <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13130224050040044120+0x6e5>
  13a686:	movups %xmm0,0x20(%rax)
  13a68a:	movups -0x112dc4(%rip),%xmm0        # 278cd <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13130224050040044120+0x6d5>
  13a691:	movups %xmm0,0x10(%rax)
  13a695:	movups -0x112ddf(%rip),%xmm0        # 278bd <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13130224050040044120+0x6c5>
  13a69c:	movups %xmm0,(%rax)
  13a69f:	mov    $0x39,%r13d
  13a6a5:	jmp    139c41 <litchi_pptx::shape::reader::Scene::read_with+0x1ae1>
  13a6aa:	mov    $0x1,%edi
  13a6af:	mov    %rbx,%rsi
  13a6b2:	call   *0xf1d78(%rip)        # 22c430 <_DYNAMIC+0x230>
  13a6b8:	jmp    13a6fc <litchi_pptx::shape::reader::Scene::read_with+0x259c>
  13a6ba:	mov    $0x1,%edi
  13a6bf:	mov    %r15,%rsi
  13a6c2:	call   *0xf1d68(%rip)        # 22c430 <_DYNAMIC+0x230>
  13a6c8:	jmp    13a6fc <litchi_pptx::shape::reader::Scene::read_with+0x259c>
  13a6ca:	mov    $0x1,%edi
  13a6cf:	mov    %rbx,%rsi
  13a6d2:	call   *0xf1d58(%rip)        # 22c430 <_DYNAMIC+0x230>
  13a6d8:	jmp    13a6fc <litchi_pptx::shape::reader::Scene::read_with+0x259c>
  13a6da:	mov    $0x1,%edi
  13a6df:	mov    $0x3e,%esi
  13a6e4:	call   *0xf1d46(%rip)        # 22c430 <_DYNAMIC+0x230>
  13a6ea:	jmp    13a6fc <litchi_pptx::shape::reader::Scene::read_with+0x259c>
  13a6ec:	mov    $0x1,%edi
  13a6f1:	mov    $0x26,%esi
  13a6f6:	call   *0xf1d34(%rip)        # 22c430 <_DYNAMIC+0x230>
  13a6fc:	ud2
  13a6fe:	mov    %rax,%r14
  13a701:	lea    0x1c8(%rsp),%rdi
  13a709:	call   93550 <core::ptr::drop_in_place<litchi_ooxml_common::mce::model::NamespaceSet>>
  13a70e:	mov    %r14,%rdi
  13a711:	call   221ef0 <_Unwind_Resume@plt>
  13a716:	mov    %rax,%r14
  13a719:	lea    0x1a0(%rsp),%rdi
  13a721:	call   161920 <core::ptr::drop_in_place<quick_xml::events::Event>>
  13a726:	jmp    13a8a4 <litchi_pptx::shape::reader::Scene::read_with+0x2744>
  13a72b:	jmp    13a78d <litchi_pptx::shape::reader::Scene::read_with+0x262d>
  13a72d:	jmp    13a7e3 <litchi_pptx::shape::reader::Scene::read_with+0x2683>
  13a732:	jmp    13a81d <litchi_pptx::shape::reader::Scene::read_with+0x26bd>
  13a737:	mov    %rax,%r14
  13a73a:	lea    0x160(%rsp),%rdi
  13a742:	call   161920 <core::ptr::drop_in_place<quick_xml::events::Event>>
  13a747:	jmp    13a88b <litchi_pptx::shape::reader::Scene::read_with+0x272b>
  13a74c:	jmp    13a79e <litchi_pptx::shape::reader::Scene::read_with+0x263e>
  13a74e:	jmp    13a78d <litchi_pptx::shape::reader::Scene::read_with+0x262d>
  13a750:	jmp    13a832 <litchi_pptx::shape::reader::Scene::read_with+0x26d2>
  13a755:	jmp    13a86a <litchi_pptx::shape::reader::Scene::read_with+0x270a>
  13a75a:	mov    %r13,0x18(%rsp)
  13a75f:	mov    %rax,%r14
  13a762:	shl    $1,%r12
  13a765:	test   %r12,%r12
  13a768:	je     13a790 <litchi_pptx::shape::reader::Scene::read_with+0x2630>
  13a76a:	mov    %rbp,%rdi
  13a76d:	call   *0xf1c85(%rip)        # 22c3f8 <free@GLIBC_2.2.5>
  13a773:	jmp    13a790 <litchi_pptx::shape::reader::Scene::read_with+0x2630>
  13a775:	mov    %rax,%r14
  13a778:	test   %rbx,%rbx
  13a77b:	je     13a7a1 <litchi_pptx::shape::reader::Scene::read_with+0x2641>
  13a77d:	mov    %r12,%rdi
  13a780:	call   *0xf1c72(%rip)        # 22c3f8 <free@GLIBC_2.2.5>
  13a786:	jmp    13a7a1 <litchi_pptx::shape::reader::Scene::read_with+0x2641>
  13a788:	mov    %r13,0x18(%rsp)
  13a78d:	mov    %rax,%r14
  13a790:	test   %rbx,%rbx
  13a793:	jle    13a88b <litchi_pptx::shape::reader::Scene::read_with+0x272b>
  13a799:	jmp    13a845 <litchi_pptx::shape::reader::Scene::read_with+0x26e5>
  13a79e:	mov    %rax,%r14
  13a7a1:	cmpq   $0x0,0x110(%rsp)
  13a7aa:	jle    13a88b <litchi_pptx::shape::reader::Scene::read_with+0x272b>
  13a7b0:	mov    0x118(%rsp),%rdi
  13a7b8:	jmp    13a885 <litchi_pptx::shape::reader::Scene::read_with+0x2725>
  13a7bd:	mov    %r13,0x140(%rsp)
  13a7c5:	mov    %rax,%r14
  13a7c8:	shl    $1,%r12
  13a7cb:	test   %r12,%r12
  13a7ce:	je     13a7e6 <litchi_pptx::shape::reader::Scene::read_with+0x2686>
  13a7d0:	mov    %rbp,%rdi
  13a7d3:	call   *0xf1c1f(%rip)        # 22c3f8 <free@GLIBC_2.2.5>
  13a7d9:	jmp    13a7e6 <litchi_pptx::shape::reader::Scene::read_with+0x2686>
  13a7db:	mov    %r13,0x140(%rsp)
  13a7e3:	mov    %rax,%r14
  13a7e6:	shl    $1,%rbx
  13a7e9:	test   %rbx,%rbx
  13a7ec:	je     13a820 <litchi_pptx::shape::reader::Scene::read_with+0x26c0>
  13a7ee:	mov    0xa8(%rsp),%rdi
  13a7f6:	call   *0xf1bfc(%rip)        # 22c3f8 <free@GLIBC_2.2.5>
  13a7fc:	jmp    13a820 <litchi_pptx::shape::reader::Scene::read_with+0x26c0>
  13a7fe:	mov    %rax,%r14
  13a801:	lea    0x30(%rsp),%rdi
  13a806:	call   1641d0 <core::ptr::drop_in_place<quick_xml::reader::Reader<&[u8]>>>
  13a80b:	jmp    13a8b1 <litchi_pptx::shape::reader::Scene::read_with+0x2751>
  13a810:	mov    %r12,0x18(%rsp)
  13a815:	mov    %r13,0x140(%rsp)
  13a81d:	mov    %rax,%r14
  13a820:	cmpq   $0x0,0x18(%rsp)
  13a826:	jle    13a88b <litchi_pptx::shape::reader::Scene::read_with+0x272b>
  13a828:	mov    0x140(%rsp),%rdi
  13a830:	jmp    13a885 <litchi_pptx::shape::reader::Scene::read_with+0x2725>
  13a832:	mov    %rax,%r14
  13a835:	mov    0xa8(%rsp),%rax
  13a83d:	shl    $1,%rax
  13a840:	test   %rax,%rax
  13a843:	je     13a88b <litchi_pptx::shape::reader::Scene::read_with+0x272b>
  13a845:	mov    0x18(%rsp),%rdi
  13a84a:	jmp    13a885 <litchi_pptx::shape::reader::Scene::read_with+0x2725>
  13a84c:	jmp    13a86f <litchi_pptx::shape::reader::Scene::read_with+0x270f>
  13a84e:	jmp    13a86f <litchi_pptx::shape::reader::Scene::read_with+0x270f>
  13a850:	jmp    13a86a <litchi_pptx::shape::reader::Scene::read_with+0x270a>
  13a852:	mov    %rax,%r14
  13a855:	lea    0x370(%rsp),%rdi
  13a85d:	call   163da0 <core::ptr::drop_in_place<litchi_ooxml_common::mce::model::Capabilities>>
  13a862:	mov    %r14,%rdi
  13a865:	call   221ef0 <_Unwind_Resume@plt>
  13a86a:	mov    %rax,%r14
  13a86d:	jmp    13a8a4 <litchi_pptx::shape::reader::Scene::read_with+0x2744>
  13a86f:	mov    %rax,%r14
  13a872:	cmpq   $0x0,0xb0(%rsp)
  13a87b:	jle    13a88b <litchi_pptx::shape::reader::Scene::read_with+0x272b>
  13a87d:	mov    0xb8(%rsp),%rdi
  13a885:	call   *0xf1b6d(%rip)        # 22c3f8 <free@GLIBC_2.2.5>
  13a88b:	cmpq   $0x0,0x90(%rsp)
  13a894:	jle    13a8a4 <litchi_pptx::shape::reader::Scene::read_with+0x2744>
  13a896:	mov    0x98(%rsp),%rdi
  13a89e:	call   *0xf1b54(%rip)        # 22c3f8 <free@GLIBC_2.2.5>
  13a8a4:	lea    0x2b0(%rsp),%rdi
  13a8ac:	call   164800 <core::ptr::drop_in_place<quick_xml::reader::ns_reader::NsReader<&[u8]>>>
  13a8b1:	cmpq   $0x0,0x1e8(%rsp)
  13a8ba:	jne    13a904 <litchi_pptx::shape::reader::Scene::read_with+0x27a4>
  13a8bc:	cmpq   $0x0,0x200(%rsp)
  13a8c5:	jne    13a91d <litchi_pptx::shape::reader::Scene::read_with+0x27bd>
  13a8c7:	lea    0x218(%rsp),%rdi
  13a8cf:	call   1645d0 <core::ptr::drop_in_place<alloc::vec::Vec<litchi_pptx::shape::reader::Active>>>
  13a8d4:	cmpq   $0x0,0x230(%rsp)
  13a8dd:	jne    13a943 <litchi_pptx::shape::reader::Scene::read_with+0x27e3>
  13a8df:	mov    0x158(%rsp),%rax
  13a8e7:	shl    $1,%rax
  13a8ea:	test   %rax,%rax
  13a8ed:	jne    13a961 <litchi_pptx::shape::reader::Scene::read_with+0x2801>
  13a8ef:	lea    0x370(%rsp),%rdi
  13a8f7:	call   163da0 <core::ptr::drop_in_place<litchi_ooxml_common::mce::model::Capabilities>>
  13a8fc:	mov    %r14,%rdi
  13a8ff:	call   221ef0 <_Unwind_Resume@plt>
  13a904:	mov    0x1f0(%rsp),%rdi
  13a90c:	call   *0xf1ae6(%rip)        # 22c3f8 <free@GLIBC_2.2.5>
  13a912:	cmpq   $0x0,0x200(%rsp)
  13a91b:	je     13a8c7 <litchi_pptx::shape::reader::Scene::read_with+0x2767>
  13a91d:	mov    0x208(%rsp),%rdi
  13a925:	call   *0xf1acd(%rip)        # 22c3f8 <free@GLIBC_2.2.5>
  13a92b:	lea    0x218(%rsp),%rdi
  13a933:	call   1645d0 <core::ptr::drop_in_place<alloc::vec::Vec<litchi_pptx::shape::reader::Active>>>
  13a938:	cmpq   $0x0,0x230(%rsp)
  13a941:	je     13a8df <litchi_pptx::shape::reader::Scene::read_with+0x277f>
  13a943:	mov    0x238(%rsp),%rdi
  13a94b:	call   *0xf1aa7(%rip)        # 22c3f8 <free@GLIBC_2.2.5>
  13a951:	mov    0x158(%rsp),%rax
  13a959:	shl    $1,%rax
  13a95c:	test   %rax,%rax
  13a95f:	je     13a8ef <litchi_pptx::shape::reader::Scene::read_with+0x278f>
  13a961:	mov    0x190(%rsp),%rdi
  13a969:	call   *0xf1a89(%rip)        # 22c3f8 <free@GLIBC_2.2.5>
  13a96f:	lea    0x370(%rsp),%rdi
  13a977:	call   163da0 <core::ptr::drop_in_place<litchi_ooxml_common::mce::model::Capabilities>>
  13a97c:	mov    %r14,%rdi
  13a97f:	call   221ef0 <_Unwind_Resume@plt>

Disassembly of section .init:

Disassembly of section .fini:

Disassembly of section .plt:
