
/home/zhuhe/code/litchi-target-0823/after-real-native:     file format elf64-x86-64


Disassembly of section .text:

0000000000138310 <litchi_pptx::shape::reader::Scene::read_with>:
  138310:	push   %rbp
  138311:	push   %r15
  138313:	push   %r14
  138315:	push   %r13
  138317:	push   %r12
  138319:	push   %rbx
  13831a:	sub    $0x3e8,%rsp
  138321:	mov    %rdi,%r14
  138324:	mov    (%rcx),%r13
  138327:	cmp    %r13,%rdx
  13832a:	jbe    13834a <litchi_pptx::shape::reader::Scene::read_with+0x3a>
  13832c:	movb   $0x9,0x8(%r14)
  138331:	mov    %r13,0x10(%r14)
  138335:	lea    -0x110f83(%rip),%rax        # 273b9 <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13171701955629477448+0x81>
  13833c:	mov    %rax,0x18(%r14)
  138340:	movq   $0x17,0x20(%r14)
  138348:	jmp    1383bc <litchi_pptx::shape::reader::Scene::read_with+0xac>
  13834a:	mov    %rcx,0xe8(%rsp)
  138352:	mov    0x8(%rcx),%r12
  138356:	mov    %r12,%rax
  138359:	shr    $0x20,%rax
  13835d:	je     1383d1 <litchi_pptx::shape::reader::Scene::read_with+0xc1>
  13835f:	mov    $0x36,%edi
  138364:	call   *0xf46de(%rip)        # 22ca48 <malloc@GLIBC_2.2.5>
  13836a:	test   %rax,%rax
  13836d:	je     13ab7a <litchi_pptx::shape::reader::Scene::read_with+0x286a>
  138373:	movups -0x110fd7(%rip),%xmm0        # 273a3 <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13171701955629477448+0x6b>
  13837a:	movups %xmm0,0x20(%rax)
  13837e:	movups -0x110ff2(%rip),%xmm0        # 27393 <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13171701955629477448+0x5b>
  138385:	movups %xmm0,0x10(%rax)
  138389:	movdqu -0x11100e(%rip),%xmm0        # 27383 <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13171701955629477448+0x4b>
  138391:	movdqu %xmm0,(%rax)
  138395:	movabs $0x6e69616d6f64206e,%rcx
  13839f:	mov    %rcx,0x2e(%rax)
  1383a3:	movb   $0x8,0x8(%r14)
  1383a8:	movq   $0x36,0x10(%r14)
  1383b0:	mov    %rax,0x18(%r14)
  1383b4:	movq   $0x36,0x20(%r14)
  1383bc:	movabs $0x8000000000000001,%rax
  1383c6:	dec    %rax
  1383c9:	mov    %rax,(%r14)
  1383cc:	jmp    13a8c3 <litchi_pptx::shape::reader::Scene::read_with+0x25b3>
  1383d1:	mov    %rdx,%rbx
  1383d4:	mov    %rsi,%r15
  1383d7:	movq   $0x0,0x148(%rsp)
  1383e3:	mov    $0xffffffffffffffd8,%rbp
  1383ea:	cmpb   $0x1,%fs:0x10(%rbp)
  1383ef:	jne    13a537 <litchi_pptx::shape::reader::Scene::read_with+0x2227>
  1383f5:	mov    %fs:0x0(%rbp),%rax
  1383fa:	mov    %fs:0x8(%rbp),%rdx
  1383ff:	lea    0x1(%rax),%rcx
  138403:	mov    %rcx,%fs:0x0(%rbp)
  138408:	movups 0x148(%rsp),%xmm0
  138410:	movdqu 0x158(%rsp),%xmm1
  138419:	movups 0x168(%rsp),%xmm2
  138421:	movaps %xmm0,0x380(%rsp)
  138429:	movdqa %xmm1,0x390(%rsp)
  138432:	movaps %xmm2,0x3a0(%rsp)
  13843a:	movups 0xeb8b7(%rip),%xmm0        # 223cf8 <anon.4cef9af9ef5ad65ef8056305e811a596.14.llvm.8137685535559598875>
  138441:	movaps %xmm0,0x350(%rsp)
  138449:	movdqu 0xeb8b7(%rip),%xmm0        # 223d08 <anon.4cef9af9ef5ad65ef8056305e811a596.14.llvm.8137685535559598875+0x10>
  138451:	movdqa %xmm0,0x360(%rsp)
  13845a:	mov    %rax,0x370(%rsp)
  138462:	mov    %rdx,0x378(%rsp)
  13846a:	lea    -0x113812(%rip),%rsi        # 24c5f <anon.bc3d94ef2394061926bee58178356194.284.llvm.5373085861489949100+0x4c67>
  138471:	lea    0x350(%rsp),%rdi
  138479:	mov    $0x38,%edx
  13847e:	call   15fea0 <litchi_ooxml_common::mce::model::Capabilities::understand_namespace>
  138483:	lea    -0x113622(%rip),%rsi        # 24e68 <anon.bc3d94ef2394061926bee58178356194.284.llvm.5373085861489949100+0x4e70>
  13848a:	lea    0x350(%rsp),%rdi
  138492:	mov    $0x38,%edx
  138497:	call   15fea0 <litchi_ooxml_common::mce::model::Capabilities::understand_namespace>
  13849c:	lea    -0x112b81(%rip),%rsi        # 25922 <anon.bc3d94ef2394061926bee58178356194.284.llvm.5373085861489949100+0x592a>
  1384a3:	lea    0x350(%rsp),%rdi
  1384ab:	mov    $0x3b,%edx
  1384b0:	call   15fea0 <litchi_ooxml_common::mce::model::Capabilities::understand_namespace>
  1384b5:	mov    0xe8(%rsp),%rax
  1384bd:	mov    0x10(%rax),%rax
  1384c1:	mov    %r13,0x3b0(%rsp)
  1384c9:	mov    %r12,0x3b8(%rsp)
  1384d1:	mov    %rax,0x3c0(%rsp)
  1384d9:	movq   $0x1000,0x3c8(%rsp)
  1384e5:	movq   $0x1000,0x3d0(%rsp)
  1384f1:	movq   $0x400,0x3d8(%rsp)
  1384fd:	movq   $0x400,0x3e0(%rsp)
  138509:	lea    0x148(%rsp),%rdi
  138511:	lea    0x350(%rsp),%rcx
  138519:	lea    0x3b0(%rsp),%r8
  138521:	mov    %r15,%rsi
  138524:	mov    %rbx,%rdx
  138527:	call   *0xf4a33(%rip)        # 22cf60 <_DYNAMIC+0x700>
  13852d:	movabs $0x8000000000000001,%r15
  138537:	mov    0x148(%rsp),%rcx
  13853f:	mov    0x150(%rsp),%rbx
  138547:	mov    0x158(%rsp),%rax
  13854f:	cmp    %r15,%rcx
  138552:	jne    13857b <litchi_pptx::shape::reader::Scene::read_with+0x26b>
  138554:	movdqu 0x160(%rsp),%xmm0
  13855d:	movb   $0x22,0x8(%r14)
  138562:	mov    %rbx,0x10(%r14)
  138566:	mov    %rax,0x18(%r14)
  13856a:	movdqu %xmm0,0x20(%r14)
  138570:	dec    %r15
  138573:	mov    %r15,(%r14)
  138576:	jmp    13a8b6 <litchi_pptx::shape::reader::Scene::read_with+0x25a6>
  13857b:	cmp    %r12,%rax
  13857e:	jbe    1385b8 <litchi_pptx::shape::reader::Scene::read_with+0x2a8>
  138580:	movb   $0x9,0x8(%r14)
  138585:	mov    %r12,0x10(%r14)
  138589:	lea    -0x111228(%rip),%rax        # 27368 <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13171701955629477448+0x30>
  138590:	mov    %rax,0x18(%r14)
  138594:	movq   $0x1b,0x20(%r14)
  13859c:	dec    %r15
  13859f:	mov    %r15,(%r14)
  1385a2:	shl    $1,%rcx
  1385a5:	test   %rcx,%rcx
  1385a8:	je     1385b3 <litchi_pptx::shape::reader::Scene::read_with+0x2a3>
  1385aa:	mov    %rbx,%rdi
  1385ad:	call   *0xf44a5(%rip)        # 22ca58 <free@GLIBC_2.2.5>
  1385b3:	jmp    13a8b6 <litchi_pptx::shape::reader::Scene::read_with+0x25a6>
  1385b8:	mov    %rcx,0x138(%rsp)
  1385c0:	mov    %rbx,0x1c8(%rsp)
  1385c8:	mov    %rax,0x1d0(%rsp)
  1385d0:	mov    0xe8(%rsp),%r13
  1385d8:	movdqu 0x0(%r13),%xmm0
  1385de:	movdqu 0x10(%r13),%xmm1
  1385e4:	movups 0x20(%r13),%xmm2
  1385e9:	movdqu %xmm0,0x1d8(%rsp)
  1385f2:	movdqu %xmm1,0x1e8(%rsp)
  1385fb:	movups %xmm2,0x1f8(%rsp)
  138603:	movq   $0x0,0x168(%rsp)
  13860f:	movq   $0x8,0x170(%rsp)
  13861b:	pxor   %xmm0,%xmm0
  13861f:	movdqu %xmm0,0x178(%rsp)
  138628:	movq   $0x1,0x188(%rsp)
  138634:	movdqu %xmm0,0x190(%rsp)
  13863d:	movq   $0x8,0x1a0(%rsp)
  138649:	movdqu %xmm0,0x1a8(%rsp)
  138652:	movdqu %xmm0,0x208(%rsp)
  13865b:	movq   $0x0,0x218(%rsp)
  138667:	movq   $0x1,0x1b8(%rsp)
  138673:	movq   $0x0,0x1c0(%rsp)
  13867f:	movq   $0x0,0x148(%rsp)
  13868b:	movq   $0x0,0x158(%rsp)
  138697:	movb   $0x0,0x220(%rsp)
  13869f:	mov    %rbx,0x2d8(%rsp)
  1386a7:	mov    %rax,0x268(%rsp)
  1386af:	mov    %rax,0x2e0(%rsp)
  1386b7:	movq   $0x0,0x288(%rsp)
  1386c3:	movq   $0x1,0x290(%rsp)
  1386cf:	movdqu %xmm0,0x298(%rsp)
  1386d8:	movq   $0x8,0x2a8(%rsp)
  1386e4:	movq   $0x0,0x2b0(%rsp)
  1386f0:	movl   $0x0,0x2b7(%rsp)
  1386fb:	movw   $0x1,0x2bb(%rsp)
  138705:	movb   $0x1,0x2bd(%rsp)
  13870d:	movdqu %xmm0,0x2be(%rsp)
  138716:	movl   $0x0,0x2cd(%rsp)
  138721:	lea    0x228(%rsp),%rdi
  138729:	mov    %rbx,0x140(%rsp)
  138731:	call   *0xf4a11(%rip)        # 22d148 <_DYNAMIC+0x8e8>
  138737:	cmpq   $0x3,0x1d0(%rsp)
  138740:	jb     138769 <litchi_pptx::shape::reader::Scene::read_with+0x459>
  138742:	mov    0x1c8(%rsp),%rax
  13874a:	cmpb   $0xef,(%rax)
  13874d:	jne    138769 <litchi_pptx::shape::reader::Scene::read_with+0x459>
  13874f:	cmpb   $0xbb,0x1(%rax)
  138753:	jne    138769 <litchi_pptx::shape::reader::Scene::read_with+0x459>
  138755:	xor    %ecx,%ecx
  138757:	cmpb   $0xbf,0x2(%rax)
  13875b:	sete   %cl
  13875e:	lea    (%rcx,%rcx,2),%rax
  138762:	mov    %rax,0x78(%rsp)
  138767:	jmp    138772 <litchi_pptx::shape::reader::Scene::read_with+0x462>
  138769:	movq   $0x0,0x78(%rsp)
  138772:	xor    %ebx,%ebx
  138774:	mov    $0x12,%eax
  138779:	movq   %rax,%xmm0
  13877e:	movdqa %xmm0,0x10(%rsp)
  138784:	mov    $0x17,%eax
  138789:	movq   %rax,%xmm0
  13878e:	movdqa %xmm0,0x300(%rsp)
  138797:	movl   $0x0,0xac(%rsp)
  1387a2:	jmp    1387b8 <litchi_pptx::shape::reader::Scene::read_with+0x4a8>
  1387a4:	data16 data16 cs nopw 0x0(%rax,%rax,1)
  1387b0:	mov    0x2c0(%rsp),%rbx
  1387b8:	add    0x78(%rsp),%rbx
  1387bd:	jb     1398d3 <litchi_pptx::shape::reader::Scene::read_with+0x15c3>
  1387c3:	testb  $0x1,0xac(%rsp)
  1387cb:	je     13885a <litchi_pptx::shape::reader::Scene::read_with+0x54a>
  1387d1:	movzwl 0x260(%rsp),%edx
  1387d9:	cmp    $0x1,%dx
  1387dd:	adc    $0xffffffff,%edx
  1387e0:	mov    %dx,0x260(%rsp)
  1387e8:	mov    0x248(%rsp),%rax
  1387f0:	mov    0x250(%rsp),%rsi
  1387f8:	mov    %rsi,%rcx
  1387fb:	shl    $0x5,%rcx
  1387ff:	lea    (%rax,%rcx,1),%r8
  138803:	xor    %edi,%edi
  138805:	data16 cs nopw 0x0(%rax,%rax,1)
  138810:	test   %rcx,%rcx
  138813:	je     138846 <litchi_pptx::shape::reader::Scene::read_with+0x536>
  138815:	add    $0xffffffffffffffe0,%rcx
  138819:	inc    %rdi
  13881c:	cmp    %dx,-0x8(%r8)
  138821:	lea    -0x20(%r8),%r8
  138825:	ja     138810 <litchi_pptx::shape::reader::Scene::read_with+0x500>
  138827:	mov    %rsi,%rdx
  13882a:	sub    %rdi,%rdx
  13882d:	inc    %rdx
  138830:	cmp    %rsi,%rdx
  138833:	jae    13885a <litchi_pptx::shape::reader::Scene::read_with+0x54a>
  138835:	mov    0x20(%rax,%rcx,1),%rax
  13883a:	cmp    0x238(%rsp),%rax
  138842:	jbe    13884a <litchi_pptx::shape::reader::Scene::read_with+0x53a>
  138844:	jmp    138852 <litchi_pptx::shape::reader::Scene::read_with+0x542>
  138846:	xor    %eax,%eax
  138848:	xor    %edx,%edx
  13884a:	mov    %rax,0x238(%rsp)
  138852:	mov    %rdx,0x250(%rsp)
  13885a:	lea    0x20(%rsp),%rdi
  13885f:	lea    0x288(%rsp),%rsi
  138867:	call   179890 <quick_xml::reader::Reader<R>::read_event_impl>
  13886c:	mov    0x20(%rsp),%rax
  138871:	lea    0x28(%rsp),%rdx
  138876:	mov    0x20(%rdx),%rcx
  13887a:	mov    %rcx,0x130(%rsp)
  138882:	lea    0xe(%r15),%rcx
  138886:	movups (%rdx),%xmm0
  138889:	movups 0x10(%rdx),%xmm1
  13888d:	movaps %xmm0,0x110(%rsp)
  138895:	movaps %xmm1,0x120(%rsp)
  13889d:	cmp    %rcx,%rax
  1388a0:	jne    139b2c <litchi_pptx::shape::reader::Scene::read_with+0x181c>
  1388a6:	mov    0x130(%rsp),%rax
  1388ae:	mov    %rax,0xa0(%rsp)
  1388b6:	movdqa 0x110(%rsp),%xmm0
  1388bf:	movdqa 0x120(%rsp),%xmm1
  1388c8:	movdqa %xmm1,0x90(%rsp)
  1388d1:	movdqa %xmm0,0x80(%rsp)
  1388da:	mov    0x80(%rsp),%r15
  1388e2:	test   %r15,%r15
  1388e5:	je     138924 <litchi_pptx::shape::reader::Scene::read_with+0x614>
  1388e7:	cmp    $0x1,%r15
  1388eb:	je     138919 <litchi_pptx::shape::reader::Scene::read_with+0x609>
  1388ed:	cmp    $0x2,%r15
  1388f1:	jne    13894a <litchi_pptx::shape::reader::Scene::read_with+0x63a>
  1388f3:	lea    0x20(%rsp),%rdi
  1388f8:	lea    0x228(%rsp),%rsi
  138900:	lea    0x88(%rsp),%rdx
  138908:	call   *0xf4e9a(%rip)        # 22d7a8 <_DYNAMIC+0xf48>
  13890e:	cmpl   $0x6,0x20(%rsp)
  138913:	jne    139d9f <litchi_pptx::shape::reader::Scene::read_with+0x1a8f>
  138919:	mov    $0x1,%al
  13891b:	mov    %eax,0xac(%rsp)
  138922:	jmp    138955 <litchi_pptx::shape::reader::Scene::read_with+0x645>
  138924:	lea    0x20(%rsp),%rdi
  138929:	lea    0x228(%rsp),%rsi
  138931:	lea    0x88(%rsp),%rdx
  138939:	call   *0xf4e69(%rip)        # 22d7a8 <_DYNAMIC+0xf48>
  13893f:	cmpl   $0x6,0x20(%rsp)
  138944:	jne    139d56 <litchi_pptx::shape::reader::Scene::read_with+0x1a46>
  13894a:	movl   $0x0,0xac(%rsp)
  138955:	mov    0x2c0(%rsp),%rbp
  13895d:	add    0x78(%rsp),%rbp
  138962:	jb     1398f4 <litchi_pptx::shape::reader::Scene::read_with+0x15e4>
  138968:	cmp    $0xa,%r15
  13896c:	ja     139266 <litchi_pptx::shape::reader::Scene::read_with+0xf56>
  138972:	lea    -0x1147c5(%rip),%rcx        # 241b4 <anon.bc3d94ef2394061926bee58178356194.284.llvm.5373085861489949100+0x41bc>
  138979:	movslq (%rcx,%r15,4),%rax
  13897d:	add    %rcx,%rax
  138980:	jmp    *%rax
  138982:	mov    %rbp,0x70(%rsp)
  138987:	lea    0x88(%rsp),%rax
  13898f:	movdqu (%rax),%xmm0
  138993:	movdqu 0x10(%rax),%xmm1
  138998:	movdqa %xmm0,0xb0(%rsp)
  1389a1:	movdqa %xmm1,0xc0(%rsp)
  1389aa:	mov    0xc0(%rsp),%rdx
  1389b2:	mov    0xc8(%rsp),%r15
  1389ba:	cmp    %rdx,%r15
  1389bd:	ja     13a5c1 <litchi_pptx::shape::reader::Scene::read_with+0x22b1>
  1389c3:	mov    0xb8(%rsp),%rbp
  1389cb:	lea    (%r15,%rbp,1),%rdx
  1389cf:	mov    0xf4542(%rip),%rax        # 22cf18 <_DYNAMIC+0x6b8>
  1389d6:	mov    (%rax),%rax
  1389d9:	mov    $0x3a,%edi
  1389de:	mov    %rbp,%rsi
  1389e1:	call   *%rax
  1389e3:	cmp    $0x1,%rax
  1389e7:	jne    139060 <litchi_pptx::shape::reader::Scene::read_with+0xd50>
  1389ed:	sub    %rbp,%rdx
  1389f0:	mov    %rdx,%rcx
  1389f3:	mov    %rbp,%rdx
  1389f6:	jmp    13906a <litchi_pptx::shape::reader::Scene::read_with+0xd5a>
  1389fb:	mov    0x88(%rsp),%rbp
  138a03:	mov    0x90(%rsp),%rbx
  138a0b:	mov    0x98(%rsp),%rcx
  138a13:	lea    0x20(%rsp),%rdi
  138a18:	lea    0x148(%rsp),%rsi
  138a20:	mov    %rbx,%rdx
  138a23:	xor    %r8d,%r8d
  138a26:	call   13f980 <litchi_pptx::shape::reader::Scanner::reject_placeholder_character_data>
  138a2b:	movzbl 0x20(%rsp),%edx
  138a30:	cmp    $0x27,%dl
  138a33:	movabs $0x8000000000000001,%r15
  138a3d:	jne    139eb7 <litchi_pptx::shape::reader::Scene::read_with+0x1ba7>
  138a43:	mov    0x1a8(%rsp),%rax
  138a4b:	test   %rax,%rax
  138a4e:	je     138b52 <litchi_pptx::shape::reader::Scene::read_with+0x842>
  138a54:	mov    0x1a0(%rsp),%rcx
  138a5c:	shl    $0x8,%rax
  138a60:	cmpq   $0x0,-0xc0(%rcx,%rax,1)
  138a69:	je     138b52 <litchi_pptx::shape::reader::Scene::read_with+0x842>
  138a6f:	lea    0xb0(%rsp),%rdi
  138a77:	lea    0x88(%rsp),%rsi
  138a7f:	call   1a2ea0 <quick_xml::encoding::Decoder::content>
  138a84:	mov    0xb0(%rsp),%r15
  138a8c:	movabs $0x8000000000000001,%rax
  138a96:	cmp    %rax,%r15
  138a99:	je     13a147 <litchi_pptx::shape::reader::Scene::read_with+0x1e37>
  138a9f:	mov    0xb8(%rsp),%r12
  138aa7:	mov    0xc0(%rsp),%rdx
  138aaf:	lea    0x110(%rsp),%rdi
  138ab7:	mov    %r12,%rsi
  138aba:	mov    %r12,0x8(%rsp)
  138abf:	call   *0xf4d73(%rip)        # 22d838 <_DYNAMIC+0xfd8>
  138ac5:	movabs $0x8000000000000001,%rax
  138acf:	add    $0x2,%rax
  138ad3:	cmp    %rax,0x110(%rsp)
  138adb:	jne    13a1bc <litchi_pptx::shape::reader::Scene::read_with+0x1eac>
  138ae1:	mov    0x118(%rsp),%r13
  138ae9:	mov    0x120(%rsp),%r12
  138af1:	mov    0x128(%rsp),%rcx
  138af9:	lea    0x20(%rsp),%rdi
  138afe:	lea    0x148(%rsp),%rsi
  138b06:	mov    %r12,%rdx
  138b09:	call   13b070 <litchi_pptx::shape::reader::Scanner::append_text>
  138b0e:	movzbl 0x20(%rsp),%edx
  138b13:	cmp    $0x27,%dl
  138b16:	jne    13a2a0 <litchi_pptx::shape::reader::Scene::read_with+0x1f90>
  138b1c:	shl    $1,%r13
  138b1f:	test   %r13,%r13
  138b22:	je     138b2d <litchi_pptx::shape::reader::Scene::read_with+0x81d>
  138b24:	mov    %r12,%rdi
  138b27:	call   *0xf3f2b(%rip)        # 22ca58 <free@GLIBC_2.2.5>
  138b2d:	shl    $1,%r15
  138b30:	test   %r15,%r15
  138b33:	mov    0xe8(%rsp),%r13
  138b3b:	movabs $0x8000000000000001,%r15
  138b45:	je     138b52 <litchi_pptx::shape::reader::Scene::read_with+0x842>
  138b47:	mov    0x8(%rsp),%rdi
  138b4c:	call   *0xf3f06(%rip)        # 22ca58 <free@GLIBC_2.2.5>
  138b52:	test   %rbp,%rbp
  138b55:	jle    13975d <litchi_pptx::shape::reader::Scene::read_with+0x144d>
  138b5b:	jmp    139754 <litchi_pptx::shape::reader::Scene::read_with+0x1444>
  138b60:	lea    0x88(%rsp),%rax
  138b68:	movdqu (%rax),%xmm0
  138b6c:	movdqu 0x10(%rax),%xmm1
  138b71:	movdqa %xmm1,0xc0(%rsp)
  138b7a:	movdqa %xmm0,0xb0(%rsp)
  138b83:	mov    0xc0(%rsp),%rax
  138b8b:	mov    0xc8(%rsp),%rdx
  138b93:	cmp    %rax,%rdx
  138b96:	ja     13a5d8 <litchi_pptx::shape::reader::Scene::read_with+0x22c8>
  138b9c:	mov    0xb8(%rsp),%r15
  138ba4:	add    %r15,%rdx
  138ba7:	mov    0xf436a(%rip),%rax        # 22cf18 <_DYNAMIC+0x6b8>
  138bae:	mov    (%rax),%rax
  138bb1:	mov    $0x3a,%edi
  138bb6:	mov    %r15,%rsi
  138bb9:	call   *%rax
  138bbb:	cmp    $0x1,%rax
  138bbf:	jne    138f51 <litchi_pptx::shape::reader::Scene::read_with+0xc41>
  138bc5:	sub    %r15,%rdx
  138bc8:	mov    %rdx,%rcx
  138bcb:	jmp    138f5c <litchi_pptx::shape::reader::Scene::read_with+0xc4c>
  138bd0:	mov    0x88(%rsp),%r13
  138bd8:	mov    0x90(%rsp),%rbx
  138be0:	mov    0x98(%rsp),%r12
  138be8:	lea    (%rbx,%r12,1),%r15
  138bec:	mov    0xf4325(%rip),%rax        # 22cf18 <_DYNAMIC+0x6b8>
  138bf3:	mov    (%rax),%rax
  138bf6:	mov    $0x3a,%edi
  138bfb:	mov    %rbx,%rsi
  138bfe:	mov    %r15,%rdx
  138c01:	mov    %r13,0xe0(%rsp)
  138c09:	call   *%rax
  138c0b:	cmp    $0x1,%rax
  138c0f:	jne    138e52 <litchi_pptx::shape::reader::Scene::read_with+0xb42>
  138c15:	sub    %rbx,%rdx
  138c18:	mov    %rdx,%rcx
  138c1b:	mov    %rbx,%rdx
  138c1e:	jmp    138e5c <litchi_pptx::shape::reader::Scene::read_with+0xb4c>
  138c23:	mov    0x88(%rsp),%r12
  138c2b:	mov    0x90(%rsp),%rbx
  138c33:	mov    0x1a8(%rsp),%rcx
  138c3b:	test   %rcx,%rcx
  138c3e:	je     138d17 <litchi_pptx::shape::reader::Scene::read_with+0xa07>
  138c44:	mov    0x1a0(%rsp),%rax
  138c4c:	shl    $0x8,%rcx
  138c50:	cmpb   $0x0,-0x80(%rax,%rcx,1)
  138c55:	jne    139c32 <litchi_pptx::shape::reader::Scene::read_with+0x1922>
  138c5b:	add    %rcx,%rax
  138c5e:	cmpq   $0x0,-0x70(%rax)
  138c63:	jne    139c32 <litchi_pptx::shape::reader::Scene::read_with+0x1922>
  138c69:	cmpb   $0x0,-0x60(%rax)
  138c6d:	jne    139c32 <litchi_pptx::shape::reader::Scene::read_with+0x1922>
  138c73:	cmpb   $0x0,-0x90(%rax)
  138c7a:	je     138c91 <litchi_pptx::shape::reader::Scene::read_with+0x981>
  138c7c:	mov    -0x88(%rax),%rcx
  138c83:	cmp    0x210(%rsp),%rcx
  138c8b:	je     13a4aa <litchi_pptx::shape::reader::Scene::read_with+0x219a>
  138c91:	cmpq   $0x0,-0xc0(%rax)
  138c99:	je     138d17 <litchi_pptx::shape::reader::Scene::read_with+0xa07>
  138c9b:	lea    0xb0(%rsp),%rdi
  138ca3:	lea    0x88(%rsp),%rsi
  138cab:	call   1a2ea0 <quick_xml::encoding::Decoder::content>
  138cb0:	mov    0xb0(%rsp),%r13
  138cb8:	movabs $0x8000000000000001,%rax
  138cc2:	cmp    %rax,%r13
  138cc5:	je     13a378 <litchi_pptx::shape::reader::Scene::read_with+0x2068>
  138ccb:	mov    0xb8(%rsp),%r15
  138cd3:	mov    0xc0(%rsp),%rcx
  138cdb:	lea    0x20(%rsp),%rdi
  138ce0:	lea    0x148(%rsp),%rsi
  138ce8:	mov    %r15,%rdx
  138ceb:	call   13b070 <litchi_pptx::shape::reader::Scanner::append_text>
  138cf0:	movzbl 0x20(%rsp),%edx
  138cf5:	cmp    $0x27,%dl
  138cf8:	jne    13a3e8 <litchi_pptx::shape::reader::Scene::read_with+0x20d8>
  138cfe:	shl    $1,%r13
  138d01:	test   %r13,%r13
  138d04:	mov    0xe8(%rsp),%r13
  138d0c:	je     138d17 <litchi_pptx::shape::reader::Scene::read_with+0xa07>
  138d0e:	mov    %r15,%rdi
  138d11:	call   *0xf3d41(%rip)        # 22ca58 <free@GLIBC_2.2.5>
  138d17:	test   %r12,%r12
  138d1a:	jle    139266 <litchi_pptx::shape::reader::Scene::read_with+0xf56>
  138d20:	mov    %rbx,%rdi
  138d23:	jmp    139260 <litchi_pptx::shape::reader::Scene::read_with+0xf50>
  138d28:	lea    0x88(%rsp),%rcx
  138d30:	mov    0x10(%rcx),%rax
  138d34:	mov    %rax,0x120(%rsp)
  138d3c:	movdqu (%rcx),%xmm0
  138d40:	movdqa %xmm0,0x110(%rsp)
  138d49:	mov    0x1a8(%rsp),%rcx
  138d51:	test   %rcx,%rcx
  138d54:	movabs $0x8000000000000001,%r15
  138d5e:	je     138e36 <litchi_pptx::shape::reader::Scene::read_with+0xb26>
  138d64:	mov    0x1a0(%rsp),%rax
  138d6c:	shl    $0x8,%rcx
  138d70:	cmpb   $0x0,-0x80(%rax,%rcx,1)
  138d75:	jne    139cbf <litchi_pptx::shape::reader::Scene::read_with+0x19af>
  138d7b:	add    %rcx,%rax
  138d7e:	cmpq   $0x0,-0x70(%rax)
  138d83:	jne    139cbf <litchi_pptx::shape::reader::Scene::read_with+0x19af>
  138d89:	cmpb   $0x0,-0x60(%rax)
  138d8d:	jne    139cbf <litchi_pptx::shape::reader::Scene::read_with+0x19af>
  138d93:	cmpb   $0x0,-0x90(%rax)
  138d9a:	je     138db1 <litchi_pptx::shape::reader::Scene::read_with+0xaa1>
  138d9c:	mov    -0x88(%rax),%rcx
  138da3:	cmp    0x210(%rsp),%rcx
  138dab:	je     13a4f1 <litchi_pptx::shape::reader::Scene::read_with+0x21e1>
  138db1:	cmpq   $0x0,-0xc0(%rax)
  138db9:	je     138e36 <litchi_pptx::shape::reader::Scene::read_with+0xb26>
  138dbb:	lea    0xb0(%rsp),%rdi
  138dc3:	lea    0x110(%rsp),%rsi
  138dcb:	call   *0xf428f(%rip)        # 22d060 <_DYNAMIC+0x800>
  138dd1:	mov    0xb0(%rsp),%rbp
  138dd9:	mov    0xb8(%rsp),%r15
  138de1:	mov    0xc0(%rsp),%rbx
  138de9:	mov    0xc8(%rsp),%rcx
  138df1:	cmp    $0x2,%rbp
  138df5:	jne    13a35a <litchi_pptx::shape::reader::Scene::read_with+0x204a>
  138dfb:	lea    0x20(%rsp),%rdi
  138e00:	lea    0x148(%rsp),%rsi
  138e08:	mov    %rbx,%rdx
  138e0b:	call   13b070 <litchi_pptx::shape::reader::Scanner::append_text>
  138e10:	movzbl 0x20(%rsp),%edx
  138e15:	cmp    $0x27,%dl
  138e18:	jne    13a450 <litchi_pptx::shape::reader::Scene::read_with+0x2140>
  138e1e:	test   %r15,%r15
  138e21:	movabs $0x8000000000000001,%r15
  138e2b:	je     138e36 <litchi_pptx::shape::reader::Scene::read_with+0xb26>
  138e2d:	mov    %rbx,%rdi
  138e30:	call   *0xf3c22(%rip)        # 22ca58 <free@GLIBC_2.2.5>
  138e36:	cmpq   $0x0,0x110(%rsp)
  138e3f:	jle    13975d <litchi_pptx::shape::reader::Scene::read_with+0x144d>
  138e45:	mov    0x118(%rsp),%rdi
  138e4d:	jmp    139757 <litchi_pptx::shape::reader::Scene::read_with+0x1447>
  138e52:	xor    %edx,%edx
  138e54:	mov    0x280(%rsp),%rcx
  138e5c:	lea    0x338(%rsp),%rdi
  138e64:	lea    0x228(%rsp),%rsi
  138e6c:	mov    %rcx,0x280(%rsp)
  138e74:	mov    $0x1,%r8d
  138e7a:	call   *0xf40a0(%rip)        # 22cf20 <_DYNAMIC+0x6c0>
  138e80:	mov    0x338(%rsp),%rax
  138e88:	mov    %rax,0x8(%rsp)
  138e8d:	mov    0x340(%rsp),%rax
  138e95:	mov    %rax,0x68(%rsp)
  138e9a:	mov    0x210(%rsp),%r13
  138ea2:	test   %r13,%r13
  138ea5:	je     139ecb <litchi_pptx::shape::reader::Scene::read_with+0x1bbb>
  138eab:	mov    %rbp,0x70(%rsp)
  138eb0:	mov    0x348(%rsp),%rbp
  138eb8:	mov    0xf4059(%rip),%rax        # 22cf18 <_DYNAMIC+0x6b8>
  138ebf:	mov    (%rax),%rax
  138ec2:	mov    $0x3a,%edi
  138ec7:	mov    %rbx,%rsi
  138eca:	mov    %r15,%rdx
  138ecd:	call   *%rax
  138ecf:	cmp    $0x1,%rax
  138ed3:	jne    139180 <litchi_pptx::shape::reader::Scene::read_with+0xe70>
  138ed9:	mov    %rdx,%rax
  138edc:	sub    %rbx,%rax
  138edf:	not    %rax
  138ee2:	add    %rax,%r12
  138ee5:	inc    %rdx
  138ee8:	mov    0x1a8(%rsp),%r15
  138ef0:	test   %r15,%r15
  138ef3:	je     139194 <litchi_pptx::shape::reader::Scene::read_with+0xe84>
  138ef9:	mov    %r15,%rsi
  138efc:	shl    $0x8,%rsi
  138f00:	add    0x1a0(%rsp),%rsi
  138f08:	cmp    $0x5,%r12
  138f0c:	jne    1393d4 <litchi_pptx::shape::reader::Scene::read_with+0x10c4>
  138f12:	mov    (%rdx),%eax
  138f14:	mov    $0x656d6163,%ecx
  138f19:	xor    %ecx,%eax
  138f1b:	movzbl 0x4(%rdx),%ecx
  138f1f:	xor    $0x6f,%ecx
  138f22:	or     %eax,%ecx
  138f24:	mov    0x8(%rsp),%r12
  138f29:	jne    139658 <litchi_pptx::shape::reader::Scene::read_with+0x1348>
  138f2f:	mov    %r12,%rax
  138f32:	movabs $0x8000000000000001,%rcx
  138f3c:	xor    %rcx,%rax
  138f3f:	xor    $0x3b,%rbp
  138f43:	or     %rax,%rbp
  138f46:	jne    139658 <litchi_pptx::shape::reader::Scene::read_with+0x1348>
  138f4c:	jmp    13941b <litchi_pptx::shape::reader::Scene::read_with+0x110b>
  138f51:	xor    %r15d,%r15d
  138f54:	mov    0x270(%rsp),%rcx
  138f5c:	lea    0x320(%rsp),%rdi
  138f64:	lea    0x228(%rsp),%rsi
  138f6c:	mov    %r15,%rdx
  138f6f:	mov    %rcx,0x270(%rsp)
  138f77:	mov    $0x1,%r8d
  138f7d:	call   *0xf3f9d(%rip)        # 22cf20 <_DYNAMIC+0x6c0>
  138f83:	mov    0x320(%rsp),%r15
  138f8b:	mov    0x328(%rsp),%rax
  138f93:	mov    %rax,0x70(%rsp)
  138f98:	mov    0x1f0(%rsp),%r12
  138fa0:	mov    0x218(%rsp),%r13
  138fa8:	cmp    $0xffffffffffffffff,%r13
  138fac:	je     13a663 <litchi_pptx::shape::reader::Scene::read_with+0x2353>
  138fb2:	lea    -0x111be9(%rip),%rax        # 273d0 <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13171701955629477448+0x98>
  138fb9:	mov    %rax,0x30(%rsp)
  138fbe:	movq   $0x12,0x38(%rsp)
  138fc7:	mov    %r12,0x28(%rsp)
  138fcc:	movb   $0x9,0x20(%rsp)
  138fd1:	lea    0x20(%rsp),%rdi
  138fd6:	call   161fb0 <core::ptr::drop_in_place<litchi_pptx::error::Error>>
  138fdb:	lea    0x1(%r13),%rax
  138fdf:	mov    %rax,0x218(%rsp)
  138fe7:	mov    $0x9,%dl
  138fe9:	cmp    %r12,%r13
  138fec:	jae    13a675 <litchi_pptx::shape::reader::Scene::read_with+0x2365>
  138ff2:	mov    0x1e8(%rsp),%r13
  138ffa:	mov    0x210(%rsp),%r12
  139002:	lea    -0x111aff(%rip),%rax        # 2750a <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13171701955629477448+0x1d2>
  139009:	mov    %rax,0x30(%rsp)
  13900e:	movq   $0x17,0x38(%rsp)
  139017:	inc    %r12
  13901a:	je     13a94e <litchi_pptx::shape::reader::Scene::read_with+0x263e>
  139020:	mov    %r13,0x28(%rsp)
  139025:	movb   $0x9,0x20(%rsp)
  13902a:	lea    0x20(%rsp),%rdi
  13902f:	call   161fb0 <core::ptr::drop_in_place<litchi_pptx::error::Error>>
  139034:	cmp    %r13,%r12
  139037:	ja     139f49 <litchi_pptx::shape::reader::Scene::read_with+0x1c39>
  13903d:	mov    0x1c0(%rsp),%rax
  139045:	test   %rax,%rax
  139048:	je     1391ee <litchi_pptx::shape::reader::Scene::read_with+0xede>
  13904e:	mov    0x1b8(%rsp),%rcx
  139056:	movzbl -0x1(%rcx,%rax,1),%eax
  13905b:	jmp    1391f0 <litchi_pptx::shape::reader::Scene::read_with+0xee0>
  139060:	xor    %edx,%edx
  139062:	mov    0x278(%rsp),%rcx
  13906a:	lea    0x2e8(%rsp),%rdi
  139072:	lea    0x228(%rsp),%rsi
  13907a:	mov    %rcx,0x278(%rsp)
  139082:	mov    $0x1,%r8d
  139088:	call   *0xf3e92(%rip)        # 22cf20 <_DYNAMIC+0x6c0>
  13908e:	mov    0x2e8(%rsp),%rax
  139096:	mov    %rax,0x8(%rsp)
  13909b:	mov    0x2f0(%rsp),%rax
  1390a3:	mov    %rax,0xe0(%rsp)
  1390ab:	mov    0x1f0(%rsp),%r12
  1390b3:	mov    0x218(%rsp),%r13
  1390bb:	cmp    $0xffffffffffffffff,%r13
  1390bf:	je     13a5f2 <litchi_pptx::shape::reader::Scene::read_with+0x22e2>
  1390c5:	mov    0x2f8(%rsp),%rax
  1390cd:	mov    %rax,0x68(%rsp)
  1390d2:	lea    -0x111d09(%rip),%rax        # 273d0 <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13171701955629477448+0x98>
  1390d9:	mov    %rax,0x30(%rsp)
  1390de:	movq   $0x12,0x38(%rsp)
  1390e7:	mov    %r12,0x28(%rsp)
  1390ec:	movb   $0x9,0x20(%rsp)
  1390f1:	lea    0x20(%rsp),%rdi
  1390f6:	call   161fb0 <core::ptr::drop_in_place<litchi_pptx::error::Error>>
  1390fb:	lea    0x1(%r13),%rax
  1390ff:	mov    %rax,0x218(%rsp)
  139107:	mov    $0x9,%dl
  139109:	cmp    %r12,%r13
  13910c:	jae    13a604 <litchi_pptx::shape::reader::Scene::read_with+0x22f4>
  139112:	mov    0x1e8(%rsp),%r13
  13911a:	mov    0x210(%rsp),%r12
  139122:	lea    -0x111c1f(%rip),%rax        # 2750a <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13171701955629477448+0x1d2>
  139129:	mov    %rax,0x30(%rsp)
  13912e:	movq   $0x17,0x38(%rsp)
  139137:	inc    %r12
  13913a:	je     13a979 <litchi_pptx::shape::reader::Scene::read_with+0x2669>
  139140:	mov    %r13,0x28(%rsp)
  139145:	movb   $0x9,0x20(%rsp)
  13914a:	lea    0x20(%rsp),%rdi
  13914f:	call   161fb0 <core::ptr::drop_in_place<litchi_pptx::error::Error>>
  139154:	cmp    %r13,%r12
  139157:	ja     139f1a <litchi_pptx::shape::reader::Scene::read_with+0x1c0a>
  13915d:	mov    0x1c0(%rsp),%rax
  139165:	test   %rax,%rax
  139168:	je     139275 <litchi_pptx::shape::reader::Scene::read_with+0xf65>
  13916e:	mov    0x1b8(%rsp),%rcx
  139176:	movzbl -0x1(%rcx,%rax,1),%eax
  13917b:	jmp    139277 <litchi_pptx::shape::reader::Scene::read_with+0xf67>
  139180:	mov    %rbx,%rdx
  139183:	mov    0x1a8(%rsp),%r15
  13918b:	test   %r15,%r15
  13918e:	jne    138ef9 <litchi_pptx::shape::reader::Scene::read_with+0xbe9>
  139194:	cmp    $0x1,%r12
  139198:	jne    13969e <litchi_pptx::shape::reader::Scene::read_with+0x138e>
  13919e:	cmpb   $0x74,(%rdx)
  1391a1:	sete   %al
  1391a4:	movabs $0x8000000000000001,%rcx
  1391ae:	mov    0x8(%rsp),%r12
  1391b3:	cmp    %rcx,%r12
  1391b6:	sete   %cl
  1391b9:	and    %al,%cl
  1391bb:	cmp    $0x1,%cl
  1391be:	jne    139653 <litchi_pptx::shape::reader::Scene::read_with+0x1343>
  1391c4:	cmp    $0x29,%rbp
  1391c8:	je     139634 <litchi_pptx::shape::reader::Scene::read_with+0x1324>
  1391ce:	cmp    $0x35,%rbp
  1391d2:	jne    139653 <litchi_pptx::shape::reader::Scene::read_with+0x1343>
  1391d8:	mov    $0x35,%edx
  1391dd:	mov    0x68(%rsp),%rdi
  1391e2:	lea    -0x114650(%rip),%rsi        # 24b99 <anon.bc3d94ef2394061926bee58178356194.284.llvm.5373085861489949100+0x4ba1>
  1391e9:	jmp    139645 <litchi_pptx::shape::reader::Scene::read_with+0x1335>
  1391ee:	xor    %eax,%eax
  1391f0:	sub    $0x8,%rsp
  1391f4:	movzbl %al,%eax
  1391f7:	lea    0x28(%rsp),%rdi
  1391fc:	lea    0x150(%rsp),%rsi
  139204:	lea    0x328(%rsp),%rdx
  13920c:	lea    0xb8(%rsp),%rcx
  139214:	mov    %rbx,%r8
  139217:	mov    %r12,%r9
  13921a:	push   %rbp
  13921b:	push   $0x1
  13921d:	push   %rax
  13921e:	call   13bf40 <litchi_pptx::shape::reader::Scanner::start_element>
  139223:	add    $0x20,%rsp
  139227:	movzbl 0x20(%rsp),%edx
  13922c:	cmp    $0x27,%dl
  13922f:	jne    13a047 <litchi_pptx::shape::reader::Scene::read_with+0x1d37>
  139235:	test   %r15,%r15
  139238:	mov    0xe8(%rsp),%r13
  139240:	jle    13924d <litchi_pptx::shape::reader::Scene::read_with+0xf3d>
  139242:	mov    0x70(%rsp),%rdi
  139247:	call   *0xf380b(%rip)        # 22ca58 <free@GLIBC_2.2.5>
  13924d:	cmpq   $0x0,0xb0(%rsp)
  139256:	jle    139266 <litchi_pptx::shape::reader::Scene::read_with+0xf56>
  139258:	mov    0xb8(%rsp),%rdi
  139260:	call   *0xf37f2(%rip)        # 22ca58 <free@GLIBC_2.2.5>
  139266:	movabs $0x8000000000000001,%r15
  139270:	jmp    13975d <litchi_pptx::shape::reader::Scene::read_with+0x144d>
  139275:	xor    %eax,%eax
  139277:	sub    $0x8,%rsp
  13927b:	movzbl %al,%eax
  13927e:	lea    0x28(%rsp),%rdi
  139283:	lea    0x150(%rsp),%rsi
  13928b:	lea    0x2f0(%rsp),%rdx
  139293:	lea    0xb8(%rsp),%rcx
  13929b:	mov    %rbx,%r8
  13929e:	mov    %r12,%r9
  1392a1:	push   0x78(%rsp)
  1392a5:	push   $0x0
  1392a7:	push   %rax
  1392a8:	call   13bf40 <litchi_pptx::shape::reader::Scanner::start_element>
  1392ad:	add    $0x20,%rsp
  1392b1:	movzbl 0x20(%rsp),%edx
  1392b6:	cmp    $0x27,%dl
  1392b9:	jne    139ff7 <litchi_pptx::shape::reader::Scene::read_with+0x1ce7>
  1392bf:	mov    0xf3c52(%rip),%rax        # 22cf18 <_DYNAMIC+0x6b8>
  1392c6:	mov    (%rax),%rax
  1392c9:	mov    $0x3a,%edi
  1392ce:	mov    %rbp,%rsi
  1392d1:	lea    (%r15,%rbp,1),%rdx
  1392d5:	call   *%rax
  1392d7:	cmp    $0x1,%rax
  1392db:	jne    1392ef <litchi_pptx::shape::reader::Scene::read_with+0xfdf>
  1392dd:	mov    %rdx,%rax
  1392e0:	sub    %rbp,%rax
  1392e3:	not    %rax
  1392e6:	add    %rax,%r15
  1392e9:	inc    %rdx
  1392ec:	mov    %rdx,%rbp
  1392ef:	cmp    $0x4,%r15
  1392f3:	jne    13934f <litchi_pptx::shape::reader::Scene::read_with+0x103f>
  1392f5:	cmpl   $0x7250766e,0x0(%rbp)
  1392fc:	sete   %al
  1392ff:	movabs $0x8000000000000001,%rcx
  139309:	cmp    %rcx,0x8(%rsp)
  13930e:	sete   %cl
  139311:	and    %al,%cl
  139313:	cmp    $0x1,%cl
  139316:	jne    13934f <litchi_pptx::shape::reader::Scene::read_with+0x103f>
  139318:	mov    0x68(%rsp),%rax
  13931d:	cmp    $0x2e,%rax
  139321:	je     139591 <litchi_pptx::shape::reader::Scene::read_with+0x1281>
  139327:	cmp    $0x3a,%rax
  13932b:	jne    13934f <litchi_pptx::shape::reader::Scene::read_with+0x103f>
  13932d:	mov    $0x3a,%edx
  139332:	mov    0xe0(%rsp),%rdi
  13933a:	lea    -0x11474a(%rip),%rsi        # 24bf7 <anon.bc3d94ef2394061926bee58178356194.284.llvm.5373085861489949100+0x4bff>
  139341:	call   *0xf3909(%rip)        # 22cc50 <bcmp@GLIBC_2.2.5>
  139347:	test   %eax,%eax
  139349:	je     1395b3 <litchi_pptx::shape::reader::Scene::read_with+0x12a3>
  13934f:	xor    %ebx,%ebx
  139351:	mov    0x1c0(%rsp),%r15
  139359:	cmp    0x1b0(%rsp),%r15
  139361:	jne    139371 <litchi_pptx::shape::reader::Scene::read_with+0x1061>
  139363:	lea    0x1b0(%rsp),%rdi
  13936b:	call   *0xf453f(%rip)        # 22d8b0 <_DYNAMIC+0x1050>
  139371:	mov    0x1b8(%rsp),%rax
  139379:	mov    %bl,(%rax,%r15,1)
  13937d:	inc    %r15
  139380:	mov    %r15,0x1c0(%rsp)
  139388:	mov    %r12,0x210(%rsp)
  139390:	cmpq   $0x0,0x8(%rsp)
  139396:	jle    1393a6 <litchi_pptx::shape::reader::Scene::read_with+0x1096>
  139398:	mov    0xe0(%rsp),%rdi
  1393a0:	call   *0xf36b2(%rip)        # 22ca58 <free@GLIBC_2.2.5>
  1393a6:	cmpq   $0x0,0xb0(%rsp)
  1393af:	mov    0xe8(%rsp),%r13
  1393b7:	movabs $0x8000000000000001,%r15
  1393c1:	jle    13975d <litchi_pptx::shape::reader::Scene::read_with+0x144d>
  1393c7:	mov    0xb8(%rsp),%rdi
  1393cf:	jmp    139757 <litchi_pptx::shape::reader::Scene::read_with+0x1447>
  1393d4:	cmp    $0x7,%r12
  1393d8:	jne    139461 <litchi_pptx::shape::reader::Scene::read_with+0x1151>
  1393de:	mov    (%rdx),%eax
  1393e0:	mov    $0x6e6b6e75,%ecx
  1393e5:	xor    %ecx,%eax
  1393e7:	mov    0x3(%rdx),%ecx
  1393ea:	mov    $0x6e776f6e,%edx
  1393ef:	xor    %edx,%ecx
  1393f1:	or     %eax,%ecx
  1393f3:	mov    0x8(%rsp),%r12
  1393f8:	jne    139658 <litchi_pptx::shape::reader::Scene::read_with+0x1348>
  1393fe:	mov    %r12,%rax
  139401:	movabs $0x8000000000000001,%rcx
  13940b:	xor    %rcx,%rax
  13940e:	xor    $0x3b,%rbp
  139412:	or     %rax,%rbp
  139415:	jne    139658 <litchi_pptx::shape::reader::Scene::read_with+0x1348>
  13941b:	mov    $0x3b,%edx
  139420:	mov    0x68(%rsp),%rdi
  139425:	mov    %rsi,%rbp
  139428:	lea    -0x113b0d(%rip),%rsi        # 25922 <anon.bc3d94ef2394061926bee58178356194.284.llvm.5373085861489949100+0x592a>
  13942f:	call   *0xf381b(%rip)        # 22cc50 <bcmp@GLIBC_2.2.5>
  139435:	mov    %rbp,%rcx
  139438:	test   %eax,%eax
  13943a:	jne    139658 <litchi_pptx::shape::reader::Scene::read_with+0x1348>
  139440:	cmpl   $0x1,-0x60(%rcx)
  139444:	jne    139658 <litchi_pptx::shape::reader::Scene::read_with+0x1348>
  13944a:	cmp    %r13,-0x58(%rcx)
  13944e:	jne    139658 <litchi_pptx::shape::reader::Scene::read_with+0x1348>
  139454:	movq   $0x0,-0x60(%rcx)
  13945c:	jmp    139658 <litchi_pptx::shape::reader::Scene::read_with+0x1348>
  139461:	cmp    $0x4,%r12
  139465:	jne    1394e5 <litchi_pptx::shape::reader::Scene::read_with+0x11d5>
  139467:	cmpl   $0x65707974,(%rdx)
  13946d:	mov    0x8(%rsp),%r12
  139472:	jne    139658 <litchi_pptx::shape::reader::Scene::read_with+0x1348>
  139478:	mov    %r12,%rax
  13947b:	movabs $0x8000000000000001,%rcx
  139485:	xor    %rcx,%rax
  139488:	xor    $0x3b,%rbp
  13948c:	or     %rax,%rbp
  13948f:	jne    139658 <litchi_pptx::shape::reader::Scene::read_with+0x1348>
  139495:	mov    $0x3b,%edx
  13949a:	mov    0x68(%rsp),%rdi
  13949f:	mov    %rsi,%rbp
  1394a2:	lea    -0x113b87(%rip),%rsi        # 25922 <anon.bc3d94ef2394061926bee58178356194.284.llvm.5373085861489949100+0x592a>
  1394a9:	call   *0xf37a1(%rip)        # 22cc50 <bcmp@GLIBC_2.2.5>
  1394af:	test   %eax,%eax
  1394b1:	jne    139658 <litchi_pptx::shape::reader::Scene::read_with+0x1348>
  1394b7:	mov    %rbp,%rcx
  1394ba:	cmpb   $0x0,-0x70(%rbp)
  1394be:	je     139658 <litchi_pptx::shape::reader::Scene::read_with+0x1348>
  1394c4:	cmp    %r13,-0x68(%rcx)
  1394c8:	jne    139658 <litchi_pptx::shape::reader::Scene::read_with+0x1348>
  1394ce:	cmpb   $0x0,-0xb(%rcx)
  1394d2:	je     13aa42 <litchi_pptx::shape::reader::Scene::read_with+0x2732>
  1394d8:	movq   $0x0,-0x70(%rcx)
  1394e0:	jmp    139658 <litchi_pptx::shape::reader::Scene::read_with+0x1348>
  1394e5:	cmp    $0x9,%r12
  1394e9:	jne    1395d0 <litchi_pptx::shape::reader::Scene::read_with+0x12c0>
  1394ef:	mov    (%rdx),%rax
  1394f2:	movabs $0x7845657079546870,%rcx
  1394fc:	xor    %rcx,%rax
  1394ff:	movzbl 0x8(%rdx),%ecx
  139503:	xor    $0x74,%rcx
  139507:	or     %rax,%rcx
  13950a:	mov    0x8(%rsp),%r12
  13950f:	jne    139658 <litchi_pptx::shape::reader::Scene::read_with+0x1348>
  139515:	mov    %r12,%rax
  139518:	movabs $0x8000000000000001,%rcx
  139522:	xor    %rcx,%rax
  139525:	xor    $0x3b,%rbp
  139529:	or     %rax,%rbp
  13952c:	jne    139658 <litchi_pptx::shape::reader::Scene::read_with+0x1348>
  139532:	mov    $0x3b,%edx
  139537:	mov    0x68(%rsp),%rdi
  13953c:	mov    %rsi,%rbp
  13953f:	lea    -0x113c24(%rip),%rsi        # 25922 <anon.bc3d94ef2394061926bee58178356194.284.llvm.5373085861489949100+0x592a>
  139546:	call   *0xf3704(%rip)        # 22cc50 <bcmp@GLIBC_2.2.5>
  13954c:	test   %eax,%eax
  13954e:	jne    139658 <litchi_pptx::shape::reader::Scene::read_with+0x1348>
  139554:	mov    %rbp,%rcx
  139557:	cmpb   $0x0,-0x80(%rbp)
  13955b:	je     139658 <litchi_pptx::shape::reader::Scene::read_with+0x1348>
  139561:	cmp    %r13,-0x78(%rcx)
  139565:	jne    139658 <litchi_pptx::shape::reader::Scene::read_with+0x1348>
  13956b:	cmpb   $0x0,-0xc(%rcx)
  13956f:	je     13aa8b <litchi_pptx::shape::reader::Scene::read_with+0x277b>
  139575:	cmpb   $0x1,-0xb(%rcx)
  139579:	jne    13aa8b <litchi_pptx::shape::reader::Scene::read_with+0x277b>
  13957f:	movq   $0x0,-0x80(%rcx)
  139587:	mov    0x8(%rsp),%r12
  13958c:	jmp    139658 <litchi_pptx::shape::reader::Scene::read_with+0x1348>
  139591:	mov    $0x2e,%edx
  139596:	mov    0xe0(%rsp),%rdi
  13959e:	lea    -0x114974(%rip),%rsi        # 24c31 <anon.bc3d94ef2394061926bee58178356194.284.llvm.5373085861489949100+0x4c39>
  1395a5:	call   *0xf36a5(%rip)        # 22cc50 <bcmp@GLIBC_2.2.5>
  1395ab:	test   %eax,%eax
  1395ad:	jne    13934f <litchi_pptx::shape::reader::Scene::read_with+0x103f>
  1395b3:	mov    $0x1,%bl
  1395b5:	mov    0x1c0(%rsp),%r15
  1395bd:	cmp    0x1b0(%rsp),%r15
  1395c5:	jne    139371 <litchi_pptx::shape::reader::Scene::read_with+0x1061>
  1395cb:	jmp    139363 <litchi_pptx::shape::reader::Scene::read_with+0x1053>
  1395d0:	cmp    $0x3,%r12
  1395d4:	jne    1397ff <litchi_pptx::shape::reader::Scene::read_with+0x14ef>
  1395da:	movzwl (%rdx),%eax
  1395dd:	xor    $0x7865,%eax
  1395e2:	movzbl 0x2(%rdx),%ecx
  1395e6:	xor    $0x74,%ecx
  1395e9:	or     %ax,%cx
  1395ec:	sete   %al
  1395ef:	movabs $0x8000000000000001,%rcx
  1395f9:	mov    0x8(%rsp),%r12
  1395fe:	cmp    %rcx,%r12
  139601:	sete   %cl
  139604:	and    %al,%cl
  139606:	cmp    $0x1,%cl
  139609:	jne    139658 <litchi_pptx::shape::reader::Scene::read_with+0x1348>
  13960b:	cmp    $0x2e,%rbp
  13960f:	je     139916 <litchi_pptx::shape::reader::Scene::read_with+0x1606>
  139615:	cmp    $0x3a,%rbp
  139619:	jne    139658 <litchi_pptx::shape::reader::Scene::read_with+0x1348>
  13961b:	mov    $0x3a,%edx
  139620:	mov    0x68(%rsp),%rdi
  139625:	mov    %rsi,%rbp
  139628:	lea    -0x114a38(%rip),%rsi        # 24bf7 <anon.bc3d94ef2394061926bee58178356194.284.llvm.5373085861489949100+0x4bff>
  13962f:	jmp    13992a <litchi_pptx::shape::reader::Scene::read_with+0x161a>
  139634:	mov    $0x29,%edx
  139639:	mov    0x68(%rsp),%rdi
  13963e:	lea    -0x114a77(%rip),%rsi        # 24bce <anon.bc3d94ef2394061926bee58178356194.284.llvm.5373085861489949100+0x4bd6>
  139645:	call   *0xf3605(%rip)        # 22cc50 <bcmp@GLIBC_2.2.5>
  13964b:	test   %eax,%eax
  13964d:	je     1397b9 <litchi_pptx::shape::reader::Scene::read_with+0x14a9>
  139653:	test   %r15,%r15
  139656:	je     1396a3 <litchi_pptx::shape::reader::Scene::read_with+0x1393>
  139658:	mov    0x1a0(%rsp),%rax
  139660:	shl    $0x8,%r15
  139664:	cmp    %r13,-0x20(%rax,%r15,1)
  139669:	jne    1396a3 <litchi_pptx::shape::reader::Scene::read_with+0x1393>
  13966b:	add    %r15,%rax
  13966e:	mov    -0x18(%rax),%edx
  139671:	lea    0x20(%rsp),%rdi
  139676:	lea    0x148(%rsp),%rsi
  13967e:	mov    0x70(%rsp),%rcx
  139683:	call   13b820 <litchi_pptx::shape::reader::Scanner::finish_shape>
  139688:	movzbl 0x20(%rsp),%edx
  13968d:	cmp    $0x27,%dl
  139690:	jne    13a250 <litchi_pptx::shape::reader::Scene::read_with+0x1f40>
  139696:	mov    0x210(%rsp),%r13
  13969e:	mov    0x8(%rsp),%r12
  1396a3:	cmpl   $0x1,0x158(%rsp)
  1396ab:	jne    1396c3 <litchi_pptx::shape::reader::Scene::read_with+0x13b3>
  1396ad:	cmp    %r13,0x160(%rsp)
  1396b5:	jne    1396c3 <litchi_pptx::shape::reader::Scene::read_with+0x13b3>
  1396b7:	movq   $0x0,0x158(%rsp)
  1396c3:	cmpl   $0x1,0x148(%rsp)
  1396cb:	movabs $0x8000000000000001,%r15
  1396d5:	jne    1396ed <litchi_pptx::shape::reader::Scene::read_with+0x13dd>
  1396d7:	cmp    %r13,0x150(%rsp)
  1396df:	jne    1396ed <litchi_pptx::shape::reader::Scene::read_with+0x13dd>
  1396e1:	movq   $0x0,0x148(%rsp)
  1396ed:	test   %r13,%r13
  1396f0:	je     139f75 <litchi_pptx::shape::reader::Scene::read_with+0x1c65>
  1396f6:	dec    %r13
  1396f9:	mov    %r13,0x210(%rsp)
  139701:	mov    0x1c0(%rsp),%rax
  139709:	test   %rax,%rax
  13970c:	je     13a094 <litchi_pptx::shape::reader::Scene::read_with+0x1d84>
  139712:	dec    %rax
  139715:	mov    %rax,0x1c0(%rsp)
  13971d:	movabs $0x8000000000000002,%rax
  139727:	cmp    %rax,%r12
  13972a:	mov    0xe8(%rsp),%r13
  139732:	jl     139744 <litchi_pptx::shape::reader::Scene::read_with+0x1434>
  139734:	test   %r12,%r12
  139737:	je     139744 <litchi_pptx::shape::reader::Scene::read_with+0x1434>
  139739:	mov    0x68(%rsp),%rdi
  13973e:	call   *0xf3314(%rip)        # 22ca58 <free@GLIBC_2.2.5>
  139744:	mov    0xe0(%rsp),%rax
  13974c:	shl    $1,%rax
  13974f:	test   %rax,%rax
  139752:	je     13975d <litchi_pptx::shape::reader::Scene::read_with+0x144d>
  139754:	mov    %rbx,%rdi
  139757:	call   *0xf32fb(%rip)        # 22ca58 <free@GLIBC_2.2.5>
  13975d:	mov    0x80(%rsp),%rax
  139765:	cmp    $0x5,%rax
  139769:	jb     1387b0 <litchi_pptx::shape::reader::Scene::read_with+0x4a0>
  13976f:	cmp    $0x9,%rax
  139773:	je     1387b0 <litchi_pptx::shape::reader::Scene::read_with+0x4a0>
  139779:	add    $0xfffffffffffffffb,%rax
  13977d:	cmp    $0x3,%rax
  139781:	ja     1387b0 <litchi_pptx::shape::reader::Scene::read_with+0x4a0>
  139787:	lea    -0x11545a(%rip),%rcx        # 24334 <anon.bc3d94ef2394061926bee58178356194.284.llvm.5373085861489949100+0x433c>
  13978e:	movslq (%rcx,%rax,4),%rax
  139792:	add    %rcx,%rax
  139795:	jmp    *%rax
  139797:	cmpq   $0x0,0x88(%rsp)
  1397a0:	jle    1387b0 <litchi_pptx::shape::reader::Scene::read_with+0x4a0>
  1397a6:	mov    0x90(%rsp),%rdi
  1397ae:	call   *0xf32a4(%rip)        # 22ca58 <free@GLIBC_2.2.5>
  1397b4:	jmp    1387b0 <litchi_pptx::shape::reader::Scene::read_with+0x4a0>
  1397b9:	test   %r15,%r15
  1397bc:	je     1396a3 <litchi_pptx::shape::reader::Scene::read_with+0x1393>
  1397c2:	mov    0x1a0(%rsp),%rax
  1397ca:	mov    %r15,%rcx
  1397cd:	shl    $0x8,%rcx
  1397d1:	cmpl   $0x1,-0xc0(%rax,%rcx,1)
  1397d9:	jne    139658 <litchi_pptx::shape::reader::Scene::read_with+0x1348>
  1397df:	add    %rcx,%rax
  1397e2:	cmp    %r13,-0xb8(%rax)
  1397e9:	jne    139658 <litchi_pptx::shape::reader::Scene::read_with+0x1348>
  1397ef:	movq   $0x0,-0xc0(%rax)
  1397fa:	jmp    139658 <litchi_pptx::shape::reader::Scene::read_with+0x1348>
  1397ff:	cmp    $0x6,%r12
  139803:	jne    13986a <litchi_pptx::shape::reader::Scene::read_with+0x155a>
  139805:	mov    (%rdx),%eax
  139807:	mov    $0x4c747865,%ecx
  13980c:	xor    %ecx,%eax
  13980e:	movzwl 0x4(%rdx),%ecx
  139812:	xor    $0x7473,%ecx
  139818:	or     %eax,%ecx
  13981a:	sete   %al
  13981d:	movabs $0x8000000000000001,%rcx
  139827:	mov    0x8(%rsp),%r12
  13982c:	cmp    %rcx,%r12
  13982f:	sete   %cl
  139832:	and    %al,%cl
  139834:	cmp    $0x1,%cl
  139837:	jne    139658 <litchi_pptx::shape::reader::Scene::read_with+0x1348>
  13983d:	cmp    $0x2e,%rbp
  139841:	je     13999b <litchi_pptx::shape::reader::Scene::read_with+0x168b>
  139847:	cmp    $0x3a,%rbp
  13984b:	jne    139658 <litchi_pptx::shape::reader::Scene::read_with+0x1348>
  139851:	mov    $0x3a,%edx
  139856:	mov    0x68(%rsp),%rdi
  13985b:	mov    %rsi,%rbp
  13985e:	lea    -0x114c6e(%rip),%rsi        # 24bf7 <anon.bc3d94ef2394061926bee58178356194.284.llvm.5373085861489949100+0x4bff>
  139865:	jmp    1399af <litchi_pptx::shape::reader::Scene::read_with+0x169f>
  13986a:	cmp    $0x1,%r12
  13986e:	je     13919e <litchi_pptx::shape::reader::Scene::read_with+0xe8e>
  139874:	cmp    $0x2,%r12
  139878:	mov    0x8(%rsp),%r12
  13987d:	jne    139658 <litchi_pptx::shape::reader::Scene::read_with+0x1348>
  139883:	cmpw   $0x6870,(%rdx)
  139888:	sete   %al
  13988b:	movabs $0x8000000000000001,%rcx
  139895:	cmp    %rcx,%r12
  139898:	sete   %cl
  13989b:	and    %al,%cl
  13989d:	cmp    $0x1,%cl
  1398a0:	jne    139658 <litchi_pptx::shape::reader::Scene::read_with+0x1348>
  1398a6:	cmp    $0x2e,%rbp
  1398aa:	je     1399ea <litchi_pptx::shape::reader::Scene::read_with+0x16da>
  1398b0:	cmp    $0x3a,%rbp
  1398b4:	jne    139658 <litchi_pptx::shape::reader::Scene::read_with+0x1348>
  1398ba:	mov    $0x3a,%edx
  1398bf:	mov    0x68(%rsp),%rdi
  1398c4:	mov    %rsi,%rbp
  1398c7:	lea    -0x114cd7(%rip),%rsi        # 24bf7 <anon.bc3d94ef2394061926bee58178356194.284.llvm.5373085861489949100+0x4bff>
  1398ce:	jmp    1399fe <litchi_pptx::shape::reader::Scene::read_with+0x16ee>
  1398d3:	lea    0x20(%rsp),%rdi
  1398d8:	call   13ff20 <litchi_pptx::shape::reader::position::{{closure}}>
  1398dd:	movzbl 0x20(%rsp),%eax
  1398e2:	mov    0x28(%rsp),%rbx
  1398e7:	cmp    $0x27,%al
  1398e9:	je     1387c3 <litchi_pptx::shape::reader::Scene::read_with+0x4b3>
  1398ef:	jmp    13a9a7 <litchi_pptx::shape::reader::Scene::read_with+0x2697>
  1398f4:	lea    0x20(%rsp),%rdi
  1398f9:	call   13ff20 <litchi_pptx::shape::reader::position::{{closure}}>
  1398fe:	movzbl 0x20(%rsp),%edx
  139903:	mov    0x28(%rsp),%rbp
  139908:	cmp    $0x27,%dl
  13990b:	je     138968 <litchi_pptx::shape::reader::Scene::read_with+0x658>
  139911:	jmp    13a9f9 <litchi_pptx::shape::reader::Scene::read_with+0x26e9>
  139916:	mov    $0x2e,%edx
  13991b:	mov    0x68(%rsp),%rdi
  139920:	mov    %rsi,%rbp
  139923:	lea    -0x114cf9(%rip),%rsi        # 24c31 <anon.bc3d94ef2394061926bee58178356194.284.llvm.5373085861489949100+0x4c39>
  13992a:	call   *0xf3320(%rip)        # 22cc50 <bcmp@GLIBC_2.2.5>
  139930:	mov    %rbp,%rcx
  139933:	test   %eax,%eax
  139935:	jne    139658 <litchi_pptx::shape::reader::Scene::read_with+0x1348>
  13993b:	cmpb   $0x0,-0x90(%rcx)
  139942:	je     139658 <litchi_pptx::shape::reader::Scene::read_with+0x1348>
  139948:	cmp    %r13,-0x88(%rcx)
  13994f:	jne    139658 <litchi_pptx::shape::reader::Scene::read_with+0x1348>
  139955:	cmpb   $0x1,-0x8(%rcx)
  139959:	jne    13aadd <litchi_pptx::shape::reader::Scene::read_with+0x27cd>
  13995f:	mov    %rcx,%rbp
  139962:	mov    -0x38(%rcx),%rax
  139966:	movabs $0x8000000000000001,%rcx
  139970:	lea    -0x1(%rcx),%r12
  139974:	cmp    %r12,%rax
  139977:	jne    139ab8 <litchi_pptx::shape::reader::Scene::read_with+0x17a8>
  13997d:	movq   $0x0,-0x90(%rbp)
  139988:	movb   $0x0,-0x8(%rbp)
  13998c:	movw   $0x0,-0x10(%rbp)
  139992:	movb   $0x0,-0xe(%rbp)
  139996:	jmp    139b1e <litchi_pptx::shape::reader::Scene::read_with+0x180e>
  13999b:	mov    $0x2e,%edx
  1399a0:	mov    0x68(%rsp),%rdi
  1399a5:	mov    %rsi,%rbp
  1399a8:	lea    -0x114d7e(%rip),%rsi        # 24c31 <anon.bc3d94ef2394061926bee58178356194.284.llvm.5373085861489949100+0x4c39>
  1399af:	call   *0xf329b(%rip)        # 22cc50 <bcmp@GLIBC_2.2.5>
  1399b5:	mov    %rbp,%rcx
  1399b8:	test   %eax,%eax
  1399ba:	jne    139658 <litchi_pptx::shape::reader::Scene::read_with+0x1348>
  1399c0:	cmpb   $0x0,-0xa0(%rcx)
  1399c7:	je     139658 <litchi_pptx::shape::reader::Scene::read_with+0x1348>
  1399cd:	cmp    %r13,-0x98(%rcx)
  1399d4:	jne    139658 <litchi_pptx::shape::reader::Scene::read_with+0x1348>
  1399da:	movq   $0x0,-0xa0(%rcx)
  1399e5:	jmp    139658 <litchi_pptx::shape::reader::Scene::read_with+0x1348>
  1399ea:	mov    $0x2e,%edx
  1399ef:	mov    0x68(%rsp),%rdi
  1399f4:	mov    %rsi,%rbp
  1399f7:	lea    -0x114dcd(%rip),%rsi        # 24c31 <anon.bc3d94ef2394061926bee58178356194.284.llvm.5373085861489949100+0x4c39>
  1399fe:	call   *0xf324c(%rip)        # 22cc50 <bcmp@GLIBC_2.2.5>
  139a04:	mov    %rbp,%rcx
  139a07:	test   %eax,%eax
  139a09:	jne    139658 <litchi_pptx::shape::reader::Scene::read_with+0x1348>
  139a0f:	cmpl   $0x1,-0xb0(%rcx)
  139a16:	jne    139658 <litchi_pptx::shape::reader::Scene::read_with+0x1348>
  139a1c:	cmp    %r13,-0xa8(%rcx)
  139a23:	jne    139658 <litchi_pptx::shape::reader::Scene::read_with+0x1348>
  139a29:	mov    %rcx,%r12
  139a2c:	cmpb   $0x0,-0xd(%rcx)
  139a30:	je     139a56 <litchi_pptx::shape::reader::Scene::read_with+0x1746>
  139a32:	cmpb   $0x0,-0xc(%r12)
  139a38:	je     13ab3b <litchi_pptx::shape::reader::Scene::read_with+0x282b>
  139a3e:	cmpb   $0x1,-0xb(%r12)
  139a44:	jne    13ab3b <litchi_pptx::shape::reader::Scene::read_with+0x282b>
  139a4a:	cmpb   $0x2,-0xa(%r12)
  139a50:	je     13ab3b <litchi_pptx::shape::reader::Scene::read_with+0x282b>
  139a56:	movq   $0x0,-0xb0(%r12)
  139a62:	movq   $0x0,-0xa0(%r12)
  139a6e:	movq   $0x0,-0x90(%r12)
  139a7a:	movb   $0x0,-0x8(%r12)
  139a80:	movl   $0x0,-0x11(%r12)
  139a89:	cmpq   $0x0,-0x38(%r12)
  139a8f:	jle    139a9c <litchi_pptx::shape::reader::Scene::read_with+0x178c>
  139a91:	mov    -0x30(%r12),%rdi
  139a96:	call   *0xf2fbc(%rip)        # 22ca58 <free@GLIBC_2.2.5>
  139a9c:	movabs $0x8000000000000001,%rax
  139aa6:	dec    %rax
  139aa9:	mov    %rax,-0x38(%r12)
  139aae:	mov    0x8(%rsp),%r12
  139ab3:	jmp    139658 <litchi_pptx::shape::reader::Scene::read_with+0x1348>
  139ab8:	cmpq   $0x1e,-0x28(%rbp)
  139abd:	jne    139af6 <litchi_pptx::shape::reader::Scene::read_with+0x17e6>
  139abf:	mov    -0x30(%rbp),%rcx
  139ac3:	movdqu (%rcx),%xmm0
  139ac7:	pcmpeqb -0x12398f(%rip),%xmm0        # 16140 <anon.58baeeae7d5fe476157d296c9f08f803.346.llvm.5844774155465566772+0x1c0>
  139acf:	movdqu 0xe(%rcx),%xmm1
  139ad4:	pcmpeqb -0x1233bc(%rip),%xmm1        # 16720 <anon.58baeeae7d5fe476157d296c9f08f803.346.llvm.5844774155465566772+0x7a0>
  139adc:	pand   %xmm0,%xmm1
  139ae0:	pmovmskb %xmm1,%ecx
  139ae4:	cmp    $0xffff,%ecx
  139aea:	jne    139af6 <litchi_pptx::shape::reader::Scene::read_with+0x17e6>
  139aec:	cmpb   $0x0,-0xe(%rbp)
  139af0:	je     13ab8a <litchi_pptx::shape::reader::Scene::read_with+0x287a>
  139af6:	movq   $0x0,-0x90(%rbp)
  139b01:	movb   $0x0,-0x8(%rbp)
  139b05:	movw   $0x0,-0x10(%rbp)
  139b0b:	movb   $0x0,-0xe(%rbp)
  139b0f:	test   %rax,%rax
  139b12:	je     139b1e <litchi_pptx::shape::reader::Scene::read_with+0x180e>
  139b14:	mov    -0x30(%rbp),%rdi
  139b18:	call   *0xf2f3a(%rip)        # 22ca58 <free@GLIBC_2.2.5>
  139b1e:	mov    %r12,-0x38(%rbp)
  139b22:	mov    0x8(%rsp),%r12
  139b27:	jmp    139658 <litchi_pptx::shape::reader::Scene::read_with+0x1348>
  139b2c:	mov    0x130(%rsp),%rcx
  139b34:	mov    %rcx,0xd8(%rsp)
  139b3c:	movdqa 0x110(%rsp),%xmm0
  139b45:	movdqa 0x120(%rsp),%xmm1
  139b4e:	movdqu %xmm1,0xc8(%rsp)
  139b57:	movdqu %xmm0,0xb8(%rsp)
  139b60:	mov    %rax,0xb0(%rsp)
  139b68:	lea    0x20(%rsp),%rdi
  139b6d:	lea    0xb0(%rsp),%rsi
  139b75:	call   *0xf3d05(%rip)        # 22d880 <_DYNAMIC+0x1020>
  139b7b:	movzbl 0x20(%rsp),%edx
  139b80:	mov    0x21(%rsp),%eax
  139b84:	mov    %eax,(%rsp)
  139b87:	mov    0x24(%rsp),%eax
  139b8b:	mov    %eax,0x3(%rsp)
  139b8f:	mov    0x28(%rsp),%rbp
  139b94:	mov    0x30(%rsp),%r13
  139b99:	movups 0x38(%rsp),%xmm0
  139b9e:	movaps %xmm0,0x10(%rsp)
  139ba3:	movups 0x48(%rsp),%xmm0
  139ba8:	movaps %xmm0,0xf0(%rsp)
  139bb0:	movdqu 0x58(%rsp),%xmm0
  139bb6:	movdqa %xmm0,0x100(%rsp)
  139bbf:	jmp    13a705 <litchi_pptx::shape::reader::Scene::read_with+0x23f5>
  139bc4:	mov    $0x3e,%ebp
  139bc9:	mov    $0x3e,%edi
  139bce:	call   *0xf2e74(%rip)        # 22ca48 <malloc@GLIBC_2.2.5>
  139bd4:	test   %rax,%rax
  139bd7:	je     13ac0c <litchi_pptx::shape::reader::Scene::read_with+0x28fc>
  139bdd:	mov    %rax,%r13
  139be0:	movups -0x1120de(%rip),%xmm0        # 27b09 <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13171701955629477448+0x7d1>
  139be7:	movups %xmm0,0x2e(%rax)
  139beb:	movups -0x1120f7(%rip),%xmm0        # 27afb <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13171701955629477448+0x7c3>
  139bf2:	movups %xmm0,0x20(%rax)
  139bf6:	movups -0x112112(%rip),%xmm0        # 27aeb <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13171701955629477448+0x7b3>
  139bfd:	movups %xmm0,0x10(%rax)
  139c01:	movups -0x11212d(%rip),%xmm0        # 27adb <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13171701955629477448+0x7a3>
  139c08:	movups %xmm0,(%rax)
  139c0b:	movq   %rbp,%xmm0
  139c10:	movdqa %xmm0,0x10(%rsp)
  139c16:	mov    $0x8,%dl
  139c18:	cmp    $0x9,%r15
  139c1c:	ja     13a6ce <litchi_pptx::shape::reader::Scene::read_with+0x23be>
  139c22:	lea    -0x115a21(%rip),%rcx        # 24208 <anon.bc3d94ef2394061926bee58178356194.284.llvm.5373085861489949100+0x4210>
  139c29:	movslq (%rcx,%r15,4),%rax
  139c2d:	add    %rcx,%rax
  139c30:	jmp    *%rax
  139c32:	mov    $0x3e,%r15d
  139c38:	mov    $0x3e,%edi
  139c3d:	call   *0xf2e05(%rip)        # 22ca48 <malloc@GLIBC_2.2.5>
  139c43:	test   %rax,%rax
  139c46:	je     13abec <litchi_pptx::shape::reader::Scene::read_with+0x28dc>
  139c4c:	mov    %rax,%r13
  139c4f:	movups -0x1121bb(%rip),%xmm0        # 27a9b <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13171701955629477448+0x763>
  139c56:	movups %xmm0,0x2e(%rax)
  139c5a:	movups -0x1121d4(%rip),%xmm0        # 27a8d <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13171701955629477448+0x755>
  139c61:	movups %xmm0,0x20(%rax)
  139c65:	movups -0x1121ef(%rip),%xmm0        # 27a7d <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13171701955629477448+0x745>
  139c6c:	movups %xmm0,0x10(%rax)
  139c70:	movups -0x11220a(%rip),%xmm0        # 27a6d <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13171701955629477448+0x735>
  139c77:	movups %xmm0,(%rax)
  139c7a:	mov    $0x3e,%ebp
  139c7f:	movd   %ebp,%xmm0
  139c83:	movdqa %xmm0,0x10(%rsp)
  139c89:	mov    $0x8,%dl
  139c8b:	test   %r12,%r12
  139c8e:	jle    139c9d <litchi_pptx::shape::reader::Scene::read_with+0x198d>
  139c90:	mov    %rbx,%rdi
  139c93:	mov    %edx,%ebx
  139c95:	call   *0xf2dbd(%rip)        # 22ca58 <free@GLIBC_2.2.5>
  139c9b:	mov    %ebx,%edx
  139c9d:	mov    0x80(%rsp),%r15
  139ca5:	cmp    $0x9,%r15
  139ca9:	ja     13a6ce <litchi_pptx::shape::reader::Scene::read_with+0x23be>
  139caf:	lea    -0x115a86(%rip),%rcx        # 24230 <anon.bc3d94ef2394061926bee58178356194.284.llvm.5373085861489949100+0x4238>
  139cb6:	movslq (%rcx,%r15,4),%rax
  139cba:	add    %rcx,%rax
  139cbd:	jmp    *%rax
  139cbf:	mov    $0x3e,%ebx
  139cc4:	mov    $0x3e,%edi
  139cc9:	call   *0xf2d79(%rip)        # 22ca48 <malloc@GLIBC_2.2.5>
  139ccf:	test   %rax,%rax
  139cd2:	je     13abfc <litchi_pptx::shape::reader::Scene::read_with+0x28ec>
  139cd8:	mov    %rax,%r13
  139cdb:	movups -0x112247(%rip),%xmm0        # 27a9b <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13171701955629477448+0x763>
  139ce2:	movups %xmm0,0x2e(%rax)
  139ce6:	movups -0x112260(%rip),%xmm0        # 27a8d <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13171701955629477448+0x755>
  139ced:	movups %xmm0,0x20(%rax)
  139cf1:	movups -0x11227b(%rip),%xmm0        # 27a7d <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13171701955629477448+0x745>
  139cf8:	movups %xmm0,0x10(%rax)
  139cfc:	movups -0x112296(%rip),%xmm0        # 27a6d <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13171701955629477448+0x735>
  139d03:	movups %xmm0,(%rax)
  139d06:	mov    $0x3e,%ebp
  139d0b:	movd   %ebp,%xmm0
  139d0f:	movdqa %xmm0,0x10(%rsp)
  139d15:	mov    $0x8,%dl
  139d17:	cmpq   $0x0,0x110(%rsp)
  139d20:	jle    139d34 <litchi_pptx::shape::reader::Scene::read_with+0x1a24>
  139d22:	mov    0x118(%rsp),%rdi
  139d2a:	mov    %edx,%ebx
  139d2c:	call   *0xf2d26(%rip)        # 22ca58 <free@GLIBC_2.2.5>
  139d32:	mov    %ebx,%edx
  139d34:	mov    0x80(%rsp),%r15
  139d3c:	cmp    $0x9,%r15
  139d40:	ja     13a6ce <litchi_pptx::shape::reader::Scene::read_with+0x23be>
  139d46:	lea    -0x115b6d(%rip),%rcx        # 241e0 <anon.bc3d94ef2394061926bee58178356194.284.llvm.5373085861489949100+0x41e8>
  139d4d:	movslq (%rcx,%r15,4),%rax
  139d51:	add    %rcx,%rax
  139d54:	jmp    *%rax
  139d56:	movdqu 0x20(%rsp),%xmm0
  139d5c:	movdqu 0x30(%rsp),%xmm1
  139d62:	movdqu %xmm1,0xc8(%rsp)
  139d6b:	movdqu %xmm0,0xb8(%rsp)
  139d74:	movabs $0x8000000000000001,%rax
  139d7e:	add    $0xd,%rax
  139d82:	mov    %rax,0xb0(%rsp)
  139d8a:	lea    0x20(%rsp),%rdi
  139d8f:	lea    0xb0(%rsp),%rsi
  139d97:	call   *0xf3ae3(%rip)        # 22d880 <_DYNAMIC+0x1020>
  139d9d:	jmp    139de6 <litchi_pptx::shape::reader::Scene::read_with+0x1ad6>
  139d9f:	movdqu 0x20(%rsp),%xmm0
  139da5:	movdqu 0x30(%rsp),%xmm1
  139dab:	movdqu %xmm1,0xc8(%rsp)
  139db4:	movdqu %xmm0,0xb8(%rsp)
  139dbd:	movabs $0x8000000000000001,%rax
  139dc7:	add    $0xd,%rax
  139dcb:	mov    %rax,0xb0(%rsp)
  139dd3:	lea    0x20(%rsp),%rdi
  139dd8:	lea    0xb0(%rsp),%rsi
  139de0:	call   *0xf3a9a(%rip)        # 22d880 <_DYNAMIC+0x1020>
  139de6:	movzbl 0x20(%rsp),%edx
  139deb:	mov    0x21(%rsp),%eax
  139def:	mov    %eax,(%rsp)
  139df2:	mov    0x24(%rsp),%eax
  139df6:	mov    %eax,0x3(%rsp)
  139dfa:	mov    0x28(%rsp),%rbp
  139dff:	mov    0x30(%rsp),%r13
  139e04:	movups 0x38(%rsp),%xmm0
  139e09:	movaps %xmm0,0x10(%rsp)
  139e0e:	movups 0x48(%rsp),%xmm0
  139e13:	movaps %xmm0,0xf0(%rsp)
  139e1b:	movdqu 0x58(%rsp),%xmm0
  139e21:	movdqa %xmm0,0x100(%rsp)
  139e2a:	jmp    13a6e8 <litchi_pptx::shape::reader::Scene::read_with+0x23d8>
  139e2f:	cmpq   $0x0,0x210(%rsp)
  139e38:	jne    139e54 <litchi_pptx::shape::reader::Scene::read_with+0x1b44>
  139e3a:	cmpq   $0x0,0x1a8(%rsp)
  139e43:	jne    139e54 <litchi_pptx::shape::reader::Scene::read_with+0x1b44>
  139e45:	cmpq   $0x0,0x1c0(%rsp)
  139e4e:	je     13a54c <litchi_pptx::shape::reader::Scene::read_with+0x223c>
  139e54:	mov    $0x26,%ebx
  139e59:	mov    $0x26,%edi
  139e5e:	call   *0xf2be4(%rip)        # 22ca48 <malloc@GLIBC_2.2.5>
  139e64:	test   %rax,%rax
  139e67:	je     13ac1e <litchi_pptx::shape::reader::Scene::read_with+0x290e>
  139e6d:	mov    %rax,%r13
  139e70:	movups -0x11234e(%rip),%xmm0        # 27b29 <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13171701955629477448+0x7f1>
  139e77:	movups %xmm0,0x10(%rax)
  139e7b:	movups -0x112369(%rip),%xmm0        # 27b19 <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13171701955629477448+0x7e1>
  139e82:	movups %xmm0,(%rax)
  139e85:	movabs $0x73746e656d656c65,%rax
  139e8f:	mov    %rax,0x1e(%r13)
  139e93:	movq   %rbx,%xmm0
  139e98:	movdqa %xmm0,0x10(%rsp)
  139e9e:	movb   $0x8,0x78(%rsp)
  139ea3:	cmpq   $0x0,0x228(%rsp)
  139eac:	jne    13a717 <litchi_pptx::shape::reader::Scene::read_with+0x2407>
  139eb2:	jmp    13a725 <litchi_pptx::shape::reader::Scene::read_with+0x2415>
  139eb7:	mov    0x21(%rsp),%eax
  139ebb:	mov    0x24(%rsp),%ecx
  139ebf:	mov    %ecx,0x3(%rsp)
  139ec3:	mov    %eax,(%rsp)
  139ec6:	jmp    13a182 <litchi_pptx::shape::reader::Scene::read_with+0x1e72>
  139ecb:	mov    $0x27,%r15d
  139ed1:	mov    $0x27,%edi
  139ed6:	call   *0xf2b6c(%rip)        # 22ca48 <malloc@GLIBC_2.2.5>
  139edc:	test   %rax,%rax
  139edf:	mov    0x8(%rsp),%r12
  139ee4:	je     13abdc <litchi_pptx::shape::reader::Scene::read_with+0x28cc>
  139eea:	mov    %rax,%r13
  139eed:	movups -0x112a39(%rip),%xmm0        # 274bb <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13171701955629477448+0x183>
  139ef4:	movups %xmm0,0x10(%rax)
  139ef8:	movups -0x112a54(%rip),%xmm0        # 274ab <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13171701955629477448+0x173>
  139eff:	movups %xmm0,(%rax)
  139f02:	movabs $0x67617420646e6520,%rax
  139f0c:	mov    %rax,0x1f(%r13)
  139f10:	mov    $0x27,%ebp
  139f15:	jmp    139fac <litchi_pptx::shape::reader::Scene::read_with+0x1c9c>
  139f1a:	lea    -0x112a17(%rip),%rax        # 2750a <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13171701955629477448+0x1d2>
  139f21:	mov    %r13,%rbp
  139f24:	mov    %rax,%r13
  139f27:	movdqa 0x300(%rsp),%xmm0
  139f30:	movdqa %xmm0,0x10(%rsp)
  139f36:	mov    $0x9,%dl
  139f38:	cmpq   $0x0,0x8(%rsp)
  139f3e:	jg     13a616 <litchi_pptx::shape::reader::Scene::read_with+0x2306>
  139f44:	jmp    13a628 <litchi_pptx::shape::reader::Scene::read_with+0x2318>
  139f49:	lea    -0x112a46(%rip),%rax        # 2750a <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13171701955629477448+0x1d2>
  139f50:	mov    %r13,%rbp
  139f53:	mov    %rax,%r13
  139f56:	movdqa 0x300(%rsp),%xmm0
  139f5f:	movdqa %xmm0,0x10(%rsp)
  139f65:	mov    $0x9,%dl
  139f67:	test   %r15,%r15
  139f6a:	jg     13a684 <litchi_pptx::shape::reader::Scene::read_with+0x2374>
  139f70:	jmp    13a693 <litchi_pptx::shape::reader::Scene::read_with+0x2383>
  139f75:	mov    $0x19,%r15d
  139f7b:	mov    $0x19,%edi
  139f80:	call   *0xf2ac2(%rip)        # 22ca48 <malloc@GLIBC_2.2.5>
  139f86:	test   %rax,%rax
  139f89:	je     13abdc <litchi_pptx::shape::reader::Scene::read_with+0x28cc>
  139f8f:	mov    %rax,%r13
  139f92:	movups -0x112a9f(%rip),%xmm0        # 274fa <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13171701955629477448+0x1c2>
  139f99:	movups %xmm0,0x9(%rax)
  139f9d:	movups -0x112ab3(%rip),%xmm0        # 274f1 <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13171701955629477448+0x1b9>
  139fa4:	movups %xmm0,(%rax)
  139fa7:	mov    $0x19,%ebp
  139fac:	movq   %rbp,%xmm0
  139fb1:	movdqa %xmm0,0x10(%rsp)
  139fb7:	mov    $0x8,%dl
  139fb9:	mov    0x110(%rsp),%eax
  139fc0:	mov    0x113(%rsp),%ecx
  139fc7:	mov    %ecx,0x3(%rsp)
  139fcb:	mov    %eax,(%rsp)
  139fce:	movdqa 0xb0(%rsp),%xmm0
  139fd7:	movdqa 0xc0(%rsp),%xmm1
  139fe0:	movdqa %xmm0,0xf0(%rsp)
  139fe9:	movdqa %xmm1,0x100(%rsp)
  139ff2:	jmp    13a0e3 <litchi_pptx::shape::reader::Scene::read_with+0x1dd3>
  139ff7:	mov    0x21(%rsp),%eax
  139ffb:	mov    0x24(%rsp),%ecx
  139fff:	mov    %ecx,0x3(%rsp)
  13a003:	mov    %eax,(%rsp)
  13a006:	mov    0x28(%rsp),%rbp
  13a00b:	mov    0x30(%rsp),%r13
  13a010:	movups 0x38(%rsp),%xmm0
  13a015:	movaps %xmm0,0x10(%rsp)
  13a01a:	movups 0x48(%rsp),%xmm0
  13a01f:	movaps %xmm0,0xf0(%rsp)
  13a027:	movdqu 0x58(%rsp),%xmm0
  13a02d:	movdqa %xmm0,0x100(%rsp)
  13a036:	cmpq   $0x0,0x8(%rsp)
  13a03c:	jg     13a616 <litchi_pptx::shape::reader::Scene::read_with+0x2306>
  13a042:	jmp    13a628 <litchi_pptx::shape::reader::Scene::read_with+0x2318>
  13a047:	mov    0x21(%rsp),%eax
  13a04b:	mov    0x24(%rsp),%ecx
  13a04f:	mov    %ecx,0x3(%rsp)
  13a053:	mov    %eax,(%rsp)
  13a056:	mov    0x28(%rsp),%rbp
  13a05b:	mov    0x30(%rsp),%r13
  13a060:	movups 0x38(%rsp),%xmm0
  13a065:	movaps %xmm0,0x10(%rsp)
  13a06a:	movups 0x48(%rsp),%xmm0
  13a06f:	movaps %xmm0,0xf0(%rsp)
  13a077:	movdqu 0x58(%rsp),%xmm0
  13a07d:	movdqa %xmm0,0x100(%rsp)
  13a086:	test   %r15,%r15
  13a089:	jg     13a684 <litchi_pptx::shape::reader::Scene::read_with+0x2374>
  13a08f:	jmp    13a693 <litchi_pptx::shape::reader::Scene::read_with+0x2383>
  13a094:	mov    $0x2b,%r15d
  13a09a:	mov    $0x2b,%edi
  13a09f:	call   *0xf29a3(%rip)        # 22ca48 <malloc@GLIBC_2.2.5>
  13a0a5:	test   %rax,%rax
  13a0a8:	je     13abdc <litchi_pptx::shape::reader::Scene::read_with+0x28cc>
  13a0ae:	mov    %rax,%r13
  13a0b1:	movups -0x11255e(%rip),%xmm0        # 27b5a <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13171701955629477448+0x822>
  13a0b8:	movups %xmm0,0x1b(%rax)
  13a0bc:	movups -0x112574(%rip),%xmm0        # 27b4f <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13171701955629477448+0x817>
  13a0c3:	movups %xmm0,0x10(%rax)
  13a0c7:	movups -0x11258f(%rip),%xmm0        # 27b3f <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13171701955629477448+0x807>
  13a0ce:	movups %xmm0,(%rax)
  13a0d1:	mov    $0x2b,%ebp
  13a0d6:	movq   %rbp,%xmm0
  13a0db:	movdqa %xmm0,0x10(%rsp)
  13a0e1:	mov    $0x8,%dl
  13a0e3:	movabs $0x8000000000000002,%rax
  13a0ed:	cmp    %rax,%r12
  13a0f0:	jl     13a108 <litchi_pptx::shape::reader::Scene::read_with+0x1df8>
  13a0f2:	test   %r12,%r12
  13a0f5:	je     13a108 <litchi_pptx::shape::reader::Scene::read_with+0x1df8>
  13a0f7:	mov    0x68(%rsp),%rdi
  13a0fc:	mov    %edx,%r15d
  13a0ff:	call   *0xf2953(%rip)        # 22ca58 <free@GLIBC_2.2.5>
  13a105:	mov    %r15d,%edx
  13a108:	mov    0xe0(%rsp),%rax
  13a110:	shl    $1,%rax
  13a113:	test   %rax,%rax
  13a116:	je     13a125 <litchi_pptx::shape::reader::Scene::read_with+0x1e15>
  13a118:	mov    %rbx,%rdi
  13a11b:	mov    %edx,%ebx
  13a11d:	call   *0xf2935(%rip)        # 22ca58 <free@GLIBC_2.2.5>
  13a123:	mov    %ebx,%edx
  13a125:	mov    0x80(%rsp),%r15
  13a12d:	cmp    $0x9,%r15
  13a131:	ja     13a6ce <litchi_pptx::shape::reader::Scene::read_with+0x23be>
  13a137:	lea    -0x115e96(%rip),%rcx        # 242a8 <anon.bc3d94ef2394061926bee58178356194.284.llvm.5373085861489949100+0x42b0>
  13a13e:	movslq (%rcx,%r15,4),%rax
  13a142:	add    %rcx,%rax
  13a145:	jmp    *%rax
  13a147:	lea    0xb8(%rsp),%rax
  13a14f:	movdqu (%rax),%xmm0
  13a153:	movdqa %xmm0,0x110(%rsp)
  13a15c:	lea    0x20(%rsp),%rdi
  13a161:	lea    0x110(%rsp),%rsi
  13a169:	call   13fb30 <litchi_pptx::shape::reader::Scanner::scan::{{closure}}>
  13a16e:	movzbl 0x20(%rsp),%edx
  13a173:	mov    0x21(%rsp),%eax
  13a177:	mov    %eax,(%rsp)
  13a17a:	mov    0x24(%rsp),%eax
  13a17e:	mov    %eax,0x3(%rsp)
  13a182:	mov    0x28(%rsp),%rax
  13a187:	mov    %rax,0x70(%rsp)
  13a18c:	mov    0x30(%rsp),%r13
  13a191:	movups 0x38(%rsp),%xmm0
  13a196:	movaps %xmm0,0x10(%rsp)
  13a19b:	movups 0x48(%rsp),%xmm0
  13a1a0:	movaps %xmm0,0xf0(%rsp)
  13a1a8:	movdqu 0x58(%rsp),%xmm0
  13a1ae:	movdqa %xmm0,0x100(%rsp)
  13a1b7:	jmp    13a321 <litchi_pptx::shape::reader::Scene::read_with+0x2011>
  13a1bc:	mov    0x130(%rsp),%rax
  13a1c4:	mov    %rax,0xd0(%rsp)
  13a1cc:	movdqu 0x110(%rsp),%xmm0
  13a1d5:	movdqu 0x120(%rsp),%xmm1
  13a1de:	movdqa %xmm1,0xc0(%rsp)
  13a1e7:	movdqa %xmm0,0xb0(%rsp)
  13a1f0:	lea    0x20(%rsp),%rdi
  13a1f5:	lea    0xb0(%rsp),%rsi
  13a1fd:	call   13fbf0 <litchi_pptx::shape::reader::Scanner::scan::{{closure}}>
  13a202:	movzbl 0x20(%rsp),%edx
  13a207:	mov    0x21(%rsp),%eax
  13a20b:	mov    %eax,(%rsp)
  13a20e:	mov    0x24(%rsp),%eax
  13a212:	mov    %eax,0x3(%rsp)
  13a216:	mov    0x28(%rsp),%rax
  13a21b:	mov    %rax,0x70(%rsp)
  13a220:	mov    0x30(%rsp),%r13
  13a225:	movups 0x38(%rsp),%xmm0
  13a22a:	movaps %xmm0,0x10(%rsp)
  13a22f:	movups 0x48(%rsp),%xmm0
  13a234:	movaps %xmm0,0xf0(%rsp)
  13a23c:	movdqu 0x58(%rsp),%xmm0
  13a242:	movdqa %xmm0,0x100(%rsp)
  13a24b:	jmp    13a30a <litchi_pptx::shape::reader::Scene::read_with+0x1ffa>
  13a250:	mov    0x21(%rsp),%eax
  13a254:	mov    0x24(%rsp),%ecx
  13a258:	mov    %ecx,0x113(%rsp)
  13a25f:	mov    %eax,0x110(%rsp)
  13a266:	mov    0x28(%rsp),%rbp
  13a26b:	mov    0x30(%rsp),%r13
  13a270:	movups 0x38(%rsp),%xmm0
  13a275:	movaps %xmm0,0x10(%rsp)
  13a27a:	movups 0x48(%rsp),%xmm0
  13a27f:	movaps %xmm0,0xb0(%rsp)
  13a287:	movdqu 0x58(%rsp),%xmm0
  13a28d:	movdqa %xmm0,0xc0(%rsp)
  13a296:	mov    0x8(%rsp),%r12
  13a29b:	jmp    139fb9 <litchi_pptx::shape::reader::Scene::read_with+0x1ca9>
  13a2a0:	mov    0x21(%rsp),%eax
  13a2a4:	mov    0x24(%rsp),%ecx
  13a2a8:	mov    %ecx,0x3(%rsp)
  13a2ac:	mov    %eax,(%rsp)
  13a2af:	mov    0x28(%rsp),%rax
  13a2b4:	mov    %rax,0x70(%rsp)
  13a2b9:	mov    0x30(%rsp),%rax
  13a2be:	mov    %rax,0x78(%rsp)
  13a2c3:	movups 0x38(%rsp),%xmm0
  13a2c8:	movaps %xmm0,0x10(%rsp)
  13a2cd:	movups 0x48(%rsp),%xmm0
  13a2d2:	movaps %xmm0,0xf0(%rsp)
  13a2da:	movdqu 0x58(%rsp),%xmm0
  13a2e0:	movdqa %xmm0,0x100(%rsp)
  13a2e9:	shl    $1,%r13
  13a2ec:	test   %r13,%r13
  13a2ef:	je     13a300 <litchi_pptx::shape::reader::Scene::read_with+0x1ff0>
  13a2f1:	mov    %r12,%rdi
  13a2f4:	mov    %edx,%r12d
  13a2f7:	call   *0xf275b(%rip)        # 22ca58 <free@GLIBC_2.2.5>
  13a2fd:	mov    %r12d,%edx
  13a300:	mov    0x78(%rsp),%r13
  13a305:	mov    0x8(%rsp),%r12
  13a30a:	shl    $1,%r15
  13a30d:	test   %r15,%r15
  13a310:	je     13a321 <litchi_pptx::shape::reader::Scene::read_with+0x2011>
  13a312:	mov    %r12,%rdi
  13a315:	mov    %edx,%r15d
  13a318:	call   *0xf273a(%rip)        # 22ca58 <free@GLIBC_2.2.5>
  13a31e:	mov    %r15d,%edx
  13a321:	test   %rbp,%rbp
  13a324:	jle    13a333 <litchi_pptx::shape::reader::Scene::read_with+0x2023>
  13a326:	mov    %rbx,%rdi
  13a329:	mov    %edx,%ebx
  13a32b:	call   *0xf2727(%rip)        # 22ca58 <free@GLIBC_2.2.5>
  13a331:	mov    %ebx,%edx
  13a333:	mov    0x80(%rsp),%r15
  13a33b:	cmp    $0x9,%r15
  13a33f:	mov    0x70(%rsp),%rbp
  13a344:	ja     13a6ce <litchi_pptx::shape::reader::Scene::read_with+0x23be>
  13a34a:	lea    -0x1160f9(%rip),%rcx        # 24258 <anon.bc3d94ef2394061926bee58178356194.284.llvm.5373085861489949100+0x4260>
  13a351:	movslq (%rcx,%r15,4),%rax
  13a355:	add    %rcx,%rax
  13a358:	jmp    *%rax
  13a35a:	movq   %rcx,%xmm0
  13a35f:	movq   %rbx,%xmm1
  13a364:	punpcklqdq %xmm0,%xmm1
  13a368:	movdqa %xmm1,0x10(%rsp)
  13a36e:	mov    $0x23,%dl
  13a370:	mov    %r15,%r13
  13a373:	jmp    139d17 <litchi_pptx::shape::reader::Scene::read_with+0x1a07>
  13a378:	lea    0xb8(%rsp),%rax
  13a380:	movdqu (%rax),%xmm0
  13a384:	movdqa %xmm0,0x110(%rsp)
  13a38d:	lea    0x20(%rsp),%rdi
  13a392:	lea    0x110(%rsp),%rsi
  13a39a:	call   13fb30 <litchi_pptx::shape::reader::Scanner::scan::{{closure}}>
  13a39f:	movzbl 0x20(%rsp),%edx
  13a3a4:	mov    0x21(%rsp),%eax
  13a3a8:	mov    %eax,(%rsp)
  13a3ab:	mov    0x24(%rsp),%eax
  13a3af:	mov    %eax,0x3(%rsp)
  13a3b3:	mov    0x28(%rsp),%rbp
  13a3b8:	mov    0x30(%rsp),%r13
  13a3bd:	movups 0x38(%rsp),%xmm0
  13a3c2:	movaps %xmm0,0x10(%rsp)
  13a3c7:	movups 0x48(%rsp),%xmm0
  13a3cc:	movaps %xmm0,0xf0(%rsp)
  13a3d4:	movdqu 0x58(%rsp),%xmm0
  13a3da:	movdqa %xmm0,0x100(%rsp)
  13a3e3:	jmp    139c8b <litchi_pptx::shape::reader::Scene::read_with+0x197b>
  13a3e8:	mov    0x21(%rsp),%eax
  13a3ec:	mov    0x24(%rsp),%ecx
  13a3f0:	mov    %ecx,0x3(%rsp)
  13a3f4:	mov    %eax,(%rsp)
  13a3f7:	mov    0x28(%rsp),%rax
  13a3fc:	mov    %rax,0x70(%rsp)
  13a401:	mov    0x30(%rsp),%rbp
  13a406:	movups 0x38(%rsp),%xmm0
  13a40b:	movaps %xmm0,0x10(%rsp)
  13a410:	movups 0x48(%rsp),%xmm0
  13a415:	movaps %xmm0,0xf0(%rsp)
  13a41d:	movdqu 0x58(%rsp),%xmm0
  13a423:	movdqa %xmm0,0x100(%rsp)
  13a42c:	shl    $1,%r13
  13a42f:	test   %r13,%r13
  13a432:	je     13a443 <litchi_pptx::shape::reader::Scene::read_with+0x2133>
  13a434:	mov    %r15,%rdi
  13a437:	mov    %edx,%r15d
  13a43a:	call   *0xf2618(%rip)        # 22ca58 <free@GLIBC_2.2.5>
  13a440:	mov    %r15d,%edx
  13a443:	mov    %rbp,%r13
  13a446:	mov    0x70(%rsp),%rbp
  13a44b:	jmp    139c8b <litchi_pptx::shape::reader::Scene::read_with+0x197b>
  13a450:	mov    0x21(%rsp),%eax
  13a454:	mov    0x24(%rsp),%ecx
  13a458:	mov    %ecx,0x3(%rsp)
  13a45c:	mov    %eax,(%rsp)
  13a45f:	mov    0x28(%rsp),%rbp
  13a464:	mov    0x30(%rsp),%r13
  13a469:	movups 0x38(%rsp),%xmm0
  13a46e:	movaps %xmm0,0x10(%rsp)
  13a473:	movups 0x48(%rsp),%xmm0
  13a478:	movaps %xmm0,0xf0(%rsp)
  13a480:	movdqu 0x58(%rsp),%xmm0
  13a486:	movdqa %xmm0,0x100(%rsp)
  13a48f:	test   %r15,%r15
  13a492:	je     139d17 <litchi_pptx::shape::reader::Scene::read_with+0x1a07>
  13a498:	mov    %rbx,%rdi
  13a49b:	mov    %edx,%ebx
  13a49d:	call   *0xf25b5(%rip)        # 22ca58 <free@GLIBC_2.2.5>
  13a4a3:	mov    %ebx,%edx
  13a4a5:	jmp    139d17 <litchi_pptx::shape::reader::Scene::read_with+0x1a07>
  13a4aa:	mov    $0x30,%r15d
  13a4b0:	mov    $0x30,%edi
  13a4b5:	call   *0xf258d(%rip)        # 22ca48 <malloc@GLIBC_2.2.5>
  13a4bb:	test   %rax,%rax
  13a4be:	je     13abec <litchi_pptx::shape::reader::Scene::read_with+0x28dc>
  13a4c4:	mov    %rax,%r13
  13a4c7:	movups -0x112a03(%rip),%xmm0        # 27acb <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13171701955629477448+0x793>
  13a4ce:	movups %xmm0,0x20(%rax)
  13a4d2:	movups -0x112a1e(%rip),%xmm0        # 27abb <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13171701955629477448+0x783>
  13a4d9:	movups %xmm0,0x10(%rax)
  13a4dd:	movups -0x112a39(%rip),%xmm0        # 27aab <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13171701955629477448+0x773>
  13a4e4:	movups %xmm0,(%rax)
  13a4e7:	mov    $0x30,%ebp
  13a4ec:	jmp    139c7f <litchi_pptx::shape::reader::Scene::read_with+0x196f>
  13a4f1:	mov    $0x30,%ebx
  13a4f6:	mov    $0x30,%edi
  13a4fb:	call   *0xf2547(%rip)        # 22ca48 <malloc@GLIBC_2.2.5>
  13a501:	test   %rax,%rax
  13a504:	je     13abfc <litchi_pptx::shape::reader::Scene::read_with+0x28ec>
  13a50a:	mov    %rax,%r13
  13a50d:	movups -0x112a49(%rip),%xmm0        # 27acb <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13171701955629477448+0x793>
  13a514:	movups %xmm0,0x20(%rax)
  13a518:	movups -0x112a64(%rip),%xmm0        # 27abb <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13171701955629477448+0x783>
  13a51f:	movups %xmm0,0x10(%rax)
  13a523:	movups -0x112a7f(%rip),%xmm0        # 27aab <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13171701955629477448+0x773>
  13a52a:	movups %xmm0,(%rax)
  13a52d:	mov    $0x30,%ebp
  13a532:	jmp    139d0b <litchi_pptx::shape::reader::Scene::read_with+0x19fb>
  13a537:	call   *0xf2a2b(%rip)        # 22cf68 <_DYNAMIC+0x708>
  13a53d:	mov    %rdx,%fs:0x8(%rbp)
  13a542:	movb   $0x1,%fs:0x10(%rbp)
  13a547:	jmp    1383ff <litchi_pptx::shape::reader::Scene::read_with+0xef>
  13a54c:	lea    0x198(%rsp),%r15
  13a554:	mov    0x168(%rsp),%rbx
  13a55c:	mov    0x170(%rsp),%rbp
  13a564:	movups 0x178(%rsp),%xmm0
  13a56c:	movaps %xmm0,0x10(%rsp)
  13a571:	lea    0x188(%rsp),%rax
  13a579:	movdqu (%rax),%xmm0
  13a57d:	movdqa %xmm0,0xf0(%rsp)
  13a586:	lea    0x228(%rsp),%rdi
  13a58e:	call   163730 <core::ptr::drop_in_place<quick_xml::name::NamespaceResolver>>
  13a593:	lea    0x288(%rsp),%rdi
  13a59b:	call   164830 <core::ptr::drop_in_place<quick_xml::reader::Reader<&[u8]>>>
  13a5a0:	mov    %r15,%rdi
  13a5a3:	call   164c30 <core::ptr::drop_in_place<alloc::vec::Vec<litchi_pptx::shape::reader::Active>>>
  13a5a8:	movb   $0x27,0x78(%rsp)
  13a5ad:	cmpq   $0x0,0x1b0(%rsp)
  13a5b6:	jne    13a827 <litchi_pptx::shape::reader::Scene::read_with+0x2517>
  13a5bc:	jmp    13a84e <litchi_pptx::shape::reader::Scene::read_with+0x253e>
  13a5c1:	lea    0xebb20(%rip),%rcx        # 2260e8 <anon.bc3d94ef2394061926bee58178356194.1849.llvm.5373085861489949100+0x580>
  13a5c8:	xor    %edi,%edi
  13a5ca:	mov    %r15,%rsi
  13a5cd:	call   *0xf260d(%rip)        # 22cbe0 <_DYNAMIC+0x380>
  13a5d3:	jmp    13ac3e <litchi_pptx::shape::reader::Scene::read_with+0x292e>
  13a5d8:	lea    0xebb09(%rip),%rcx        # 2260e8 <anon.bc3d94ef2394061926bee58178356194.1849.llvm.5373085861489949100+0x580>
  13a5df:	xor    %edi,%edi
  13a5e1:	mov    %rdx,%rsi
  13a5e4:	mov    %rax,%rdx
  13a5e7:	call   *0xf25f3(%rip)        # 22cbe0 <_DYNAMIC+0x380>
  13a5ed:	jmp    13ac3e <litchi_pptx::shape::reader::Scene::read_with+0x292e>
  13a5f2:	mov    $0x12,%eax
  13a5f7:	movq   %rax,%xmm0
  13a5fc:	movdqa %xmm0,0x10(%rsp)
  13a602:	mov    $0x9,%dl
  13a604:	mov    %r12,%rbp
  13a607:	lea    -0x11323e(%rip),%r13        # 273d0 <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13171701955629477448+0x98>
  13a60e:	cmpq   $0x0,0x8(%rsp)
  13a614:	jle    13a628 <litchi_pptx::shape::reader::Scene::read_with+0x2318>
  13a616:	mov    0xe0(%rsp),%rdi
  13a61e:	mov    %edx,%ebx
  13a620:	call   *0xf2432(%rip)        # 22ca58 <free@GLIBC_2.2.5>
  13a626:	mov    %ebx,%edx
  13a628:	cmpq   $0x0,0xb0(%rsp)
  13a631:	jle    13a645 <litchi_pptx::shape::reader::Scene::read_with+0x2335>
  13a633:	mov    0xb8(%rsp),%rdi
  13a63b:	mov    %edx,%ebx
  13a63d:	call   *0xf2415(%rip)        # 22ca58 <free@GLIBC_2.2.5>
  13a643:	mov    %ebx,%edx
  13a645:	mov    0x80(%rsp),%r15
  13a64d:	cmp    $0x9,%r15
  13a651:	ja     13a6ce <litchi_pptx::shape::reader::Scene::read_with+0x23be>
  13a653:	lea    -0x11638a(%rip),%rcx        # 242d0 <anon.bc3d94ef2394061926bee58178356194.284.llvm.5373085861489949100+0x42d8>
  13a65a:	movslq (%rcx,%r15,4),%rax
  13a65e:	add    %rcx,%rax
  13a661:	jmp    *%rax
  13a663:	mov    $0x12,%eax
  13a668:	movq   %rax,%xmm0
  13a66d:	movdqa %xmm0,0x10(%rsp)
  13a673:	mov    $0x9,%dl
  13a675:	mov    %r12,%rbp
  13a678:	lea    -0x1132af(%rip),%r13        # 273d0 <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13171701955629477448+0x98>
  13a67f:	test   %r15,%r15
  13a682:	jle    13a693 <litchi_pptx::shape::reader::Scene::read_with+0x2383>
  13a684:	mov    0x70(%rsp),%rdi
  13a689:	mov    %edx,%ebx
  13a68b:	call   *0xf23c7(%rip)        # 22ca58 <free@GLIBC_2.2.5>
  13a691:	mov    %ebx,%edx
  13a693:	cmpq   $0x0,0xb0(%rsp)
  13a69c:	jle    13a6b0 <litchi_pptx::shape::reader::Scene::read_with+0x23a0>
  13a69e:	mov    0xb8(%rsp),%rdi
  13a6a6:	mov    %edx,%ebx
  13a6a8:	call   *0xf23aa(%rip)        # 22ca58 <free@GLIBC_2.2.5>
  13a6ae:	mov    %ebx,%edx
  13a6b0:	mov    0x80(%rsp),%r15
  13a6b8:	cmp    $0x9,%r15
  13a6bc:	ja     13a6ce <litchi_pptx::shape::reader::Scene::read_with+0x23be>
  13a6be:	lea    -0x116445(%rip),%rcx        # 24280 <anon.bc3d94ef2394061926bee58178356194.284.llvm.5373085861489949100+0x4288>
  13a6c5:	movslq (%rcx,%r15,4),%rax
  13a6c9:	add    %rcx,%rax
  13a6cc:	jmp    *%rax
  13a6ce:	add    $0xfffffffffffffffb,%r15
  13a6d2:	cmp    $0x4,%r15
  13a6d6:	ja     13a705 <litchi_pptx::shape::reader::Scene::read_with+0x23f5>
  13a6d8:	lea    -0x1163e7(%rip),%rax        # 242f8 <anon.bc3d94ef2394061926bee58178356194.284.llvm.5373085861489949100+0x4300>
  13a6df:	movslq (%rax,%r15,4),%rcx
  13a6e3:	add    %rax,%rcx
  13a6e6:	jmp    *%rcx
  13a6e8:	cmpq   $0x0,0x88(%rsp)
  13a6f1:	jle    13a705 <litchi_pptx::shape::reader::Scene::read_with+0x23f5>
  13a6f3:	mov    0x90(%rsp),%rdi
  13a6fb:	mov    %edx,%ebx
  13a6fd:	call   *0xf2355(%rip)        # 22ca58 <free@GLIBC_2.2.5>
  13a703:	mov    %ebx,%edx
  13a705:	mov    %dl,0x78(%rsp)
  13a709:	mov    %rbp,%rbx
  13a70c:	cmpq   $0x0,0x228(%rsp)
  13a715:	je     13a725 <litchi_pptx::shape::reader::Scene::read_with+0x2415>
  13a717:	mov    0x230(%rsp),%rdi
  13a71f:	call   *0xf2333(%rip)        # 22ca58 <free@GLIBC_2.2.5>
  13a725:	cmpq   $0x0,0x240(%rsp)
  13a72e:	je     13a73e <litchi_pptx::shape::reader::Scene::read_with+0x242e>
  13a730:	mov    0x248(%rsp),%rdi
  13a738:	call   *0xf231a(%rip)        # 22ca58 <free@GLIBC_2.2.5>
  13a73e:	cmpq   $0x0,0x288(%rsp)
  13a747:	je     13a757 <litchi_pptx::shape::reader::Scene::read_with+0x2447>
  13a749:	mov    0x290(%rsp),%rdi
  13a751:	call   *0xf2301(%rip)        # 22ca58 <free@GLIBC_2.2.5>
  13a757:	cmpq   $0x0,0x2a0(%rsp)
  13a760:	je     13a770 <litchi_pptx::shape::reader::Scene::read_with+0x2460>
  13a762:	mov    0x2a8(%rsp),%rdi
  13a76a:	call   *0xf22e8(%rip)        # 22ca58 <free@GLIBC_2.2.5>
  13a770:	cmpq   $0x0,0x168(%rsp)
  13a779:	je     13a789 <litchi_pptx::shape::reader::Scene::read_with+0x2479>
  13a77b:	mov    0x170(%rsp),%rdi
  13a783:	call   *0xf22cf(%rip)        # 22ca58 <free@GLIBC_2.2.5>
  13a789:	mov    %r13,%rbp
  13a78c:	cmpq   $0x0,0x180(%rsp)
  13a795:	je     13a7a5 <litchi_pptx::shape::reader::Scene::read_with+0x2495>
  13a797:	mov    0x188(%rsp),%rdi
  13a79f:	call   *0xf22b3(%rip)        # 22ca58 <free@GLIBC_2.2.5>
  13a7a5:	mov    0x1a0(%rsp),%r15
  13a7ad:	mov    0x1a8(%rsp),%r12
  13a7b5:	test   %r12,%r12
  13a7b8:	je     13a800 <litchi_pptx::shape::reader::Scene::read_with+0x24f0>
  13a7ba:	lea    0xd0(%r15),%r13
  13a7c1:	jmp    13a7dc <litchi_pptx::shape::reader::Scene::read_with+0x24cc>
  13a7c3:	data16 data16 data16 cs nopw 0x0(%rax,%rax,1)
  13a7d0:	add    $0x100,%r13
  13a7d7:	dec    %r12
  13a7da:	je     13a800 <litchi_pptx::shape::reader::Scene::read_with+0x24f0>
  13a7dc:	cmpq   $0x0,-0x20(%r13)
  13a7e1:	jle    13a7ed <litchi_pptx::shape::reader::Scene::read_with+0x24dd>
  13a7e3:	mov    -0x18(%r13),%rdi
  13a7e7:	call   *0xf226b(%rip)        # 22ca58 <free@GLIBC_2.2.5>
  13a7ed:	cmpq   $0x0,-0x8(%r13)
  13a7f2:	jle    13a7d0 <litchi_pptx::shape::reader::Scene::read_with+0x24c0>
  13a7f4:	mov    0x0(%r13),%rdi
  13a7f8:	call   *0xf225a(%rip)        # 22ca58 <free@GLIBC_2.2.5>
  13a7fe:	jmp    13a7d0 <litchi_pptx::shape::reader::Scene::read_with+0x24c0>
  13a800:	cmpq   $0x0,0x198(%rsp)
  13a809:	je     13a814 <litchi_pptx::shape::reader::Scene::read_with+0x2504>
  13a80b:	mov    %r15,%rdi
  13a80e:	call   *0xf2244(%rip)        # 22ca58 <free@GLIBC_2.2.5>
  13a814:	cmpq   $0x0,0x1b0(%rsp)
  13a81d:	mov    0xe8(%rsp),%r13
  13a825:	je     13a835 <litchi_pptx::shape::reader::Scene::read_with+0x2525>
  13a827:	mov    0x1b8(%rsp),%rdi
  13a82f:	call   *0xf2223(%rip)        # 22ca58 <free@GLIBC_2.2.5>
  13a835:	movzbl 0x78(%rsp),%esi
  13a83a:	cmp    $0x27,%sil
  13a83e:	movabs $0x8000000000000001,%rdx
  13a848:	jne    13a8d8 <litchi_pptx::shape::reader::Scene::read_with+0x25c8>
  13a84e:	mov    0x138(%rsp),%rax
  13a856:	movaps 0xf0(%rsp),%xmm0
  13a85e:	movaps %xmm0,0x310(%rsp)
  13a866:	movups %xmm0,0x20(%r14)
  13a86b:	mov    %rbx,(%r14)
  13a86e:	mov    %rbp,0x8(%r14)
  13a872:	movaps 0x10(%rsp),%xmm0
  13a877:	movups %xmm0,0x10(%r14)
  13a87c:	mov    %rax,0x30(%r14)
  13a880:	mov    0x140(%rsp),%rax
  13a888:	mov    %rax,0x38(%r14)
  13a88c:	mov    0x268(%rsp),%rax
  13a894:	mov    %rax,0x40(%r14)
  13a898:	movups 0x0(%r13),%xmm0
  13a89d:	movups 0x10(%r13),%xmm1
  13a8a2:	movups 0x20(%r13),%xmm2
  13a8a7:	movups %xmm0,0x48(%r14)
  13a8ac:	movups %xmm1,0x58(%r14)
  13a8b1:	movups %xmm2,0x68(%r14)
  13a8b6:	lea    0x350(%rsp),%rdi
  13a8be:	call   164400 <core::ptr::drop_in_place<litchi_ooxml_common::mce::model::Capabilities>>
  13a8c3:	mov    %r14,%rax
  13a8c6:	add    $0x3e8,%rsp
  13a8cd:	pop    %rbx
  13a8ce:	pop    %r12
  13a8d0:	pop    %r13
  13a8d2:	pop    %r14
  13a8d4:	pop    %r15
  13a8d6:	pop    %rbp
  13a8d7:	ret
  13a8d8:	mov    (%rsp),%eax
  13a8db:	mov    0x3(%rsp),%ecx
  13a8df:	mov    %ecx,0xc(%r14)
  13a8e3:	mov    %eax,0x9(%r14)
  13a8e7:	movaps 0xf0(%rsp),%xmm0
  13a8ef:	movaps 0x100(%rsp),%xmm1
  13a8f7:	movaps %xmm0,0x310(%rsp)
  13a8ff:	movups %xmm1,0x40(%r14)
  13a904:	movaps 0x310(%rsp),%xmm0
  13a90c:	movups %xmm0,0x30(%r14)
  13a911:	mov    %sil,0x8(%r14)
  13a915:	mov    %rbx,0x10(%r14)
  13a919:	mov    %rbp,0x18(%r14)
  13a91d:	movaps 0x10(%rsp),%xmm0
  13a922:	movups %xmm0,0x20(%r14)
  13a927:	dec    %rdx
  13a92a:	mov    %rdx,(%r14)
  13a92d:	mov    0x140(%rsp),%rbx
  13a935:	mov    0x138(%rsp),%rcx
  13a93d:	shl    $1,%rcx
  13a940:	test   %rcx,%rcx
  13a943:	je     1385b3 <litchi_pptx::shape::reader::Scene::read_with+0x2a3>
  13a949:	jmp    1385aa <litchi_pptx::shape::reader::Scene::read_with+0x29a>
  13a94e:	mov    $0x17,%eax
  13a953:	movq   %rax,%xmm0
  13a958:	movdqa %xmm0,0x10(%rsp)
  13a95e:	lea    -0x11345b(%rip),%rax        # 2750a <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13171701955629477448+0x1d2>
  13a965:	mov    %r13,%rbp
  13a968:	mov    %rax,%r13
  13a96b:	test   %r15,%r15
  13a96e:	jg     13a684 <litchi_pptx::shape::reader::Scene::read_with+0x2374>
  13a974:	jmp    13a693 <litchi_pptx::shape::reader::Scene::read_with+0x2383>
  13a979:	mov    $0x17,%eax
  13a97e:	movq   %rax,%xmm0
  13a983:	movdqa %xmm0,0x10(%rsp)
  13a989:	lea    -0x113486(%rip),%rax        # 2750a <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13171701955629477448+0x1d2>
  13a990:	mov    %r13,%rbp
  13a993:	mov    %rax,%r13
  13a996:	cmpq   $0x0,0x8(%rsp)
  13a99c:	jg     13a616 <litchi_pptx::shape::reader::Scene::read_with+0x2306>
  13a9a2:	jmp    13a628 <litchi_pptx::shape::reader::Scene::read_with+0x2318>
  13a9a7:	mov    %al,0x78(%rsp)
  13a9ab:	mov    0x21(%rsp),%eax
  13a9af:	mov    0x24(%rsp),%ecx
  13a9b3:	mov    %ecx,0x3(%rsp)
  13a9b7:	mov    %eax,(%rsp)
  13a9ba:	mov    0x30(%rsp),%r13
  13a9bf:	movups 0x38(%rsp),%xmm0
  13a9c4:	movaps %xmm0,0x10(%rsp)
  13a9c9:	movups 0x48(%rsp),%xmm0
  13a9ce:	movaps %xmm0,0xf0(%rsp)
  13a9d6:	movdqu 0x58(%rsp),%xmm0
  13a9dc:	movdqa %xmm0,0x100(%rsp)
  13a9e5:	cmpq   $0x0,0x228(%rsp)
  13a9ee:	jne    13a717 <litchi_pptx::shape::reader::Scene::read_with+0x2407>
  13a9f4:	jmp    13a725 <litchi_pptx::shape::reader::Scene::read_with+0x2415>
  13a9f9:	mov    0x21(%rsp),%eax
  13a9fd:	mov    0x24(%rsp),%ecx
  13aa01:	mov    %ecx,0x3(%rsp)
  13aa05:	mov    %eax,(%rsp)
  13aa08:	mov    0x30(%rsp),%r13
  13aa0d:	movups 0x38(%rsp),%xmm0
  13aa12:	movaps %xmm0,0x10(%rsp)
  13aa17:	movups 0x48(%rsp),%xmm0
  13aa1c:	movaps %xmm0,0xf0(%rsp)
  13aa24:	movdqu 0x58(%rsp),%xmm0
  13aa2a:	movdqa %xmm0,0x100(%rsp)
  13aa33:	cmp    $0x9,%r15
  13aa37:	jbe    139c22 <litchi_pptx::shape::reader::Scene::read_with+0x1912>
  13aa3d:	jmp    13a6ce <litchi_pptx::shape::reader::Scene::read_with+0x23be>
  13aa42:	mov    $0x2c,%r15d
  13aa48:	mov    $0x2c,%edi
  13aa4d:	call   *0xf1ff5(%rip)        # 22ca48 <malloc@GLIBC_2.2.5>
  13aa53:	test   %rax,%rax
  13aa56:	je     13ac30 <litchi_pptx::shape::reader::Scene::read_with+0x2920>
  13aa5c:	mov    %rax,%r13
  13aa5f:	movups -0x11317d(%rip),%xmm0        # 278e9 <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13171701955629477448+0x5b1>
  13aa66:	movups %xmm0,0x1c(%rax)
  13aa6a:	movups -0x113194(%rip),%xmm0        # 278dd <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13171701955629477448+0x5a5>
  13aa71:	movups %xmm0,0x10(%rax)
  13aa75:	movups -0x1131af(%rip),%xmm0        # 278cd <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13171701955629477448+0x595>
  13aa7c:	movups %xmm0,0x0(%r13)
  13aa81:	mov    $0x2c,%ebp
  13aa86:	jmp    13ab1f <litchi_pptx::shape::reader::Scene::read_with+0x280f>
  13aa8b:	mov    $0x37,%r15d
  13aa91:	mov    $0x37,%edi
  13aa96:	call   *0xf1fac(%rip)        # 22ca48 <malloc@GLIBC_2.2.5>
  13aa9c:	test   %rax,%rax
  13aa9f:	je     13ac30 <litchi_pptx::shape::reader::Scene::read_with+0x2920>
  13aaa5:	mov    %rax,%r13
  13aaa8:	movups -0x113059(%rip),%xmm0        # 27a56 <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13171701955629477448+0x71e>
  13aaaf:	movups %xmm0,0x20(%rax)
  13aab3:	movups -0x113074(%rip),%xmm0        # 27a46 <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13171701955629477448+0x70e>
  13aaba:	movups %xmm0,0x10(%rax)
  13aabe:	movups -0x11308f(%rip),%xmm0        # 27a36 <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13171701955629477448+0x6fe>
  13aac5:	movups %xmm0,(%rax)
  13aac8:	movabs $0x617461646174656d,%rax
  13aad2:	mov    %rax,0x2f(%r13)
  13aad6:	mov    $0x37,%ebp
  13aadb:	jmp    13ab1f <litchi_pptx::shape::reader::Scene::read_with+0x280f>
  13aadd:	mov    $0x2e,%r15d
  13aae3:	mov    $0x2e,%edi
  13aae8:	call   *0xf1f5a(%rip)        # 22ca48 <malloc@GLIBC_2.2.5>
  13aaee:	test   %rax,%rax
  13aaf1:	je     13ac30 <litchi_pptx::shape::reader::Scene::read_with+0x2920>
  13aaf7:	mov    %rax,%r13
  13aafa:	movups -0x113197(%rip),%xmm0        # 2796a <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13171701955629477448+0x632>
  13ab01:	movups %xmm0,0x1e(%rax)
  13ab05:	movups -0x1131b0(%rip),%xmm0        # 2795c <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13171701955629477448+0x624>
  13ab0c:	movups %xmm0,0x10(%rax)
  13ab10:	movups -0x1131cb(%rip),%xmm0        # 2794c <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13171701955629477448+0x614>
  13ab17:	movups %xmm0,(%rax)
  13ab1a:	mov    $0x2e,%ebp
  13ab1f:	movq   %rbp,%xmm0
  13ab24:	movdqa %xmm0,0x10(%rsp)
  13ab2a:	mov    $0x8,%dl
  13ab2c:	movabs $0x8000000000000001,%r12
  13ab36:	jmp    139fb9 <litchi_pptx::shape::reader::Scene::read_with+0x1ca9>
  13ab3b:	mov    $0x2c,%r15d
  13ab41:	mov    $0x2c,%edi
  13ab46:	call   *0xf1efc(%rip)        # 22ca48 <malloc@GLIBC_2.2.5>
  13ab4c:	test   %rax,%rax
  13ab4f:	je     13ac30 <litchi_pptx::shape::reader::Scene::read_with+0x2920>
  13ab55:	mov    %rax,%r13
  13ab58:	movups -0x113172(%rip),%xmm0        # 279ed <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13171701955629477448+0x6b5>
  13ab5f:	movups %xmm0,0x1c(%rax)
  13ab63:	movups -0x113189(%rip),%xmm0        # 279e1 <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13171701955629477448+0x6a9>
  13ab6a:	movups %xmm0,0x10(%rax)
  13ab6e:	movups -0x1131a4(%rip),%xmm0        # 279d1 <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13171701955629477448+0x699>
  13ab75:	jmp    13aa7c <litchi_pptx::shape::reader::Scene::read_with+0x276c>
  13ab7a:	mov    $0x1,%edi
  13ab7f:	mov    $0x36,%esi
  13ab84:	call   *0xf1f06(%rip)        # 22ca90 <_DYNAMIC+0x230>
  13ab8a:	mov    $0x39,%r15d
  13ab90:	mov    $0x39,%edi
  13ab95:	call   *0xf1ead(%rip)        # 22ca48 <malloc@GLIBC_2.2.5>
  13ab9b:	test   %rax,%rax
  13ab9e:	je     13ac30 <litchi_pptx::shape::reader::Scene::read_with+0x2920>
  13aba4:	mov    %rax,%r13
  13aba7:	movups -0x113188(%rip),%xmm0        # 27a26 <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13171701955629477448+0x6ee>
  13abae:	movups %xmm0,0x29(%rax)
  13abb2:	movups -0x11319c(%rip),%xmm0        # 27a1d <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13171701955629477448+0x6e5>
  13abb9:	movups %xmm0,0x20(%rax)
  13abbd:	movups -0x1131b7(%rip),%xmm0        # 27a0d <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13171701955629477448+0x6d5>
  13abc4:	movups %xmm0,0x10(%rax)
  13abc8:	movups -0x1131d2(%rip),%xmm0        # 279fd <anon.7d1f7c43824a99e7625b6b83caf5366f.6940.llvm.13171701955629477448+0x6c5>
  13abcf:	movups %xmm0,(%rax)
  13abd2:	mov    $0x39,%ebp
  13abd7:	jmp    13ab1f <litchi_pptx::shape::reader::Scene::read_with+0x280f>
  13abdc:	mov    $0x1,%edi
  13abe1:	mov    %r15,%rsi
  13abe4:	call   *0xf1ea6(%rip)        # 22ca90 <_DYNAMIC+0x230>
  13abea:	jmp    13ac3e <litchi_pptx::shape::reader::Scene::read_with+0x292e>
  13abec:	mov    $0x1,%edi
  13abf1:	mov    %r15,%rsi
  13abf4:	call   *0xf1e96(%rip)        # 22ca90 <_DYNAMIC+0x230>
  13abfa:	jmp    13ac3e <litchi_pptx::shape::reader::Scene::read_with+0x292e>
  13abfc:	mov    $0x1,%edi
  13ac01:	mov    %rbx,%rsi
  13ac04:	call   *0xf1e86(%rip)        # 22ca90 <_DYNAMIC+0x230>
  13ac0a:	jmp    13ac3e <litchi_pptx::shape::reader::Scene::read_with+0x292e>
  13ac0c:	mov    $0x1,%edi
  13ac11:	mov    $0x3e,%esi
  13ac16:	call   *0xf1e74(%rip)        # 22ca90 <_DYNAMIC+0x230>
  13ac1c:	jmp    13ac3e <litchi_pptx::shape::reader::Scene::read_with+0x292e>
  13ac1e:	mov    $0x1,%edi
  13ac23:	mov    $0x26,%esi
  13ac28:	call   *0xf1e62(%rip)        # 22ca90 <_DYNAMIC+0x230>
  13ac2e:	jmp    13ac3e <litchi_pptx::shape::reader::Scene::read_with+0x292e>
  13ac30:	mov    $0x1,%edi
  13ac35:	mov    %r15,%rsi
  13ac38:	call   *0xf1e52(%rip)        # 22ca90 <_DYNAMIC+0x230>
  13ac3e:	ud2
  13ac40:	jmp    13ad76 <litchi_pptx::shape::reader::Scene::read_with+0x2a66>
  13ac45:	mov    %rax,%r14
  13ac48:	lea    0x148(%rsp),%rdi
  13ac50:	call   93700 <core::ptr::drop_in_place<litchi_ooxml_common::mce::model::NamespaceSet>>
  13ac55:	mov    %r14,%rdi
  13ac58:	call   222550 <_Unwind_Resume@plt>
  13ac5d:	jmp    13acd1 <litchi_pptx::shape::reader::Scene::read_with+0x29c1>
  13ac5f:	jmp    13ad06 <litchi_pptx::shape::reader::Scene::read_with+0x29f6>
  13ac64:	jmp    13ad4f <litchi_pptx::shape::reader::Scene::read_with+0x2a3f>
  13ac69:	jmp    13aca5 <litchi_pptx::shape::reader::Scene::read_with+0x2995>
  13ac6b:	jmp    13acd1 <litchi_pptx::shape::reader::Scene::read_with+0x29c1>
  13ac6d:	jmp    13ad26 <litchi_pptx::shape::reader::Scene::read_with+0x2a16>
  13ac72:	jmp    13adb9 <litchi_pptx::shape::reader::Scene::read_with+0x2aa9>
  13ac77:	jmp    13adde <litchi_pptx::shape::reader::Scene::read_with+0x2ace>
  13ac7c:	mov    %rax,%r14
  13ac7f:	shl    $1,%r13
  13ac82:	test   %r13,%r13
  13ac85:	je     13acd4 <litchi_pptx::shape::reader::Scene::read_with+0x29c4>
  13ac87:	mov    %r15,%rdi
  13ac8a:	call   *0xf1dc8(%rip)        # 22ca58 <free@GLIBC_2.2.5>
  13ac90:	jmp    13acd4 <litchi_pptx::shape::reader::Scene::read_with+0x29c4>
  13ac92:	mov    %rax,%r14
  13ac95:	test   %r15,%r15
  13ac98:	je     13aca8 <litchi_pptx::shape::reader::Scene::read_with+0x2998>
  13ac9a:	mov    %rbx,%rdi
  13ac9d:	call   *0xf1db5(%rip)        # 22ca58 <free@GLIBC_2.2.5>
  13aca3:	jmp    13aca8 <litchi_pptx::shape::reader::Scene::read_with+0x2998>
  13aca5:	mov    %rax,%r14
  13aca8:	mov    $0x1,%r15b
  13acab:	cmpq   $0x0,0x110(%rsp)
  13acb4:	jle    13acc4 <litchi_pptx::shape::reader::Scene::read_with+0x29b4>
  13acb6:	mov    0x118(%rsp),%rdi
  13acbe:	call   *0xf1d94(%rip)        # 22ca58 <free@GLIBC_2.2.5>
  13acc4:	xor    %edx,%edx
  13acc6:	mov    $0x1,%bl
  13acc8:	mov    $0x1,%al
  13acca:	mov    $0x1,%cl
  13accc:	jmp    13ae48 <litchi_pptx::shape::reader::Scene::read_with+0x2b38>
  13acd1:	mov    %rax,%r14
  13acd4:	mov    $0x1,%r15b
  13acd7:	test   %r12,%r12
  13acda:	jle    13ace5 <litchi_pptx::shape::reader::Scene::read_with+0x29d5>
  13acdc:	mov    %rbx,%rdi
  13acdf:	call   *0xf1d73(%rip)        # 22ca58 <free@GLIBC_2.2.5>
  13ace5:	xor    %ecx,%ecx
  13ace7:	mov    $0x1,%bl
  13ace9:	mov    $0x1,%al
  13aceb:	jmp    13ae46 <litchi_pptx::shape::reader::Scene::read_with+0x2b36>
  13acf0:	mov    %rax,%r14
  13acf3:	shl    $1,%r13
  13acf6:	test   %r13,%r13
  13acf9:	je     13ad09 <litchi_pptx::shape::reader::Scene::read_with+0x29f9>
  13acfb:	mov    %r12,%rdi
  13acfe:	call   *0xf1d54(%rip)        # 22ca58 <free@GLIBC_2.2.5>
  13ad04:	jmp    13ad09 <litchi_pptx::shape::reader::Scene::read_with+0x29f9>
  13ad06:	mov    %rax,%r14
  13ad09:	shl    $1,%r15
  13ad0c:	test   %r15,%r15
  13ad0f:	je     13ad52 <litchi_pptx::shape::reader::Scene::read_with+0x2a42>
  13ad11:	mov    0x8(%rsp),%rdi
  13ad16:	call   *0xf1d3c(%rip)        # 22ca58 <free@GLIBC_2.2.5>
  13ad1c:	jmp    13ad52 <litchi_pptx::shape::reader::Scene::read_with+0x2a42>
  13ad1e:	mov    %rax,%r14
  13ad21:	jmp    13aedb <litchi_pptx::shape::reader::Scene::read_with+0x2bcb>
  13ad26:	mov    %rax,%r14
  13ad29:	movabs $0x8000000000000002,%rax
  13ad33:	cmp    %rax,0x8(%rsp)
  13ad38:	jl     13ad79 <litchi_pptx::shape::reader::Scene::read_with+0x2a69>
  13ad3a:	cmpq   $0x0,0x8(%rsp)
  13ad40:	je     13ad79 <litchi_pptx::shape::reader::Scene::read_with+0x2a69>
  13ad42:	mov    0x68(%rsp),%rdi
  13ad47:	call   *0xf1d0b(%rip)        # 22ca58 <free@GLIBC_2.2.5>
  13ad4d:	jmp    13ad79 <litchi_pptx::shape::reader::Scene::read_with+0x2a69>
  13ad4f:	mov    %rax,%r14
  13ad52:	mov    $0x1,%r15b
  13ad55:	test   %rbp,%rbp
  13ad58:	jle    13ad63 <litchi_pptx::shape::reader::Scene::read_with+0x2a53>
  13ad5a:	mov    %rbx,%rdi
  13ad5d:	call   *0xf1cf5(%rip)        # 22ca58 <free@GLIBC_2.2.5>
  13ad63:	xor    %eax,%eax
  13ad65:	mov    $0x1,%bl
  13ad67:	jmp    13ae44 <litchi_pptx::shape::reader::Scene::read_with+0x2b34>
  13ad6c:	jmp    13adfe <litchi_pptx::shape::reader::Scene::read_with+0x2aee>
  13ad71:	jmp    13ae21 <litchi_pptx::shape::reader::Scene::read_with+0x2b11>
  13ad76:	mov    %rax,%r14
  13ad79:	mov    $0x1,%r15b
  13ad7c:	mov    0xe0(%rsp),%rax
  13ad84:	shl    $1,%rax
  13ad87:	test   %rax,%rax
  13ad8a:	je     13ad95 <litchi_pptx::shape::reader::Scene::read_with+0x2a85>
  13ad8c:	mov    %rbx,%rdi
  13ad8f:	call   *0xf1cc3(%rip)        # 22ca58 <free@GLIBC_2.2.5>
  13ad95:	xor    %esi,%esi
  13ad97:	mov    $0x1,%bl
  13ad99:	mov    $0x1,%al
  13ad9b:	mov    $0x1,%cl
  13ad9d:	mov    $0x1,%dl
  13ad9f:	jmp    13ae4b <litchi_pptx::shape::reader::Scene::read_with+0x2b3b>
  13ada4:	mov    %rax,%r14
  13ada7:	test   %r15,%r15
  13adaa:	jle    13ae01 <litchi_pptx::shape::reader::Scene::read_with+0x2af1>
  13adac:	mov    0x70(%rsp),%rdi
  13adb1:	call   *0xf1ca1(%rip)        # 22ca58 <free@GLIBC_2.2.5>
  13adb7:	jmp    13ae01 <litchi_pptx::shape::reader::Scene::read_with+0x2af1>
  13adb9:	mov    %rax,%r14
  13adbc:	mov    $0x1,%r15b
  13adbf:	mov    $0x1,%bl
  13adc1:	jmp    13ae42 <litchi_pptx::shape::reader::Scene::read_with+0x2b32>
  13adc3:	mov    %rax,%r14
  13adc6:	cmpq   $0x0,0x8(%rsp)
  13adcc:	jle    13ae24 <litchi_pptx::shape::reader::Scene::read_with+0x2b14>
  13adce:	mov    0xe0(%rsp),%rdi
  13add6:	call   *0xf1c7c(%rip)        # 22ca58 <free@GLIBC_2.2.5>
  13addc:	jmp    13ae24 <litchi_pptx::shape::reader::Scene::read_with+0x2b14>
  13adde:	mov    %rax,%r14
  13ade1:	jmp    13aece <litchi_pptx::shape::reader::Scene::read_with+0x2bbe>
  13ade6:	mov    %rax,%r14
  13ade9:	lea    0x350(%rsp),%rdi
  13adf1:	call   164400 <core::ptr::drop_in_place<litchi_ooxml_common::mce::model::Capabilities>>
  13adf6:	mov    %r14,%rdi
  13adf9:	call   222550 <_Unwind_Resume@plt>
  13adfe:	mov    %rax,%r14
  13ae01:	mov    $0x1,%r15b
  13ae04:	cmpq   $0x0,0xb0(%rsp)
  13ae0d:	jle    13ae1d <litchi_pptx::shape::reader::Scene::read_with+0x2b0d>
  13ae0f:	mov    0xb8(%rsp),%rdi
  13ae17:	call   *0xf1c3b(%rip)        # 22ca58 <free@GLIBC_2.2.5>
  13ae1d:	xor    %ebx,%ebx
  13ae1f:	jmp    13ae42 <litchi_pptx::shape::reader::Scene::read_with+0x2b32>
  13ae21:	mov    %rax,%r14
  13ae24:	mov    $0x1,%bl
  13ae26:	cmpq   $0x0,0xb0(%rsp)
  13ae2f:	jle    13ae3f <litchi_pptx::shape::reader::Scene::read_with+0x2b2f>
  13ae31:	mov    0xb8(%rsp),%rdi
  13ae39:	call   *0xf1c19(%rip)        # 22ca58 <free@GLIBC_2.2.5>
  13ae3f:	xor    %r15d,%r15d
  13ae42:	mov    $0x1,%al
  13ae44:	mov    $0x1,%cl
  13ae46:	mov    $0x1,%dl
  13ae48:	mov    $0x1,%sil
  13ae4b:	mov    0x80(%rsp),%rdi
  13ae53:	cmp    $0x9,%rdi
  13ae57:	ja     13aea0 <litchi_pptx::shape::reader::Scene::read_with+0x2b90>
  13ae59:	lea    -0x116b54(%rip),%r8        # 2430c <anon.bc3d94ef2394061926bee58178356194.284.llvm.5373085861489949100+0x4314>
  13ae60:	movslq (%r8,%rdi,4),%rdi
  13ae64:	add    %r8,%rdi
  13ae67:	jmp    *%rdi
  13ae69:	cmpq   $0x0,0x88(%rsp)
  13ae72:	setg   %al
  13ae75:	test   %al,%r15b
  13ae78:	jne    13aec0 <litchi_pptx::shape::reader::Scene::read_with+0x2bb0>
  13ae7a:	jmp    13aece <litchi_pptx::shape::reader::Scene::read_with+0x2bbe>
  13ae7c:	cmpq   $0x0,0x88(%rsp)
  13ae85:	setg   %al
  13ae88:	test   %al,%bl
  13ae8a:	jne    13aec0 <litchi_pptx::shape::reader::Scene::read_with+0x2bb0>
  13ae8c:	jmp    13aece <litchi_pptx::shape::reader::Scene::read_with+0x2bbe>
  13ae8e:	cmpq   $0x0,0x88(%rsp)
  13ae97:	setg   %al
  13ae9a:	test   %al,%dl
  13ae9c:	jne    13aec0 <litchi_pptx::shape::reader::Scene::read_with+0x2bb0>
  13ae9e:	jmp    13aece <litchi_pptx::shape::reader::Scene::read_with+0x2bbe>
  13aea0:	lea    0x80(%rsp),%rdi
  13aea8:	call   161f80 <core::ptr::drop_in_place<quick_xml::events::Event>>
  13aead:	jmp    13aece <litchi_pptx::shape::reader::Scene::read_with+0x2bbe>
  13aeaf:	cmpq   $0x0,0x88(%rsp)
  13aeb8:	setg   %al
  13aebb:	test   %al,%sil
  13aebe:	je     13aece <litchi_pptx::shape::reader::Scene::read_with+0x2bbe>
  13aec0:	mov    0x90(%rsp),%rdi
  13aec8:	call   *0xf1b8a(%rip)        # 22ca58 <free@GLIBC_2.2.5>
  13aece:	lea    0x228(%rsp),%rdi
  13aed6:	call   163730 <core::ptr::drop_in_place<quick_xml::name::NamespaceResolver>>
  13aedb:	lea    0x288(%rsp),%rdi
  13aee3:	call   164830 <core::ptr::drop_in_place<quick_xml::reader::Reader<&[u8]>>>
  13aee8:	cmpq   $0x0,0x168(%rsp)
  13aef1:	jne    13af3b <litchi_pptx::shape::reader::Scene::read_with+0x2c2b>
  13aef3:	cmpq   $0x0,0x180(%rsp)
  13aefc:	jne    13af54 <litchi_pptx::shape::reader::Scene::read_with+0x2c44>
  13aefe:	lea    0x198(%rsp),%rdi
  13af06:	call   164c30 <core::ptr::drop_in_place<alloc::vec::Vec<litchi_pptx::shape::reader::Active>>>
  13af0b:	cmpq   $0x0,0x1b0(%rsp)
  13af14:	jne    13af7a <litchi_pptx::shape::reader::Scene::read_with+0x2c6a>
  13af16:	mov    0x138(%rsp),%rax
  13af1e:	shl    $1,%rax
  13af21:	test   %rax,%rax
  13af24:	jne    13af98 <litchi_pptx::shape::reader::Scene::read_with+0x2c88>
  13af26:	lea    0x350(%rsp),%rdi
  13af2e:	call   164400 <core::ptr::drop_in_place<litchi_ooxml_common::mce::model::Capabilities>>
  13af33:	mov    %r14,%rdi
  13af36:	call   222550 <_Unwind_Resume@plt>
  13af3b:	mov    0x170(%rsp),%rdi
  13af43:	call   *0xf1b0f(%rip)        # 22ca58 <free@GLIBC_2.2.5>
  13af49:	cmpq   $0x0,0x180(%rsp)
  13af52:	je     13aefe <litchi_pptx::shape::reader::Scene::read_with+0x2bee>
  13af54:	mov    0x188(%rsp),%rdi
  13af5c:	call   *0xf1af6(%rip)        # 22ca58 <free@GLIBC_2.2.5>
  13af62:	lea    0x198(%rsp),%rdi
  13af6a:	call   164c30 <core::ptr::drop_in_place<alloc::vec::Vec<litchi_pptx::shape::reader::Active>>>
  13af6f:	cmpq   $0x0,0x1b0(%rsp)
  13af78:	je     13af16 <litchi_pptx::shape::reader::Scene::read_with+0x2c06>
  13af7a:	mov    0x1b8(%rsp),%rdi
  13af82:	call   *0xf1ad0(%rip)        # 22ca58 <free@GLIBC_2.2.5>
  13af88:	mov    0x138(%rsp),%rax
  13af90:	shl    $1,%rax
  13af93:	test   %rax,%rax
  13af96:	je     13af26 <litchi_pptx::shape::reader::Scene::read_with+0x2c16>
  13af98:	mov    0x140(%rsp),%rdi
  13afa0:	call   *0xf1ab2(%rip)        # 22ca58 <free@GLIBC_2.2.5>
  13afa6:	lea    0x350(%rsp),%rdi
  13afae:	call   164400 <core::ptr::drop_in_place<litchi_ooxml_common::mce::model::Capabilities>>
  13afb3:	mov    %r14,%rdi
  13afb6:	call   222550 <_Unwind_Resume@plt>
  13afbb:	cmpq   $0x0,0x88(%rsp)
  13afc4:	setg   %cl
  13afc7:	test   %cl,%al
  13afc9:	jne    13aec0 <litchi_pptx::shape::reader::Scene::read_with+0x2bb0>
  13afcf:	jmp    13aece <litchi_pptx::shape::reader::Scene::read_with+0x2bbe>
  13afd4:	cmpq   $0x0,0x88(%rsp)
  13afdd:	setg   %al
  13afe0:	test   %al,%cl
  13afe2:	jne    13aec0 <litchi_pptx::shape::reader::Scene::read_with+0x2bb0>
  13afe8:	jmp    13aece <litchi_pptx::shape::reader::Scene::read_with+0x2bbe>

Disassembly of section .init:

Disassembly of section .fini:

Disassembly of section .plt:
