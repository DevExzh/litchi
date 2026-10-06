
/tmp/litchi-performance-0925/measurement-target-v4-final/release/litchi-perf-baseline:     file format elf64-x86-64


Disassembly of section .text:

00000000037d4140 <_ZN4sha26sha2567x86_sha8compress17hd36fc2477c3c8f36E>:
 37d4140:	f3 0f 6f 07          	movdqu (%rdi),%xmm0
 37d4144:	f3 0f 6f 4f 10       	movdqu 0x10(%rdi),%xmm1
 37d4149:	66 0f 70 c0 b1       	pshufd $0xb1,%xmm0,%xmm0
 37d414e:	66 44 0f 70 e1 1b    	pshufd $0x1b,%xmm1,%xmm12
 37d4154:	66 44 0f 6f e8       	movdqa %xmm0,%xmm13
 37d4159:	66 45 0f 3a 0f ec 08 	palignr $0x8,%xmm12,%xmm13
 37d4160:	66 44 0f 3a 0e e0 f0 	pblendw $0xf0,%xmm0,%xmm12
 37d4167:	48 85 d2             	test   %rdx,%rdx
 37d416a:	0f 84 66 03 00 00    	je     37d44d6 <_ZN4sha26sha2567x86_sha8compress17hd36fc2477c3c8f36E+0x396>
 37d4170:	48 c1 e2 06          	shl    $0x6,%rdx
 37d4174:	48 01 f2             	add    %rsi,%rdx
 37d4177:	66 0f 6f 2d 41 eb 95 	movdqa -0x36a14bf(%rip),%xmm5        # 132cc0 <anon.24a7ac914408ca04882f525f221f6366.471.llvm.3105813448066715807+0x310>
 37d417e:	fc 
 37d417f:	66 44 0f 6f 0d 78 7e 	movdqa -0x36a8188(%rip),%xmm9        # 12c000 <anon.532fb113304ac58b7999dd10d83ca2c4.31.llvm.16951954313665096363+0x3a0>
 37d4186:	95 fc 
 37d4188:	66 44 0f 6f 15 2f aa 	movdqa -0x36a55d1(%rip),%xmm10        # 12ebc0 <anon.517ed6af453157268891d104078c2924.472.llvm.10030043671302692207+0x140>
 37d418f:	95 fc 
 37d4191:	66 44 0f 6f 1d c6 9c 	movdqa -0x36a633a(%rip),%xmm11        # 12de60 <anon.296bd47d3b35cec32007788f89377efb.160.llvm.16556089365447441167+0x270>
 37d4198:	95 fc 
 37d419a:	66 0f 1f 44 00 00    	nopw   0x0(%rax,%rax,1)
 37d41a0:	f3 0f 6f 3e          	movdqu (%rsi),%xmm7
 37d41a4:	f3 44 0f 6f 76 10    	movdqu 0x10(%rsi),%xmm14
 37d41aa:	f3 0f 6f 56 20       	movdqu 0x20(%rsi),%xmm2
 37d41af:	66 0f 38 00 fd       	pshufb %xmm5,%xmm7
 37d41b4:	66 44 0f 38 00 f5    	pshufb %xmm5,%xmm14
 37d41ba:	66 0f 6f c7          	movdqa %xmm7,%xmm0
 37d41be:	66 0f fe 05 2a 90 95 	paddd  -0x36a6fd6(%rip),%xmm0        # 12d1f0 <anon.b781ef986be5a1e41dc71150f7e03496.2.llvm.9622472577055389530+0xa0>
 37d41c5:	fc 
 37d41c6:	66 41 0f 6f dc       	movdqa %xmm12,%xmm3
 37d41cb:	41 0f 38 cb dd       	sha256rnds2 %xmm0,%xmm13,%xmm3
 37d41d0:	66 0f 70 c0 0e       	pshufd $0xe,%xmm0,%xmm0
 37d41d5:	66 41 0f 6f e5       	movdqa %xmm13,%xmm4
 37d41da:	0f 38 cb e3          	sha256rnds2 %xmm0,%xmm3,%xmm4
 37d41de:	66 41 0f 6f c6       	movdqa %xmm14,%xmm0
 37d41e3:	66 0f fe 05 25 6b 95 	paddd  -0x36a94db(%rip),%xmm0        # 12ad10 <anon.2df47c5299e257657e8d615b695e6e82.51.llvm.14721412269600395664+0x20>
 37d41ea:	fc 
 37d41eb:	66 0f 38 00 d5       	pshufb %xmm5,%xmm2
 37d41f0:	0f 38 cb dc          	sha256rnds2 %xmm0,%xmm4,%xmm3
 37d41f4:	66 0f 70 c0 0e       	pshufd $0xe,%xmm0,%xmm0
 37d41f9:	0f 38 cb e3          	sha256rnds2 %xmm0,%xmm3,%xmm4
 37d41fd:	66 0f 6f c2          	movdqa %xmm2,%xmm0
 37d4201:	66 0f fe 05 17 d3 95 	paddd  -0x36a2ce9(%rip),%xmm0        # 131520 <anon.24a7ac914408ca04882f525f221f6366.1095.llvm.3105813448066715807+0x2c0>
 37d4208:	fc 
 37d4209:	f3 44 0f 6f 46 30    	movdqu 0x30(%rsi),%xmm8
 37d420f:	0f 38 cb dc          	sha256rnds2 %xmm0,%xmm4,%xmm3
 37d4213:	66 0f 70 c0 0e       	pshufd $0xe,%xmm0,%xmm0
 37d4218:	66 44 0f 38 00 c5    	pshufb %xmm5,%xmm8
 37d421e:	0f 38 cb e3          	sha256rnds2 %xmm0,%xmm3,%xmm4
 37d4222:	66 41 0f 6f c0       	movdqa %xmm8,%xmm0
 37d4227:	66 0f fe 05 a1 19 96 	paddd  -0x369e65f(%rip),%xmm0        # 135bd0 <anon.749e18c22614ce031527bed8eb1cf111.542.llvm.12148391067192879684+0x150>
 37d422e:	fc 
 37d422f:	66 0f 70 c8 0e       	pshufd $0xe,%xmm0,%xmm1
 37d4234:	41 0f 38 cc fe       	sha256msg1 %xmm14,%xmm7
 37d4239:	0f 38 cb dc          	sha256rnds2 %xmm0,%xmm4,%xmm3
 37d423d:	66 41 0f 6f f0       	movdqa %xmm8,%xmm6
 37d4242:	66 0f 3a 0f f2 04    	palignr $0x4,%xmm2,%xmm6
 37d4248:	66 0f 6f c1          	movdqa %xmm1,%xmm0
 37d424c:	0f 38 cb e3          	sha256rnds2 %xmm0,%xmm3,%xmm4
 37d4250:	66 0f fe f7          	paddd  %xmm7,%xmm6
 37d4254:	41 0f 38 cd f0       	sha256msg2 %xmm8,%xmm6
 37d4259:	66 0f 6f c6          	movdqa %xmm6,%xmm0
 37d425d:	66 0f fe 05 4b e4 95 	paddd  -0x36a1bb5(%rip),%xmm0        # 1326b0 <anon.778e880568a74592c22c79f0d6cff3c3.32.llvm.13910578690857602878+0xc0>
 37d4264:	fc 
 37d4265:	66 0f 70 c8 0e       	pshufd $0xe,%xmm0,%xmm1
 37d426a:	44 0f 38 cc f2       	sha256msg1 %xmm2,%xmm14
 37d426f:	0f 38 cb dc          	sha256rnds2 %xmm0,%xmm4,%xmm3
 37d4273:	66 0f 6f fe          	movdqa %xmm6,%xmm7
 37d4277:	66 41 0f 3a 0f f8 04 	palignr $0x4,%xmm8,%xmm7
 37d427e:	66 0f 6f c1          	movdqa %xmm1,%xmm0
 37d4282:	0f 38 cb e3          	sha256rnds2 %xmm0,%xmm3,%xmm4
 37d4286:	66 41 0f fe fe       	paddd  %xmm14,%xmm7
 37d428b:	0f 38 cd fe          	sha256msg2 %xmm6,%xmm7
 37d428f:	66 0f 6f c7          	movdqa %xmm7,%xmm0
 37d4293:	66 0f fe 05 45 19 96 	paddd  -0x369e6bb(%rip),%xmm0        # 135be0 <anon.749e18c22614ce031527bed8eb1cf111.542.llvm.12148391067192879684+0x160>
 37d429a:	fc 
 37d429b:	66 0f 70 c8 0e       	pshufd $0xe,%xmm0,%xmm1
 37d42a0:	41 0f 38 cc d0       	sha256msg1 %xmm8,%xmm2
 37d42a5:	0f 38 cb dc          	sha256rnds2 %xmm0,%xmm4,%xmm3
 37d42a9:	66 44 0f 6f ff       	movdqa %xmm7,%xmm15
 37d42ae:	66 44 0f 3a 0f fe 04 	palignr $0x4,%xmm6,%xmm15
 37d42b5:	66 0f 6f c1          	movdqa %xmm1,%xmm0
 37d42b9:	0f 38 cb e3          	sha256rnds2 %xmm0,%xmm3,%xmm4
 37d42bd:	66 44 0f fe fa       	paddd  %xmm2,%xmm15
 37d42c2:	44 0f 38 cd ff       	sha256msg2 %xmm7,%xmm15
 37d42c7:	66 41 0f 6f c7       	movdqa %xmm15,%xmm0
 37d42cc:	66 0f fe 05 5c d2 95 	paddd  -0x36a2da4(%rip),%xmm0        # 131530 <anon.24a7ac914408ca04882f525f221f6366.1095.llvm.3105813448066715807+0x2d0>
 37d42d3:	fc 
 37d42d4:	0f 38 cb dc          	sha256rnds2 %xmm0,%xmm4,%xmm3
 37d42d8:	44 0f 38 cc c6       	sha256msg1 %xmm6,%xmm8
 37d42dd:	66 41 0f 6f d7       	movdqa %xmm15,%xmm2
 37d42e2:	66 0f 3a 0f d7 04    	palignr $0x4,%xmm7,%xmm2
 37d42e8:	66 41 0f fe d0       	paddd  %xmm8,%xmm2
 37d42ed:	41 0f 38 cd d7       	sha256msg2 %xmm15,%xmm2
 37d42f2:	66 0f 70 c0 0e       	pshufd $0xe,%xmm0,%xmm0
 37d42f7:	66 44 0f 6f c2       	movdqa %xmm2,%xmm8
 37d42fc:	66 44 0f fe 05 3b de 	paddd  -0x36a21c5(%rip),%xmm8        # 132140 <anon.79a5a50476ce45ad8014bea9c03f713d.496.llvm.5232561714008784323+0x3c0>
 37d4303:	95 fc 
 37d4305:	0f 38 cb e3          	sha256rnds2 %xmm0,%xmm3,%xmm4
 37d4309:	66 41 0f 6f c0       	movdqa %xmm8,%xmm0
 37d430e:	0f 38 cb dc          	sha256rnds2 %xmm0,%xmm4,%xmm3
 37d4312:	0f 38 cc f7          	sha256msg1 %xmm7,%xmm6
 37d4316:	66 0f 6f ca          	movdqa %xmm2,%xmm1
 37d431a:	66 41 0f 3a 0f cf 04 	palignr $0x4,%xmm15,%xmm1
 37d4321:	66 0f fe ce          	paddd  %xmm6,%xmm1
 37d4325:	0f 38 cd ca          	sha256msg2 %xmm2,%xmm1
 37d4329:	66 41 0f 70 c0 0e    	pshufd $0xe,%xmm8,%xmm0
 37d432f:	66 0f 6f f1          	movdqa %xmm1,%xmm6
 37d4333:	66 0f fe 35 25 00 96 	paddd  -0x369ffdb(%rip),%xmm6        # 134360 <anon.24a7ac914408ca04882f525f221f6366.561.llvm.3105813448066715807+0x2d0>
 37d433a:	fc 
 37d433b:	0f 38 cb e3          	sha256rnds2 %xmm0,%xmm3,%xmm4
 37d433f:	66 0f 6f c6          	movdqa %xmm6,%xmm0
 37d4343:	0f 38 cb dc          	sha256rnds2 %xmm0,%xmm4,%xmm3
 37d4347:	41 0f 38 cc ff       	sha256msg1 %xmm15,%xmm7
 37d434c:	66 44 0f 6f f1       	movdqa %xmm1,%xmm14
 37d4351:	66 44 0f 3a 0f f2 04 	palignr $0x4,%xmm2,%xmm14
 37d4358:	66 44 0f fe f7       	paddd  %xmm7,%xmm14
 37d435d:	44 0f 38 cd f1       	sha256msg2 %xmm1,%xmm14
 37d4362:	66 0f 70 c6 0e       	pshufd $0xe,%xmm6,%xmm0
 37d4367:	66 41 0f 6f f6       	movdqa %xmm14,%xmm6
 37d436c:	44 0f 38 cc fa       	sha256msg1 %xmm2,%xmm15
 37d4371:	0f 38 cb e3          	sha256rnds2 %xmm0,%xmm3,%xmm4
 37d4375:	66 41 0f 6f fe       	movdqa %xmm14,%xmm7
 37d437a:	66 0f 3a 0f f9 04    	palignr $0x4,%xmm1,%xmm7
 37d4380:	66 41 0f fe ff       	paddd  %xmm15,%xmm7
 37d4385:	66 0f fe 35 c3 dd 95 	paddd  -0x36a223d(%rip),%xmm6        # 132150 <anon.79a5a50476ce45ad8014bea9c03f713d.496.llvm.5232561714008784323+0x3d0>
 37d438c:	fc 
 37d438d:	66 44 0f 70 fe 0e    	pshufd $0xe,%xmm6,%xmm15
 37d4393:	41 0f 38 cd fe       	sha256msg2 %xmm14,%xmm7
 37d4398:	66 0f 6f c6          	movdqa %xmm6,%xmm0
 37d439c:	0f 38 cb dc          	sha256rnds2 %xmm0,%xmm4,%xmm3
 37d43a0:	0f 38 cc d1          	sha256msg1 %xmm1,%xmm2
 37d43a4:	66 0f 6f f7          	movdqa %xmm7,%xmm6
 37d43a8:	66 41 0f 3a 0f f6 04 	palignr $0x4,%xmm14,%xmm6
 37d43af:	66 0f fe f2          	paddd  %xmm2,%xmm6
 37d43b3:	66 0f 6f d7          	movdqa %xmm7,%xmm2
 37d43b7:	66 0f fe 15 91 9a 95 	paddd  -0x36a656f(%rip),%xmm2        # 12de50 <anon.296bd47d3b35cec32007788f89377efb.160.llvm.16556089365447441167+0x260>
 37d43be:	fc 
 37d43bf:	66 41 0f 6f c7       	movdqa %xmm15,%xmm0
 37d43c4:	0f 38 cb e3          	sha256rnds2 %xmm0,%xmm3,%xmm4
 37d43c8:	0f 38 cd f7          	sha256msg2 %xmm7,%xmm6
 37d43cc:	41 0f 38 cc ce       	sha256msg1 %xmm14,%xmm1
 37d43d1:	66 0f 6f c2          	movdqa %xmm2,%xmm0
 37d43d5:	0f 38 cb dc          	sha256rnds2 %xmm0,%xmm4,%xmm3
 37d43d9:	66 44 0f 6f c6       	movdqa %xmm6,%xmm8
 37d43de:	66 44 0f 3a 0f c7 04 	palignr $0x4,%xmm7,%xmm8
 37d43e5:	66 44 0f fe c1       	paddd  %xmm1,%xmm8
 37d43ea:	66 0f 70 c2 0e       	pshufd $0xe,%xmm2,%xmm0
 37d43ef:	66 0f 6f ce          	movdqa %xmm6,%xmm1
 37d43f3:	66 0f fe 0d 75 93 95 	paddd  -0x36a6c8b(%rip),%xmm1        # 12d770 <anon.4a508b472db87571703eb0fa39eceb22.401.llvm.4366785517890424898+0x110>
 37d43fa:	fc 
 37d43fb:	0f 38 cb e3          	sha256rnds2 %xmm0,%xmm3,%xmm4
 37d43ff:	44 0f 38 cd c6       	sha256msg2 %xmm6,%xmm8
 37d4404:	44 0f 38 cc f7       	sha256msg1 %xmm7,%xmm14
 37d4409:	66 0f 6f c1          	movdqa %xmm1,%xmm0
 37d440d:	0f 38 cb dc          	sha256rnds2 %xmm0,%xmm4,%xmm3
 37d4411:	66 41 0f 6f d0       	movdqa %xmm8,%xmm2
 37d4416:	66 0f 3a 0f d6 04    	palignr $0x4,%xmm6,%xmm2
 37d441c:	66 41 0f fe d6       	paddd  %xmm14,%xmm2
 37d4421:	66 0f 70 c1 0e       	pshufd $0xe,%xmm1,%xmm0
 37d4426:	41 0f 38 cd d0       	sha256msg2 %xmm8,%xmm2
 37d442b:	0f 38 cc fe          	sha256msg1 %xmm6,%xmm7
 37d442f:	0f 38 cb e3          	sha256rnds2 %xmm0,%xmm3,%xmm4
 37d4433:	66 44 0f 6f f2       	movdqa %xmm2,%xmm14
 37d4438:	66 45 0f 3a 0f f0 04 	palignr $0x4,%xmm8,%xmm14
 37d443f:	41 0f 38 cc f0       	sha256msg1 %xmm8,%xmm6
 37d4444:	66 41 0f 6f c0       	movdqa %xmm8,%xmm0
 37d4449:	66 0f fe 05 5f a7 95 	paddd  -0x36a58a1(%rip),%xmm0        # 12ebb0 <anon.517ed6af453157268891d104078c2924.472.llvm.10030043671302692207+0x130>
 37d4450:	fc 
 37d4451:	0f 38 cb dc          	sha256rnds2 %xmm0,%xmm4,%xmm3
 37d4455:	66 0f 70 c0 0e       	pshufd $0xe,%xmm0,%xmm0
 37d445a:	66 44 0f fe f7       	paddd  %xmm7,%xmm14
 37d445f:	44 0f 38 cd f2       	sha256msg2 %xmm2,%xmm14
 37d4464:	0f 38 cb e3          	sha256rnds2 %xmm0,%xmm3,%xmm4
 37d4468:	66 41 0f 6f ce       	movdqa %xmm14,%xmm1
 37d446d:	66 0f 3a 0f ca 04    	palignr $0x4,%xmm2,%xmm1
 37d4473:	66 0f 6f c2          	movdqa %xmm2,%xmm0
 37d4477:	66 41 0f fe c1       	paddd  %xmm9,%xmm0
 37d447c:	0f 38 cb dc          	sha256rnds2 %xmm0,%xmm4,%xmm3
 37d4480:	66 0f 70 c0 0e       	pshufd $0xe,%xmm0,%xmm0
 37d4485:	0f 38 cb e3          	sha256rnds2 %xmm0,%xmm3,%xmm4
 37d4489:	66 0f fe ce          	paddd  %xmm6,%xmm1
 37d448d:	41 0f 38 cd ce       	sha256msg2 %xmm14,%xmm1
 37d4492:	66 41 0f 6f c6       	movdqa %xmm14,%xmm0
 37d4497:	66 41 0f fe c2       	paddd  %xmm10,%xmm0
 37d449c:	0f 38 cb dc          	sha256rnds2 %xmm0,%xmm4,%xmm3
 37d44a0:	66 0f 70 c0 0e       	pshufd $0xe,%xmm0,%xmm0
 37d44a5:	0f 38 cb e3          	sha256rnds2 %xmm0,%xmm3,%xmm4
 37d44a9:	66 41 0f fe cb       	paddd  %xmm11,%xmm1
 37d44ae:	66 0f 6f c1          	movdqa %xmm1,%xmm0
 37d44b2:	0f 38 cb dc          	sha256rnds2 %xmm0,%xmm4,%xmm3
 37d44b6:	66 0f 70 c1 0e       	pshufd $0xe,%xmm1,%xmm0
 37d44bb:	0f 38 cb e3          	sha256rnds2 %xmm0,%xmm3,%xmm4
 37d44bf:	66 44 0f fe ec       	paddd  %xmm4,%xmm13
 37d44c4:	66 44 0f fe e3       	paddd  %xmm3,%xmm12
 37d44c9:	48 83 c6 40          	add    $0x40,%rsi
 37d44cd:	48 39 d6             	cmp    %rdx,%rsi
 37d44d0:	0f 85 ca fc ff ff    	jne    37d41a0 <_ZN4sha26sha2567x86_sha8compress17hd36fc2477c3c8f36E+0x60>
 37d44d6:	66 41 0f 70 c5 1b    	pshufd $0x1b,%xmm13,%xmm0
 37d44dc:	66 41 0f 70 cc b1    	pshufd $0xb1,%xmm12,%xmm1
 37d44e2:	66 0f 6f d0          	movdqa %xmm0,%xmm2
 37d44e6:	66 0f 3a 0e d1 f0    	pblendw $0xf0,%xmm1,%xmm2
 37d44ec:	66 0f 3a 0f c8 08    	palignr $0x8,%xmm0,%xmm1
 37d44f2:	f3 0f 7f 17          	movdqu %xmm2,(%rdi)
 37d44f6:	f3 0f 7f 4f 10       	movdqu %xmm1,0x10(%rdi)
 37d44fb:	c3                   	ret
