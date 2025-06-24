# Porting Progress: wvWave to WvWaveDotNet

This document tracks the progress of porting each C source file from the original wvWave project to C# in the WvWaveDotNet library.

## Legend

- [ ] Not started
- [~] In progress
- [x] Completed

## File Conversion Checklist

| C Source File         | Status    | Notes                |
|---------------------- |-----------|----------------------|
| anld.c                | [ ]       |                      |
| anlv.c                | [ ]       |                      |
| asumy.c               | [ ]       |                      |
| asumyi.c              | [ ]       |                      |
| atrd.c                | [ ]       |                      |
| basename.c            | [ ]       |                      |
| bintree.c             | [ ]       |                      |
| bkd.c                 | [ ]       |                      |
| bkf.c                 | [ ]       |                      |
| bkl.c                 | [ ]       |                      |
| blip.c                | [ ]       |                      |
| brc.c                 | [ ]       |                      |
| bte.c                 | [ ]       |                      |
| bx.c                  | [ ]       |                      |
| chp.c                 | [ ]       |                      |
| clx.c                 | [ ]       |                      |
| crc32.c               | [ ]       |                      |
| dcs.c                 | [ ]       |                      |
| decode_complex.c      | [ ]       |                      |
| decode_simple.c       | [ ]       |                      |
| decompresswmf.c       | [ ]       |                      |
| decrypt95.c           | [ ]       |                      |
| decrypt97.c           | [ ]       |                      |
| dogrid.c              | [ ]       |                      |
| dop.c                 | [ ]       |                      |
| doptypography.c       | [ ]       |                      |
| dttm.c                | [ ]       |                      |
| error.c               | [ ]       |                      |
| escher.c              | [ ]       |                      |
| fbse.c                | [ ]       |                      |
| fdoa.c                | [ ]       |                      |
| ffn.c                 | [ ]       |                      |
| fib.c                 | [ ]       |                      |
| field.c               | [ ]       |                      |
| filetime.c            | [ ]       |                      |
| fkp.c                 | [ ]       |                      |
| fld.c                 | [ ]       |                      |
| font.c                | [ ]       |                      |
| fopt.c                | [ ]       |                      |
| frd.c                 | [ ]       |                      |
| fspa.c                | [ ]       |                      |
| ftxbxs.c              | [ ]       |                      |
| generic.c             | [ ]       |                      |
| getopt.c              | [ ]       |                      |
| getopt1.c             | [ ]       |                      |
| isbidi.c              | [ ]       |                      |
| laolareplace.c        | [ ]       |                      |
| lfo.c                 | [ ]       |                      |
| list.c                | [ ]       |                      |
| lspd.c                | [ ]       |                      |
| lst.c                 | [ ]       |                      |
| lvl.c                 | [ ]       |                      |
| md5.c                 | [ ]       |                      |
| mtextra.c             | [ ]       |                      |
| numrm.c               | [ ]       |                      |
| olst.c                | [ ]       |                      |
| pap.c                 | [ ]       |                      |
| pcd.c                 | [ ]       |                      |
| pgd.c                 | [ ]       |                      |
| phe.c                 | [ ]       |                      |
| picf.c                | [ ]       |                      |
| plcf.c                | [ ]       |                      |
| prm.c                 | [ ]       |                      |
| rc4.c                 | [ ]       |                      |
| reasons.c             | [ ]       |                      |
| roman.c               | [ ]       |                      |
| rr.c                  | [ ]       |                      |
| rs.c                  | [ ]       |                      |
| sed.c                 | [ ]       |                      |
| sep.c                 | [ ]       |                      |
| shd.c                 | [ ]       |                      |
| sprm.c                | [ ]       |                      |
| strcasecmp.c          | [ ]       |                      |
| sttbf.c               | [ ]       |                      |
| stylesheet.c          | [ ]       |                      |
| support.c             | [ ]       |                      |
| symbol.c              | [ ]       |                      |
| table.c               | [ ]       |                      |
| tap.c                 | [ ]       |                      |
| tbd.c                 | [ ]       |                      |
| tc.c                  | [ ]       |                      |
| text.c                | [ ]       |                      |
| tlp.c                 | [ ]       |                      |
| twips.c               | [ ]       |                      |
| unicode.c             | [ ]       |                      |
| utf.c                 | [ ]       |                      |
| version.c             | [ ]       |                      |
| winmmap.c             | [ ]       |                      |
| wkb.c                 | [ ]       |                      |
| wvConfig.c            | [ ]       |                      |
| wvConvert.c           | [ ]       |                      |
| wvHtmlEngine.c        | [ ]       |                      |
| wvRTF.c               | [ ]       |                      |
| wvSummary.c           | [ ]       |                      |
| wvTextEngine.c        | [ ]       |                      |
| wvVersion.c           | [ ]       |                      |
| wvWare.c              | [ ]       |                      |
| wvparse.c             | [ ]       |                      |
| xst.c                 | [ ]       |                      |
