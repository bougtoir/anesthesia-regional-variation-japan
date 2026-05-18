# Khmer inscription vocabulary pilot analysis

This pilot uses the public SEAlang Corpus of Khmer Inscriptions query interface
to test whether religious vocabulary associated with Mahayana, general Buddhist,
Theravada/Pali, and Brahmanical comparators can be tracked over time.

Source: http://sealang.net/ok/corpus-top.htm and http://sealang.net/classic/khmer/cki-bot.htm

## Caveats

- SEAlang describes the corpus as a demonstration version: texts are incomplete, not fully proofread, and era/region divisions are rough.
- This script queries selected lexical forms only; it does not lemmatize Sanskrit/Pali/Old Khmer variants exhaustively.
- Counts should be treated as exploratory signals for follow-up philological work, not as final historical proof.

## Top queried terms

| term | group | note | hits | dated_hits | first_year | last_year |
| --- | --- | --- | --- | --- | --- | --- |
| liṅga | Shaiva | Śaiva object vocabulary comparator | 28 | 29 | 680.0 | 1118.0 |
| deva | Brahmanical | General Indic divine vocabulary | 20 | 25 | 578.0 | 1178.0 |
| viṣṇu | Vaishnava | Vaiṣṇava vocabulary comparator | 19 | 21 | 878.0 | 1178.0 |
| śiva | Shaiva | Śaiva vocabulary comparator | 8 | 20 | 578.0 | 1178.0 |
| buddha | Buddhist-general | General Buddhist term | 5 | 5 | 1178.0 | 1361.0 |
| lokeśvara | Mahayana | Avalokiteśvara/Lokeśvara cult | 5 | 6 | 972.0 | 1026.0 |
| brāhmaṇa | Brahmanical | Brahmin vocabulary comparator | 4 | 4 | 678.0 | 1361.0 |
| maheśvara | Shaiva | Śaiva deity title comparator | 4 | 4 | 878.0 | 1306.0 |
| vajra | Mahayana | Vajrayāna/tantric marker | 4 | 3 | 878.0 | 1145.0 |
| saṅgha | Buddhist-general | Buddhist monastic community | 3 | 4 | 1107.0 | 1278.0 |
| thera | Theravada/Pali | Thera/Theravāda-adjacent vocabulary | 2 | 2 | 1278.0 | 1308.0 |
| bodhisattva | Mahayana | Bodhisattva terminology | 1 | 1 | 1178.0 | 1178.0 |

## Counts by vocabulary group

| group | count |
| --- | --- |
| Shaiva | 53 |
| Brahmanical | 29 |
| Vaishnava | 21 |
| Buddhist-general | 11 |
| Mahayana | 11 |
| Theravada/Pali | 3 |

## Mahayana-oriented queried terms

| term | group | note | hits | dated_hits | first_year | last_year |
| --- | --- | --- | --- | --- | --- | --- |
| vajra | Mahayana | Vajrayāna/tantric marker | 4 | 3 | 878.0 | 1145.0 |
| lokeśvara | Mahayana | Avalokiteśvara/Lokeśvara cult | 5 | 6 | 972.0 | 1026.0 |
| bodhisattva | Mahayana | Bodhisattva terminology | 1 | 1 | 1178.0 | 1178.0 |
| mahābodhi | Mahayana | Bodhi/Buddhist cult vocabulary | 1 | 1 | 1178.0 | 1178.0 |

## Theravada/Pali-oriented queried terms

| term | group | note | hits | dated_hits | first_year | last_year |
| --- | --- | --- | --- | --- | --- | --- |
| thera | Theravada/Pali | Thera/Theravāda-adjacent vocabulary | 2 | 2 | 1278.0 | 1308.0 |
| sāsana | Theravada/Pali | Pali sāsana / institutional vocabulary | 1 | 1 | 1361.0 | 1361.0 |
| bhikkhu | Theravada/Pali | Pali monastic term | 0 | 0 |  |  |

## Example contexts

| term | group | inscription | date | context |
| --- | --- | --- | --- | --- |
| lokeśvara | Mahayana | K.168:{3} | 972 | ekādaśamukha nu vraḥ kaṃmrateṅ ’añ lokeśvara nu {4} vraḥ kaṃmrateṅ ’añ |
| lokeśvara | Mahayana | K.168:{10} | 972 | nā {11} vraḥ kaṃmrateṅ ’añ lokeśvara ○ srū ta kh’val neḥ |
| lokeśvara | Mahayana | K.240S-2:{7} | 979 | khñuṃ ta vraḥ kamrateṅ ’añ lokeśvara ○ gho kapur {8} tai |
| lokeśvara | Mahayana | K.158:{C2} | 1003 | nu vraḥ {C3} kaṃmrateṅ ’añ lokeśvara {C4} ta praśasta ○ bhūmi |
| lokeśvara | Mahayana | K.230:{C15} | 1026 | neḥh saṃ sit {C16} ’arccā lokeśvara jmaḥ kamrate{C17} ṅ ’añ śrī |
| bodhisattva | Mahayana | K.294:{1} | 1178-1277 | mahābodhi kuraba{2} bṛkṣa [○] vraḥ bodhisattva jā vraḥ indra paṃre ta |
| mahābodhi | Mahayana | K.294:{1} | 1178-1277 | kamrateṅ ’añ dharmmadarṣi ○ vraḥ mahābodhi kuraba{2} bṛkṣa [○] vraḥ bodhisattva |
| vajra | Mahayana | K.809N:{45} | 878-887 | cap pi hau \| sī vajra \| tai kanteṃ \| {46} |
| vajra | Mahayana | K.366:{C7} | 1139 | □ □ □ [ga]ṇa samrit vajra 1 trayvaṅ=gaṃ 3 cirā dhūpa |
| vajra | Mahayana | K.200:{B3} | 1145 | kamandalu samrit mvāy trayvaṅ prām vajra vya{B4} [r jvan] kriyā vraḥ |
| buddha | Buddhist-general | K.294:{1} | 1178-1277 | {1} vraḥ buddha kamrateṅ ’añ dharmmadarṣi ○ vraḥ |
| buddha | Buddhist-general | K.294:{2} | 1178-1277 | vraḥ indra paṃre ta vraḥ buddha kamra{3} teṅ ’añ nu pañcāṅgikatūryya |
| buddha | Buddhist-general | K.413:{A46} | 1361 | . . . . braḥ buddha . . . . . |
| buddha | Buddhist-general | K.413:{B48} | 1361 | pvas ta sāsana braḥ {B49} buddha kamrateṅ ’añ ruv neḥ ’añ |
| buddha | Buddhist-general | K.413:{B50} | 1361 | ’añ ’aṃpān jā {B51} braḥ buddha pi nāṃ satva phoṅ chloṅ |
| stūpa | Buddhist-general | K.177:{40} | 1278-1477 | tiya saṅgāyanā {41} coṅ braḥ stūpa prā{42} sāda sarap sthā{43} panā |
| saṅgha | Buddhist-general | K.258:{A38} | 1107 | 2 liṅ 18 vraḥ uttarā{A39} saṅgha 1 liṅ 1 śaṅkha ’so |
| saṅgha | Buddhist-general | K.177:{19} | 1278-1477 | {20} gi saṅgharāja nu kray saṅgha phoṅ {21} nāṃ mahāpurusa dau |
| saṅgha | Buddhist-general | K.177:{25} | 1278-1477 | sila{26} viryyādhika bodhisambhāra tarāp {27} saṅgha paricāra mahāpurusa {28} hai sādhu |
| buddhasāsa | Buddhist-general | K.177:{10} | 1278-1477 | sa{11} rvvakusala taṃkal braḥ {12} buddhasāsa parātha kusa{13} la mahājana phoṅ |
| sāsana | Theravada/Pali | K.413:{B48} | 1361 | punya ti ’añ pvas ta sāsana braḥ {B49} buddha kamrateṅ ’añ |
| thera | Theravada/Pali | K.177:{33} | 1278-1477 | piy kamrateṅ {34} ’añ muggaliputtatissa{35} thera karuṅ śrī dharmmāso{36} ka mvay |
| thera | Theravada/Pali | K.754B:{8} | 1308 | śrīndramahādeva prasāda ta ma{9} hāsvāmī thera śrī śrīndramaulīdeva jā saṅkalpa paṃre |
| śiva | Shaiva | K.430:{5} | 578-677 | . . . . {6} śiva □ davā . . . |
| śiva | Shaiva | K.46:{A3} | 578-677 | . . [’aṃnoy] {A4} [mrat]āñ śiva kñuṃ vraḥ ku ’□ □ |
