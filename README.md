# })MS2RSS*�����쯿�

VBA�(W_})<8MarketSpeed2nRSS APIK�*��������k֗WCSVbg��Y�����

## ��

Sn������o})<8nMarketSpeed2MS2	nRSSReal-time Spreadsheet Support	_��;(Wf*�������֗WCSVbg��Y�Excel VBA�������gY

## ;j_�

- **p����**: ���:�gp�Ē ��
- **�j�.**: 1�5�15�30�60��k��
- **��**: ����B���Wf���֗
- **CSV��**: ��jOHLCVbg��
- **2Wh:**: �뿤�n�2Wh��h:
- **��������**: 4B���ĳ��<

## ǣ����

```
rakuten-ms2rss-collector/
   README.md
   docs/                     # ɭ����
      ms2rss/              # MS2RSS API���
      vba-guide.md         # VBA(�լ��
   vba/                     # VBA������
      src/
         modules/         # VBA����
         forms/           # ����թ��
         classes/         # ������
      templates/           # Excel������
      tests/               # ƹ�(���
   output/                  # ��ա��
      csv/                 # CSVա��
      logs/                # ��ա��
   config/                  # -�ա��
   tools/                   # �z���
       pdf-reader/          # PDF�֊���
```

## Łj��

- **Microsoft Excel**: VBA��HOffice 2016�M�h	
- **})<8�**: MarketSpeed2n)(Q
- **MarketSpeed2**: RSS_�L	�

## ��������Ȣ��

1. Sn�ݸ�꒯���~_o������
2. `vba/templates/StockDataCollector.xlsm` ��O
3. ޯ�n�L�	�
4. MarketSpeed2�w�WRSS_��	�kY�

## (��

### �,�jD�

1. Excelա�뒋M�����թ���w�
2. �ĳ��e��7203, 6758, 9984	
3. �h�.�x�
4. ��Hթ����
5. �Lܿ�g���֗��

### �ĳ��b

```
7203        # q<�թ��	
7203.T      # q<:
7203.JAX    # JAX4
7203.JNX    # JNX4
7203.CHJ    # Chi-X4
```

### ��CSVb

```csv
DateTime,Open,High,Low,Close,Volume
2025-01-14 09:00:00,2500,2520,2495,2510,150000
2025-01-14 09:01:00,2510,2525,2505,2520,120000
```

## ��

- MarketSpeed2nRSS_�Lc8k�\WfD�ŁLB�~Y
- 4B�o ����֗k6PLB�~Y
- '����֗BoAPI6Pk�WfO`UD
- ,j�gn(oAjƹȒ��WfO`UD

## 餻�

MIT License

## M��

Sn��Ȧ��oY��v�gЛU�fD~Y��$���P�kdDf�zo n����D~[�T�n��gT)(O`UD

## �.

а1J�_�9�n�Ho Issues ~gJXDW~Y

## �#��

- [})<8 MarketSpeed2](https://www.rakuten-sec.co.jp/marketspeed/)
- [MS2RSS API ɭ����](./docs/ms2rss/)