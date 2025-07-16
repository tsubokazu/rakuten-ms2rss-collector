# VBA�\�[�X�R�[�h - �C���|�[�g�菇

## �T�v

���̃t�H���_�ɂ́A�y�VMS2RSS�����f�[�^�R���N�^�[�̂��ׂĂ�VBA�\�[�X�R�[�h���܂܂�Ă��܂��B

## �t�@�C���\��

### ? modules/ - VBA���W���[��
| �t�@�C���� | ���� | ��v�֐� |
|------------|------|----------|
| **MainModule.bas** | ���C���G���g���[�|�C���g | `ShowMainForm()`, `QuickTest()` |
| **WorksheetMacros.bas** | ���[�N�V�[�g�{�^���p�}�N�� | `StartDataCollection()` �Ȃ� |
| **DataCollector.bas** | �f�[�^�擾�G���W�� | `CollectStockData()` |
| **CSVExporter.bas** | CSV�o�͋@�\ | `ExportStockDataToCSV()` |
| **Utils.bas** | ���[�e�B���e�B�E���O | `LogMessage()`, `ValidateTimeFrame()` |

### ? forms/ - ���[�U�[�t�H�[��
| �t�@�C���� | ���� |
|------------|------|
| **MainForm.frm** | ���C��GUI�t�H�[�� |

### ? classes/ - �N���X���W���[��
| �t�@�C���� | ���� |
|------------|------|
| **StockData.cls** | �����f�[�^�\���N���X |
| **Configuration.cls** | �ݒ�Ǘ��N���X |

## Excel�ւ̃C���|�[�g�菇

### 1. �V����Excel�t�@�C�����쐬
1. Microsoft Excel���N��
2. �V�����u�b�N���쐬
3. �t�@�C������`StockDataCollector.xlsm`�Ƃ��ĕۑ��i�}�N���L���u�b�N�`���j

### 2. VBA�G�f�B�^���J��
1. `Alt + F11`��������VBA�G�f�B�^���J��
2. �v���W�F�N�g�G�N�X�v���[���[��VBAProject���m�F

### 3. �Q�Ɛݒ��ǉ�
1. VBA�G�f�B�^�Łu�c�[���v���u�Q�Ɛݒ�v��I��
2. �ȉ��̍��ڂɃ`�F�b�N������F
   - ? Microsoft Office 16.0 Object Library
   - ? Microsoft Forms 2.0 Object Library
   - ? Microsoft Windows Common Controls 6.0 (SP6)
   - ? Microsoft Windows Common Controls-2 6.0 (SP6)

### 4. ���W���[�����C���|�[�g

#### �W�����W���[�� (.bas)
1. �v���W�F�N�g�G�N�X�v���[���[�ŉE�N���b�N
2. �u�t�@�C���̃C���|�[�g�v��I��
3. �ȉ��̃t�@�C�������ԂɃC���|�[�g�F
   ```
   modules/MainModule.bas
   modules/WorksheetMacros.bas
   modules/DataCollector.bas
   modules/CSVExporter.bas
   modules/Utils.bas
   ```

#### ���[�U�[�t�H�[�� (.frm)
1. �v���W�F�N�g�G�N�X�v���[���[�ŉE�N���b�N
2. �u�t�@�C���̃C���|�[�g�v��I��
3. `forms/MainForm.frm`���C���|�[�g

#### �N���X���W���[�� (.cls)
1. �v���W�F�N�g�G�N�X�v���[���[�ŉE�N���b�N
2. �u�t�@�C���̃C���|�[�g�v��I��
3. �ȉ��̃t�@�C�����C���|�[�g�F
   ```
   classes/StockData.cls
   classes/Configuration.cls
   ```

### 5. ���[�N�V�[�g�̐ݒ�

#### Sheet1�̐ݒ�
1. Sheet1��I�����A�ȉ��̂悤�ɐݒ�F

```
A1: �y�VMS2RSS�����f�[�^�R���N�^�[ v1.0
A3: [�f�[�^���W�J�n] (�{�^��)
A5: [�N�C�b�N�e�X�g] (�{�^��)
A7: [�ڑ��e�X�g] (�{�^��)
A9: [�ݒ�\��] (�{�^��)
A11: [�w���v] (�{�^��)

C3: [�o�̓t�H���_���J��] (�{�^��)
C5: [���O�t�H���_���J��] (�{�^��)
C7: [�o�[�W�������] (�{�^��)
```

#### �{�^���̃}�N�����蓖��
�e�{�^���Ɉȉ��̃}�N�������蓖�āF

| �{�^���� | �}�N���� |
|----------|----------|
| �f�[�^���W�J�n | `StartDataCollection` |
| �N�C�b�N�e�X�g | `RunQuickTest` |
| �ڑ��e�X�g | `TestConnection` |
| �ݒ�\�� | `DisplaySettings` |
| �w���v | `ShowHelp` |
| �o�̓t�H���_���J�� | `OpenOutputFolder` |
| ���O�t�H���_���J�� | `OpenLogFolder` |
| �o�[�W������� | `AboutApp` |

## ��{�I�Ȏg�p���@

### 1. �A�v���P�[�V�����N��
```vba
' ���C���t�H�[����\��
Sub Test_ShowMainForm()
    Call ShowMainForm
End Sub
```

### 2. �N�C�b�N�e�X�g���s
```vba
' �ڑ��ƃf�[�^�擾�̃e�X�g
Sub Test_QuickTest()
    Call QuickTest
End Sub
```

### 3. �v���O��������̒��ڎ��s
```vba
Sub Test_DirectCall()
    Dim result As Boolean
    
    ' �g���^�����Ԃ�5�����f�[�^��1�T�ԕ��擾
    result = CollectStockData("7203", "5M", Date-7, Date)
    
    If result Then
        MsgBox "�f�[�^�擾����"
    Else
        MsgBox "�f�[�^�擾���s"
    End If
End Sub
```

## ��v�֐����t�@�����X

### ShowMainForm()
���C��GUI�t�H�[����\�����ăf�[�^���W���J�n

### CollectStockData(stockCode, timeFrame, startDate, endDate)
- **stockCode**: �����R�[�h�i"7203", "7203.T" �Ȃǁj
- **timeFrame**: ����i"1M", "5M", "15M", "30M", "60M", "D"�j
- **startDate**: �J�n��
- **endDate**: �I����
- **�߂�l**: Boolean�i������True�j

### CollectMultipleStocks(stockCodes, timeFrame, startDate, endDate)
���������̈ꊇ�f�[�^�擾
- **stockCodes**: �J���}��؂�̖����R�[�h�i"7203,6758,9984"�j

## �g���u���V���[�e�B���O

### �悭����G���[

1. **�u�v���V�[�W����������܂���v**
   - ���W���[�����������C���|�[�g����Ă��邩�m�F
   - �Q�Ɛݒ肪�������ݒ肳��Ă��邩�m�F

2. **�uRSS�֐����G���[��Ԃ��܂��v**
   - MarketSpeed2���N�����Ă��邩�m�F
   - RSS�@�\���L���ɂȂ��Ă��邩�m�F

3. **�u�t�@�C�����ۑ��ł��܂���v**
   - �o�̓t�H���_�����݂��邩�m�F
   - �t�H���_�̏������݌������m�F

### �f�o�b�O���@

1. **�X�e�b�v���s**: F8�L�[�ōs�P�ʎ��s
2. **�u���[�N�|�C���g**: F9�L�[�Őݒ�
3. **�C�~�f�B�G�C�g�E�B���h�E**: Ctrl+G�ŕ\��
4. **���O�m�F**: `output/logs/`�t�H���_�̃��O�t�@�C��

## ���ӎ���

- �}�N���̃Z�L�����e�B�ݒ�ŁA�}�N���̎��s�������Ă�������
- MarketSpeed2��RSS�@�\���L���ɂȂ��Ă���K�v������܂�
- ��ʃf�[�^�擾���͏������Ԃ�������ꍇ������܂�
- �{�Ԋ��ł̎g�p�O�ɏ\���ȃe�X�g�����{���Ă�������

## �J�X�^�}�C�Y

### �V��������̒ǉ�
`Utils.bas`��`ValidateTimeFrame`�֐����C��

### �V�����s��̒ǉ�
`DataCollector.bas`��`ValidateStockCode`�֐����C��

### UI�\�����ڂ̕ύX
`MainForm.frm`�̃f�U�C�����C��

�ڍׂȃJ�X�^�}�C�Y���@�́A`docs/vba-guide.md`���Q�Ƃ��Ă��������B