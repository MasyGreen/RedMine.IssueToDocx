#Convert RedMine Issue to DOCX

1. Set settings in **config.cfg** (run once to create struct file)
   1. host, IP or DNS name (example: **http://192.168.1.1**)
   2. apikey, *RedMine - User - API key*, RESTAPI must by On (example: **aldjfoeiwgj9348gn348**)
   2. issuesid, convert Issue ID list split *";"* (example: **1;2;114;9123**)
3. Run

##Sample config.cfg
```
[Settings]
host = http://192.168.1.1
apikey = dq3inqgnqe8igqngninkkvekmviewrgir9384
issuesid = 1677;318
```

#Result
Util crete files name *"Issue - [Issue ID].docx"* in new Folder

#How use pyinstaller whit HtmlToDocx

1. Run in CMD: *pyinstaller -F -i "Icon.ico" RedMineIssueToDocx.py*
2. Edit **RedMineIssueToDocx.spec**? 
   1. Add at first line **import sys**
   2. Add at first line **from os import path**
   3. Update section in Analysis, datas=**[(path.join("c:\\Python38\\Lib\\site-packages","docx","templates"), "docx/templates")]** where *c:\\Python38\\Lib\\site-packages* is your path to *docx/templates* folder
3. Run in CMD: *pyinstaller RedMineIssueToDocx.spec*
4. Result in **dist\RedMineIssueToDocx.exe**

##Sample RedMineIssueToDocx.spec
```
# -*- mode: python ; coding: utf-8 -*-

import sys
from os import path

block_cipher = None


a = Analysis(['RedMineIssueToDocx.py'],
             pathex=[],
             binaries=[],
             datas=[(path.join("c:\\Python38\\Lib\\site-packages","docx","templates"), "docx/templates")],
             hiddenimports=[],
             hookspath=[],
             hooksconfig={},
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher,
             noarchive=False)
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)

exe = EXE(pyz,
          a.scripts,
          a.binaries,
          a.zipfiles,
          a.datas,  
          [],
          name='RedMineIssueToDocx',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          upx_exclude=[],
          runtime_tmpdir=None,
          console=True,
          disable_windowed_traceback=False,
          target_arch=None,
          codesign_identity=None,
          entitlements_file=None , icon='Icon.ico')

```


