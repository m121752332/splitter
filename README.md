# pysplitter
資料分隔工具:
  使用 python 3.8 or python3.9 搭配 PyQt5 協作開發而成

1. 編譯畫面
  * `pyuic5 pysplitter.ui -o pysplitter.py`
  * `pyrcc5 pysplitter.qrc -o pysplitter_rc.py`
-----
2. 調整程式
  * coding splitter.py
-----
3. 編譯成EXE檔
 * pyinstaller -F -icon="C:\python\splitter\logo.ico" splitter.py
 * pyinstaller -F --icon="logo.ico" --noconsole splitter.py
 * conda create -n splitter_env python==3.8
 * conda activate splitter_env  #激活虛擬環境
 * conda deactivate             #退出虛擬環境
 
 畫面圖
 ------
 ![image](https://user-images.githubusercontent.com/45743812/184643137-aaab6f93-5778-485b-b9ef-5a33841e29db.png)
