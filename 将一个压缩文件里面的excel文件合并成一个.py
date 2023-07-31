from pathlib import Path
from zipfile import ZipFile
import pandas as pd


zip_path = Path(r'C:\Users\pdcfi\Desktop\压缩包').glob('*.zip')  # 只获取zip后缀的压缩文件
to_path = Path(r'C:\Users\pdcfi\Desktop\res')

# 逐个读取目录中压缩文件
for file in zip_path:
    # 将一个压缩文件里面的excel文件合并成一个
    with ZipFile(file) as zipf:
        df = pd.concat(pd.read_excel(zipf.open(i)) for i in zipf.namelist())
        # 合并后的一个表保存到目标目录中
        df.to_excel(to_path.joinpath(f'{file.stem}.xlsx'), index=False)
