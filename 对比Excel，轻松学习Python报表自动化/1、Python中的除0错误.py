import numpy as np
import pandas as pd
df = pd.DataFrame({"col1": [-1, 2, 3]})
# df.sort_values(by=["col1", "col2"], ascending=[False, True])
df["new_col"] = df["col1"] / 0
# df["new_col"] = df["new_col"].replace([np.inf, -np.inf], 0)  # 替换
df1 = df[df["new_col"] != np.inf]
print(df1)

