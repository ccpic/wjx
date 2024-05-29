import pandas as pd
import numpy as np

D_MAP_REGION = {
    "天津‐天津市‐天津医科大学朱宪彝纪念医院(天津医科大学代谢病医院)": "北区"
}

D_MAP_ITEM = {
    "新诊断的患者每个都测，复诊患者固定频率测量（例如：一月一次）": "新诊每个都测，复诊固定频率",
    "只有当患者主诉有贫血相关症状时或当CKD有进展时测": "只当有症状或CKD进展时测",
    "每个患者固定频率测量（例如：一月一次）": "每个患者固定频率",
    "每个患者每次就诊都测": "每个患者每次就诊都测",
}


def drop_outlier(df: pd.DataFrame, col: str, iqr_index: float = 3) -> pd.DataFrame:
    # 计算第一四分位数（Q1）和第三四分位数（Q3）
    Q1 = df[col].quantile(0.25)
    Q3 = df[col].quantile(0.75)

    # 计算四分位数间距（IQR）
    IQR = Q3 - Q1

    # 定义上下限（通常为1.5倍的IQR）
    lower_bound = Q1 - iqr_index * IQR
    upper_bound = Q3 + iqr_index * IQR

    # 去除极端值
    df_filtered = df[(df[col] >= lower_bound) & (df[col] <= upper_bound)]

    # 找到被去除的值
    df_outlier = df[~df.index.isin(df_filtered.index)]

    print(f"根据[{col}]列去除极端值：")
    print(df_outlier[["医院", "姓名", col]])

    return df_filtered


def clean_data(file_path: str) -> pd.DataFrame:
    df = pd.read_excel(file_path)

    # 去掉无用列
    df.drop(["总分"], axis=1, inplace=True)

    # 简化列名
    df_q = pd.read_excel("设置.xlsx", sheet_name="题目映射")
    df.rename(columns=df_q.set_index("原始列名")["简化列名"].to_dict(), inplace=True)

    # 百分比字段转换为小数
    for col in [
        "门诊患者中CKD占比",
        "病房患者中CKD占比",
        "门诊CKD患者中ND占比",
        "病房CKD患者中ND占比",
        "门诊ND-CKD患者中3-5期占比",
        "病房ND-CKD患者中3-5期占比",
        "Hb>110g/L的患者比例",
        "Hb101-110g/L的患者比例",
        "Hb91-100g/L的患者比例",
        "Hb81-90g/L的患者比例",
        "Hb≤80g/L的患者比例",
        "HIF总体使用比例",
    ]:
        df[col] = df[col] / 100

    # 详细计算病人数相关
    df["门诊+病房患者数"] = df["门诊患者数"] + df["病房患者数"]
    df["门诊CKD患者数"] = df["门诊患者数"] * df["门诊患者中CKD占比"]
    df["病房CKD患者数"] = df["病房患者数"] * df["病房患者中CKD占比"]
    df["门诊+病房CKD患者数"] = df["门诊CKD患者数"] + df["病房CKD患者数"]
    df["门诊+病房患者中CKD占比"] = df["门诊+病房CKD患者数"] / df["门诊+病房患者数"]
    df["门诊ND-CKD患者数"] = df["门诊CKD患者数"] * df["门诊CKD患者中ND占比"]
    df["病房ND-CKD患者数"] = df["病房CKD患者数"] * df["病房CKD患者中ND占比"]
    df["门诊+病房ND-CKD患者数"] = df["门诊ND-CKD患者数"] + df["病房ND-CKD患者数"]
    df["门诊+病房CKD患者中ND占比"] = (
        df["门诊+病房ND-CKD患者数"] / df["门诊+病房CKD患者数"]
    )
    df["门诊ND-CKD1-2期患者数"] = df["门诊ND-CKD患者数"] * (
        1 - df["门诊ND-CKD患者中3-5期占比"]
    )
    df["病房ND-CKD1-2期患者数"] = df["病房ND-CKD患者数"] * (
        1 - df["病房ND-CKD患者中3-5期占比"]
    )
    df["门诊+病房ND-CKD1-2期患者数"] = (
        df["门诊ND-CKD1-2期患者数"] + df["病房ND-CKD1-2期患者数"]
    )
    df["门诊+病房ND-CKD患者中1-2期占比"] = (
        df["门诊+病房ND-CKD1-2期患者数"] / df["门诊+病房ND-CKD患者数"]
    )
    df["门诊ND-CKD3-5期患者数"] = (
        df["门诊ND-CKD患者数"] * df["门诊ND-CKD患者中3-5期占比"]
    )
    df["病房ND-CKD3-5期患者数"] = (
        df["病房ND-CKD患者数"] * df["病房ND-CKD患者中3-5期占比"]
    )
    df["门诊+病房ND-CKD3-5期患者数"] = (
        df["门诊ND-CKD3-5期患者数"] + df["病房ND-CKD3-5期患者数"]
    )
    df["门诊+病房ND-CKD患者中3-5期占比"] = (
        df["门诊+病房ND-CKD3-5期患者数"] / df["门诊+病房ND-CKD患者数"]
    )

    # 匹配内部架构
    df_internal = pd.read_excel("设置.xlsx", sheet_name="内部架构")
    df["目标名称"] = df["医院"].apply(lambda x: x.split("‐")[-1])
    df["大区"] = df["目标名称"].map(df_internal.set_index("目标名称")["大区"])
    df["大区"] = df["医院"].map(D_MAP_REGION).fillna(df["大区"])

    #  去除极端值
    for col in ["门诊患者数", "病房患者数"]:
        df = drop_outlier(df, col)

    # 简化部分字段值的文本
    df["Hb测量时机"] = df["Hb测量时机"].map(D_MAP_ITEM).fillna("其他")

    return df


if __name__ == "__main__":
    df = clean_data("265857608_按文本_ND-CKD患者肾性贫血治疗观念调研_107_90.xlsx")
    df.to_excel("cleaned.xlsx", index=False)
