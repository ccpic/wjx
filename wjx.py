from os import path
import sys
import pandas as pd

sys.path.append(path.abspath("../chart_class"))
from typing import List, Dict, Optional
import matplotlib.pyplot as plt
from figure import GridFigure
from data_clean import clean_data


class Result:
    def __init__(
        self,
        df: pd.DataFrame,
        col_question: str,
        qtype: str,
    ):
        self.df = df
        self.data = df[col_question]
        self.qtype = qtype
        self.col_question = col_question
        self.total_n = len(self.data)
        self.valid_n = self.data.count()


class ResultSingleChoice(Result):

    def __init__(
        self,
        df: pd.DataFrame,
        col_question: str,
        qtype: str = "单选",
        weights: Optional[Dict[str, float]] = None,
    ):
        super().__init__(
            df,
            col_question,
            qtype,
        )
        self.weights = weights

    def get_n(self, col_breakout: Optional[str] = None) -> pd.Series:
        if col_breakout:
            return self.df.groupby(col_breakout)[self.col_question].count()
        else:
            return self.valid_n

    def get_stats(
        self,
        col_breakout: Optional[str] = None,
        percentage: bool = True,
        sorter: Optional[List[str]] = None,
        add_base: bool = True,
    ) -> pd.DataFrame:

        stats_total = self.data.value_counts()
        if percentage:
            stats_total = stats_total.div(self.valid_n)

        stats_total.sort_values(ascending=False, inplace=True)
        print(stats_total)

        if col_breakout:
            stats_breakout = (
                self.df[[col_breakout, self.col_question]]
                .groupby(col_breakout)[self.col_question]
                .value_counts()
                .unstack(fill_value=0)
                .T
            ).reindex(stats_total.index)
            count = stats_breakout.sum()
            if percentage:
                stats_breakout = stats_breakout.div(count)

            if add_base:
                stats_breakout.columns = (
                    stats_breakout.columns + "\n(n=" + count.astype(str) + ")"
                )

        stats = stats_breakout if col_breakout else stats_total
        if sorter:
            stats = stats.reindex(sorter)

        return stats

    def weighted_avg(
        self,
        col_breakout: Optional[str] = None,
        add_base: bool = True,
    ) -> float:

        if col_breakout:
            try:
                stats_breakout = self.get_stats(col_breakout, percentage=False, add_base=False)
                count = stats_breakout.sum()

                weighted_avg = self.df.groupby(col_breakout)[self.col_question].apply(
                    lambda x: x.map(self.weights).mean()
                )
                
                print(weighted_avg)
                if add_base:
                    weighted_avg.index = (
                        stats_breakout.columns + "\n(n=" + count.astype(str) + ")"
                    )
                    
                return weighted_avg

            except Exception:
                return None
        else:
            try:
                return self.data.map(self.weights).mean()
            except Exception:
                return None


class ResultMultipleChoice(Result):

    def __init__(
        self,
        df: pd.DataFrame,
        col_question: str,
        qtype: str = "多选",
        delimiter: str = "┋",
    ):
        super().__init__(
            df,
            col_question,
            qtype,
        )
        self.delimiter = delimiter

    def get_stats(
        self, col_breakout: Optional[str], sorter: Optional[List[str]] = None
    ) -> pd.DataFrame:
        stats = pd.DataFrame()
        stats["计数"] = self.data.str.split(self.delimiter).explode().value_counts()
        stats["百分比"] = stats["计数"] / self.valid_n

        if col_breakout:
            stats_breakout = (
                self.df.groupby(col_breakout)[self.col_question]
                .apply(
                    lambda x: x.str.split(self.delimiter)
                    .explode()
                    .value_counts(normalize=True)
                )
                .unstack()
                .fillna(0)
            )
            stats_breakout.columns = stats_breakout.columns.map(
                lambda x: f"{self.col_question}={x}"
            )
            stats = stats.join(stats_breakout)

        if sorter:
            stats = stats.reindex(sorter)

        return stats

    def plot(
        self,
        sorter: Optional[List[str]] = None,
        width: float = 15,
        height: float = 6,
        fontsize: float = 12,
    ) -> str:
        f = plt.figure(
            FigureClass=GridFigure,
            width=width,
            height=height,
            fontsize=fontsize,
            style={
                "title": f"{self.column_name}\n({self.qtype}, n={self.valid_n})",
            },
        )

        f.plot(
            kind="barh",
            data=self.get_stats(sorter=sorter)["百分比"],
            fmt="{:.1%}",
            ax_index=0,
            style={
                "remove_xticks": True,
                "show_legend": False,
            },
            label_threshold=0,
        )

        return f.save()


class ResultNumericValue(Result):

    def __init__(
        self,
        data: pd.DataFrame,
        col_question: str,
        qtype: str = "数值填空",
    ):
        super().__init__(
            data,
            col_question,
            qtype,
        )

    def get_stats(self, col_breakout: Optional[str] = None) -> pd.DataFrame:
        stats = pd.Series(dtype=float)
        stats["平均值"] = self.data.mean()
        stats["标准差"] = self.data.std()
        stats["最小值"] = self.data.min()
        stats["25%分位数"] = self.data.quantile(0.25)
        stats["中位数"] = self.data.median()
        stats["75%分位数"] = self.data.quantile(0.75)
        stats["最大值"] = self.data.max()

        if col_breakout:
            stats = (
                self.df[[col_breakout, self.col_question]]
                .groupby(col_breakout)
                .agg(["count", "mean"])
            )
            stats.index = stats.index + "\n(n=" + stats["count"].astype(str) + ")"

        return stats

    def get_stats_by_bins(self, bins: List[float]) -> pd.DataFrame:
        stats = pd.DataFrame()
        stats["计数"] = pd.cut(self.data, bins).value_counts().sort_index()
        stats["百分比"] = stats["计数"] / self.valid_n
        stats["百分比"] = stats["百分比"].map(lambda x: f"{x:.1%}")

        return stats

    def plot(
        self,
        width: float = 15,
        height: float = 6,
        fmt: str = "{:.0f}",
        fontsize: float = 12,
    ) -> str:
        f = plt.figure(
            FigureClass=GridFigure,
            width=width,
            height=height,
            fontsize=fontsize,
            style={
                "title": f"{self.column_name}\n({self.qtype}, n={self.valid_n})",
            },
        )

        f.plot(
            kind="hist",
            data=self.data,
            ax_index=0,
            fmt=fmt,
            style={
                "ylabel": "频数",
                "xlabel": self.column_name,
                "show_legend": False,
            },
        )

        return f.save()


if __name__ == "__main__":
    df = clean_data("265857608_按文本_ND-CKD患者肾性贫血治疗观念调研_107_90.xlsx")
    q1 = ResultSingleChoice(
        df,
        "基线>110g/L新诊患者HIF使用比例",
        weights={
            "＜20%": 0.1,
            "20-40%": 0.3,
            "40-60%": 0.5,
            "60-80%": 0.7,
            "＞80%": 0.9,
        },
    )
    q1.get_stats("大区")
