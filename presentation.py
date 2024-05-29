from os import path
import sys
import pandas as pd

sys.path.append(path.abspath("../chart_class"))
from figure import GridFigure
from ppt import PPT, SlideContent
from pptx.util import Inches, Pt, Cm
from typing import List, Union, Optional
import matplotlib.pyplot as plt
from wjx import ResultNumericValue, ResultSingleChoice, ResultMultipleChoice
from data_clean import clean_data

try:
    from typing import Literal
except ImportError:
    from typing_extensions import Literal

DICT_WEIGHTS = {
    "＜20%": 0.1,
    "20-40%": 0.3,
    "40-60%": 0.5,
    "60-80%": 0.7,
    "＞80%": 0.9,
}

DICT_COLOR_BY_SOURCE = {
    "门诊": "navy",
    "病房": "crimson",
    "门诊+病房": "purple",
}

D_BREAKOUT = {"大区": ["东1区", "东2区", "中区", "北区", "南区", "西区"]}
D_LAYOUT = {6: (3, 2)}


class PPT_survey(PPT):

    def add_content_standard(
        self,
        result: Union[ResultSingleChoice, ResultMultipleChoice, ResultNumericValue],
        col_breakout: Optional[str] = None,
        width: float = 8,
        height: float = 6,
        fontsize: float = 12,
    ) -> SlideContent:
        c = self.add_content_slide()
        c.set_title(result.col_question)

        f = plt.figure(
            FigureClass=GridFigure,
            width=width,
            height=height,
            fontsize=fontsize,
            style={
                "title": f"{result.col_question} - 全国\n({result.qtype}, n={result.valid_n})",
            },
        )

        f.plot(
            kind="barh",
            data=result.get_stats(),
            fmt="{:.1%}",
            ax_index=0,
            style={
                "remove_xticks": True,
                "show_legend": False,
                "xlabel": (
                    f"加权平均：{result.weighted_avg():.1%}"
                    if result.weighted_avg()
                    else None
                ),
            },
            label_threshold=0,
        )

        c.add_image(
            f.save(),
            width=c.body.width / 2 * 0.9,
            loc=c.body.fraction(dimension="width", frac_n=2, index=1).center,
        )

        if col_breakout:
            f = plt.figure(
                FigureClass=GridFigure,
                width=8,
                height=6,
                sharex=True,
                nrows=D_LAYOUT.get(len(D_BREAKOUT[col_breakout]))[0],
                ncols=D_LAYOUT.get(len(D_BREAKOUT[col_breakout]))[1],
                fontsize=fontsize - 1,
                style={
                    "title": f"{result.col_question} - 分{col_breakout}\n({result.qtype}, n={result.valid_n})",
                    "label_outer": True,
                },
            )

            df = result.get_stats(col_breakout=col_breakout)
            for i, bk in enumerate(df.columns):
                f.plot(
                    kind="barh",
                    data=df[bk],
                    fmt="{:.1%}",
                    ax_index=i,
                    style={
                        "title": bk,
                        "remove_xticks": True,
                        "show_legend": False,
                        "xlabel": (
                            f"加权平均：{result.weighted_avg(col_breakout=bk).loc[bk]:.1%}"
                            if result.weighted_avg(col_breakout=bk)
                            else None
                        ),
                    },
                    label_threshold=0,
                )

            c.add_image(
                f.save(),
                width=c.body.width / 2 * 0.9,
                loc=c.body.fraction(dimension="width", frac_n=2, index=2).center,
            )

        return c

    def add_content_slide_in_and_out(
        self,
        col_name: str,
        breakout: Optional[str] = None,
    ) -> SlideContent:
        d_map = {
            "患者中CKD占比": "CKD患者数",
            "CKD患者中ND占比": "ND-CKD患者数",
            "ND-CKD患者中3-5期占比": "ND-CKD3-5期患者数",
        }

        d_map_question = {
            "患者数": "1. 在过去的半年，您平均每月诊治的患者数约为：门诊____人；病房____人",
            "患者中CKD占比": "2. 在您诊治的所有患者中，CKD患者的比例约为：门诊____%；病房____%",
            "CKD患者中ND占比": "3. 在您诊治的CKD患者中，非透析的比例约为：门诊____%；病房____%",
            "ND-CKD患者中3-5期占比": "4. 在您诊治的非透析CKD患者中，CKD3-5期的患者比例约为：门诊____%；病房____%",
        }

        c = self.add_content_slide()

        if breakout is None:
            if "占比" in col_name:
                title = f"门诊/病房{col_name} & {d_map[col_name]} - 全国"
            else:
                title = f"门诊/病房{col_name} - 全国"
            c.set_title(title)

            patients_n = ResultNumericValue(df, f"门诊+病房{col_name}")

            f = plt.figure(
                FigureClass=GridFigure,
                width=15,
                height=6,
                ncols=3,
                nrows=2 if "占比" in col_name else 1,
                fontsize=12,
                style={
                    "title": f"{title}\n({patients_n.qtype}, n={patients_n.valid_n})",
                    # "label_outer": True,
                },
            )

            for i, source in enumerate(["门诊", "病房", "门诊+病房"]):
                patients_n = ResultNumericValue(df, f"{source}{col_name}")
                f.plot(
                    kind="hist",
                    data=patients_n.data,
                    ax_index=i,
                    fmt="{:.1%}" if "占比" in col_name else "{:.0f}",
                    style={
                        "ylabel": "频数",
                        "xlabel": col_name,
                        "title": source,
                        "show_legend": False,
                        "hide_top_right_spines": True,
                    },
                    color_kde=DICT_COLOR_BY_SOURCE[source],
                    color_mean=DICT_COLOR_BY_SOURCE[source],
                    color_median=DICT_COLOR_BY_SOURCE[source],
                )

                if "占比" in col_name:
                    patients_n = ResultNumericValue(df, f"{source}{d_map[col_name]}")
                    f.plot(
                        kind="hist",
                        data=patients_n.data,
                        ax_index=i + 3,
                        style={
                            "ylabel": "频数",
                            "xlabel": d_map[col_name],
                            "title": source,
                            "show_legend": False,
                            "hide_top_right_spines": True,
                        },
                        color_kde=DICT_COLOR_BY_SOURCE[source],
                        color_mean=DICT_COLOR_BY_SOURCE[source],
                        color_median=DICT_COLOR_BY_SOURCE[source],
                    )
        else:
            if "占比" in col_name:
                title = f"门诊/病房{col_name} & {d_map[col_name]} - 分{breakout}"
            else:
                title = f"门诊/病房{col_name} - 分{breakout}"
            c.set_title(title)

            patients_n = ResultNumericValue(df, f"门诊+病房{col_name}")
            f = plt.figure(
                FigureClass=GridFigure,
                width=15,
                height=6,
                ncols=3,
                nrows=2 if "占比" in col_name else 1,
                fontsize=11,
                style={
                    "title": f"{title}\n({patients_n.qtype}, n={patients_n.valid_n})",
                    # "label_outer": True,
                },
            )

            for i, source in enumerate(["门诊", "病房", "门诊+病房"]):
                patients_n = ResultNumericValue(df, f"{source}{col_name}")
                f.plot(
                    kind="bar",
                    data=patients_n.get_stats(col_breakout="大区"),
                    ax_index=i,
                    fmt="{:.1%}" if "占比" in col_name else "{:.0f}",
                    style={
                        "ylabel": f"{col_name}（平均值)",
                        "title": source,
                        "show_legend": False,
                        "remove_yticks": True,
                    },
                    color_bar=DICT_COLOR_BY_SOURCE[source],
                )

                if "占比" in col_name:
                    patients_n = ResultNumericValue(df, f"{source}{d_map[col_name]}")
                    f.plot(
                        kind="bar",
                        data=patients_n.get_stats(col_breakout="大区"),
                        ax_index=i + 3,
                        style={
                            "ylabel": f"{d_map[col_name]}（平均值)",
                            "title": source,
                            "show_legend": False,
                            "remove_yticks": True,
                        },
                        color_bar=DICT_COLOR_BY_SOURCE[source],
                    )

        c.add_image(
            f.save(),
            width=c.body.width * 0.9,
            height=None,
            loc=c.body.center,
        )

        self.add_question(c, d_map_question.get(col_name, ""))

        return c

    def add_question(self, slide: SlideContent, text: str):
        slide.add_text(
            text,
            width=slide.body.width * 0.9,
            height=Cm(1),
            loc=slide.body.bottom_mid,
            anchor="center",
            font_italic=True,
            font_size=Pt(10),
        )


if __name__ == "__main__":
    df = clean_data("265857608_按文本_ND-CKD患者肾性贫血治疗观念调研_107_90.xlsx")
    print(df.columns)

    p = PPT_survey("template.pptx")

    # """门诊病房患者数"""
    # p.add_content_slide_in_and_out("患者数")
    # p.add_content_slide_in_and_out("患者数", breakout="大区")

    # """门诊/病房患者中CKD占比"""
    # p.add_content_slide_in_and_out("患者中CKD占比")
    # p.add_content_slide_in_and_out("患者中CKD占比", breakout="大区")

    # """门诊/病房CKD患者中ND占比"""
    # p.add_content_slide_in_and_out("CKD患者中ND占比")
    # p.add_content_slide_in_and_out("CKD患者中ND占比", breakout="大区")

    # """门诊/病房ND-CKD患者中3-5期占比"""
    # p.add_content_slide_in_and_out(
    #     "ND-CKD患者中3-5期占比",
    # )
    # p.add_content_slide_in_and_out("ND-CKD患者中3-5期占比", breakout="大区")

    """ Hb测量时机和贫血发生率 """
    p.add_content_standard(ResultSingleChoice(df, "Hb测量时机"), col_breakout="大区")

    # anema_prevalance_1_2 = ResultSingleChoice(
    #     df["ND-CKD1-2期合并肾性贫血比例"], weights=DICT_WEIGHTS
    # )
    # anema_prevalance_3_5 = ResultSingleChoice(
    #     df["ND-CKD3-5期合并肾性贫血比例"], weights=DICT_WEIGHTS
    # )
    # c.add_image(
    #     anema_prevalance_1_2.plot(
    #         width=6, height=3, sorter=["＞80%", "60-80%", "40-60%", "20-40%", "＜20%"]
    #     ),
    #     width=c.body.width / 3 * 0.9,
    #     loc=c.body.fraction(dimension="width", frac_n=2, index=2)
    #     .fraction(dimension="height", frac_n=2, index=1)
    #     .center,
    # )
    # c.add_image(
    #     anema_prevalance_3_5.plot(width=6, height=3),
    #     width=c.body.width / 3 * 0.9,
    #     loc=c.body.fraction(dimension="width", frac_n=2, index=2)
    #     .fraction(dimension="height", frac_n=2, index=2)
    #     .center,
    # )

    # """ 每月相关病人数推算 """

    # c = p.add_content_slide()
    # c.set_title("每月相关病人数推算")

    # d_map = {
    #     "门诊": "navy",
    #     "病房": "crimson",
    #     "门诊+病房": "purple",
    # }

    # f = plt.figure(
    #     FigureClass=GridFigure,
    #     width=15,
    #     height=6,
    #     ncols=3,
    #     nrows=2,
    #     fontsize=11,
    #     style={
    #         "title": f"每月相关病人数推算\n({anema_prevalance_3_5.qtype}, n={anema_prevalance_3_5.valid_n})"
    #     },
    # )

    # for i, source in enumerate(d_map.keys()):
    #     df_funnel = pd.DataFrame()
    #     df_funnel.loc["患者数", "患者数"] = ResultNumericValue(
    #         df[f"{source}患者数"]
    #     ).get_stats()["平均值"]
    #     df_funnel.loc["CKD患者数", "患者数"] = ResultNumericValue(
    #         df[f"{source}CKD患者数"]
    #     ).get_stats()["平均值"]
    #     df_funnel.loc["ND-CKD患者数", "患者数"] = ResultNumericValue(
    #         df[f"{source}ND-CKD患者数"]
    #     ).get_stats()["平均值"]
    #     df_funnel.loc["ND-CKD3-5期患者数", "患者数"] = ResultNumericValue(
    #         df[f"{source}ND-CKD3-5期患者数"]
    #     ).get_stats()["平均值"]
    #     df_funnel.loc["ND-CKD3-5期合并贫血患者数", "患者数"] = (
    #         ResultSingleChoice(
    #             df["ND-CKD3-5期合并肾性贫血比例"],
    #             weights=DICT_WEIGHTS,
    #         ).weighted_avg()
    #         * df_funnel.loc["ND-CKD3-5期患者数", "患者数"]
    #     )

    #     f.plot(
    #         kind="funnel",
    #         data=df_funnel,
    #         ax_index=i,
    #         style={"title": source},
    #         color=d_map[source],
    #         show_label=True if i == 0 else False,
    #         bbox=None,
    #     )

    # for i, source in enumerate(d_map.keys()):
    #     df_funnel = pd.DataFrame()
    #     df_funnel.loc["患者数", "患者数"] = ResultNumericValue(
    #         df[f"{source}患者数"]
    #     ).get_stats()["平均值"]
    #     df_funnel.loc["CKD患者数", "患者数"] = ResultNumericValue(
    #         df[f"{source}CKD患者数"]
    #     ).get_stats()["平均值"]
    #     df_funnel.loc["ND-CKD患者数", "患者数"] = ResultNumericValue(
    #         df[f"{source}ND-CKD患者数"]
    #     ).get_stats()["平均值"]
    #     df_funnel.loc["ND-CKD1-2期患者数", "患者数"] = ResultNumericValue(
    #         df[f"{source}ND-CKD1-2期患者数"]
    #     ).get_stats()["平均值"]
    #     df_funnel.loc["ND-CKD1-2期合并贫血患者数", "患者数"] = (
    #         ResultSingleChoice(
    #             df["ND-CKD1-2期合并肾性贫血比例"],
    #             weights=DICT_WEIGHTS,
    #         ).weighted_avg()
    #         * df_funnel.loc["ND-CKD1-2期患者数", "患者数"]
    #     )

    #     f.plot(
    #         kind="funnel",
    #         data=df_funnel,
    #         ax_index=i + 3,
    #         style={"title": source},
    #         color=d_map[source],
    #         show_label=True if i == 0 else False,
    #         bbox=None,
    #     )

    # c.add_image(
    #     f.save(),
    #     width=c.body.width * 0.9,
    #     height=None,
    #     loc=c.body.center,
    # )

    # """ 就诊肾性贫血患者的Hb值分布 """

    # c = p.add_content_slide()
    # c.set_title("就诊肾性贫血患者的Hb值分布")

    # hb_distribution = pd.DataFrame()
    # hb_range = [
    #     "Hb>110g/L的患者比例",
    #     "Hb101-110g/L的患者比例",
    #     "Hb91-100g/L的患者比例",
    #     "Hb81-90g/L的患者比例",
    #     "Hb≤80g/L的患者比例",
    # ]
    # for hb in hb_range:
    #     hb_distribution.loc[hb, "比例"] = ResultNumericValue(df[hb]).get_stats()[
    #         "平均值"
    #     ]

    # f = plt.figure(
    #     FigureClass=GridFigure,
    #     width=15,
    #     height=6,
    #     fontsize=12,
    #     style={
    #         "title": f"就诊肾性贫血患者的Hb值分布\n({ResultNumericValue(df[hb_range[0]]).qtype}, n={ResultNumericValue(df[hb_range[0]]).valid_n})",
    #     },
    # )

    # f.plot(
    #     kind="barh",
    #     data=hb_distribution,
    #     fmt="{:.1%}",
    #     ax_index=0,
    #     style={},
    # )

    # c.add_image(
    #     f.save(),
    #     width=c.body.width * 0.9,
    #     height=None,
    #     loc=c.body.center,
    # )

    # """ ND-CKD患者的HIF-PHI总体使用比例 & 治疗启动时机 """

    # c = p.add_content_slide()
    # c.set_title("ND-CKD患者的HIF-PHI总体使用比例 & 治疗启动时机")

    # hif_usage = ResultNumericValue(df["HIF总体使用比例"])
    # c.add_image(
    #     hif_usage.plot(width=8, height=7, fmt="{:.1%}"),
    #     width=c.body.width / 2 * 0.9,
    #     loc=c.body.fraction(dimension="width", frac_n=2, index=1).center,
    # )

    # hif_timing = ResultSingleChoice(df["HIF治疗启动时机"])
    # c.add_image(
    #     hif_timing.plot(
    #         sorter=["低于110g/L", "低于105g/L", "低于100g/L", "低于95g/L", "低于90g/L"],
    #         width=8,
    #         height=7,
    #     ),
    #     width=c.body.width / 2 * 0.9,
    #     loc=c.body.fraction(dimension="width", frac_n=2, index=2).center,
    # )

    # """ 新诊断不同Hb基线患者的HIF-PHI使用比例 """

    # c = p.add_content_slide()
    # c.set_title("新诊断不同Hb基线患者的HIF-PHI使用比例")

    # hb_range = [
    #     ">110g/L",
    #     "101-110g/L",
    #     "91-100g/L",
    #     "81-90g/L",
    #     "≤80g/L",
    # ]
    # f = plt.figure(
    #     FigureClass=GridFigure,
    #     width=15,
    #     height=6,
    #     ncols=len(hb_range),
    #     fontsize=12,
    #     style={
    #         "title": f"新诊断不同Hb基线患者的HIF-PHI使用比例\n({ResultSingleChoice(df['基线>110g/L新诊患者HIF使用比例']).qtype}, n={ResultSingleChoice(df['基线>110g/L新诊患者HIF使用比例']).valid_n})",
    #         "label_outer": True,
    #     },
    # )

    # for hb in hb_range:
    #     hif_usage = ResultSingleChoice(
    #         df[f"基线{hb}新诊患者HIF使用比例"], weights=DICT_WEIGHTS
    #     )
    #     f.plot(
    #         kind="barh",
    #         data=hif_usage.get_stats(
    #             sorter=["＞80%", "60-80%", "40-60%", "20-40%", "＜20%"]
    #         )["百分比"],
    #         ax_index=hb_range.index(hb),
    #         fmt="{:.1%}",
    #         style={
    #             "title": hb,
    #             "ylabel": "HIF-PHI使用比例",
    #             "xlabel": f"加权平均值: {hif_usage.weighted_avg():.1%}",
    #             "show_legend": False,
    #             "remove_xticks": True,
    #         },
    #         label_threshold=0,
    #     )

    # c.add_image(
    #     f.save(),
    #     width=c.body.width * 0.9,
    #     height=None,
    #     loc=c.body.center,
    # )

    # """ HIF治疗2个月的预期 """

    # c = p.add_content_slide()
    # c.set_title("HIF治疗2个月的预期")

    # expectation_2m = ResultSingleChoice(df["HIF治疗2个月的预期"])

    # c.add_image(
    #     expectation_2m.plot(width=15, height=6),
    #     width=c.body.width * 0.9,
    #     height=None,
    #     loc=c.body.center,
    # )

    # """ HIF治疗2个月不达预期的处理方式 """

    # c = p.add_content_slide()
    # c.set_title("HIF治疗2个月不达预期的处理方式")

    # behavior_below_expectation = ResultSingleChoice(
    #     df["HIF治疗2个月不达预期的处理方式"]
    # )

    # c.add_image(
    #     behavior_below_expectation.plot(width=15, height=6),
    #     width=c.body.width * 0.9,
    #     height=None,
    #     loc=c.body.center,
    # )

    # """ 处方罗沙司他的顾虑 """

    # c = p.add_content_slide()
    # c.set_title("处方罗沙司他的顾虑")

    # roxa_concern = ResultMultipleChoice(df["处方罗沙司他的顾虑"])

    # c.add_image(
    #     roxa_concern.plot(width=15, height=6),
    #     width=c.body.width * 0.9,
    #     height=None,
    #     loc=c.body.center,
    # )

    # """ 恩那罗最吸引人的特点 """

    # c = p.add_content_slide()
    # c.set_title("恩那罗最吸引人的特点")

    # enar_most_attractive = ResultSingleChoice(
    #     df["恩那罗最吸引人的特点（除价格和服药方式）"]
    # )

    # c.add_image(
    #     enar_most_attractive.plot(width=15, height=6),
    #     width=c.body.width * 0.9,
    #     height=None,
    #     loc=c.body.center,
    # )

    p.save("test.pptx")
