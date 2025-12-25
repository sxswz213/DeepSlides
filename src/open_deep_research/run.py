import csv, asyncio, pathlib, traceback
from tqdm import tqdm   
from time import sleep
# import os

# # HTTP 代理
# os.environ["http_proxy"] = "http://127.0.0.1:7890"
# os.environ["https_proxy"] = "http://127.0.0.1:7890"
from dotenv import load_dotenv
load_dotenv()
                    # ★ 1. 引入 tqdm
from graph import graph, Configuration, ReportStateInput,RunnableConfig
         
config = RunnableConfig

async def run_one(row: dict):
    """
    把 CSV 中的一行转换成输入状态并执行流程
    """
    try:
        # —— 1. 构造初始 State —— #
        init_state = ReportStateInput(
            topic=row["Topic"],
            image_path=row.get("image_path", "") or None,
            style=row.get("style", "") or None,
            presentation_minutes=int(row.get("presentation_minutes", 12)),
        )
        # —— 2. 触发执行 —— #
        result: dict = await graph.ainvoke(init_state, timeout=12000)
        
        # —— 3. 后处理（保存日志、结果路径等）—— #
        # out_dir = pathlib.Path("saves") / init_state.topic
        # out_dir.mkdir(parents=True, exist_ok=True)
        # # 保存最终报告
        # (out_dir / "final_report.md").write_text(result.get("final_report", ""), encoding="utf-8")
        # 如果有最终 PPT
        # final_ppt = result.get("final_ppt_path")
        # if final_ppt:
        #     # 拷贝或移动到输出目录
        #     pathlib.Path(final_ppt).rename(out_dir / pathlib.Path(final_ppt).name)

        # print(f"✅  [{init_state.topic}] 完成")

    except Exception as e:
        print(f"❌  [{row.get('Topic')}] 执行失败: {e}")
        traceback.print_exc()

async def main(csv_path="data/unused_topics_with_pdf.csv", start=0, max_rows=100):
    with open(csv_path, newline='', encoding="utf-8") as f:
        rows = list(csv.DictReader(f))

    if max_rows > 0:
        rows = rows[start:start + max_rows]

    total = len(rows)
    print(f"👉 本次 CSV 共 {total} 条记录，将依次生成报告 …")
    # ★ 2. 在 for 循环外包一层 tqdm
    for row in tqdm(rows, desc="生成报告进度"):
        out_dir = pathlib.Path("saves_test") / "outlines" / row["Topic"]
        if out_dir.exists():
            print(f"⚠️  [{row['Topic']}] 已存在目录，跳过。")
            continue
        sleep(61)
        await run_one(row)

if __name__ == "__main__":
    asyncio.run(main())
