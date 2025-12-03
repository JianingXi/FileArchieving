import os
import asyncio
from pathlib import Path
from playwright.async_api import async_playwright

# ====== 使用 Playwright 将本地 HTML（含微信图片）渲染为 PDF ======

async def convert_html_to_pdf_with_playwright(html_path, pdf_path):
    abs_html_path = Path(html_path).resolve()
    file_url = abs_html_path.as_uri()

    print(f"▶ 正在渲染（含图片）: {html_path}")

    async with async_playwright() as pw:
        browser = await pw.chromium.launch(headless=True)

        context = await browser.new_context(
            locale="zh-CN",
            java_script_enabled=True,
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/120.0 Safari/537.36"
            ),
            extra_http_headers={
                "Referer": "https://mp.weixin.qq.com/"  # 微信图片防盗链关键
            }
        )

        page = await context.new_page()

        # ---- 全局请求放行（兼容旧版 Playwright：route.continue_）----
        await context.route("**/*", lambda route: route.continue_())

        # ---- 打开本地 HTML 文件 ----
        await page.goto(file_url, wait_until="load")
        await page.wait_for_load_state("networkidle")

        # ---- 强制加载 lazyload 图片 ----
        await page.evaluate("""
            document.querySelectorAll('img').forEach(img => {
                if (img.dataset && img.dataset.src) {
                    img.src = img.dataset.src;
                }
            });
        """)

        await page.wait_for_load_state("networkidle")

        # ---- 导出 PDF ----
        await page.pdf(
            path=pdf_path,
            format="A4",
            print_background=True,
            margin={"top": "20mm", "bottom": "20mm", "left": "15mm", "right": "15mm"}
        )

        await browser.close()

    print(f"✔ 完成: {pdf_path}")
    return True


# ========= 批处理主函数（保持你的原始命名） =========

def convert_html_files_in_directory(directory):
    html_files = []
    for root, dirs, files in os.walk(directory):
        for f in files:
            if f.lower().endswith(('.html', '.htm')):
                html_files.append(os.path.join(root, f))

    if not html_files:
        print("目录下没有 HTML 文件")
        return

    async def _runner():
        for html_path in html_files:
            pdf_path = os.path.splitext(html_path)[0] + ".pdf"
            try:
                await convert_html_to_pdf_with_playwright(html_path, pdf_path)
            except Exception as e:
                print(f"❌ 失败: {html_path} → {e}")

    asyncio.run(_runner())


# ========= 启动入口 =========

if __name__ == "__main__":
    target_directory = r"C:\MyDocument\ToDoList\D20_ToDailyNotice"
    print("开始批量转换 HTML → PDF（含图片）")
    convert_html_files_in_directory(target_directory)
    print("全部完成！")
