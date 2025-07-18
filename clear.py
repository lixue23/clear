import pandas as pd
import streamlit as st
from io import BytesIO
import base64
import os
import sys
from datetime import datetime

# === 安全获取DeepSeek API密钥 ===
deepseek_api_key = None

# 1. 首先尝试从环境变量获取
if 'DEEPSEEK_API_KEY' in os.environ:
    deepseek_api_key = os.environ['DEEPSEEK_API_KEY']

# 2. 尝试从st.secrets获取（使用异常处理）
try:
    # 只有在Streamlit环境中才尝试访问st.secrets
    if hasattr(st, 'secrets') and 'DEEPSEEK_API_KEY' in st.secrets:
        deepseek_api_key = st.secrets['DEEPSEEK_API_KEY']
except Exception:
    pass  # 忽略错误

# 3. 如果以上都失败，尝试从.env文件加载
if not deepseek_api_key and os.path.exists('.env'):
    try:
        from dotenv import load_dotenv

        load_dotenv()
        deepseek_api_key = os.getenv('DEEPSEEK_API_KEY')
    except ImportError:
        pass
    except Exception:
        pass

# 检查关键依赖
try:
    from st_aggrid import AgGrid, GridOptionsBuilder, DataReturnMode, GridUpdateMode
except ImportError:
    st.error("缺少关键依赖: streamlit-aggrid! 请确保requirements.txt中包含该包")
    st.stop()

try:
    from openai import OpenAI
except ImportError:
    st.error("缺少关键依赖: openai! 请确保requirements.txt中包含该包")
    st.stop()

# === 主应用代码 ===
st.set_page_config(page_title="清洗服务记录转换工具", page_icon="🧹", layout="wide")
st.title("🧹 清洗服务记录转换工具")
st.markdown("""
将无序繁杂的清洗服务记录文本转换为结构化的表格数据，并导出为Excel文件。
""")

# 在侧边栏显示API密钥状态
with st.sidebar:
    st.subheader("API密钥状态")

    # 显示系统时间（解决时间同步问题）
    st.caption(f"系统时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

    # 添加手动输入API密钥的选项
    manual_key = st.text_input("手动输入API密钥", type="password", key="manual_api_key")
    if manual_key:
        # 如果用户手动输入了密钥，则使用它
        deepseek_api_key = manual_key

    if deepseek_api_key:
        # 显示部分密钥用于验证
        masked_key = f"{deepseek_api_key[:8]}...{deepseek_api_key[-4:]}" if len(deepseek_api_key) > 12 else "****"
        st.info(f"当前密钥: {masked_key}")

        # 检查密钥格式
        if not deepseek_api_key.startswith("sk-") or len(deepseek_api_key) < 40:
            st.error("⚠️ API密钥格式无效")
            st.info("密钥应以'sk-'开头，长度至少40字符")
        elif " " in deepseek_api_key:
            st.warning("密钥包含空格，已自动清理")
            deepseek_api_key = deepseek_api_key.strip()

        if st.button("重新加载密钥"):
            st.experimental_rerun()
    else:
        st.error("API密钥未配置!")
        st.info("请设置环境变量 DEEPSEEK_API_KEY 或手动输入密钥")
        st.markdown("""
        **本地配置方法:**
        1. 创建 `.env` 文件并添加:
           ```
           DEEPSEEK_API_KEY=sk-your-api-key
           ```
        2. 或在运行前设置环境变量:
           ```bash
           export DEEPSEEK_API_KEY=sk-your-api-key
           streamlit run data.py
           ```
        """)

# 示例文本
sample_text = """
张雨浪 凡尔赛 下午 融创 凡尔赛领馆四期 16栋27-7 15223355185 空调内外机清洗 有异味，可能要全拆洗180，外机在室外150，内机高温蒸汽洗58  未支付 这个要翻外墙，什么时候来

李雪霜 华宇 寸滩派出所楼上 2栋9-8 13983014034 挂机加氟+1空调清洗 加氟一共299 清洗50 未支付

王师傅 龙湖源著 8栋12-3 13800138000 空调维修 不制冷 加氟200 已支付 需要周末上门

刘工 恒大御景半岛 3栋2单元501 13512345678 中央空调深度清洗 全拆洗380 已支付 业主周日下午在家
"""

# 文本输入区域
with st.expander("📝 输入清洗服务记录文本", expanded=True):
    input_text = st.text_area("请输入清洗服务记录（每行一条记录）:",
                              value=sample_text,
                              height=300,
                              placeholder="请输入清洗服务记录文本...")

    # 添加示例下载按钮
    st.download_button("📥 下载示例文本",
                       sample_text,
                       file_name="清洗服务记录示例.txt")

columns = ['师傅', '项目', '地址', '房号', '客户姓名', '电话号码', '服务内容', '费用', '支付状态', '备注']

# 处理按钮
if st.button("🚀 转换文本为表格", use_container_width=True):
    st.session_state['reset_table'] = True

    if not input_text.strip():
        st.warning("请输入清洗服务记录文本！")
        st.stop()

    # 检查API密钥
    if not deepseek_api_key:
        st.error("缺少DeepSeek API密钥！请按照侧边栏说明配置")
        st.stop()

    # 测试API密钥有效性
    try:
        # 尝试多个API端点
        endpoints = [
            "https://api.deepseek.com",
            "https://api.deepseek.com/v1",
            "https://api.deepseek.cc"
        ]

        success = False
        error_messages = []

        for endpoint in endpoints:
            try:
                client = OpenAI(api_key=deepseek_api_key, base_url=endpoint)
                test_response = client.chat.completions.create(
                    model="deepseek-chat",
                    messages=[{"role": "user", "content": "测试"}],
                    max_tokens=5
                )
                if test_response.choices:
                    st.sidebar.success(f"API端点可用: {endpoint}")
                    success = True
                    break
            except Exception as e:
                error_messages.append(f"{endpoint}: {str(e)}")

        if not success:
            raise Exception("所有API端点测试失败")

    except Exception as e:
        st.error(f"API密钥验证失败: {str(e)}")
        st.info("请检查：")
        st.info("1. API密钥是否正确且未过期")
        st.info("2. 密钥是否完整复制（以'sk-'开头）")
        st.info("3. 访问 https://platform.deepseek.com 检查账户状态")

        with st.expander("详细错误信息"):
            for msg in error_messages:
                st.error(msg)

        st.stop()

    # 限制最大记录数
    max_records = 50
    line_count = len(input_text.strip().split('\n'))
    if line_count > max_records:
        st.warning(f"一次最多处理{max_records}条记录（当前{line_count}条），请分批处理")
        st.stop()

    # 使用成功的端点
    client = OpenAI(api_key=deepseek_api_key, base_url=endpoint)

    with st.spinner("正在解析服务记录，请稍候..."):
        try:
            response = client.chat.completions.create(
                model="deepseek-chat",
                messages=[
                    {"role": "system", "content": """
                        你是一个文本解析专家，负责将无序的清洗服务记录文本转换为结构化的表格数据。请根据以下规则处理输入文本，并输出清晰的JSON格式。

                        ### 解析规则:
                        1. 自动识别11位电话号码
                        2. 识别"未支付"和"已支付"状态
                        3. 提取费用信息（如180元）
                        4. 识别房号格式（如16栋27-7）
                        5. 开头的中文名字作为师傅姓名
                        6. 剩余内容分割为项目和服务内容

                        ### 输出格式:
                        请将解析结果输出为JSON格式，包含以下字段:
                        - 师傅: 师傅姓名
                        - 项目: 项目名称
                        - 地址: 地址
                        - 房号: 房号
                        - 客户姓名: 客户姓名
                        - 电话号码: 电话号码
                        - 服务内容: 服务内容
                        - 费用: 费用
                        - 支付状态: 支付状态
                        - 备注: 备注

                        ### 支持的文本格式示例:
                        张雨浪 凡尔赛 下午 融创 凡尔赛领馆四期 16栋27-7 15223355185 空调内外机清洗 有异味，可能要全拆洗180，外机在室外150，内机高温蒸汽洗58 未支付 这个要翻外墙，什么时候来
                        李雪霜 华宇 寸滩派出所楼上 2栋9-8 13983014034 挂机加氟+1空调清洗 加氟一共299 清洗50 未支付
                        王师傅 龙湖源著 8栋12-3 13800138000 空调维修 不制冷 加氟200 已支付 需要周末上门

                        ### 输出结果格式示例:
                        [
                            {
                                "师傅": "张雨浪",
                                "项目": "空调内外机清洗",
                                "地址": "融创 凡尔赛领馆四期",
                                "房号": "16栋27-7",
                                "客户姓名": "凡尔赛",
                                "电话号码": "15223355185",
                                "服务内容": "有异味，可能要全拆洗180，外机在室外150，内机高温蒸汽洗58",
                                "费用": "180元+150元+58元=388元",
                                "支付状态": "未支付",
                                "备注": "这个要翻外墙，什么时候来"
                            },
                            {
                                "师傅": "李雪霜",
                                "项目": "挂机加氟+1空调清洗",
                                "地址": "寸滩派出所楼上",
                                "房号": "2栋9-8",
                                "客户姓名": "华宇",
                                "电话号码": "13983014034",
                                "服务内容": "加氟一共299 清洗50",
                                "费用": "299元+50元=349元",
                                "支付状态": "未支付",
                                "备注": ""
                            },
                            {
                                "师傅": "王师傅",
                                "项目": "空调维修",
                                "地址": "龙湖源著",
                                "房号": "8栋12-3",
                                "客户姓名": "",
                                "电话号码": "13800138000",
                                "服务内容": "不制冷 加氟200",
                                "费用": "200元",
                                "支付状态": "已支付",
                                "备注": ""
                            }
                        ]

                        ## 注意事项:
                        - 请确保输出的JSON格式正确，方便后续处理。
                        - 如果无法解析某条记录，请返回空对象或空列表，并在备注中说明原因。
                        - 返回的格式必须严格遵循上述示例格式的字符串，不要携带任何额外的文本或说明，包括```json```。
                        - 如果没有指定属性的值，请将该值设置为空字符串。
                        - 返回的结果要确保能直接被python的eval函数解析为列表或字典格式。
                    """},
                    {"role": "user", "content": "请解析以下清洗服务记录文本并输出为JSON格式:\n" + input_text},
                ],
                stream=False
            )
        except Exception as e:
            st.error(f"API调用失败: {str(e)}")
            st.info("建议尝试：")
            st.info("1. 检查DeepSeek平台状态")
            st.info("2. 稍后重试")
            st.info("3. 联系DeepSeek支持")
            st.stop()

    # 解析响应内容
    data = []
    errors = []
    if not response.choices or not response.choices[0].message.content:
        st.error("未能解析出任何记录，请检查输入格式！")
        st.stop()

    try:
        parsed_data = eval(response.choices[0].message.content)
        if isinstance(parsed_data, list):
            for record in parsed_data:
                if isinstance(record, dict):
                    data.append([
                        record.get('师傅', ''),
                        record.get('项目', ''),
                        record.get('地址', ''),
                        record.get('房号', ''),
                        record.get('客户姓名', ''),
                        record.get('电话号码', ''),
                        record.get('服务内容', ''),
                        record.get('费用', ''),
                        record.get('支付状态', ''),
                        record.get('备注', '')
                    ])
                else:
                    errors.append(f"第 {len(data) + 1} 条记录格式错误: {record}")
        else:
            errors.append("解析结果不是列表格式，请检查输入文本！")
    except Exception as e:
        errors.append(f"解析失败: {str(e)}")

    if data:
        st.session_state['df'] = pd.DataFrame(data, columns=columns)
        st.session_state['reset_table'] = False
        st.success(f"成功解析 {len(data)} 条记录！")
    else:
        st.error("未能解析出任何记录，请检查输入格式！")
        if errors:
            st.warning(f"共发现 {len(errors)} 条解析错误")
            for error in errors:
                st.error(error)

# 只要 session_state['df'] 存在就显示可编辑表格
if 'df' in st.session_state:

    st.subheader("清洗服务记录表格（可编辑）")

    gb = GridOptionsBuilder.from_dataframe(st.session_state['df'])
    gb.configure_default_column(editable=True)
    gb.configure_grid_options(domLayout='normal')
    grid_options = gb.build()

    grid_response = AgGrid(
        st.session_state['df'],
        gridOptions=grid_options,
        data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
        update_mode=GridUpdateMode.MODEL_CHANGED,
        fit_columns_on_grid_load=True,
        enable_enterprise_modules=False,
        allow_unsafe_jscode=True,
        use_container_width=True
    )

    # 保存编辑后的 DataFrame
    st.session_state['df'] = pd.DataFrame(grid_response['data'])
    df = st.session_state['df']

    # 添加统计信息
    col1, col2, col3 = st.columns(3)
    col1.metric("总记录数", len(df))
    payment_counts = df['支付状态'].value_counts()
    if not payment_counts.empty:
        col2.metric("未支付数量", payment_counts.get('未支付', 0))
        col3.metric("已支付数量", payment_counts.get('已支付', 0))

    # 导出Excel功能
    st.subheader("导出数据")
    output = BytesIO()
    try:
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='清洗服务记录')
            workbook = writer.book
            worksheet = writer.sheets['清洗服务记录']
            for idx, col in enumerate(df.columns):
                max_len = max(df[col].astype(str).map(len).max(), len(col)) + 2
                worksheet.set_column(idx, idx, max_len)
            format_red = workbook.add_format({'bg_color': '#FFC7CE'})
            format_green = workbook.add_format({'bg_color': '#C6EFCE'})
            worksheet.conditional_format(1, 7, len(df), 7, {
                'type': 'text',
                'criteria': 'containing',
                'value': '未支付',
                'format': format_red
            })
            worksheet.conditional_format(1, 7, len(df), 7, {
                'type': 'text',
                'criteria': 'containing',
                'value': '已支付',
                'format': format_green
            })
            worksheet.freeze_panes(1, 0)
            worksheet.autofilter(0, 0, len(df), len(df.columns) - 1)
    except Exception as e:
        st.warning(f"高级Excel格式设置失败: {str(e)}，使用基础导出")
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='清洗服务记录')
    excel_data = output.getvalue()
    b64 = base64.b64encode(excel_data).decode()
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="清洗服务记录.xlsx">⬇️ 下载Excel文件</a>'
    st.markdown(href, unsafe_allow_html=True)

# 使用说明
st.divider()
st.subheader("使用说明")
st.markdown("""
1. 在文本框中输入清洗服务记录（每行一条记录）
2. 使用自动编辑功能预处理文本
3. 点击 **🚀 转换文本为表格** 按钮
4. 查看解析后的表格数据
5. 点击 **⬇️ 下载Excel文件** 导出数据

### 密钥配置
在使用前，请设置DeepSeek API密钥：
1. **本地开发**：创建 `.env` 文件并添加：DEEPSEEK_API_KEY=sk-your-api-key

text
2. **Streamlit Cloud**：在部署设置中添加密钥（Secrets）
3. **手动输入**：在侧边栏手动输入API密钥

### API密钥问题排查：
- 确保密钥以 `sk-` 开头
- 检查密钥是否过期或被撤销
- 确认密钥在DeepSeek平台有效
- 密钥不应包含多余空格或换行符

### 支持的文本格式示例:
张雨浪 凡尔赛 下午 融创 凡尔赛领馆四期 16栋27-7 15223355185 空调内外机清洗 有异味，可能要全拆洗180，外机在室外150，内机高温蒸汽洗58 未支付 这个要翻外墙，什么时候来

李雪霜 华宇 寸滩派出所楼上 2栋9-8 13983014034 挂机加氟+1空调清洗 加氟一共299 清洗50 未支付

王师傅 龙湖源著 8栋12-3 13800138000 空调维修 不制冷 加氟200 已支付 需要周末上门

### 解析规则:
1. 自动识别11位电话号码
2. 识别"未支付"和"已支付"状态
3. 提取费用信息（如180元）
4. 识别房号格式（如16栋27-7）
5. 开头的中文名字作为师傅姓名
6. 剩余内容分割为项目和服务内容
""")

# 页脚
st.divider()
st.caption("© 2025 清洗服务记录转换工具 | 使用Python和Streamlit构建")