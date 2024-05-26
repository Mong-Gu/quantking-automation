import os
import pandas as pd
from lxml import etree

def extract_data_from_html(html_file):
    # HTML 파일을 읽고 lxml.html로 파싱합니다.
    with open(html_file, 'r', encoding='utf-8') as file:
        tree = etree.parse(file, etree.HTMLParser())

    # XPath 표현식을 사용하여 데이터를 추출합니다.
    def safe_extract(xpath_expr):
        result = tree.xpath(xpath_expr)
        if result:
            if isinstance(result[0], etree._Element):
                return result[0].text.strip()
            else:
                return result[0].strip()
        return None

    data = {
        "title": safe_extract('/html/body/div/h1/text()'),
        "period": safe_extract('/html/body/div/h1/dt/text()'),
        "Strategy CAGR": safe_extract('/html/body/div/div[2]/table/tbody/tr[5]/td[2]'),
        "Benchmark CAGR": safe_extract('/html/body/div/div[2]/table/tbody/tr[5]/td[3]'),
        "Strategy MDD": safe_extract('/html/body/div/div[2]/table/tbody/tr[15]/td[2]'),
        "Benchmark MDD": safe_extract('/html/body/div/div[2]/table/tbody/tr[15]/td[3]'),
        "Strategy MTD": safe_extract('/html/body/div/div[2]/table/tbody/tr[42]/td[2]'),
        "Benchmark MTD": safe_extract('/html/body/div/div[2]/table/tbody/tr[42]/td[3]'),
        "Strategy 3M": safe_extract('/html/body/div/div[2]/table/tbody/tr[43]/td[2]'),
        "Benchmark 3M": safe_extract('/html/body/div/div[2]/table/tbody/tr[43]/td[3]'),
        "Strategy 6M": safe_extract('/html/body/div/div[2]/table/tbody/tr[44]/td[2]'),
        "Benchmark 6M": safe_extract('/html/body/div/div[2]/table/tbody/tr[44]/td[3]'),
        "Strategy 1YTD": safe_extract('/html/body/div/div[2]/table/tbody/tr[45]/td[2]'),
        "Benchmark 1YTD": safe_extract('/html/body/div/div[2]/table/tbody/tr[45]/td[3]'),
        "universe": safe_extract('/html/body/div/div[3]/table/tbody/tr[1]/td[2]'),
        "factor": safe_extract('/html/body/div/div[3]/table/tbody/tr[4]/td[2]'),
        "trade fee": safe_extract('/html/body/div/div[3]/table/tbody/tr[6]/td[2]'),
        "rebalancing period": safe_extract('/html/body/div/div[3]/table/tbody/tr[7]/td[2]'),
        "filter": safe_extract('/html/body/div/div[3]/table/tbody/tr[8]/td[2]')
    }

    return data

def write_to_excel(data, excel_file):
    # 데이터프레임으로 변환합니다.
    df = pd.DataFrame([data])

    if os.path.exists(excel_file):
        # 엑셀 파일이 존재하면, 기존 파일을 로드하고 데이터를 추가합니다.
        existing_df = pd.read_excel(excel_file)
        df = pd.concat([existing_df, df], ignore_index=True)

    # 데이터프레임을 엑셀 파일에 씁니다.
    df.to_excel(excel_file, index=False)

def main():
    # 'html_files' 폴더 경로를 지정합니다.
    html_folder = os.path.join(os.getcwd(), 'html_files')

    # 'html_files' 폴더에 있는 모든 HTML 파일을 찾습니다.
    html_files = [f for f in os.listdir(html_folder) if f.endswith('.html')]
    all_data = []

    for html_file in html_files:
        # 'html_files' 폴더 내의 파일 경로를 생성합니다.
        file_path = os.path.join(html_folder, html_file)
        data = extract_data_from_html(file_path)
        all_data.append(data)

    # 데이터프레임으로 변환합니다.
    df = pd.DataFrame(all_data)

    # 데이터프레임을 엑셀 파일에 씁니다. 기존 파일을 덮어씁니다.
    excel_file = 'output.xlsx'
    df.to_excel(excel_file, index=False)

if __name__ == "__main__":
    main()
