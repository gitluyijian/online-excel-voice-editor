import pandas as pd
import os


class ExcelHandler:
    def read_excel(self, filepath):
        try:
            # 读取Excel文件
            df = pd.read_excel(filepath)

            # 获取列名
            columns = df.columns.tolist()

            # 获取数据
            data = df.values.tolist()

            return {
                'columns': columns,
                'data': data
            }
        except Exception as e:
            print(f"Error reading Excel file: {str(e)}")
            return None

    def save_excel(self, data, filepath):
        try:
            # 确保数据格式正确
            if not isinstance(data, dict) or 'columns' not in data or 'data' not in data:
                print("Invalid data format")
                return False

            # 创建DataFrame
            df = pd.DataFrame(data['data'], columns=data['columns'])

            # 确保文件扩展名为 .xlsx
            filepath = os.path.splitext(filepath)[0] + '.xlsx'

            # 使用 openpyxl 引擎保存为 xlsx 格式
            df.to_excel(filepath, index=False, engine='openpyxl')

            return True
        except Exception as e:
            print(f"Error saving Excel file: {str(e)}")
            return False 