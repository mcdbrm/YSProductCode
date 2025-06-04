import sys
import datetime
import time
from openpyxl import Workbook, load_workbook
from collections import defaultdict
import logging
import warnings
from typing import Dict, List, Optional, Set, Tuple, Union, Any

# 配置日志和警告
warnings.filterwarnings("ignore")
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

class ProductCodeGenerator:
    """商品编码生成器"""
    
    def __init__(self):
        self.processed_codes: Set[str] = set()  # 存储已处理的商品编码
        self.brand_workbooks: Dict[str, Workbook] = {}  # 各品牌的工作簿
        self.size_chart_worksheet: Optional[Workbook] = None  # 回力童装号型对照表
        
    @staticmethod
    def wait_for_exit() -> None:
        """等待用户输入后退出程序"""
        logging.info("按下Enter回车键以退出程序...")
        input()
        sys.exit()
    
    @staticmethod
    def find_duplicate_items(item_list: List[List]) -> List[List]:
        """查找列表中的重复项并统计出现次数"""
        count_dict = defaultdict(int)
        for item in item_list:
            count_dict[tuple(item)] += 1
        return [[list(key), value] for key, value in count_dict.items()]
    
    def load_excel_file(self, file_path: str, sheet_names: Optional[List[str]] = None, 
                      data_only: bool = True) -> Union[Workbook, Tuple[Workbook, Dict]]:
        """加载Excel文件"""
        logging.info(f'正在打开文件: {file_path}')
        start_time = time.perf_counter()
        try:
            workbook = load_workbook(file_path, data_only=data_only)
            end_time = time.perf_counter()  
            file_name = file_path.split(".")[0]
            logging.info(f'{file_name}.xlsx 加载用时: {end_time - start_time:.6f}秒')
            
            if sheet_names:
                worksheets = {name: workbook[name] for name in sheet_names}
                return workbook, worksheets
            return workbook
        except Exception as e:
            logging.error(f'{file_path} 加载失败: {str(e)}')
            self.wait_for_exit()
    
    @staticmethod
    def get_unique_brands(worksheet) -> List[str]:
        """从商品信息表中获取品牌列表"""
        return list({worksheet.cell(row, 2).value 
                   for row in range(2, worksheet.max_row + 1) 
                   if worksheet.cell(row, 2).value is not None})
    
    @staticmethod
    def find_matching_value(worksheet, lookup_col: int, lookup_val: str, 
                          return_col: int) -> Optional[str]:
        """在工作表中查找匹配的值"""
        for row in range(2, worksheet.max_row + 1):
            if worksheet.cell(row, lookup_col).value == lookup_val:
                return worksheet.cell(row, return_col).value
        return None
    
    @staticmethod
    def validate_required_fields(product: Dict, required_keys: List[str]) -> None:
        """验证商品数据是否包含必需字段"""
        for key in required_keys:
            if product.get(key) is None:
                logging.error(f"商品数据缺失: {key} 字段")
                raise ValueError(f"缺少必需的字段: {key}")
    
    def process_single_product(self, product_info: Dict, worksheet) -> None:
        """处理单件装商品信息"""
        for row in range(2, worksheet.max_row + 1):
            if worksheet.cell(row, 4).value == product_info.get('品类'):
                product_info.update({
                    '商品分类': worksheet.cell(row, 3).value,
                    '季节': worksheet.cell(row, 1).value
                })
                break
        self.validate_required_fields(product_info, ['商品分类', '季节'])
    
    def process_combo_product(self, product_info: Dict, info_worksheet, combo_worksheet) -> None:
        """处理多件装商品信息"""
        for row in range(2, info_worksheet.max_row + 1):
            if info_worksheet.cell(row, 4).value == product_info.get('品类'):
                product_info.update({
                    '商品分类': info_worksheet.cell(row, 3).value,
                    '季节': info_worksheet.cell(row, 1).value,
                    '单件组合装款式编码': info_worksheet.cell(row, 6).value
                })
                break
        self.validate_required_fields(product_info, ['商品分类', '季节', '单件组合装款式编码'])
        
        # 查找组合装款式商品编码
        product_info['组合装款式商品编码'] = self.find_matching_value(
            combo_worksheet, 4, product_info.get('组合装款式编码'), 3)
        self.validate_required_fields(product_info, ['组合装款式商品编码'])
    
    def process_print_info(self, print_worksheet, product_info: Dict, side: str) -> None:
        """处理印花信息（前或后）"""
        position_key = f'位置-{side}'
        print_name_key = f'印花名称-{side}'
        print_code_key = f'印花编码-{side}'
        position_code_key = f'位置代码-{side}'

        # 如果没有印花名称则跳过
        if product_info.get(print_name_key) is None:
            return

        # 根据位置生成印花编码
        if product_info.get(position_key) in ["胸", "裤"]:
            product_info[print_code_key] = f"{product_info.get(print_name_key)}X"
            position_col, code_col = 10, 11
        else:
            product_info[print_code_key] = f"{product_info.get(print_name_key)}D"
            position_col, code_col = 7, 8

        # 查找位置代码
        product_info[position_code_key] = self.find_matching_value(
            print_worksheet, position_col, product_info.get(position_key), code_col)
        self.validate_required_fields(product_info, [position_code_key])

        # 查找印花名称
        product_info[print_name_key] = self.find_matching_value(
            print_worksheet, 4, product_info.get(print_code_key), 1)
        self.validate_required_fields(product_info, [print_name_key])
    
    def fetch_brand_data(self, product_info: Dict, sheet_name: str, 
                       specification_brand: str, brand_code: str, 
                       column_range: Tuple[int, int]) -> List:
        """获取品牌特定数据"""
        if product_info.get('组合形式') == '多件装':
            temp_value = product_info.get('组合装款式商品编码')
            product_info['组合装款式商品编码'] = f"{temp_value[:4]}{brand_code}{temp_value[-1:]}"

        try:
            worksheet = self.brand_workbooks[brand_code][sheet_name]
        except Exception as e:
            logging.error(f"{brand_code} 吊牌信息表加载失败: {sheet_name}: {str(e)}")
            self.wait_for_exit()

        # 在品牌工作表中查找匹配的数据
        for data_row in range(2, worksheet.max_row + 1):
            if (worksheet.cell(data_row, 1).value == product_info.get('单件组合装款式编码') and 
                worksheet.cell(data_row, 2).value == specification_brand):
                start_col, end_col = column_range
                return [worksheet.cell(data_row, col).value for col in range(start_col, end_col)]
        
        logging.error(f"{brand_code} 吊牌信息汇总表查询失败-《{product_info.get('品类')}》分表中找不到{product_info.get('单件组合装款式编码')} {specification_brand}")
        self.wait_for_exit()
    
    def add_combo_record(self, product_info: Dict, combo_code: str, 
                       product_name: str, entity_code: str, specification: str, 
                       code: str, trademark_data: List, worksheet) -> None:
        """添加组合装记录到工作表"""
        record_data = [
            product_info.get('单件组合装款式编码'),
            combo_code,
            product_name,
            entity_code,
            '成品',
            specification,
            code,
            1,
            0,
            'YS'
        ] + trademark_data

        unique_id = combo_code + code
        if unique_id not in self.processed_codes:
            self.processed_codes.add(unique_id)
            worksheet.append(record_data)
    
    @staticmethod
    def generate_print_codes(product_info: Dict, side_prefix: str, color: str) -> Tuple[str, str]:
        """生成印花相关临时编码"""
        position = '' if product_info.get(f'位置-{side_prefix}') == '大图' else product_info.get(f'位置-{side_prefix}')
        code1 = f"{product_info.get(f'印花名称-{side_prefix}')}{product_info.get(f'位置代码-{side_prefix}')}/{color}"
        code2 = f"{product_info.get(f'印花名称-{side_prefix}')}{position}/{color}"
        return code1, code2
    
    @staticmethod
    def store_print_codes(position_color_list: List, color_list: List, 
                        color_value_list: List, code1: str, code2: str, color: str) -> None:
        """存储印花编码到对应列表"""
        position_color_list.append(code1)
        color_list.append(code2)
        color_value_list.append(color)
    
    def generate_product_codes(self) -> None:
        """主处理流程 - 生成商品编码"""
        # 打开所有必要的工作簿
        info_wb, info_ws_dict = self.load_excel_file('商品编码信息表.xlsx', ['商品编码信息表1'])
        product_info_ws = info_ws_dict['商品编码信息表1']

        base_wb, base_ws_dict = self.load_excel_file('资料生成器.xlsx', 
            ['附表1印花基础资料', '附表2单品基础资料', '附表3多件装基础信息', '大货称重表'])
        print_info_ws = base_ws_dict['附表1印花基础资料']
        product_base_ws = base_ws_dict['附表2单品基础资料']
        combo_info_ws = base_ws_dict['附表3多件装基础信息']
        weight_info_ws = base_ws_dict['大货称重表']

        relation_wb, _ = self.load_excel_file('商品对应关系.xlsx', ['Sheet1'])
        relation_ws = relation_wb['Sheet1']
        
        # 加载品牌工作簿
        brand_list = self.get_unique_brands(product_info_ws)
        self.load_brand_files(brand_list)

        # 创建结果工作簿
        result_wb = Workbook()
        single_general_ws = result_wb.active
        single_general_ws.title = '单品-普通资料'
        self.setup_single_general_sheet(single_general_ws)
        
        single_combo_ws = result_wb.create_sheet('单品-组合构成')
        self.setup_single_combo_sheet(single_combo_ws)
        
        multi_combo_ws = result_wb.create_sheet('多件装-组合构成')
        self.setup_multi_combo_sheet(multi_combo_ws)
        
        # 关系对照表从第二行开始填充
        relation_row = 2
        multi_combo_temp_list = []

        # 处理每一行商品信息
        for row_index in range(4, product_info_ws.max_row + 1):
            if product_info_ws.cell(row_index, 2).value is None:
                continue
                
            logging.info(f'处理第 {row_index} 行数据')
            products = self._extract_products_from_row(product_info_ws, row_index)
            
            for product in products:
                try:
                    self._process_product(
                        product, print_info_ws, product_base_ws, combo_info_ws, weight_info_ws, 
                        single_general_ws, single_combo_ws, 
                        multi_combo_temp_list, relation_ws, 
                        relation_row
                    )
                    relation_row += 1
                except Exception as e:
                    logging.error(f"处理商品时出错: {str(e)}")
                    continue

        # 处理多件装组合
        if multi_combo_temp_list:
            duplicates = self.find_duplicate_items(multi_combo_temp_list)
            for duplicate in duplicates:
                combo_data = duplicate[0]
                combo_data[-3] = duplicate[1]  # 设置数量
                multi_combo_ws.append(combo_data)

        # 保存结果文件
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        logging.info('商品编码生成完成')
        logging.info(f'保存编码生成表: 编码生成表_{timestamp}.xlsx')
        result_wb.save(f'编码生成表_{timestamp}.xlsx')

        logging.info(f'保存关系对应表: 关系对应生成表_{timestamp}.xlsx')
        relation_wb.save(f'关系对应生成表_{timestamp}.xlsx')

    def load_brand_files(self, brand_list: List[str]) -> None:
        """加载品牌相关Excel文件"""
        if 'HL' in brand_list:
            hl_wb, hl_ws_dict = self.load_excel_file('回力吊牌信息汇总表.xlsx', ['回力童装号型对照表'])
            self.brand_workbooks['HL'] = hl_wb
            self.size_chart_worksheet = hl_ws_dict['回力童装号型对照表']

        if 'BD' in brand_list:
            self.brand_workbooks['BD'] = self.load_excel_file('巴帝吊牌信息汇总表.xlsx')

        if 'SE' in brand_list:
            self.brand_workbooks['SE'] = self.load_excel_file('少宜吊牌信息汇总表.xlsx')
        
        if 'ML' in brand_list:
            self.brand_workbooks['ML_male'] = self.load_excel_file('菲尔吊牌信息汇总表-男童.xlsx')
            self.brand_workbooks['ML_female'] = self.load_excel_file('菲尔吊牌信息汇总表-女童.xlsx')

        if 'JW' in brand_list:
            self.brand_workbooks['JW_male'] = self.load_excel_file('真维斯吊牌信息汇总表-男童.xlsx')
            self.brand_workbooks['JW_female'] = self.load_excel_file('真维斯吊牌信息汇总表-女童.xlsx')

    def setup_single_general_sheet(self, worksheet) -> None:
        """初始化单品普通资料表"""
        headers = [
            '款式编码', '商品编码', '商品名', '分类', '颜色及规格', '重量', '品牌', '虚拟分类', '国标码',
            '其它属性1', '其它属性2', '其它属性3', '其它属性4', '其它属性5', '其它属性6', '其它属性7',
            '其它属性8', '其它属性9', '其它属性10'
        ]
        worksheet.append(headers)

    def setup_single_combo_sheet(self, worksheet) -> None:
        """初始化单品组合构成表"""
        headers = [
            '组合款式编码', '组合商品编码', '组合商品名称', '组合商品实体编码', '虚拟分类',
            '组合颜色规格', '商品编码', '数量', '应占售价', '品牌', '组合装国标码',
            '其它属性1', '其它属性2', '其它属性3', '其它属性4', '其它属性5', '其它属性6',
            '其它属性7', '其它属性8', '其它属性9', '其它属性10'
        ]
        worksheet.append(headers)

    def setup_multi_combo_sheet(self, worksheet) -> None:
        """初始化多件装组合构成表"""
        headers = [
            '组合款式编码', '组合商品编码', '组合商品名称', '虚拟分类', '组合颜色规格',
            '商品编码', '数量', '应占售价', '品牌'
        ]
        worksheet.append(headers)

    def _extract_products_from_row(self, worksheet, row_index: int) -> List[Dict]:
        """从工作表的行中提取商品信息"""
        products = []
        count = 0
        
        for col in range(7, 11):  # 遍历G到J列
            if worksheet.cell(row_index, col).value is None or worksheet.cell(row_index, col).value == 0:
                continue
                
            product_data = {
                '品类': worksheet.cell(row_index, col).value,
                '颜色': worksheet.cell(row_index, 5+col+count*4).value,
                '印花名称-前': str(worksheet.cell(row_index, 6+col+count*4).value) 
                            if worksheet.cell(row_index, 6+col+count*4).value is not None else None,
                '位置-前': worksheet.cell(row_index, 7+col+count*4).value,
                '印花名称-后': str(worksheet.cell(row_index, 8+col+count*4).value) 
                            if worksheet.cell(row_index, 8+col+count*4).value is not None else None,
                '位置-后': worksheet.cell(row_index, 9+col+count*4).value,
                '单件组合装款式编码': worksheet.cell(row_index, 5).value,
                '组合装款式编码': worksheet.cell(row_index, 6).value,
                '品牌': worksheet.cell(row_index, 2).value,
                '组合形式': worksheet.cell(row_index, 3).value,
                '性别': worksheet.cell(row_index, 4).value,
                '尺码': worksheet.cell(row_index, 11).value,
                '数量': ''
            }
            count += 1
            products.append(product_data)
            
        return products

    def _process_product(self, product_info: Dict, print_worksheet, base_worksheet, combo_worksheet, weight_worksheet, 
                      single_general_ws, single_combo_ws, 
                      multi_combo_list, relation_ws, 
                      relation_row: int) -> None:
        """处理单个商品信息"""
        logging.info(product_info)
        
        # 处理商品类型（单件装/多件装）
        if product_info.get('组合形式') == '单件装':
            self.process_single_product(product_info, base_worksheet)
        elif product_info.get('组合形式') == '多件装':
            self.process_combo_product(product_info, base_worksheet, combo_worksheet)

        # 处理印花信息
        self.process_print_info(print_worksheet, product_info, '前')
        self.process_print_info(print_worksheet, product_info, '后')

        # 处理每个尺码
        for size in product_info.get('尺码').split('/'):
            self._process_product_size(product_info, size, weight_worksheet, single_general_ws, 
                                    single_combo_ws, multi_combo_list, 
                                    relation_ws, relation_row)

    def _process_product_size(self, product_info: Dict, size: str, weight_worksheet, single_general_ws, 
                           single_combo_ws, multi_combo_list, 
                           relation_ws, relation_row: int) -> None:
        """处理特定尺码的商品"""
        # 生成商品编码和名称
        (product_code, product_name, combo_code, 
         entity_code, hl_trademark) = self._generate_codes_and_names(product_info, size)
            
        specification = f"{product_info.get('颜色')};{size}"
        brand_spec = f"{product_info.get('颜色')};{size}-{product_info.get('品牌')}"
        single_combo_code = f"纯色/{product_info.get('品类')}/{product_info.get('颜色')}/{size}"

        # 获取商品重量
        self._get_product_weight(product_info, weight_worksheet, size)

        # 获取品牌特定数据
        trademark_data = self._get_brand_specific_info(product_info, brand_spec, size, hl_trademark)

        # 添加到单品普通资料表
        self._add_to_single_general_sheet(product_info, product_code, product_name, 
                                        brand_spec, trademark_data, single_general_ws)

        # 处理单件装
        if product_info.get('组合形式') == '单件装':
            relation_ws.cell(relation_row, 8).value = product_code

        # 处理印花组合关系
        self._add_print_combinations(product_info, combo_code, product_name, 
                                  entity_code, specification, single_combo_code, 
                                  trademark_data, single_combo_ws)

        # 处理多件装组合关系
        if product_info.get('组合形式') == '多件装':
            self._add_combo_product_to_list(product_info, size, product_code, 
                                         multi_combo_list, 
                                         relation_ws, relation_row)

    def _generate_codes_and_names(self, product_info: Dict, size: str) -> Tuple[str, str, str, str, str]:
        """生成商品编码和名称"""
        front_pos = '' if product_info.get('位置-前') == '大图' else product_info.get('位置-前')
        back_pos = '' if product_info.get('位置-后') == '大图' else product_info.get('位置-后')
        
        # 仅有前印花
        if product_info.get('印花编码-前') and not product_info.get('印花编码-后'):
            product_code = f"{product_info.get('印花编码-前')}{front_pos}/{product_info.get('品类')}/{product_info.get('颜色')}/{size}-{product_info.get('品牌')}"
            product_name = f"{product_info.get('印花名称-前')}{front_pos}/{product_info.get('品类')}/{product_info.get('颜色')}/{size}-{product_info.get('品牌')}"
            combo_code = f"{product_info.get('印花编码-前')}{front_pos}/{product_info.get('品牌')}{product_info.get('品类')}/{product_info.get('颜色')}/{size}"
            hl_trademark = f"{product_info.get('印花编码-前')}{front_pos}/{product_info.get('品类')}/{product_info.get('颜色')}"
            
        # 仅有后印花
        elif not product_info.get('印花编码-前') and product_info.get('印花编码-后'):
            product_code = f"{product_info.get('印花编码-后')}{back_pos}/{product_info.get('品类')}/{product_info.get('颜色')}/{size}-{product_info.get('品牌')}"
            product_name = f"{product_info.get('印花名称-后')}{back_pos}/{product_info.get('品类')}/{product_info.get('颜色')}/{size}-{product_info.get('品牌')}"
            combo_code = f"{product_info.get('印花编码-后')}{back_pos}/{product_info.get('品牌')}{product_info.get('品类')}/{product_info.get('颜色')}/{size}"
            hl_trademark = f"{product_info.get('印花编码-后')}{back_pos}/{product_info.get('品类')}/{product_info.get('颜色')}"
            
        # 前后印花都有
        elif product_info.get('印花编码-前') and product_info.get('印花编码-后'):
            product_code = f"{product_info.get('印花编码-前')}{front_pos}_{product_info.get('印花编码-后')}{back_pos}/{product_info.get('品类')}/{product_info.get('颜色')}/{size}-{product_info.get('品牌')}"
            product_name = f"{product_info.get('印花名称-前')}{front_pos}_{product_info.get('印花名称-后')}{back_pos}/{product_info.get('品类')}/{product_info.get('颜色')}/{size}-{product_info.get('品牌')}"
            combo_code = f"{product_info.get('印花编码-前')}{front_pos}_{product_info.get('印花编码-后')}{back_pos}/{product_info.get('品牌')}{product_info.get('品类')}/{product_info.get('颜色')}/{size}"
            hl_trademark = f"{product_info.get('印花编码-前')}{front_pos}_{product_info.get('印花编码-后')}{back_pos}/{product_info.get('品类')}/{product_info.get('颜色')}"
            
        # CP品牌特殊处理
        if product_info.get('品牌') == 'CP':
            combo_code = combo_code.replace('CP', '')
            
        entity_code = product_code
        return product_code, product_name, combo_code, entity_code, hl_trademark

    def _get_product_weight(self, product_info: Dict, worksheet, size: str) -> None:
        """从称重表中获取商品重量"""
        for row in range(2, worksheet.max_row + 1):
            if (worksheet.cell(row, 1).value == product_info.get('品类') and 
                str(worksheet.cell(row, 2).value) == str(size)):
                product_info['重量'] = worksheet.cell(row, 3).value
                break
                
        if product_info.get('重量') is None:
            logging.error(f"称重表查询失败: {product_info.get('品类')} {size}")
            self.wait_for_exit()

    def _get_brand_specific_info(self, product_info: Dict, brand_spec: str, size: str, hl_trademark: str) -> List:
        """获取品牌特定信息"""
        trademark_data = []
        brand = product_info.get('品牌')
        
        if brand == 'HL':
            trademark_data = self._process_HL_brand_info(product_info, brand_spec, size, hl_trademark)
        elif brand == 'BD':
            trademark_data = self.fetch_brand_data(product_info, product_info.get('品类'), brand_spec, 'BD', (3, 12))
        elif brand == 'SE':
            trademark_data = self.fetch_brand_data(product_info, product_info.get('品类'), brand_spec, 'SE', (3, 12))
        elif brand == 'ML':
            gender_key = 'ML_male' if product_info.get('性别') == '男' else 'ML_female'
            trademark_data = self.fetch_brand_data(product_info, product_info.get('品类'), brand_spec, gender_key, (3, 14))
        elif brand == 'JW':
            gender_key = 'JW_male' if product_info.get('性别') == '男' else 'JW_female'
            trademark_data = self.fetch_brand_data(product_info, product_info.get('品类'), brand_spec, gender_key, (3, 14))
                
        return trademark_data

    def _process_HL_brand_info(self, product_info: Dict, brand_spec: str, size: str, hl_trademark: str) -> List:
        """处理回力品牌特定信息"""
        # 多件装特殊处理
        if product_info.get('组合形式') == '多件装':
            temp_value = product_info.get('组合装款式商品编码')
            product_info['组合装款式商品编码'] = f"{temp_value[:4]}HL{temp_value[-1:]}"

        try:
            brand_ws = self.brand_workbooks['HL'][product_info.get('品类')]
        except Exception as e:
            logging.error(f"回力吊牌信息表加载失败: {product_info.get('品类')}: {str(e)}")
            self.wait_for_exit()

        # 查找尺码类型
        size_type = ''
        for row in range(2, self.size_chart_worksheet.max_row + 1):
            if (self.size_chart_worksheet.cell(row, 1).value == product_info.get('品类') and 
                self.size_chart_worksheet.cell(row, 2).value == int(size)):
                size_type = self.size_chart_worksheet.cell(row, 3).value
                break
                
        if not size_type:
            logging.error(f"回力童装号型对照表查询失败: {product_info.get('品类')}/{size}")
            self.wait_for_exit()

        # 获取商标数据
        for row in range(2, brand_ws.max_row + 1):
            if brand_ws.cell(row, 1).value == hl_trademark:
                trademark_data = [brand_ws.cell(row, col).value for col in range(2, 10)]
                trademark_data.insert(1, size_type)  # 插入尺码类型
                trademark_data.insert(6, size)       # 插入尺码
                return trademark_data
                
        logging.error('回力吊牌信息汇总表查询失败')
        self.wait_for_exit()

    def _add_to_single_general_sheet(self, product_info: Dict, product_code: str, 
                                   product_name: str, brand_spec: str, 
                                   trademark_data: List, worksheet) -> None:
        """添加记录到单品普通资料表"""
        record_data = [
            product_info.get('单件组合装款式编码'),
            product_code,
            product_name,
            product_info.get('商品分类'),
            brand_spec,
            product_info.get('重量'),
            'YS',
            product_info.get('季节')
        ] + trademark_data

        if product_code not in self.processed_codes:
            self.processed_codes.add(product_code)
            worksheet.append(record_data)

    def _add_print_combinations(self, product_info: Dict, combo_code: str, 
                              product_name: str, entity_code: str, 
                              specification: str, single_combo_code: str, 
                              trademark_data: List, worksheet) -> None:
        """添加印花组合关系到组合构成表"""
        # 仅有前印花
        if product_info.get('印花编码-前') and not product_info.get('印花编码-后'):
            self.add_combo_record(product_info, combo_code, product_name, 
                                entity_code, specification, product_info.get('印花编码-前'), 
                                trademark_data, worksheet)
            self.add_combo_record(product_info, combo_code, product_name, 
                                entity_code, specification, single_combo_code, 
                                trademark_data, worksheet)
                                  
        # 仅有后印花
        elif not product_info.get('印花编码-前') and product_info.get('印花编码-后'):
            self.add_combo_record(product_info, combo_code, product_name, 
                                entity_code, specification, product_info.get('印花编码-后'), 
                                trademark_data, worksheet)
            self.add_combo_record(product_info, combo_code, product_name, 
                                entity_code, specification, single_combo_code, 
                                trademark_data, worksheet)
                                  
        # 前后印花都有
        elif product_info.get('印花编码-前') and product_info.get('印花编码-后'):
            self.add_combo_record(product_info, combo_code, product_name, 
                                entity_code, specification, product_info.get('印花编码-前'), 
                                trademark_data, worksheet)
            self.add_combo_record(product_info, combo_code, product_name, 
                                entity_code, specification, product_info.get('印花编码-后'), 
                                trademark_data, worksheet)
            self.add_combo_record(product_info, combo_code, product_name, 
                                entity_code, specification, single_combo_code, 
                                trademark_data, worksheet)

    def _add_combo_product_to_list(self, product_info: Dict, size: str, product_code: str,
                                combo_list: List, 
                                relation_ws, relation_row: int) -> None:
        """添加多件装商品到临时列表"""
        position_color_list = []
        color_list = []
        color_value_list = []

        # 处理印花编码及颜色
        color = product_info.get('颜色')
        
        # 仅有前印花
        if product_info.get('印花编码-前') and not product_info.get('印花编码-后'):
            code1, code2 = self.generate_print_codes(product_info, '前', color)
            self.store_print_codes(position_color_list, color_list, 
                                color_value_list, code1, code2, color)
                            
        # 仅有后印花
        elif not product_info.get('印花编码-前') and product_info.get('印花编码-后'):
            code1, code2 = self.generate_print_codes(product_info, '后', color)
            self.store_print_codes(position_color_list, color_list, 
                                color_value_list, code1, code2, color)
                            
        # 前后印花都有
        elif product_info.get('印花编码-前') and product_info.get('印花编码-后'):
            code1 = f"{product_info.get('印花名称-前')}{product_info.get('位置代码-前')}_{product_info.get('印花名称-后')}{product_info.get('位置代码-后')}/{color}"
            code2 = f"{product_info.get('印花名称-前')}{product_info.get('位置-前')}_{product_info.get('印花名称-后')}{product_info.get('位置-后')}/{color}"
            self.store_print_codes(position_color_list, color_list, 
                                color_value_list, code1, code2, color)

        position_color_str = '-'.join(position_color_list)
        color_str = '-'.join(color_list)
        combo_color = '-'.join(color_value_list)

        # 设置关系表对应值
        combo_code = f"{product_info.get('组合装款式商品编码')}-{position_color_str}-{size}"
        relation_ws.cell(relation_row, 8).value = combo_code

        # 添加到多件装组合临时列表
        combo_list.append([
            product_info.get('组合装款式编码'),
            combo_code,
            f"{product_info.get('组合装款式商品编码')}-{color_str}-{size}",
            '成品',
            f"{combo_color};{size}",
            product_code,  # 使用传入的商品编码
            1,
            0,
            'YS'
        ])

if __name__ == "__main__":
    logging.info('*************************************')
    logging.info('*       YS商品编码生成器 v1.0        *')
    logging.info('*************************************')

    try:
        generator = ProductCodeGenerator()
        generator.generate_product_codes()
    except Exception as e:
        logging.error(f"程序运行出错: {str(e)}", exc_info=True)
    finally:
        logging.info("程序执行完毕")
        logging.info("按下Enter回车键退出...")
        input()
