# -*- coding: UTF-8 -*-

import argparse
import openpyxl

# NOTE: 表格中必须要有的key
REQUIRED_KEY_CATEGORY = u'品牌'
REQUIRED_KEY_MODEL = u'型号'
REQUIRED_KEY_PRODECT_TYPE = u'产品类型'

REQUIRED_KEY_SELL_COUNT = u'零售量'
REQUIRED_KEY_SINGLE_PRICE = u'单价'

REQUIRED_KEY_REGION = u'工贸'

REQUIRED_KEYS = [
	REQUIRED_KEY_CATEGORY,
	REQUIRED_KEY_MODEL,
	REQUIRED_KEY_PRODECT_TYPE,
	REQUIRED_KEY_SELL_COUNT,
	REQUIRED_KEY_SINGLE_PRICE,
	REQUIRED_KEY_REGION,
]

# NOTE: 后处理的一些key
POST_KEY_SALES = 'post_key_sales'
POST_KEY_NAME = 'post_key_name'

# NOTE: 价位分段
PRICE_SEGMENT = [
	2000,
	2500,
	3000,
	3500,
	4000,
	4500,
	5000,
	5500,
	6000,
	8000,
	10000,
	15000,
	20000,
	30000,
]

# NOTE: 分类的key
CATEGORY_KEYS = [REQUIRED_KEY_CATEGORY, REQUIRED_KEY_PRODECT_TYPE]

# NOTE:

def getPriceSeg(price):
	# 获取单价属于哪个加个区间
	index = None
	for price_seg_index in xrange(len(PRICE_SEGMENT)):
		if price <= PRICE_SEGMENT[price_seg_index]:
			index = price_seg_index
			break

	if index == 0:
		key = '%s-'%(PRICE_SEGMENT[0])
	elif index is None:
		key = '%s+'%(PRICE_SEGMENT[-1])
	else:
		key = '%s-%s'%(PRICE_SEGMENT[index-1], PRICE_SEGMENT[index])
	return key

def getExcelData(args):
	# NOTE: 读取excel表格数据
	workbook = openpyxl.load_workbook(filename=args.excel)
	sheet = workbook[args.sheet]

	# NOTE: datas 为表中的数据
	datas = tuple(sheet.rows)

	# NOTE: 将表格中的数据读到内存中
	title_name_2_index, title_index_2_name = {}, {}
	raw_datas = {}
	for index in range(len(datas)):
		items = datas[index]

		# 是否是title行
		is_title_flag = True if index == 0 else False
		tmp_raw_data = {}
		for tmp_index in xrange(len(items)):
			item = items[tmp_index]
			value = item.value

			if is_title_flag:
				title_name_2_index[value] = tmp_index
				title_index_2_name[tmp_index] = value
				continue

			title_name = title_index_2_name[tmp_index]
			tmp_raw_data[title_name] = value

		if is_title_flag:
			continue

		# NOTE: 计算份额
		tmp_raw_data[REQUIRED_KEY_SELL_COUNT] = int(tmp_raw_data[REQUIRED_KEY_SELL_COUNT])
		tmp_raw_data[REQUIRED_KEY_SINGLE_PRICE] = int(tmp_raw_data[REQUIRED_KEY_SINGLE_PRICE])
		tmp_raw_data[POST_KEY_SALES] = tmp_raw_data[REQUIRED_KEY_SELL_COUNT]*tmp_raw_data[REQUIRED_KEY_SINGLE_PRICE]
		tmp_raw_data[POST_KEY_NAME] = "%s %s"%(tmp_raw_data[REQUIRED_KEY_CATEGORY], tmp_raw_data[REQUIRED_KEY_MODEL])

		raw_datas[index] = tmp_raw_data

	# NOTE: 检查分类的key是否都存在
	for key in REQUIRED_KEYS:
		is_exist = key in title_name_2_index
		if not is_exist:
			raise RuntimeError("required_key(%s) is absent")

	return raw_datas

def categoryExcelData(raw_datas):
	# 将excel表中的数据分类

	# NOTE: 按照分类的key进行分类
	classified_data = {} # 分类之后的数据
	for talb_id, table_data in raw_datas.iteritems():
		for catetory_key in CATEGORY_KEYS:
			price_seg = getPriceSeg(table_data[REQUIRED_KEY_SINGLE_PRICE])
			classified_data \
				.setdefault(table_data[REQUIRED_KEY_REGION], {}) \
				.setdefault(price_seg, {}) \
				.setdefault(catetory_key, {}) \
				.setdefault(table_data[catetory_key], []).append(
					[talb_id, table_data[POST_KEY_SALES]]
				)
	return classified_data

if __name__ == '__main__':
	parser = argparse.ArgumentParser()
	parser.add_argument("--excel", type=str, required=False,
				default='./wr1.xlsx', help='excel path')
	parser.add_argument("--sheet", type=str, required=False,
				default='Sheet1', help='sheet name')
	args = parser.parse_args()

	print("[ARGS] excel(%s) - sheet(%s)" % (args.excel, args.sheet))

	# NOTE: 读取excel中的数据
	raw_datas = getExcelData(args)

	# NOTE: 按照分类的key进行分类
	classified_data = categoryExcelData(raw_datas) # 分类之后的数据

	for region, region_data in classified_data.iteritems():
		# 值处理杭州的
		if region != u'杭州':
			continue
		for price_seg, price_seg_data in region_data.iteritems():
			for cat_key, cat_data in price_seg_data.iteritems():
					for cat_name, details in cat_data.iteritems():
						print("[%s][%s][%s] %s - (%s)"%(region, price_seg, cat_key, cat_name, details))