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

def getMarketSharesForPriceSeg(args, sorted_classified_data, raw_datas):
	# NOTE: 计算每个分段的占比以及
	# 目标品牌在某个价格段所占的份额
	sales_data = {}
	for region, region_data in sorted_classified_data.iteritems():
		# NOTE: 只处理特定地区的
		if region != args.target_region:
			continue

		for price_seg, price_seg_data in region_data.iteritems():
			tmp_total_sales = 0
			tmp_target_sales = 0
			for cat, details in price_seg_data[REQUIRED_KEY_CATEGORY].iteritems():
				for item in details:
					table_id, sales = item
					catgory = raw_datas[table_id][REQUIRED_KEY_CATEGORY]

					tmp_total_sales += sales
					if catgory == args.target_category:
						tmp_target_sales += sales
			sales_data[price_seg] = {
				'total': tmp_total_sales,
				'target': tmp_target_sales,
				'target_market_shares': float(float(tmp_target_sales)/float(tmp_total_sales))
			}

	# NOTE: 计算总的市场销售额
	total_sales = 0
	for price_seg, item in sales_data.iteritems():
		total_sales += item['total']

	ret_dict = {}
	for price_seg, item in sales_data.iteritems():
		ret_dict[price_seg] = {
			'total_market_shares': float(item['total'])/float(total_sales),
			'target_market_shares': item['target_market_shares'],
		}
	return ret_dict

def getMarketSharesForTotal(args, sorted_classified_data, raw_datas,
	get_target_cat=True):
	# 整体占比
	sales_data = []
	for region, region_data in sorted_classified_data.iteritems():
		# NOTE: 只处理特定地区的
		if region != args.target_region:
			continue
		for price_seg, price_seg_data in region_data.iteritems():
			for cat, details in price_seg_data[REQUIRED_KEY_CATEGORY].iteritems():
				for item in details:
					table_id, sales = item
					catgory = raw_datas[table_id][REQUIRED_KEY_CATEGORY]
					sales_data.append([table_id, price_seg, catgory, sales])

	total_sales = 0 # 销售总额度
	for item in sales_data:
		total_sales += item[-1]

	for index in xrange(len(sales_data)):
		sales_data[index].append(float(sales_data[index][-1])/float(total_sales))

	sorted_sales_data = sorted(sales_data, key=lambda s: s[-1])

	ret_dict = {}
	for index in xrange(len(sorted_sales_data)):
		item = sorted_sales_data[index]
		table_id, price_seg, catgory, sales, percent = item
		raw_data = raw_datas[table_id]
		if get_target_cat and raw_data[REQUIRED_KEY_CATEGORY] != args.target_category:
			continue
		ret_dict.setdefault(price_seg, []).append([table_id, raw_data[POST_KEY_NAME], sales, percent, index+1])

	return ret_dict

if __name__ == '__main__':
	parser = argparse.ArgumentParser()
	parser.add_argument("--excel", type=str, required=False,
				default='./wr1.xlsx', help='excel path')
	parser.add_argument("--sheet", type=str, required=False,
				default='Sheet1', help='sheet name')
	parser.add_argument("--target_region", type=unicode, required=False,
				default=u'杭州', help='target city')
	parser.add_argument("--target_category", type=unicode, required=False,
				default=u'海尔', help='target category')
	args = parser.parse_args()

	print("[ARGS] excel(%s) - sheet(%s)" % (args.excel, args.sheet))

	# NOTE: 读取excel中的数据
	raw_datas = getExcelData(args)

	# NOTE: 按照分类的key进行分类
	classified_data = categoryExcelData(raw_datas) # 分类之后的数据

	# NOTE: 排序
	sorted_classified_data = {} # 排序之后的数据
	for region, region_data in classified_data.iteritems():
		for price_seg, price_seg_data in region_data.iteritems():
			for cat_key in CATEGORY_KEYS:
				for cat, details in price_seg_data[cat_key].iteritems():
					sorted_data = sorted(details, key=lambda s: s[1], reverse=True)
					sorted_classified_data.setdefault(region, {}) \
						.setdefault(price_seg, {}) \
						.setdefault(cat_key, {})[cat] = sorted_data

	# NOTE: 每个加个段占市场的总份额 & 目标品牌在每个段的占比
	market_shares_for_all_price_seg = getMarketSharesForPriceSeg(args, sorted_classified_data, raw_datas)
	for price_seg, item in market_shares_for_all_price_seg.iteritems():
		print("[111] %s - %s "%(price_seg, item))

	# NOTE: 整体对比
	market_shares_for_total = getMarketSharesForTotal(args, sorted_classified_data, raw_datas, get_target_cat=False)
	for price_seg, item in market_shares_for_total.iteritems():
		print("[222] %s - %s "%(price_seg, item))

	# NOTE: 目标品牌
	market_shares_for_target = getMarketSharesForTotal(args, sorted_classified_data, raw_datas, get_target_cat=True)
	for price_seg, item in market_shares_for_target.iteritems():
		print("[333] %s - %s "%(price_seg, item))