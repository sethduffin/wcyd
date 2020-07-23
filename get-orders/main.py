from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
import extensions
import os,sys,time,pyperclip,signal,string,json
import local
from copy import deepcopy

driver = webdriver.Chrome()
action = ActionChains(driver)
driver.implicitly_wait(3)

extensions.set_driver(driver,action)

orders = []
product_totals = {}
all_products = local.products

used_products = []

size_product = {
	"S":"",
	"M":"",
	"L":"",
	"XL":"",
	"XXL":"",
}

default_product = ""

color_size_product_1 = {
	"Black": deepcopy(size_product),
	"Dark Grey": deepcopy(size_product),
	"Navy": deepcopy(size_product),
	"Olive Green": deepcopy(size_product),
	"Light Pink": deepcopy(size_product),
	"Light Blue": deepcopy(size_product),
	"Light Grey": deepcopy(size_product),
	"Sand": deepcopy(size_product),
}

color_size_product_2 = {
	"Black": deepcopy(size_product),
	"Navy": deepcopy(size_product),
	"Forest Green": deepcopy(size_product),
}

def signal_handler(sig, frame):
    error("Manual Exit",False)

signal.signal(signal.SIGINT, signal_handler)

def error(e="Unkown",pause=True):
	try:
		exc_type, exc_obj, exc_tb = sys.exc_info()
		fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
		print("Error: %s (%s %s)" % (e,fname,exc_tb.tb_lineno))
	except:
		print("Error: %s" % (e))
	if pause:
		input("-- Enter to Quit -- ")
	driver.quit()
	time.sleep(1)
	quit()

class Order:
	def __init__(self): 
		pass


def run():
	driver.get("https://my.ecwid.com/store/28113010#orders")
	driver.find("name","email")[0].send(local.email)
	driver.find("name","password")[0].send(local.password)
	driver.find("id","SIF.sIB").click()
	time.sleep(6)

	order_count = 0

	wait = 0
	while wait < 5:
		key = driver.find("class","order__number",True)
		if len(key) > 0:
			order_count = int(driver.find("class","order__number",True)[0].text)
			break
		time.sleep(.1)
		wait += 0.1
	else:
		error("Couldn't get order count")

	init_product_totals()
	get_orders(order_count)
	get_used_products()
	write_csv()

def get_orders(count):
	try:
		done = False
		index = 0
		if local.start_order:
			index = local.start_order-1
		while not done:
			num = index+1
			if num <= count:
				driver.get("https://my.ecwid.com/store/28113010#order:id="+str(num)+"&return=orders")

				wait = 0
				while wait < 5:
					key = driver.find("id","XG",True)
					if len(key) > 0:
						if key[0].text == str(num):
							break
					time.sleep(.1)
					wait += 0.1
				else:
					error("Couldn't get order #%s" % (num))


				order = Order()
				orders.append(order)
				
				order.number = '=HYPERLINK("https://my.ecwid.com/store/28113010#order:id=%s&return=orders","%s")' % (num,num)
				order.date = driver.find("class","order-details__date").text
				order.customer = driver.find("id","1W").text
				order.email = driver.find("id","rG").text
				order.phone = driver.find("id","7Q").text
				
				p = driver.find('id','aJ').find("class","ecwid-Person").find("tag","div")
				order.address1 = p[1].text
				order.city = p[2].text.split(',')[0]
				s = p[2].text.split(',')[1].split(' ')
				order.state = s[1]
				order.zip = s[2]
				order.country = p[3].text

				order.delivery = driver.find("class","order-details__shipping").find("tag","strong").find("class","gwt-InlineLabel")[0].text
				if order.delivery == "Cache Valley Deliver":
					order.delivery = "CVD"
				if order.delivery == "U.S.P.S. Priority Mail":
					order.delivery = "USPS"
				order.payment = driver.find("class","order-details__payment").find("tag","strong")[0].find("class","gwt-InlineLabel").text
				if order.payment == "Credit or debit card":
					order.payment = "Card"
				order.item_total = driver.find("id","NO").text
				order.shipping_total = driver.find("id","vx").text
				order.tax_total = driver.find("class","order-detailed-taxes").up().find('class','gwt-Label')[1].text
				order.total = driver.find('id',"Bo").text
				order.products = get_products()
			else:
				done = True
			index += 1
	except Exception as e:
			error(e)

def init_product_totals():
	try:
		for name in list(all_products.keys()):
			product_type_name = all_products[name]
			product_totals[name] = deepcopy(eval(product_type_name))

			if product_type_name == 'default_product':
				product_totals[name] = 0
			elif product_type_name is 'size_product':
				for size in list(product_totals[name].keys()):
					product_totals[name][size] = 0
			elif 'color_size_product' in product_type_name:
				for color,sizes in product_totals[name].items():
					for size in list(sizes.keys()):
						product_totals[name][color][size] = 0
			else:
				error('Invalid product type specified: '+name)
	except Exception as e:
		error(e)


def get_products():
	try:
		products = {}
		for web_product in driver.find("class","order-details-products-list__product",True):

			name = web_product.find("class","order-details-product__name").text.replace("WCYD ","")

			if name in all_products.keys():
				product_type_name = all_products[name]
				if not name in products:
					products[name] = deepcopy(eval(product_type_name))

				quantity = web_product.find("class","product-cost__multiplier").text
				if not quantity: 
					quantity = '1'

				if product_type_name is 'default_product':
					products[name] = quantity
					product_totals[name] += int(quantity)

				elif product_type_name is 'size_product':
					size = web_product.find("text~","Size:").up().find("class","product-attribute__value").text
					products[name][size] = quantity
					product_totals[name][size] += int(quantity)

				elif 'color_size_product' in product_type_name:
					color = web_product.find("text~","Color:").up().find("class","product-attribute__value").text
					size = web_product.find("text~","Size:").up().find("class","product-attribute__value").text
					products[name][color][size] = quantity
					product_totals[name][color][size] += int(quantity)

				else:
					error('Invalid product type specified: '+name)

			else:
				error("Product not specified: "+name)

		return products
	except Exception as e:
		error(e)

def get_used_products():
	try:
		global used_products
		used_products = deepcopy(product_totals)
		for name, product in product_totals.items():
			product_type_name = all_products[name]

			if product_type_name is 'default_product':
				if product is 0:
					del used_products[name]

			elif product_type_name is 'size_product':
				empty = True
				for size, value in product.items():
					if value is 0:
						del used_products[name][size]
					else:
						empty = False
				if empty:
					del	used_products[name]

			elif 'color_size_product' in product_type_name: 
				color_empty = True
				for color, sizes in product.items():
					size_empty = True
					for size, value in sizes.items():
						if value is 0:
							del used_products[name][color][size]
						else:
							size_empty = False
					if size_empty:
						del used_products[name][color]
					else:
						color_empty = False
				if color_empty:
					del used_products[name]

	except Exception as e:
		error(e)

def letter(i):
	up = string.ascii_uppercase
	the_letter = ''
				
	mult = 1
	while (i > mult*26):
		mult += 1
	else:
		if mult is 1:
			the_letter = up[i-(mult*26)-1]
		else:
			the_letter = up[mult-2]+up[i-(mult*26)-1]

	return the_letter

def write_csv():
	try:
		blank_row = list(map(lambda i: "", range(16)))

		row1 = ["","Info","","","","","","","","","","","Cost","","",""]
		row2 = blank_row
		row3 = ["#","Date","Customer","Email","Phone","Address 1","City","State","Zip","Country","Delivery","Payment","Items","Shipping","Tax","Total"]
		rows = [row1,row2,row3]

		for product_name in list(used_products.keys()):
			product_type_name = all_products[product_name]

			lists = [[],[],[]]
			product_type = used_products[product_name]

			if product_type_name is 'default_product':
				lists[0].append(product_name)

			elif product_type_name is 'size_product':				
				for size in list(product_type.keys()): 
					lists[0].append("")
					lists[2].append(size)
				lists[0][0] = product_name

			elif 'color_size_product' in product_type_name:
				for color,sizes in product_type.items():
					color_list = []
					for size in list(sizes.keys()):
						lists[0].append("")
						color_list.append("")
						lists[2].append(size)
					color_list[0] = color
					lists[1] += color_list

				lists[0][0] = product_name

			for i, the_list in enumerate(lists):
				try:
					rows[i] += the_list
				except:
					rows.append(blank_row + the_list)
		
		for i, order in enumerate(orders):
			the_order = [
				order.number,
				order.date,
				order.customer,
				order.email,
				"\'"+order.phone,
				order.address1,
				order.city,
				order.state,
				order.zip,
				order.country,
				order.delivery,
				order.payment,
				order.item_total,
				order.shipping_total,
				order.tax_total,
				order.total
			]

			for product_name in list(used_products.keys()):
				product_type_name = all_products[product_name]

				if product_type_name is 'default_product':
					if product_name in order.products:
						the_order.append(order.products[product_name])
					else:
						the_order.append("")

				elif product_type_name is 'size_product':				
					for size in list(used_products[product_name].keys()):
						if product_name in order.products:
								the_order.append(order.products[product_name][size])
						else:
							the_order.append("")

				elif 'color_size_product' in product_type_name:
					for color,sizes in used_products[product_name].items():
						for size in list(sizes.keys()):
							if product_name in order.products:
								the_order.append(order.products[product_name][color][size])
							else:
								the_order.append("")

			row_num = i+4
			column_let = letter(len(the_order))
			formula = '=SUM(Q%s:%s%s)' % (row_num,column_let,row_num)
			the_order += ['',formula]

			rows.append(the_order)
		
		rows.append([""])

		totals = []
		for i,label in enumerate(rows[0],1):
			if i == 0:
				totals.append("=A%s-A4" % (len(orders)+3))
			elif i == 3:
				totals.append("Totals")
			elif i <= 12:
				totals.append("")
			else:				
				formula = 'SUM(%s4:%s%s)' % (letter(i),letter(i),len(orders)+3)
				totals.append('=IF(%s<>0,%s,"")' % (formula,formula))

		column_let = letter(len(rows[0]))
		row_num = len(orders)+5

		formula = '=SUM(Q%s:%s%s)' % (row_num,column_let,row_num)
		totals += ['',formula]
				
				
		rows.append(totals)

		csv = ""
		for row in rows:
			csv += "	".join(row)+"\n"

		for full, sku in local.skus.items():
			csv = csv.replace(full,sku)

		pyperclip.copy(csv)
	except Exception as e:
			error(e)

try:
	run()
except Exception as e:
	error(e)
else:
	driver.quit()