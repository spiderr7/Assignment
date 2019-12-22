import xlrd
from xlrd import open_workbook

# Import Elasticsearch package
from elasticsearch import Elasticsearch

# Connect to the elastic cluster
es=Elasticsearch([{'host':'101.53.136.181','port':9200}])

book = open_workbook('beginner_assignment01.xlsx')
sheet = book.sheet_by_index(0)

#Create new index   
request_body = {
	    "settings" : {
	        "number_of_shards": 1,
	        "number_of_replicas": 1
	    },
	    'mappings': {
	        'default': {
	            'properties': {
	                'Product_Name': {'type': 'text'},
	                'Model_Name': {'type': 'text'},
	                'Product_Serial_No': {'type': 'text'},
	                'Group_Associated': {'type': 'text'},
	                'product_MRP': {'type': 'text'},
	            }}}
	}
print("creating 'product_listing' index...")
es.indices.create(index = 'product_listing', body = request_body)

# read header values into the list    
keys = [sheet.cell(0, col_index).value for col_index in range(sheet.ncols)]

#Prepare data
bulk_data = []
for row_index in range(1, sheet.nrows):
    data_dict = {keys[col_index]: sheet.cell(row_index, col_index).value 
         for col_index in range(sheet.ncols)}
    op_dict = {
	        "index": {
	            "_index": 'product_listing',
	            "_type": 'default',
	            "_id": data_dict['Product Name']
	        }
	    }
    bulk_data.append(op_dict)
    bulk_data.append(data_dict)
#print(bulk_data)
    
#Finally, the data is ready to be input to the elasticsearch index
res = es.bulk(index = 'product_listing', body = bulk_data)

# check data is in there, and structure in there
es.search(body={"query": {"match_all": {}}}, index = 'product_listing')
es.indices.get_mapping(index = 'product_listing')

