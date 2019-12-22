import xlrd
from xlrd import open_workbook

#Import Elasticsearch package
from elasticsearch import Elasticsearch

# Connect to the elastic cluster
es=Elasticsearch([{'host':'101.53.136.181','port':9200}])

book = open_workbook('beginner_assignment01.xlsx')
sheet = book.sheet_by_index(1)

#Create new index 
request_body = {
	    "settings" : {
	        "number_of_shards": 1,
	        "number_of_replicas": 1
	    },
	    'mappings': {
	        'default': {
	            'properties': {
	                'group_name': {'type': 'text'},
	                'group_description': {'type': 'text'},
	                'isActive': {'type': 'text'},
	            }}}
	}
print("creating 'group_listing' index...")
es.indices.create(index = 'group_listing', body = request_body)



# read header values into the list    
keys = [sheet.cell(0, col_index).value for col_index in range(sheet.ncols)]

#Prepare data    
bulk1_data = []
for row_index in range(1, sheet.nrows):
    data_dict1 = {keys[col_index]: sheet.cell(row_index, col_index).value 
         for col_index in range(sheet.ncols)}
    op_dict1 = {
	        "index": {
	            "_index": 'group_listing',
	            "_type": 'default',
	            "_id": data_dict1['group name']
	        }
	    }
    bulk1_data.append(op_dict1)
    bulk1_data.append(data_dict1)

#print(bulk1_data)
    
#Finally, the data is ready to be input to the elasticsearch index
res = es.bulk(index = 'group_listing', body = bulk1_data)

# check data is in there, and structure in there
es.search(body={"query": {"match_all": {}}}, index = 'group_listing')
es.indices.get_mapping(index = 'group_listing')





	

	
