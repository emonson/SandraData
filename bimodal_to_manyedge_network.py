import numpy as np
import pandas as pd
import os
from scipy.sparse import coo_matrix, triu

# TODO: have columns be arguments passed to routine, along with file name

data_dir = '/Users/emonson/Dropbox/Sandra'
# data_file = 'DocsRelationsV2.xlsx'
# sheet_name = 'Sheet2-Tableau'
# person_col = 'Person-Role'
# item_col = 'Item_ID'
# out_file = 'DocsRelationsV2_edges.tsv'

data_file = 'Paintings_Shipments_AGI_v3_UpdatesRefine_V4.xls'
sheet_name = 'PaintingsShipments'
person_col = 'AGENT_NAME'
item_col = 'ID_Reg'
out_file = 'PersonShipments_manyedges.tsv'
attributes_list = ['ID_Reg', 'YEAR', 'FOLIO', '#Pntgs', 'Paintings', 'Sculptures', 'Prints', 'Books', 'Painters_Mats', 'Religious_goods', 'Gral_Merch', 'Religious_Order']
any_not_null_cols = ['Religious_Order']

xl = pd.ExcelFile(os.path.join(data_dir, data_file))
df = xl.parse(sheet_name)

# Only keeping rows with non-null person_col entries
df = df[pd.notnull( df[person_col] )]

# Find all of the unique person names
p_unique = df[person_col].unique()
# Associate an integer index with each unique person name
p_lookup = dict( zip(p_unique, np.arange(len(p_unique))) )
# Build a "reverse dictionary" to quickly look up a name based on its integer index
# (could do this with a list instead)
p_rev_lookup = dict( zip(np.arange(len(p_unique)), p_unique) )
# Create the full array of indices 
# based on how the person names actually show up in the data
p_idxs = [p_lookup[pp] for pp in df[person_col]]

# Do the same process for the items/shipments/transactions that the multiple people 
# were involved with
i_unique = df[item_col].unique()
i_lookup = dict( zip(i_unique, np.arange(len(i_unique))) )
i_rev_lookup = dict( zip(np.arange(len(i_unique)), i_unique) )
i_idxs = [i_lookup[ii] for ii in df[item_col]]

bimodal_edge_mtx = coo_matrix((np.ones_like(p_idxs),(p_idxs,i_idxs)), shape=(len(p_unique),len(i_unique))).tocsr()

# Now need to generate a copy of the spreadsheet with unique entries for the item/shipment/transaction
gg = df.groupby(df[item_col])
any_not_null = lambda x: pd.notnull(x).any()
# For most of the columns we'll assume all entries are duplicates in our desired attribute rows
dfg = gg[attributes_list].first()
# But for some we may want to apply other aggregation functions
dfg[any_not_null_cols] = gg[any_not_null_cols].aggregate(any_not_null)

graph_edges_str = []
edge_count = 0

# Now do explicit double loop through persons
for p1 in range(len(p_unique)):
    print p1, len(p_unique)
    for p2 in range(p1+1,len(p_unique)):
        pp_arr = np.multiply(bimodal_edge_mtx[p1,:].toarray().flatten(),bimodal_edge_mtx[p2,:].toarray().flatten())
        for item_id in i_unique[pp_arr > 0]:
            edge = dfg.loc[item_id,:].to_dict()
            edge['source'] = p_rev_lookup[p1]
            edge['target'] = p_rev_lookup[p2]
            edge['weight'] = 1
            edge['id'] = edge_count
            edge['type'] = 'Undirected'
            graph_edges_str.append(edge)
            edge_count = edge_count + 1
        

# graph_edges_str = [{'source':p_rev_lookup[r], 'target':p_rev_lookup[c], 'weight':int(v), 'id':int(i), 'type':'Undirected'} for i,(r,c,v) in enumerate(rcv) if r != c]

grdf = pd.DataFrame(graph_edges_str)
grdf.index.name = 'idx'
# grdf.to_excel(os.path.join(data_dir, out_file))
grdf.to_csv(os.path.join(data_dir, out_file), sep='\t', encoding='utf-8')