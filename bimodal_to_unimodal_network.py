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
out_file = 'PersonShipments_edges.tsv'


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

unimodal_edge_mtx = bimodal_edge_mtx * bimodal_edge_mtx.T

# Only recording one direction so need to specify bidirectional in gephi
triu_uni_edge_coo = triu(unimodal_edge_mtx.tocoo())
rcv = zip(triu_uni_edge_coo.row, triu_uni_edge_coo.col, triu_uni_edge_coo.data)

# vt is the value of the edge coming from the target back towards the source
# which we need for some edge rendering schemes
# excluding self-edges
# graph_edges = [{'source':int(r), 'target':int(c), 'v':int(v), 'vt':int(unimodal_edge_mtx[c,r]), 'i':int(i)} for i,(r,c,v) in enumerate(zip(uni_edge_coo.row, uni_edge_coo.col, uni_edge_coo.data)) if r != c]

graph_edges = [{'source':int(r), 'target':int(c), 'weight':int(v), 'id':int(i), 'type':'Undirected'} for i,(r,c,v) in enumerate(rcv) if r != c]
graph_edges_str = [{'source':p_rev_lookup[r], 'target':p_rev_lookup[c], 'weight':int(v), 'id':int(i), 'type':'Undirected'} for i,(r,c,v) in enumerate(rcv) if r != c]

grdf = pd.DataFrame(graph_edges_str)
grdf.index.name = 'idx'
# grdf.to_excel(os.path.join(data_dir, out_file))
grdf.to_csv(os.path.join(data_dir, out_file), sep='\t', encoding='utf-8')