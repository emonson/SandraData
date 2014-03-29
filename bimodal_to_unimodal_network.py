import numpy as N
import pandas as pd
import os
from scipy.sparse import coo_matrix, triu

# TODO: have columns be arguments passed to routine, along with file name

data_dir = '/Users/emonson/Dropbox/Sandra'
data_file = 'DocsRelationsV2.xlsx'
out_file = 'DocsRelationsV2_edges.tsv'
person_col = 'Person-Role'
item_col = 'Item_ID'
sheet_name = 'Sheet2-Tableau'

xl = pd.ExcelFile(os.path.join(data_dir, data_file))
df = xl.parse(sheet_name)

p_unique = df[person_col].unique()
p_lookup = dict( zip(p_unique, N.arange(len(p_unique))) )
p_rev_lookup = dict( zip(N.arange(len(p_unique)), p_unique) )
p_idxs = [p_lookup[pp] for pp in df[person_col]]

i_unique = df[item_col].unique()
i_lookup = dict( zip(i_unique, N.arange(len(i_unique))) )
# i_rev_lookup = dict( zip(N.arange(len(i_unique)), i_unique) )
i_idxs = [i_lookup[ii] for ii in df[item_col]]

bimodal_edge_mtx = coo_matrix((N.ones_like(p_idxs),(p_idxs,i_idxs)), shape=(len(p_unique),len(i_unique))).tocsr()

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
# grdf.to_excel(os.path.join(data_dir, out_file))
grdf.to_csv(os.path.join(data_dir, out_file), sep='\t', encoding='utf-8')