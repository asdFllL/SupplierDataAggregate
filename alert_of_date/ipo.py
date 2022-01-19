import akshare as ak
ipo = ak.stock_em_xgsglb(market="全部股票")
print(ipo.head())
ipo.to_csv('ipo.csv')