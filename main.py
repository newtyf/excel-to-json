import pandas as pd
import json

class RowData:
  def __init__(self, Nom, Dep, Dis, Dir, Est, Hor, Sab):
    self.Nom = Nom
    self.Dep = Dep
    self.Dis = Dis
    self.Dir = Dir
    self.Est = Est
    self.Hor = Hor
    self.Sab = Sab
    # self.Emb = Emb

df = pd.read_excel("./agencias.xlsx", skiprows=5)

column_names = df.columns

index_agencia = column_names.get_loc("AGENCIA")
index_provincia = column_names.get_loc("PROVINCIAS")
index_distrito = column_names.get_loc("DISTRITO")
index_direccion = column_names.get_loc("DIRECCIÓN")
index_estado = column_names.get_loc("Estado")
index_hor = column_names.get_loc("HR Lun-Vie")
index_hor_sab = column_names.get_loc("HR Sab")

list_agencias = []

for row in df.values.tolist():
  agencia = RowData(row[index_agencia], row[index_provincia], row[index_distrito], row[index_direccion], row[index_estado], row[index_hor], row[index_hor_sab])
  list_agencias.append(agencia)

final_json = json.dumps([a.__dict__ for a in list_agencias], ensure_ascii=False)

with open('data.json', 'w', encoding='utf-8') as f:
  f.write(final_json)