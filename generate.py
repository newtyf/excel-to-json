import pandas as pd
import json
import argparse

class Agency:
  def __init__(self, Nom, Dep, Dis, Dir, Est, Hor, Sab, Emb):
    self.Nom = Nom
    self.Dep = Dep
    self.Dis = Dis
    self.Dir = Dir
    self.Est = Est
    self.Hor = Hor
    self.Sab = Sab
    self.Emb = Emb

parser = argparse.ArgumentParser(description='Generador de json desde un excel de agencias')
parser.add_argument('-f', '--file', type=str, required=True, help='Ruta del archivo excel de agencias')
parser.add_argument('-o', '--output', type=str, default='./agencias-abiertas.json', required=False, help='Ruta donde se generar el archivo')

args = parser.parse_args();

df = pd.read_excel(args.file, skiprows=5)

column_names = df.columns
column_name_embosadora = column_names[column_names.str.contains("Embozadora")].tolist()

def generate(rows: list, output_path: str = 'data.json'):
  list_agencies = []
  for row in rows:
    agency = Agency(row[index_agencia], row[index_provincia], row[index_distrito], row[index_direccion], row[index_estado], row[index_hor], row[index_hor_sab], row[index_embosadora])
    list_agencies.append(agency)

  final_json = json.dumps([a.__dict__ for a in list_agencies], ensure_ascii=False, indent=2)
  with open(output_path, 'w', encoding='utf-8') as f:
    f.write(final_json)

try:
  index_agencia = column_names.get_loc("AGENCIA")
  index_provincia = column_names.get_loc("PROVINCIAS")
  index_distrito = column_names.get_loc("DISTRITO")
  index_direccion = column_names.get_loc("DIRECCIÓN")
  index_estado = column_names.get_loc("Estado")
  index_hor = column_names.get_loc("HR Lun-Vie")
  index_hor_sab = column_names.get_loc("HR Sab")
  index_embosadora = column_names.get_loc(column_name_embosadora[0])

  if __name__ == '__main__' :
    generate(df.values.tolist(), args.output)
except:
  print('El archivo excel no tiene las columnas necesarias para generar el json')
