# -*- coding: utf-8 -*-
__title__ = "Listar Categorias de Famílias"

from Autodesk.Revit.DB import *
from pyrevit import revit

doc = revit.doc

# Coletar todos os símbolos de família
collector = FilteredElementCollector(doc)\
    .OfClass(FamilySymbol)

# Criar um conjunto para armazenar categorias únicas
categorias = set()

for symbol in collector:
    categoria = symbol.Category
    if categoria:
        categorias.add((categoria.Id.IntegerValue, categoria.Name))

# Listar as categorias encontradas
print("Categorias encontradas:")
for cat_id, cat_name in sorted(categorias, key=lambda x: x[1]):
    print("ID: {}, Nome: {}".format(cat_id, cat_name))
