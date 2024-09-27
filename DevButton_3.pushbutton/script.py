# -*- coding: utf-8 -*-
__title__ = "Criar Parede"
__doc__ = """Versão: 1.0
Data: 15.07.2024
_____________________________________________________________________
Descrição:
Este script cria uma parede reta no Revit usando a API do Revit.
_____________________________________________________________________
Como usar:
- Clique no botão para criar a parede.
_____________________________________________________________________
Autor: Seu Nome"""

# Importações necessárias
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from Autodesk.Revit.Exceptions import InvalidOperationException

# Importações do pyRevit
from pyrevit import revit, forms

# Variáveis do documento
doc = __revit__.ActiveUIDocument.Document  # type: Document
uidoc = __revit__.ActiveUIDocument
app = __revit__.Application


# Função principal
def criar_parede():
    # Definir os pontos de início e fim da parede
    pt_inicio = XYZ(0, 0, 0)
    pt_fim = XYZ(10, 0, 0)
    linha = Line.CreateBound(pt_inicio, pt_fim)

    # Obter o nível ativo
    nivel = doc.ActiveView.GenLevel

    # Obter o tipo de parede padrão
    tipo_parede_id = doc.GetDefaultElementTypeId(ElementTypeGroup.WallType)
    tipo_parede = doc.GetElement(tipo_parede_id)

    # Definir altura e deslocamento
    altura_parede = 3.0  # Altura em pés (1 pé ≈ 0,3048 metros)
    deslocamento_base = 0.0

    # Iniciar uma transação
    try:
        with revit.Transaction("Criar Parede"):
            # Criar a parede
            parede = Wall.Create(doc, linha, tipo_parede.Id, nivel.Id, altura_parede, deslocamento_base, False, False)
        forms.alert("Parede criada com sucesso!")
    except Exception as e:
        forms.alert("Erro ao criar a parede:\n{}".format(e))


# Executar a função principal
if __name__ == "__main__":
    criar_parede()
