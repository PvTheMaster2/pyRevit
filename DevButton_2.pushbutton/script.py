# -*- coding: utf-8 -*-
__title__ = "Inserir Tomada"
__doc__ = """Versão: 1.7
Data: 19.07.2024
_____________________________________________________________________
Descrição:
Este script insere uma tomada elétrica na parede selecionada,
permitindo escolher a altura, posição horizontal e face (frontal/traseira).
_____________________________________________________________________
Como usar:
- Clique no botão e siga as instruções.
_____________________________________________________________________
Autor: Seu Nome"""

# Importações necessárias
import math  # Importação do módulo math no escopo global

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from Autodesk.Revit.UI.Selection import ObjectType
from Autodesk.Revit.Exceptions import InvalidOperationException
from Autodesk.Revit.DB.Structure import StructuralType

# Importações do pyRevit
from pyrevit import revit, forms, script

# Variáveis do documento
doc = __revit__.ActiveUIDocument.Document  # type: Document
uidoc = __revit__.ActiveUIDocument
app = __revit__.Application

# Funções auxiliares
def selecionar_familia_tomada():
    # Coletar todos os símbolos de família da categoria "Dispositivos elétricos"
    collector = FilteredElementCollector(doc) \
        .OfClass(FamilySymbol) \
        .OfCategory(BuiltInCategory.OST_ElectricalFixtures)

    tomadas = []
    for symbol in collector:
        # Obter o nome da família usando BuiltInParameter
        family_name_param = symbol.get_Parameter(BuiltInParameter.ALL_MODEL_FAMILY_NAME)
        if family_name_param and family_name_param.HasValue:
            family_name = family_name_param.AsString()
        else:
            family_name = "Sem Família"

        # Obter o nome do símbolo usando BuiltInParameter
        symbol_name_param = symbol.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME)
        if symbol_name_param and symbol_name_param.HasValue:
            symbol_name = symbol_name_param.AsString()
        else:
            symbol_name = "Sem Nome"

        # Filtrar por famílias que contenham "Tomada" no nome
        if "Tomada" in family_name or "Tomada" in symbol_name:
            tomadas.append(symbol)

    if not tomadas:
        forms.alert("Nenhuma família de tomadas encontrada no projeto.", exitscript=True)

    # Criar um dicionário de opções
    tomadas_dict = {}
    for tomada in tomadas:
        try:
            # Obter o nome da família
            family_name_param = tomada.get_Parameter(BuiltInParameter.ALL_MODEL_FAMILY_NAME)
            if family_name_param and family_name_param.HasValue:
                family_name = family_name_param.AsString()
            else:
                family_name = "Sem Família"

            # Obter o nome do símbolo
            symbol_name_param = tomada.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME)
            if symbol_name_param and symbol_name_param.HasValue:
                symbol_name = symbol_name_param.AsString()
            else:
                symbol_name = "Sem Nome"

            display_name = "{} : {}".format(family_name, symbol_name)
            tomadas_dict[display_name] = tomada
        except Exception as e:
            # Se ocorrer um erro, ignorar este elemento
            pass

    if not tomadas_dict:
        forms.alert("Nenhuma família de tomadas válida encontrada.", exitscript=True)

    # Ordenar os nomes para exibição
    tomadas_nomes_ordenados = sorted(tomadas_dict.keys())

    # Permitir que o usuário selecione uma tomada
    tomada_selecionada_nome = forms.SelectFromList.show(
        tomadas_nomes_ordenados,
        title='Selecione uma Tomada',
        button_name='Selecionar',
        multiselect=False
    )

    if not tomada_selecionada_nome:
        forms.alert("Nenhuma tomada selecionada.", exitscript=True)

    tomada_selecionada = tomadas_dict[tomada_selecionada_nome]

    # Ativar o símbolo da família, se necessário
    if not tomada_selecionada.IsActive:
        with revit.Transaction("Ativar Família"):
            tomada_selecionada.Activate()
            doc.Regenerate()

    return tomada_selecionada

def selecionar_parede():
    # Permitir que o usuário selecione uma parede
    sel = uidoc.Selection
    referencia = sel.PickObject(ObjectType.Element, 'Selecione a parede onde a tomada será inserida.')
    parede = doc.GetElement(referencia.ElementId)

    if not isinstance(parede, Wall):
        forms.alert("O elemento selecionado não é uma parede.", exitscript=True)

    return parede

def obter_ponto_insercao(parede):
    # Obter a altura desejada do usuário
    altura_metros_input = forms.ask_for_string(
        prompt="Insira a altura da tomada em metros:",
        title="Altura da Tomada",
        default="1.10"
    )
    try:
        altura_metros = float(altura_metros_input.replace(',', '.'))
    except ValueError:
        forms.alert("Entrada inválida. Usando altura padrão de 1.10 metros.")
        altura_metros = 1.10

    # Converter metros para pés
    altura_pes = altura_metros * 3.28084

    # Obter a localização da parede
    loc_curve = parede.Location
    if not isinstance(loc_curve, LocationCurve):
        forms.alert("Não foi possível obter a localização da parede.", exitscript=True)

    # Obter a curva (linha) central da parede
    curva = loc_curve.Curve

    # Solicitar ao usuário a distância ao longo da parede
    distancia_metros_input = forms.ask_for_string(
        prompt="Insira a distância ao longo da parede em metros (0 para início, deixe em branco para meio):",
        title="Posição Horizontal",
        default=""
    )
    if distancia_metros_input.strip() == "":
        # Usar o ponto médio se o usuário não inserir nada
        parametro_normalizado = 0.5
    else:
        try:
            distancia_metros = float(distancia_metros_input.replace(',', '.'))
            comprimento_parede = curva.Length  # Comprimento em pés
            comprimento_metros = comprimento_parede / 3.28084
            parametro_normalizado = distancia_metros / comprimento_metros
            parametro_normalizado = max(0.0, min(1.0, parametro_normalizado))  # Garantir que esteja entre 0 e 1
        except ValueError:
            forms.alert("Entrada inválida. Usando o ponto médio da parede.")
            parametro_normalizado = 0.5

    # Obter o ponto ao longo da curva
    ponto_na_curva = curva.Evaluate(parametro_normalizado, True)

    # Manter o ponto de inserção no nível base (Z = 0)
    ponto_insercao = XYZ(ponto_na_curva.X, ponto_na_curva.Y, 0)

    # Obter a espessura da parede
    espessura_parede = parede.WallType.Width  # Em pés

    # Perguntar ao usuário em qual face deseja inserir a tomada
    face_opcoes = ['Frontal', 'Traseira']
    face_selecionada = forms.SelectFromList.show(
        face_opcoes,
        title='Selecione a Face da Parede',
        button_name='Selecionar',
        multiselect=False
    )

    rotacionar_180 = False

    if not face_selecionada:
        forms.alert("Nenhuma face selecionada. Usando a posição central da parede.")

    else:
        # Calcular o vetor normal da parede
        curva_parede = loc_curve.Curve
        # Obter o vetor direcional da parede
        direcao_parede = (curva_parede.GetEndPoint(1) - curva_parede.GetEndPoint(0)).Normalize()
        # Calcular o vetor normal (perpendicular) à parede no plano XY
        vetor_normal = XYZ(-direcao_parede.Y, direcao_parede.X, 0).Normalize()

        # Calcular o deslocamento (metade da espessura da parede)
        deslocamento = (espessura_parede / 2)

        if face_selecionada == 'Frontal':
            # Deslocar na direção do vetor normal
            ponto_insercao += vetor_normal * deslocamento
        elif face_selecionada == 'Traseira':
            # Deslocar na direção oposta ao vetor normal
            ponto_insercao -= vetor_normal * deslocamento
            # Marcar para rotacionar 180 graus
            rotacionar_180 = True

    return ponto_insercao, altura_pes, rotacionar_180

def inserir_tomada(parede, tomada_selecionada, ponto_insercao, altura_pes, rotacionar_180):
    # Inserir a tomada usando a parede como host
    with revit.Transaction("Inserir Tomada"):
        tomada_instancia = doc.Create.NewFamilyInstance(
            ponto_insercao,
            tomada_selecionada,
            parede,
            StructuralType.NonStructural
        )

        # Ajustar a orientação da tomada para ficar alinhada com a parede
        # Obter a direção da parede
        loc_curve = parede.Location
        curva = loc_curve.Curve
        direcao_parede = (curva.GetEndPoint(1) - curva.GetEndPoint(0)).Normalize()
        # Calcular o ângulo entre a direção da parede e o eixo X
        angulo = XYZ.BasisX.AngleTo(direcao_parede)
        # Determinar o sentido da rotação
        cross = XYZ.BasisX.CrossProduct(direcao_parede)
        if cross.Z < 0:
            angulo = -angulo

        # Se a face selecionada for "Traseira", rotacionar 180 graus
        if rotacionar_180:
            angulo += math.pi  # Adicionar 180 graus em radianos

        # Criar o eixo de rotação
        eixo_rotacao = Line.CreateBound(ponto_insercao, ponto_insercao + XYZ.BasisZ)
        # Rotacionar a tomada
        ElementTransformUtils.RotateElement(
            doc,
            tomada_instancia.Id,
            eixo_rotacao,
            angulo
        )

        # Definir a altura (elevação) da tomada
        sucesso = False

        # Lista de possíveis nomes de parâmetros que controlam a altura
        nomes_parametros_altura = [
            "Offset",
            "Deslocamento",
            "Elevação",
            "Elevação do Ponto",
            "Sill Height",
            "Head Height",
            "Height",
            "Base Offset",
            "Top Offset"
        ]

        for nome_param in nomes_parametros_altura:
            offset_param = tomada_instancia.LookupParameter(nome_param)
            if offset_param and not offset_param.IsReadOnly:
                try:
                    offset_param.Set(altura_pes)
                    sucesso = True
                    break
                except Exception as e:
                    # Ignorar erros e tentar o próximo parâmetro
                    pass

        if not sucesso:
            # Como último recurso, ajustar a posição Z do ponto
            location = tomada_instancia.Location
            if isinstance(location, LocationPoint):
                point = location.Point
                new_point = XYZ(point.X, point.Y, altura_pes)
                location.Point = new_point

    forms.alert("Tomada inserida com sucesso!")

def inserir_tomada_na_parede():
    # Selecionar a família de tomada
    tomada_selecionada = selecionar_familia_tomada()

    # Selecionar a parede
    parede = selecionar_parede()

    # Obter o ponto de inserção, altura e informação de rotação
    ponto_insercao, altura_pes, rotacionar_180 = obter_ponto_insercao(parede)

    # Inserir a tomada
    inserir_tomada(parede, tomada_selecionada, ponto_insercao, altura_pes, rotacionar_180)

# Executar o script
if __name__ == "__main__":
    inserir_tomada_na_parede()
