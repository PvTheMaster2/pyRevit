# -*- coding: utf-8 -*-
__title__ = "Multiplas Tomadas"
__doc__ = """Versão: 1.6
_____________________________________________________________________
Descrição:
Este script insere múltiplas tomadas elétricas na parede selecionada,
permitindo escolher a quantidade, altura, intervalo e face (frontal/traseira).
Antes da inserção final, o script exibe uma pré-visualização das posições das tomadas.
_____________________________________________________________________
Como usar:
- Clique no botão e siga as instruções.
_____________________________________________________________________
Autor: Seu Nome"""

# Importações necessárias
import clr
clr.AddReference('System.Windows.Forms')
from System.Windows.Forms import DialogResult, MessageBox, MessageBoxButtons, MessageBoxIcon

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
    """Permite que o usuário selecione uma família de tomada elétrica."""
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
    """Permite que o usuário selecione uma parede."""
    sel = uidoc.Selection
    try:
        referencia = sel.PickObject(ObjectType.Element, 'Selecione a parede onde as tomadas serão inseridas.')
        parede = doc.GetElement(referencia.ElementId)
    except InvalidOperationException:
        forms.alert("Nenhuma parede selecionada.", exitscript=True)

    if not isinstance(parede, Wall):
        forms.alert("O elemento selecionado não é uma parede.", exitscript=True)

    return parede

def obter_parametros_usuario(parede):
    """Obtém os parâmetros do usuário para a inserção das tomadas."""
    # Obter a altura desejada do usuário
    altura_metros_input = forms.ask_for_string(
        prompt="Insira a altura das tomadas em metros:",
        title="Altura das Tomadas",
        default="1.10"
    )
    try:
        altura_metros = float(altura_metros_input.replace(',', '.'))
    except ValueError:
        forms.alert("Entrada inválida. Usando altura padrão de 1.10 metros.")
        altura_metros = 1.10

    # Obter o número de tomadas
    numero_tomadas_input = forms.ask_for_string(
        prompt="Insira o número de tomadas a serem inseridas:",
        title="Número de Tomadas",
        default="1"
    )
    try:
        numero_tomadas = int(numero_tomadas_input)
        if numero_tomadas < 1:
            forms.alert("Número inválido. Usando 1 tomada.")
            numero_tomadas = 1
    except ValueError:
        forms.alert("Entrada inválida. Usando 1 tomada.")
        numero_tomadas = 1

    # Obter o comprimento do intervalo
    intervalo_metros_input = forms.ask_for_string(
        prompt="Insira o comprimento do intervalo em metros (deixe em branco para usar o comprimento total da parede):",
        title="Comprimento do Intervalo",
        default=""
    )
    try:
        if intervalo_metros_input.strip() == "":
            # Usar o comprimento total da parede
            loc_curve = parede.Location
            curva = loc_curve.Curve
            comprimento_parede = curva.Length  # Comprimento em pés
            intervalo_metros = comprimento_parede / 3.28084  # Converter para metros
        else:
            intervalo_metros = float(intervalo_metros_input.replace(',', '.'))
            if intervalo_metros <= 0:
                forms.alert("Comprimento inválido. Usando o comprimento total da parede.")
                loc_curve = parede.Location
                curva = loc_curve.Curve
                comprimento_parede = curva.Length  # Comprimento em pés
                intervalo_metros = comprimento_parede / 3.28084  # Converter para metros
    except ValueError:
        forms.alert("Entrada inválida. Usando o comprimento total da parede.")
        loc_curve = parede.Location
        curva = loc_curve.Curve
        comprimento_parede = curva.Length  # Comprimento em pés
        intervalo_metros = comprimento_parede / 3.28084  # Converter para metros

    # Perguntar ao usuário em qual face deseja inserir as tomadas
    face_opcoes = ['Frontal', 'Traseira']
    face_selecionada = forms.SelectFromList.show(
        face_opcoes,
        title='Selecione a Face da Parede',
        button_name='Selecionar',
        multiselect=False
    )

    if not face_selecionada:
        forms.alert("Nenhuma face selecionada. Usando a posição central da parede.")
        face_selecionada = None  # Indica que não foi selecionada

    return altura_metros, numero_tomadas, intervalo_metros, face_selecionada

def calcular_pontos_insercao(parede, altura_metros, numero_tomadas, intervalo_metros, face_selecionada):
    """Calcula os pontos de inserção das tomadas."""
    # Converter metros para pés
    altura_pes = altura_metros * 3.28084
    intervalo_pes = intervalo_metros * 3.28084

    # Obter a localização da parede
    loc_curve = parede.Location
    if not isinstance(loc_curve, LocationCurve):
        forms.alert("Não foi possível obter a localização da parede.", exitscript=True)

    # Obter a curva (linha) central da parede
    curva = loc_curve.Curve
    comprimento_parede = curva.Length

    # Se o intervalo for maior que o comprimento da parede, ajustar
    if intervalo_pes > comprimento_parede:
        intervalo_pes = comprimento_parede

    # Definir o ponto inicial para o intervalo (centralizado)
    direcao_parede = (curva.GetEndPoint(1) - curva.GetEndPoint(0)).Normalize()
    centro_parede = curva.Evaluate(0.5, True)
    deslocamento_inicio = (intervalo_pes / 2) * (-direcao_parede)
    ponto_inicial = centro_parede + deslocamento_inicio

    # Calcular o espaçamento entre as tomadas
    if numero_tomadas == 1:
        espacamento = 0
    else:
        espacamento = intervalo_pes / (numero_tomadas - 1)

    # Lista para armazenar os pontos de inserção
    pontos_insercao = []

    for i in range(numero_tomadas):
        distancia = espacamento * i
        ponto_na_parede = ponto_inicial + direcao_parede * distancia
        # Ajustar o ponto para a altura desejada
        ponto_insercao = XYZ(ponto_na_parede.X, ponto_na_parede.Y, ponto_na_parede.Z + altura_pes)
        pontos_insercao.append(ponto_insercao)

    # Obter a espessura da parede
    espessura_parede = parede.WallType.Width  # Em pés

    if face_selecionada:
        # Calcular o vetor normal da parede
        vetor_normal = XYZ(-direcao_parede.Y, direcao_parede.X, 0).Normalize()

        # Calcular o deslocamento (metade da espessura da parede)
        deslocamento = (espessura_parede / 2)

        if face_selecionada == 'Frontal':
            # Deslocar na direção do vetor normal
            deslocamento_vetor = vetor_normal * deslocamento
        elif face_selecionada == 'Traseira':
            # Deslocar na direção oposta ao vetor normal
            deslocamento_vetor = vetor_normal * (-deslocamento)
        else:
            deslocamento_vetor = XYZ(0, 0, 0)
    else:
        deslocamento_vetor = XYZ(0, 0, 0)

    # Aplicar o deslocamento a todos os pontos
    pontos_insercao = [ponto + deslocamento_vetor for ponto in pontos_insercao]

    return pontos_insercao, direcao_parede

def criar_preview(pontos_insercao, direcao_parede):
    """Cria elementos de pré-visualização das posições das tomadas."""
    preview_ids = []

    vetor_perpendicular = XYZ(-direcao_parede.Y, direcao_parede.X, 0).Normalize()

    with revit.Transaction("Criar Preview"):
        for ponto in pontos_insercao:
            # Criar uma pequena linha horizontal perpendicular à direção da parede
            p1 = ponto - (vetor_perpendicular * 0.2)  # 0.2 pés (~0.06 metros)
            p2 = ponto + (vetor_perpendicular * 0.2)

            # Criar um SketchPlane que contém a linha
            try:
                plano = Plane.CreateByThreePoints(p1, p2, ponto)
                sketch_plane = SketchPlane.Create(doc, plano)
            except Exception as e:
                # Se o plano não puder ser criado, usar um plano padrão (XY)
                plano = Plane.CreateByNormalAndOrigin(XYZ.BasisZ, ponto)
                sketch_plane = SketchPlane.Create(doc, plano)

            # Criar a linha de pré-visualização
            linha_preview = Line.CreateBound(p1, p2)

            # Criar um ModelCurve para a pré-visualização
            try:
                model_line = doc.Create.NewModelCurve(linha_preview, sketch_plane)
                preview_ids.append(model_line.Id)
            except Exception as e:
                # Se ocorrer um erro na criação da ModelCurve, ignorar este ponto
                pass

    return preview_ids

def remover_preview(preview_ids):
    """Remove os elementos de pré-visualização."""
    with revit.Transaction("Remover Preview"):
        for elem_id in preview_ids:
            try:
                doc.Delete(elem_id)
            except Exception as e:
                # Se ocorrer um erro na deleção, ignorar este elemento
                pass

def inserir_tomadas(parede, tomada_selecionada, pontos_insercao):
    """Insere as tomadas nas posições calculadas."""
    with revit.Transaction("Inserir Tomadas"):
        for ponto_insercao in pontos_insercao:
            try:
                # Inserir a tomada usando a parede como host
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

                # Criar o eixo de rotação (linha vertical passando pelo ponto de inserção)
                eixo_rotacao = Line.CreateBound(ponto_insercao, ponto_insercao + XYZ.BasisZ)

                # Rotacionar a tomada
                ElementTransformUtils.RotateElement(
                    doc,
                    tomada_instancia.Id,
                    eixo_rotacao,
                    angulo
                )
            except Exception as e:
                # Se ocorrer um erro na inserção ou rotação, ignorar esta tomada
                pass

def inserir_tomadas_na_parede():
    """Função principal para inserir tomadas na parede com pré-visualização."""
    try:
        # Selecionar a família de tomada
        tomada_selecionada = selecionar_familia_tomada()

        # Selecionar a parede
        parede = selecionar_parede()

        # Obter os parâmetros do usuário
        altura_metros, numero_tomadas, intervalo_metros, face_selecionada = obter_parametros_usuario(parede)

        # Calcular os pontos de inserção e a direção da parede
        pontos_insercao, direcao_parede = calcular_pontos_insercao(
            parede, altura_metros, numero_tomadas, intervalo_metros, face_selecionada
        )

        # Criar pré-visualização
        preview_ids = criar_preview(pontos_insercao, direcao_parede)

        # Atualizar a vista para garantir que os ModelCurves apareçam
        uidoc.RefreshActiveView()

        # Perguntar ao usuário se deseja confirmar a inserção
        resultado = MessageBox.Show(
            "Deseja inserir as tomadas nas posições marcadas?",
            "Confirmar Inserção",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Question
        )

        # Remover pré-visualização
        remover_preview(preview_ids)

        if resultado == DialogResult.Yes:
            # Inserir as tomadas
            inserir_tomadas(parede, tomada_selecionada, pontos_insercao)
            # Confirmar as alterações
            forms.alert("Tomadas inseridas com sucesso!")
        else:
            forms.alert("Inserção cancelada pelo usuário.")

    except Exception as e:
        forms.alert("Ocorreu um erro: {}".format(e))

# Executar o script
if __name__ == "__main__":
    inserir_tomadas_na_parede()
