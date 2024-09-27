# -*- coding: utf-8 -*-
__title__ = "Inserir Múltiplas Tomadas com Parâmetros Elétricos"
__doc__ = """Versão: 2.8
Data: 25.09.2024
_____________________________________________________________________
Descrição:
Este script insere múltiplas tomadas elétricas na parede selecionada,
permitindo escolher a quantidade, altura, intervalo e face (frontal/traseira).
As tomadas inseridas podem ser atribuídas a um circuito elétrico, que
pode ser conectado a um painel selecionado pelo usuário dentre os disponíveis.
_____________________________________________________________________
Como usar:
- Clique no botão e siga as instruções.
_____________________________________________________________________
Autor: Seu Nome"""

# Importações necessárias
import clr
import math
import traceback

# Adicionar referências à Revit API
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')

# Importar as classes necessárias do Revit API
from Autodesk.Revit.DB import (
    FilteredElementCollector,
    FamilySymbol,
    BuiltInCategory,
    BuiltInParameter,
    XYZ,
    Line,
    Transaction,
    ElementTransformUtils,
    StorageType,
    Wall,
    LocationCurve,
    FamilyInstance,
    LocationPoint,
    Plane,
    SketchPlane,
    ElementId,
    Color,
    OverrideGraphicSettings,
)
from Autodesk.Revit.DB.Electrical import ElectricalSystem, ElectricalSystemType
from Autodesk.Revit.DB.Structure import StructuralType
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Selection import ObjectType
from Autodesk.Revit.Exceptions import InvalidOperationException

# Importações do pyRevit
from pyrevit import revit, forms, script

# Importar System.Windows.Forms para caixas de diálogo personalizadas
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
from System.Windows.Forms import (
    Form,
    Label,
    TextBox,
    Button,
    ComboBox,
    DialogResult,
    FormStartPosition,
    MessageBox,
    MessageBoxButtons,
    MessageBoxIcon,
    AutoSizeMode,
    AnchorStyles,
)
from System.Drawing import Point, Size

# Importar List do .NET
clr.AddReference("System")
from System.Collections.Generic import List

# Variáveis do documento
doc = __revit__.ActiveUIDocument.Document  # Documento ativo do Revit
uidoc = __revit__.ActiveUIDocument  # Documento UI ativo


# Função para selecionar a família de tomada
def selecionar_familia_tomada():
    """Permite que o usuário selecione uma família de tomada elétrica."""
    # Lista de categorias para procurar famílias de tomadas
    categorias = [
        BuiltInCategory.OST_ElectricalFixtures,
        BuiltInCategory.OST_ElectricalEquipment,
        BuiltInCategory.OST_GenericModel,  # Adicione outras categorias se necessário
    ]

    all_outlets = []
    for categoria in categorias:
        collector = FilteredElementCollector(doc).OfClass(FamilySymbol).OfCategory(categoria)
        coletados = list(collector)
        all_outlets.extend(coletados)

        # Mensagens de depuração
        if coletados:
            try:
                nomes_familias = ", ".join([
                    "{} : {}".format(
                        symbol.get_Parameter(BuiltInParameter.ALL_MODEL_FAMILY_NAME).AsString(),
                        symbol.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString()
                    )
                    for symbol in coletados
                    if symbol.get_Parameter(BuiltInParameter.ALL_MODEL_FAMILY_NAME) and symbol.get_Parameter(
                        BuiltInParameter.ALL_MODEL_TYPE_NAME)
                ])
            except Exception:
                nomes_familias = "Erro ao coletar nomes das famílias."

            forms.alert("Categoria {}: {} símbolos coletados.\nFamílias: {}".format(
                categoria, len(coletados), nomes_familias))
        else:
            forms.alert("Categoria {}: Nenhum símbolo coletado.".format(categoria))

    tomadas = []
    for symbol in all_outlets:
        try:
            # Obter o nome da família usando BuiltInParameter
            family_name_param = symbol.get_Parameter(BuiltInParameter.ALL_MODEL_FAMILY_NAME)
            family_name = (
                family_name_param.AsString()
                if family_name_param and family_name_param.HasValue
                else "Sem Família"
            )

            # Obter o nome do símbolo usando BuiltInParameter
            symbol_name_param = symbol.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME)
            symbol_name = (
                symbol_name_param.AsString()
                if symbol_name_param and symbol_name_param.HasValue
                else "Sem Nome"
            )

            # Filtrar por famílias que contenham "tomada" ou "outlet" no nome (case-insensitive)
            if ("tomada" in family_name.lower() or "outlet" in family_name.lower() or
                    "tomada" in symbol_name.lower() or "outlet" in symbol_name.lower()):
                tomadas.append(symbol)
        except Exception as e:
            # Mensagem de erro para depuração
            forms.alert("Erro ao processar símbolo: {}".format(e))

    if not tomadas:
        # Para depuração: listar todas as famílias encontradas
        try:
            todas_familias = "\n".join(
                [
                    "{} : {}".format(
                        symbol.get_Parameter(BuiltInParameter.ALL_MODEL_FAMILY_NAME).AsString(),
                        symbol.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString()
                    )
                    for symbol in all_outlets
                    if symbol.get_Parameter(BuiltInParameter.ALL_MODEL_FAMILY_NAME) and symbol.get_Parameter(
                    BuiltInParameter.ALL_MODEL_TYPE_NAME)
                ]
            )
        except Exception:
            todas_familias = "Erro ao coletar nomes das famílias."

        forms.alert(
            "Nenhuma família de tomadas encontrada no projeto.\n\nFamílias encontradas:\n{}".format(todas_familias),
            exitscript=True
        )

    # Criar um dicionário de opções
    tomadas_dict = {}
    for tomada in tomadas:
        try:
            # Obter o nome da família
            family_name_param = tomada.get_Parameter(BuiltInParameter.ALL_MODEL_FAMILY_NAME)
            family_name = (
                family_name_param.AsString()
                if family_name_param and family_name_param.HasValue
                else "Sem Família"
            )

            # Obter o nome do símbolo
            symbol_name_param = tomada.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME)
            symbol_name = (
                symbol_name_param.AsString()
                if symbol_name_param and symbol_name_param.HasValue
                else "Sem Nome"
            )

            display_name = "{} : {}".format(family_name, symbol_name)
            tomadas_dict[display_name] = tomada
        except Exception:
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
        multiselect=False,
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


# Função para selecionar a parede
def selecionar_parede():
    """Permite que o usuário selecione uma parede."""
    sel = uidoc.Selection
    try:
        referencia = sel.PickObject(
            ObjectType.Element,
            'Selecione a parede onde as tomadas serão inseridas.'
        )
        parede = doc.GetElement(referencia.ElementId)
    except InvalidOperationException:
        forms.alert("Nenhuma parede selecionada.", exitscript=True)

    if not isinstance(parede, Wall):
        forms.alert("O elemento selecionado não é uma parede.", exitscript=True)

    return parede


# Função para selecionar tensão e fases (mover para fora de obter_parametros_usuario)
def get_tensao_e_fases():
    """Permite ao usuário selecionar o sistema de tensão e o número de fases."""

    class VoltageForm(Form):
        def __init__(self):
            self.Text = 'Seleção de Tensão e Número de Fases'
            self.Width = 400
            self.Height = 250
            self.StartPosition = FormStartPosition.CenterScreen
            self.AutoSize = True
            self.AutoSizeMode = AutoSizeMode.GrowAndShrink
            self.results = None

            y = 10
            dy = 30

            # Selecionar o sistema de tensão
            self.label_sistema_tensao = Label()
            self.label_sistema_tensao.Text = 'Selecione o sistema de tensão:'
            self.label_sistema_tensao.Location = Point(10, y)
            self.label_sistema_tensao.Width = 350
            self.Controls.Add(self.label_sistema_tensao)

            y += dy
            self.combobox_sistema_tensao = ComboBox()
            self.combobox_sistema_tensao.Items.Add('220/380 V')
            self.combobox_sistema_tensao.Items.Add('127/220 V')
            self.combobox_sistema_tensao.SelectedIndex = 0
            self.combobox_sistema_tensao.Location = Point(10, y)
            self.Controls.Add(self.combobox_sistema_tensao)

            y += dy
            # Selecionar o número de fases
            self.label_num_fases = Label()
            self.label_num_fases.Text = 'Selecione o número de fases:'
            self.label_num_fases.Location = Point(10, y)
            self.label_num_fases.Width = 350
            self.Controls.Add(self.label_num_fases)

            y += dy
            self.combobox_num_fases = ComboBox()
            self.combobox_num_fases.Items.Add('1')
            self.combobox_num_fases.Items.Add('2')
            self.combobox_num_fases.Items.Add('3')
            self.combobox_num_fases.SelectedIndex = 0
            self.combobox_num_fases.Location = Point(10, y)
            self.Controls.Add(self.combobox_num_fases)

            y += dy + 10
            # OK and Cancel buttons
            self.button_ok = Button()
            self.button_ok.Text = 'OK'
            self.button_ok.Location = Point(220, y)
            self.button_ok.Click += self.ok_clicked
            self.Controls.Add(self.button_ok)

            self.button_cancel = Button()
            self.button_cancel.Text = 'Cancelar'
            self.button_cancel.Location = Point(300, y)
            self.button_cancel.Click += self.cancel_clicked
            self.Controls.Add(self.button_cancel)

        def ok_clicked(self, sender, event):
            self.results = {
                'sistema_tensao': self.combobox_sistema_tensao.SelectedItem,
                'numero_fases': int(self.combobox_num_fases.SelectedItem),
            }
            self.DialogResult = DialogResult.OK
            self.Close()

        def cancel_clicked(self, sender, event):
            self.DialogResult = DialogResult.Cancel
            self.Close()

    form = VoltageForm()
    result = form.ShowDialog()
    if result != DialogResult.OK:
        forms.alert("Entrada cancelada pelo usuário.", exitscript=True)

    sistema_tensao = form.results['sistema_tensao']
    numero_fases = form.results['numero_fases']

    # Determinar a tensão de acordo com o sistema de tensão e número de fases
    if sistema_tensao == '220/380 V':
        if numero_fases == 1:
            tensao = 220.0  # Tensão fase-neutro
        else:
            tensao = 380.0  # Tensão entre fases
    elif sistema_tensao == '127/220 V':
        if numero_fases == 1:
            tensao = 127.0  # Tensão fase-neutro
        else:
            tensao = 220.0  # Tensão entre fases
    else:
        tensao = 0.0

    return tensao, numero_fases


def obter_parametros_usuario(parede):
    """Obtém todos os parâmetros do usuário para a inserção das tomadas."""

    class InputForm(Form):
        def __init__(self):
            self.Text = 'Parâmetros de Inserção'
            self.Width = 400
            self.Height = 700  # Ajustar altura
            self.StartPosition = FormStartPosition.CenterScreen
            self.AutoSize = True  # Ajustar automaticamente
            self.AutoSizeMode = AutoSizeMode.GrowAndShrink
            self.results = None

            # Labels and Controls
            y = 10
            dy = 30

            # Altura das tomadas
            self.label_altura = Label()
            self.label_altura.Text = 'Insira a altura das tomadas em metros:'
            self.label_altura.Location = Point(10, y)
            self.label_altura.Width = 350
            self.Controls.Add(self.label_altura)

            y += dy
            self.textbox_altura = TextBox()
            self.textbox_altura.Text = '1.10'
            self.textbox_altura.Location = Point(10, y)
            self.Controls.Add(self.textbox_altura)

            y += dy
            # Número de tomadas
            self.label_num_tomadas = Label()
            self.label_num_tomadas.Text = 'Insira o número de tomadas a serem inseridas:'
            self.label_num_tomadas.Location = Point(10, y)
            self.label_num_tomadas.Width = 350
            self.Controls.Add(self.label_num_tomadas)

            y += dy
            self.textbox_num_tomadas = TextBox()
            self.textbox_num_tomadas.Text = '1'
            self.textbox_num_tomadas.Location = Point(10, y)
            self.Controls.Add(self.textbox_num_tomadas)

            y += dy
            # Intervalo
            self.label_intervalo = Label()
            self.label_intervalo.Text = 'Insira o comprimento do intervalo em metros (deixe em branco para usar o comprimento total da parede):'
            self.label_intervalo.Location = Point(10, y)
            self.label_intervalo.Size = Size(350, 30)
            self.Controls.Add(self.label_intervalo)

            y += dy + 10
            self.textbox_intervalo = TextBox()
            self.textbox_intervalo.Text = ''
            self.textbox_intervalo.Location = Point(10, y)
            self.Controls.Add(self.textbox_intervalo)

            y += dy
            # Face da parede
            self.label_face = Label()
            self.label_face.Text = 'Selecione a face da parede:'
            self.label_face.Location = Point(10, y)
            self.label_face.Width = 350
            self.Controls.Add(self.label_face)

            y += dy
            self.combobox_face = ComboBox()
            self.combobox_face.Items.Add('Frontal')
            self.combobox_face.Items.Add('Traseira')
            self.combobox_face.SelectedIndex = 0
            self.combobox_face.Location = Point(10, y)
            self.Controls.Add(self.combobox_face)

            y += dy + 10
            # Parâmetros elétricos
            self.label_elet = Label()
            self.label_elet.Text = 'Parâmetros Elétricos'
            self.label_elet.Location = Point(10, y)
            self.label_elet.Width = 350
            self.Controls.Add(self.label_elet)

            y += dy
            # Potência Aparente
            self.label_potencia = Label()
            self.label_potencia.Text = 'Potência Aparente (VA):'
            self.label_potencia.Location = Point(10, y)
            self.label_potencia.Width = 350
            self.Controls.Add(self.label_potencia)

            y += dy
            self.textbox_potencia = TextBox()
            self.textbox_potencia.Text = '1000'
            self.textbox_potencia.Location = Point(10, y)
            self.Controls.Add(self.textbox_potencia)

            y += dy
            # Fator de Potência
            self.label_fp = Label()
            self.label_fp.Text = 'Fator de Potência (cos φ):'
            self.label_fp.Location = Point(10, y)
            self.label_fp.Width = 350
            self.Controls.Add(self.label_fp)

            y += dy
            self.textbox_fp = TextBox()
            self.textbox_fp.Text = '0.8'
            self.textbox_fp.Location = Point(10, y)
            self.Controls.Add(self.textbox_fp)

            y += dy + 10
            # Seleção de Sistema de Tensão e Número de Fases
            self.button_voltage = Button()
            self.button_voltage.Text = 'Configurar Tensão e Fases'
            self.button_voltage.Location = Point(10, y)
            self.button_voltage.Click += self.configure_voltage_phases
            self.Controls.Add(self.button_voltage)

            y += dy + 10
            # OK and Cancel buttons
            self.button_ok = Button()
            self.button_ok.Text = 'OK'
            self.button_ok.Location = Point(220, y)
            self.button_ok.Click += self.ok_clicked
            self.Controls.Add(self.button_ok)

            self.button_cancel = Button()
            self.button_cancel.Text = 'Cancelar'
            self.button_cancel.Location = Point(300, y)
            self.button_cancel.Click += self.cancel_clicked
            self.Controls.Add(self.button_cancel)

            # Variáveis para armazenar tensão e fases
            self.voltage = 0.0
            self.number_of_phases = 1

        def configure_voltage_phases(self, sender, event):
            # Chama a função para selecionar tensão e fases
            voltage, phases = get_tensao_e_fases()
            self.voltage = voltage
            self.number_of_phases = phases
            MessageBox.Show(
                "Tensão configurada para {} V com {} fases.".format(self.voltage, self.number_of_phases),
                "Configuração Concluída",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information
            )

        def ok_clicked(self, sender, event):
            self.results = {
                'altura': self.textbox_altura.Text,
                'numero_tomadas': self.textbox_num_tomadas.Text,
                'intervalo': self.textbox_intervalo.Text,
                'face': self.combobox_face.SelectedItem,
                'potencia_aparente': self.textbox_potencia.Text,
                'fator_potencia': self.textbox_fp.Text,
                'tensao': self.voltage,
                'fases': self.number_of_phases,
            }
            self.DialogResult = DialogResult.OK
            self.Close()

        def cancel_clicked(self, sender, event):
            self.DialogResult = DialogResult.Cancel
            self.Close()

    form = InputForm()
    result = form.ShowDialog()
    if result != DialogResult.OK:
        forms.alert("Entrada cancelada pelo usuário.", exitscript=True)

    # Processar os valores inseridos
    try:
        altura_metros = float(form.results['altura'].replace(',', '.'))
    except ValueError:
        forms.alert("Entrada inválida para a altura. Usando 1.10 metros.")
        altura_metros = 1.10

    try:
        numero_tomadas = int(form.results['numero_tomadas'])
        if numero_tomadas < 1:
            raise ValueError
    except ValueError:
        forms.alert("Entrada inválida para o número de tomadas. Usando 1 tomada.")
        numero_tomadas = 1

    intervalo_input = form.results['intervalo']
    if intervalo_input.strip() == '':
        loc_curve = parede.Location
        curva = loc_curve.Curve
        comprimento_parede = curva.Length  # Comprimento em pés
        intervalo_metros = comprimento_parede / 3.28084  # Converter para metros
    else:
        try:
            intervalo_metros = float(intervalo_input.replace(',', '.'))
            if intervalo_metros <= 0:
                raise ValueError
        except ValueError:
            forms.alert("Entrada inválida para o intervalo. Usando comprimento total da parede.")
            loc_curve = parede.Location
            curva = loc_curve.Curve
            comprimento_parede = curva.Length  # Comprimento em pés
            intervalo_metros = comprimento_parede / 3.28084  # Converter para metros

    face_selecionada = form.results['face']

    # Parâmetros elétricos
    try:
        potencia_aparente = float(form.results['potencia_aparente'].replace(',', '.'))
    except ValueError:
        forms.alert("Entrada inválida para Potência Aparente. Usando 1000 VA.")
        potencia_aparente = 1000.0

    try:
        fator_potencia = float(form.results['fator_potencia'].replace(',', '.'))
        if not (0 < fator_potencia <= 1):
            raise ValueError
    except ValueError:
        forms.alert("Entrada inválida para Fator de Potência. Usando 0.8.")
        fator_potencia = 0.8

    # Tensão e número de fases já configurados no formulário
    tensao = form.results['tensao']
    numero_fases = form.results['fases']

    parametros_elet = (potencia_aparente, fator_potencia, tensao, numero_fases)

    return (
        altura_metros,
        numero_tomadas,
        intervalo_metros,
        face_selecionada,
        parametros_elet,
    )


def calcular_pontos_insercao(
        parede, altura_metros, numero_tomadas, intervalo_metros, face_selecionada
):
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
        ponto_insercao = XYZ(
            ponto_na_parede.X, ponto_na_parede.Y, ponto_na_parede.Z + altura_pes
        )
        pontos_insercao.append(ponto_insercao)

    # Obter a espessura da parede
    espessura_parede = parede.WallType.Width  # Em pés

    if face_selecionada:
        # Calcular o vetor normal da parede
        vetor_normal = XYZ(-direcao_parede.Y, direcao_parede.X, 0).Normalize()
        # Calcular o deslocamento (metade da espessura da parede)
        deslocamento = espessura_parede / 2
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
                plano_normal = XYZ.BasisZ
                plano = Plane.CreateByNormalAndOrigin(plano_normal, p1)
                sketch_plane = SketchPlane.Create(doc, plano)
            except Exception:
                # Se o plano não puder ser criado, usar um plano padrão (XY)
                plano = Plane.CreateByNormalAndOrigin(XYZ.BasisZ, p1)
                sketch_plane = SketchPlane.Create(doc, plano)
            # Criar a linha de pré-visualização
            linha_preview = Line.CreateBound(p1, p2)
            # Criar um ModelCurve para a pré-visualização
            try:
                model_line = doc.Create.NewModelCurve(linha_preview, sketch_plane)
                # Aplicar uma cor diferente para destacar a pré-visualização (opcional)
                ogs = OverrideGraphicSettings()
                ogs.SetProjectionLineColor(Color(255, 0, 0))  # Vermelho
                doc.ActiveView.SetElementOverrides(model_line.Id, ogs)
                preview_ids.append(model_line.Id)
            except Exception:
                pass
    return preview_ids


def remover_preview(preview_ids):
    """Remove os elementos de pré-visualização."""
    with revit.Transaction("Remover Preview"):
        for elem_id in preview_ids:
            try:
                doc.Delete(elem_id)
            except Exception:
                pass


def inserir_tomadas(parede, tomada_selecionada, pontos_insercao, face_selecionada, parametros_elet):
    """Insere as tomadas nas posições calculadas."""
    potencia_aparente, fator_potencia, tensao, numero_fases = parametros_elet

    # Lista para armazenar as instâncias de tomadas inseridas
    tomadas_inseridas = []

    with revit.Transaction("Inserir Tomadas"):
        for ponto_insercao in pontos_insercao:
            try:
                # Inserir a tomada usando a parede como host
                tomada_instancia = doc.Create.NewFamilyInstance(
                    ponto_insercao,
                    tomada_selecionada,
                    parede,
                    StructuralType.NonStructural,
                )

                # Ajustar a orientação da tomada para ficar alinhada com a parede
                # Obter a direção da parede
                loc_curve = parede.Location
                curva = loc_curve.Curve
                direcao_parede = (
                        curva.GetEndPoint(1) - curva.GetEndPoint(0)
                ).Normalize()
                # Calcular o ângulo entre a direção da parede e o eixo X
                angulo = XYZ.BasisX.AngleTo(direcao_parede)
                # Determinar o sentido da rotação
                cross = XYZ.BasisX.CrossProduct(direcao_parede)
                if cross.Z < 0:
                    angulo = -angulo

                # Rotacionar 180 graus se a face for 'Traseira'
                if face_selecionada == 'Traseira':
                    angulo += math.pi  # Adicionar 180 graus em radianos

                # Criar o eixo de rotação (linha vertical passando pelo ponto de inserção)
                eixo_rotacao = Line.CreateBound(
                    ponto_insercao, ponto_insercao + XYZ.BasisZ
                )
                # Rotacionar a tomada
                ElementTransformUtils.RotateElement(
                    doc, tomada_instancia.Id, eixo_rotacao, angulo
                )

                # Ajustar a altura usando o parâmetro 'Elevação do Ponto'
                sucesso_altura = False
                param_elevacao = tomada_instancia.LookupParameter('Elevação do Ponto')
                if param_elevacao and not param_elevacao.IsReadOnly:
                    param_elevacao.Set(ponto_insercao.Z)
                    sucesso_altura = True

                if not sucesso_altura:
                    # Ajustar diretamente a posição Z
                    location = tomada_instancia.Location
                    if isinstance(location, LocationCurve):
                        pass
                    elif isinstance(location, LocationPoint):
                        point = location.Point
                        new_point = XYZ(point.X, point.Y, ponto_insercao.Z)
                        location.Point = new_point

                # Definir os parâmetros elétricos na instância da família
                # Potência Aparente (VA)
                parametro_S = tomada_instancia.LookupParameter('Potência Aparente (VA)')
                if parametro_S and parametro_S.StorageType == StorageType.Double:
                    parametro_S.Set(potencia_aparente)
                # Fator de Potência
                parametro_cos_phi = tomada_instancia.LookupParameter('Fator de Potência')
                if parametro_cos_phi and parametro_cos_phi.StorageType == StorageType.Double:
                    parametro_cos_phi.Set(fator_potencia)

                tomadas_inseridas.append(tomada_instancia)

            except Exception as e:
                print("Erro ao inserir tomada: {}".format(e))
                pass

    return tomadas_inseridas


def obter_paineis_eletricos():
    """Obtém todos os painéis elétricos disponíveis no projeto."""
    collector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_ElectricalEquipment).OfClass(
        FamilyInstance)
    paineis = {}
    for painel in collector:
        try:
            # Obter o nome do painel
            painel_nome = painel.Name
            paineis[painel_nome] = painel
        except Exception:
            pass
    return paineis


def criar_circuito_eletrico(tomadas_inseridas, tensao, numero_fases, parametros_elet):
    """Cria um circuito elétrico com as tomadas inseridas e ajusta os parâmetros."""
    potencia_aparente, fator_potencia, _, _ = parametros_elet

    try:
        with revit.Transaction("Criar Circuito Elétrico"):
            # Obter os ElementIds das tomadas
            tomadas_ids = [t.Id for t in tomadas_inseridas]
            # Criar uma lista de ElementIds
            elementos_ids = List[ElementId](tomadas_ids)
            # Definir o tipo de sistema elétrico (PowerCircuit)
            sistema_tipo = ElectricalSystemType.PowerCircuit
            # Criar o circuito elétrico
            circuito = ElectricalSystem.Create(doc, elementos_ids, sistema_tipo)
            if circuito:
                forms.alert("Circuito elétrico criado com sucesso.", exitscript=False)
                # Opcional: Solicitar ao usuário para selecionar um painel
                # e atribuir o circuito ao painel
                atribuir_painel = MessageBox.Show(
                    "Deseja atribuir o circuito a um painel?",
                    "Atribuir Painel",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question,
                )
                if atribuir_painel == DialogResult.Yes:
                    # Obter os painéis disponíveis
                    paineis = obter_paineis_eletricos()
                    if paineis:
                        painel_selecionado_nome = forms.SelectFromList.show(
                            sorted(paineis.keys()),
                            title='Selecione um Painel',
                            button_name='Selecionar',
                            multiselect=False,
                        )
                        if painel_selecionado_nome:
                            painel = paineis[painel_selecionado_nome]
                            circuito.SelectPanel(painel)
                            forms.alert("Circuito atribuído ao painel selecionado.", exitscript=False)

                            # Definir a tensão e número de fases do circuito com base nos valores fornecidos
                            voltage_param = circuito.get_Parameter(BuiltInParameter.RBS_ELEC_VOLTAGE)
                            if voltage_param and voltage_param.StorageType == StorageType.Double:
                                voltage_param.Set(tensao)

                            number_of_poles_param = circuito.get_Parameter(BuiltInParameter.RBS_ELEC_NUMBER_OF_POLES)
                            if number_of_poles_param and number_of_poles_param.StorageType == StorageType.Integer:
                                number_of_poles_param.Set(numero_fases)

                            # Definir a Potência Aparente e o Fator de Potência no circuito
                            apparent_load_param = circuito.get_Parameter(BuiltInParameter.RBS_ELEC_APPARENT_LOAD)
                            if apparent_load_param and apparent_load_param.StorageType == StorageType.Double:
                                apparent_load_param.Set(potencia_aparente)

                            power_factor_param = circuito.get_Parameter(BuiltInParameter.RBS_ELEC_POWER_FACTOR)
                            if power_factor_param and power_factor_param.StorageType == StorageType.Double:
                                power_factor_param.Set(fator_potencia)

                            # Regenerar o documento para atualizar parâmetros calculados
                            doc.Regenerate()

                            # Agora, Revit deve atualizar automaticamente o parâmetro "Dados elétricos"

                        else:
                            forms.alert("Nenhum painel selecionado. Circuito não atribuído.", exitscript=False)
                    else:
                        forms.alert("Nenhum painel elétrico encontrado no projeto.", exitscript=False)
                else:
                    forms.alert("Circuito não será atribuído a nenhum painel.", exitscript=False)
            else:
                forms.alert("Não foi possível criar o circuito elétrico.", exitscript=False)
    except Exception as e:
        tb = traceback.format_exc()
        forms.alert("Erro ao criar circuito elétrico:\n{}".format(tb))


def inserir_tomadas_na_parede():
    """Função principal para inserir tomadas na parede com pré-visualização."""
    try:
        # Selecionar a família de tomada
        tomada_selecionada = selecionar_familia_tomada()
        # Selecionar a parede
        parede = selecionar_parede()
        # Obter os parâmetros do usuário
        parametros = obter_parametros_usuario(parede)
        if parametros is None:
            forms.alert("Falha ao obter parâmetros do usuário.", exitscript=True)
        (
            altura_metros,
            numero_tomadas,
            intervalo_metros,
            face_selecionada,
            parametros_elet,
        ) = parametros
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
            MessageBoxIcon.Question,
        )
        # Remover pré-visualização
        remover_preview(preview_ids)
        if resultado == DialogResult.Yes:
            # Inserir as tomadas
            tomadas_inseridas = inserir_tomadas(
                parede,
                tomada_selecionada,
                pontos_insercao,
                face_selecionada,
                parametros_elet,
            )

            # Perguntar ao usuário se deseja criar um circuito
            if tomadas_inseridas:
                criar_circuito = MessageBox.Show(
                    "Deseja criar um circuito elétrico para as tomadas inseridas?",
                    "Criar Circuito",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question,
                )

                if criar_circuito == DialogResult.Yes:
                    # Criar o circuito elétrico
                    criar_circuito_eletrico(tomadas_inseridas, parametros_elet[2], parametros_elet[3], parametros_elet)
                else:
                    forms.alert("Circuito não será criado.", exitscript=False)
            else:
                forms.alert("Nenhuma tomada foi inserida.", exitscript=False)
        else:
            forms.alert("Inserção cancelada pelo usuário.")
    except Exception as e:
        tb = traceback.format_exc()
        forms.alert("Ocorreu um erro:\n{}".format(tb))


# Executar o script
if __name__ == "__main__":
    inserir_tomadas_na_parede()
