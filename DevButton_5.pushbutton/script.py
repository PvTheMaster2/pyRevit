# -*- coding: utf-8 -*-
__title__ = "Inserir Tomada com Parâmetros Elétricos"
__doc__ = """Versão: 2.2
_____________________________________________________________________
Descrição:
Este script insere uma tomada elétrica na parede selecionada,
permitindo escolher a altura, posição horizontal e face (frontal/traseira).
Adicionalmente, coleta parâmetros elétricos como potência aparente,
fator de potência, tensão e número de fases para calcular a potência ativa.
_____________________________________________________________________
Como usar:
- Clique no botão e siga as instruções.
_____________________________________________________________________
Autor: Seu Nome"""

# Importações necessárias
import clr
import traceback  # Para capturar o traceback completo

# Importar apenas as classes necessárias de Autodesk.Revit.DB
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
    LocationCurve  # Importação correta de LocationCurve
)

# Importar StructuralType da sub-biblioteca Structure
from Autodesk.Revit.DB.Structure import StructuralType

# Importações adicionais do Revit
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Selection import ObjectType
from Autodesk.Revit.Exceptions import InvalidOperationException

# Importações do pyRevit
from pyrevit import revit, forms, script

# Importar System.Windows.Forms para caixas de diálogo personalizadas
clr.AddReference('System.Windows.Forms')
from System.Windows.Forms import DialogResult, MessageBox, MessageBoxButtons, \
    MessageBoxIcon  # Adicionado MessageBoxIcon

# Variáveis do documento
doc = revit.doc  # Documento ativo do Revit
uidoc = revit.uidoc  # Documento UI ativo


# Funções auxiliares

def selecionar_familia_tomada():
    """Permite que o usuário selecione uma família de tomada elétrica."""
    output = script.get_output()
    output.print_md("### Iniciando seleção da família de tomada.")

    # Coletar todos os símbolos de família da categoria "Dispositivos elétricos"
    collector = FilteredElementCollector(doc) \
        .OfClass(FamilySymbol) \
        .OfCategory(BuiltInCategory.OST_ElectricalFixtures)

    tomadas = []
    for symbol in collector:
        try:
            # Obter o nome da família usando BuiltInParameter
            family_name_param = symbol.get_Parameter(BuiltInParameter.ALL_MODEL_FAMILY_NAME)
            family_name = family_name_param.AsString() if family_name_param and family_name_param.HasValue else "Sem Família"

            # Obter o nome do símbolo usando BuiltInParameter
            symbol_name_param = symbol.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME)
            symbol_name = symbol_name_param.AsString() if symbol_name_param and symbol_name_param.HasValue else "Sem Nome"

            # Filtrar por famílias que contenham "Tomada" no nome
            if "Tomada" in family_name or "Tomada" in symbol_name:
                tomadas.append(symbol)
        except Exception as e:
            output.print_md("**Erro ao processar símbolo de família:** {}".format(e))
            pass

    if not tomadas:
        forms.alert("Nenhuma família de tomadas encontrada no projeto.", exitscript=True)

    # Criar um dicionário de opções
    tomadas_dict = {}
    for tomada in tomadas:
        try:
            # Obter o nome da família
            family_name_param = tomada.get_Parameter(BuiltInParameter.ALL_MODEL_FAMILY_NAME)
            family_name = family_name_param.AsString() if family_name_param and family_name_param.HasValue else "Sem Família"

            # Obter o nome do símbolo
            symbol_name_param = tomada.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME)
            symbol_name = symbol_name_param.AsString() if symbol_name_param and symbol_name_param.HasValue else "Sem Nome"

            display_name = "{} : {}".format(family_name, symbol_name)
            tomadas_dict[display_name] = tomada
        except Exception as e:
            output.print_md("**Erro ao processar tomada:** {}".format(e))
            pass

    if not tomadas_dict:
        forms.alert("Nenhuma família de tomadas válida encontrada.", exitscript=True)

    # Ordenar os nomes para exibição
    tomadas_nomes_ordenados = sorted(tomadas_dict.keys())

    # Permitir que o usuário selecione uma tomada
    output.print_md("### Exibindo lista de tomadas para seleção.")
    tomada_selecionada_nome = forms.SelectFromList.show(
        tomadas_nomes_ordenados,
        title='Selecione uma Tomada',
        button_name='Selecionar',
        multiselect=False
    )

    if not tomada_selecionada_nome:
        forms.alert("Nenhuma tomada selecionada.", exitscript=True)

    tomada_selecionada = tomadas_dict[tomada_selecionada_nome]

    # Adicionar logs para depuração
    output.print_md("### Tipo de tomada_selecionada: {}".format(type(tomada_selecionada)))
    output.print_md("### Atributos disponíveis: {}".format(dir(tomada_selecionada)))

    # Ativar o símbolo da família, se necessário
    if not tomada_selecionada.IsActive:
        output.print_md("### Ativando o símbolo da família da tomada.")
        with Transaction(doc, "Ativar Família") as t:
            t.Start()
            tomada_selecionada.Activate()
            doc.Regenerate()
            t.Commit()

    # Verificar se o parâmetro 'ALL_MODEL_TYPE_NAME' está disponível para obter o nome
    try:
        symbol_name_param = tomada_selecionada.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME)
        symbol_name = symbol_name_param.AsString() if symbol_name_param and symbol_name_param.HasValue else "Sem Nome"
        output.print_md("### Família de tomada selecionada: {}".format(symbol_name))
    except Exception as e:
        output.print_md("### O objeto tomada_selecionada não possui o atributo 'Name'. Erro: {}".format(e))

    return tomada_selecionada


def selecionar_parede():
    """Permite que o usuário selecione uma parede."""
    output = script.get_output()
    output.print_md("### Iniciando seleção da parede.")
    sel = uidoc.Selection
    try:
        referencia = sel.PickObject(ObjectType.Element, 'Selecione a parede onde a tomada será inserida.')
        parede = doc.GetElement(referencia.ElementId)
        output.print_md("### Parede selecionada: {}".format(parede.Name))
    except InvalidOperationException:
        forms.alert("Nenhuma parede selecionada.", exitscript=True)

    if not isinstance(parede, Wall):
        forms.alert("O elemento selecionado não é uma parede.", exitscript=True)

    return parede


def obter_parametros_elec():
    """Obtém os parâmetros elétricos do usuário utilizando ask_for_string."""
    output = script.get_output()
    output.print_md("### Iniciando coleta de parâmetros elétricos.")

    # Definir os campos e valores padrão
    campos = [
        ("Potência Aparente (VA):", "1000"),
        ("Fator de Potência (cos φ):", "0.8"),
        ("Tensão (V):", "127"),
        ("Número de Fases (1 para monofásico, 3 para trifásico):", "1")
    ]

    # Usar uma caixa de diálogo para coletar os valores individualmente
    dialog_inputs = []
    for prompt, default in campos:
        try:
            input_value = forms.ask_for_string(
                prompt=prompt,
                title='Parâmetros Elétricos da Tomada',
                default=default
            )
            if input_value is None:
                forms.alert("Entrada cancelada pelo usuário.", exitscript=True)
            dialog_inputs.append(input_value)
        except Exception as e:
            forms.alert("Erro ao coletar entrada: {}".format(e))
            output.print_md("### Erro ao coletar entrada: {}".format(e))
            return None

    # Processar as entradas
    try:
        potencia_aparente_input = dialog_inputs[0]
        potencia_aparente = float(potencia_aparente_input.replace(',', '.'))
        output.print_md("### Potência Aparente (S): {} VA".format(potencia_aparente))
    except (ValueError, IndexError):
        forms.alert("Entrada inválida para Potência Aparente. Usando 1000 VA.")
        potencia_aparente = 1000.0
        output.print_md("### Potência Aparente (S): 1000 VA (Padrão)")

    try:
        fator_potencia_input = dialog_inputs[1]
        fator_potencia = float(fator_potencia_input.replace(',', '.'))
        if not (0 < fator_potencia <= 1):
            raise ValueError
        output.print_md("### Fator de Potência (cos φ): {}".format(fator_potencia))
    except (ValueError, IndexError):
        forms.alert("Entrada inválida para Fator de Potência. Usando 0.8.")
        fator_potencia = 0.8
        output.print_md("### Fator de Potência (cos φ): 0.8 (Padrão)")

    try:
        tensao_input = dialog_inputs[2]
        tensao = float(tensao_input.replace(',', '.'))
        output.print_md("### Tensão (V): {} V".format(tensao))
    except (ValueError, IndexError):
        forms.alert("Entrada inválida para Tensão. Usando 127 V.")
        tensao = 127.0
        output.print_md("### Tensão (V): 127 V (Padrão)")

    try:
        numero_fases_input = dialog_inputs[3]
        numero_fases = int(numero_fases_input)
        if numero_fases not in [1, 3]:
            raise ValueError
        output.print_md("### Número de Fases: {}".format(numero_fases))
    except (ValueError, IndexError):
        forms.alert("Entrada inválida para Número de Fases. Usando 1 (monofásico).")
        numero_fases = 1
        output.print_md("### Número de Fases: 1 (Monofásico) (Padrão)")

    # Calcular a potência ativa (P)
    potencia_ativa = potencia_aparente * fator_potencia  # P = S * cos φ
    output.print_md("### Potência Ativa (P): {} W".format(potencia_ativa))

    # Mostrar os cálculos ao usuário
    forms.alert(
        "Potência Ativa (P): {} W".format(potencia_ativa),
        "Cálculo da Potência Ativa"
    )

    return potencia_aparente, fator_potencia, tensao, numero_fases, potencia_ativa


def obter_ponto_insercao(parede):
    """Obtém o ponto de inserção da tomada na parede."""
    output = script.get_output()
    output.print_md("### Iniciando obtenção do ponto de inserção.")

    # Obter a altura desejada do usuário
    altura_metros_input = forms.ask_for_string(
        prompt="Insira a altura da tomada em metros:",
        title="Altura da Tomada",
        default="1.10"
    )
    try:
        altura_metros = float(altura_metros_input.replace(',', '.'))
        output.print_md("### Altura da Tomada: {} metros".format(altura_metros))
    except ValueError:
        forms.alert("Entrada inválida. Usando altura padrão de 1.10 metros.")
        altura_metros = 1.10
        output.print_md("### Altura da Tomada: 1.10 metros (Padrão)")

    # Converter metros para pés
    altura_pes = altura_metros * 3.28084
    output.print_md("### Altura da Tomada em Pés: {:.2f} pés".format(altura_pes))

    # Obter a localização da parede
    loc_curve = parede.Location
    if not isinstance(loc_curve, LocationCurve):
        forms.alert("Não foi possível obter a localização da parede.", exitscript=True)

    # Obter a curva (linha) central da parede
    curva = loc_curve.Curve
    output.print_md("### Comprimento da Parede: {:.2f} pés".format(curva.Length))

    # Solicitar ao usuário a distância ao longo da parede
    distancia_metros_input = forms.ask_for_string(
        prompt="Insira a distância ao longo da parede em metros (0 para início, deixe em branco para meio):",
        title="Posição Horizontal",
        default=""
    )
    if distancia_metros_input.strip() == "":
        # Usar o ponto médio se o usuário não inserir nada
        parametro_normalizado = 0.5
        output.print_md("### Posição Horizontal: Meio da Parede")
    else:
        try:
            distancia_metros = float(distancia_metros_input.replace(',', '.'))
            comprimento_parede = curva.Length  # Comprimento em pés
            comprimento_metros = comprimento_parede / 3.28084
            parametro_normalizado = distancia_metros / comprimento_metros
            parametro_normalizado = max(0.0, min(1.0, parametro_normalizado))  # Garantir que esteja entre 0 e 1
            output.print_md("### Posição Horizontal: {:.2f} metros (Normalizado: {:.2f})".format(distancia_metros,
                                                                                                 parametro_normalizado))
        except ValueError:
            forms.alert("Entrada inválida. Usando o ponto médio da parede.")
            parametro_normalizado = 0.5
            output.print_md("### Posição Horizontal: Meio da Parede (Padrão)")

    # Obter o ponto ao longo da curva
    ponto_na_curva = curva.Evaluate(parametro_normalizado, True)
    output.print_md(
        "### Ponto na Curva: ({:.2f}, {:.2f}, {:.2f})".format(ponto_na_curva.X, ponto_na_curva.Y, ponto_na_curva.Z))

    # Ajustar o ponto para a altura desejada
    ponto_insercao = XYZ(ponto_na_curva.X, ponto_na_curva.Y, ponto_na_curva.Z + altura_pes)
    output.print_md(
        "### Ponto de Inserção Ajustado: ({:.2f}, {:.2f}, {:.2f})".format(ponto_insercao.X, ponto_insercao.Y,
                                                                          ponto_insercao.Z))

    # Obter a espessura da parede
    espessura_parede = parede.WallType.Width  # Em pés
    output.print_md("### Espessura da Parede: {:.2f} pés".format(espessura_parede))

    # Perguntar ao usuário em qual face deseja inserir a tomada
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
        output.print_md("### Face da Parede: Meio (Padrão)")
    else:
        # Calcular o vetor normal da parede
        direcao_parede = (curva.GetEndPoint(1) - curva.GetEndPoint(0)).Normalize()
        vetor_normal = XYZ(-direcao_parede.Y, direcao_parede.X, 0).Normalize()
        output.print_md("### Direção da Parede: ({:.2f}, {:.2f}, {:.2f})".format(direcao_parede.X, direcao_parede.Y,
                                                                                 direcao_parede.Z))
        output.print_md(
            "### Vetor Normal: ({:.2f}, {:.2f}, {:.2f})".format(vetor_normal.X, vetor_normal.Y, vetor_normal.Z))

        # Calcular o deslocamento (metade da espessura da parede)
        deslocamento = (espessura_parede / 2)
        output.print_md("### Deslocamento: {:.2f} pés".format(deslocamento))

        if face_selecionada == 'Frontal':
            # Deslocar na direção do vetor normal
            ponto_insercao += vetor_normal * deslocamento
            output.print_md("### Face Selecionada: Frontal (Deslocando na direção do vetor normal)")
        elif face_selecionada == 'Traseira':
            # Deslocar na direção oposta ao vetor normal
            ponto_insercao -= vetor_normal * deslocamento
            output.print_md("### Face Selecionada: Traseira (Deslocando na direção oposta ao vetor normal)")

        output.print_md(
            "### Ponto de Inserção Final: ({:.2f}, {:.2f}, {:.2f})".format(ponto_insercao.X, ponto_insercao.Y,
                                                                           ponto_insercao.Z))

    return ponto_insercao


def inserir_tomada(parede, tomada_selecionada, ponto_insercao, parametros_elec):
    """Insere a tomada nas posições calculadas e armazena os parâmetros elétricos."""
    if parametros_elec is None:
        forms.alert("Parâmetros elétricos não foram coletados corretamente.", exitscript=True)

    potencia_aparente, fator_potencia, tensao, numero_fases, potencia_ativa = parametros_elec

    # Inserir a tomada usando a parede como host
    try:
        with Transaction(doc, "Inserir Tomada") as t:
            t.Start()
            tomada_instancia = doc.Create.NewFamilyInstance(
                ponto_insercao,
                tomada_selecionada,
                parede,
                StructuralType.NonStructural  # Uso direto de StructuralType
            )
            output = script.get_output()
            output.print_md("### Tomada inserida.")

            # Ajustar a orientação da tomada para ficar alinhada com a parede
            # Obter a direção da parede
            loc_curve = parede.Location
            curva = loc_curve.Curve
            direcao_parede = (curva.GetEndPoint(1) - curva.GetEndPoint(0)).Normalize()
            output.print_md("### Direção da Parede: ({:.2f}, {:.2f}, {:.2f})".format(direcao_parede.X, direcao_parede.Y,
                                                                                     direcao_parede.Z))

            # Calcular o ângulo entre a direção da parede e o eixo X
            angulo = XYZ.BasisX.AngleTo(direcao_parede)
            output.print_md("### Ângulo Calculado: {:.2f} radianos".format(angulo))

            # Determinar o sentido da rotação
            cross = XYZ.BasisX.CrossProduct(direcao_parede)
            if cross.Z < 0:
                angulo = -angulo
                output.print_md("### Sentido da Rotação: Anti-horário")
            else:
                output.print_md("### Sentido da Rotação: Horário")

            # Criar o eixo de rotação (linha vertical passando pelo ponto de inserção)
            eixo_rotacao = Line.CreateBound(ponto_insercao, ponto_insercao + XYZ.BasisZ)
            output.print_md("### Eixo de Rotação Criado.")

            # Rotacionar a tomada dentro da transação
            ElementTransformUtils.RotateElement(
                doc,
                tomada_instancia.Id,
                eixo_rotacao,
                angulo
            )
            output.print_md("### Tomada rotacionada.")

            # **Armazenar os parâmetros elétricos na instância da família (se aplicável)**
            # Verifique se a família possui os parâmetros correspondentes antes de tentar defini-los
            try:
                # Potência Aparente (VA)
                parametro_S = tomada_instancia.LookupParameter("Potência Aparente (VA)")
                if parametro_S and parametro_S.StorageType == StorageType.Double:
                    parametro_S.Set(potencia_aparente)
                    output.print_md("### Parâmetro 'Potência Aparente (VA)' definido: {}".format(potencia_aparente))
                else:
                    output.print_md("### Parâmetro 'Potência Aparente (VA)' não encontrado ou tipo incorreto.")

                # Fator de Potência
                parametro_cos_phi = tomada_instancia.LookupParameter("Fator de Potência")
                if parametro_cos_phi and parametro_cos_phi.StorageType == StorageType.Double:
                    parametro_cos_phi.Set(fator_potencia)
                    output.print_md("### Parâmetro 'Fator de Potência' definido: {}".format(fator_potencia))
                else:
                    output.print_md("### Parâmetro 'Fator de Potência' não encontrado ou tipo incorreto.")

                # Tensão (V)
                parametro_V = tomada_instancia.LookupParameter("Tensão (V)")
                if parametro_V and parametro_V.StorageType == StorageType.Double:
                    parametro_V.Set(tensao)
                    output.print_md("### Parâmetro 'Tensão (V)' definido: {}".format(tensao))
                else:
                    output.print_md("### Parâmetro 'Tensão (V)' não encontrado ou tipo incorreto.")

                # N° de Fases
                parametro_fases = tomada_instancia.LookupParameter("N° de Fases")
                if parametro_fases and parametro_fases.StorageType == StorageType.Integer:
                    parametro_fases.Set(numero_fases)
                    output.print_md("### Parâmetro 'N° de Fases' definido: {}".format(numero_fases))
                else:
                    output.print_md("### Parâmetro 'N° de Fases' não encontrado ou tipo incorreto.")

                # Potência Ativa (W)
                parametro_P = tomada_instancia.LookupParameter("Potência Ativa (W)")
                if parametro_P and parametro_P.StorageType == StorageType.Double:
                    output.print_md("### Parâmetro 'Potência Ativa (W)' é read-only ou não encontrado.")
            except Exception as e:
                # Se ocorrer um erro ao definir os parâmetros, exibir uma mensagem e continuar
                output.print_md("### Erro ao definir parâmetros elétricos: {}".format(e))
                forms.alert("Erro ao definir parâmetros elétricos: {}".format(e))

            # Obter o MEPModel da instância para conectar elementos
            try:
                mep_model = tomada_instancia.MEPModel
                if mep_model:
                    output.print_md("### MEPModel acessado com sucesso.")
                    # Aqui você pode continuar com a criação do circuito, etc.
                else:
                    output.print_md("### A instância de tomada não possui MEPModel.")
            except Exception as e:
                output.print_md("### Erro ao acessar MEPModel: {}".format(e))
                forms.alert("Erro ao acessar MEPModel: {}".format(e))

            # Commit da transação após todas as modificações
            t.Commit()

            # Garantir que o documento seja regenerado
            doc.Regenerate()

    except Exception as e:
        tb = traceback.format_exc()
        forms.alert("Erro ao inserir a tomada:\n{}".format(tb))
        script.get_output().print_md("### Erro ao inserir a tomada:\n{}".format(tb))
        return

    # Exibir uma caixa de diálogo de confirmação ao usuário
    resultado = MessageBox.Show(
        "Tomada inserida com sucesso! Deseja inserir outra tomada?",
        "Confirmação",
        MessageBoxButtons.YesNo,
        MessageBoxIcon.Question  # Agora, MessageBoxIcon está definido
    )

    if resultado == DialogResult.Yes:
        # Reiniciar o processo
        main()
    else:
        script.get_output().print_md("### Processo finalizado pelo usuário.")
        forms.alert("Processo finalizado.", exitscript=True)


def main():
    try:
        script.get_output().print_md("### Iniciando o script de inserção de tomada.")
        # Selecionar a família de tomada
        tomada_selecionada = selecionar_familia_tomada()

        # Selecionar a parede
        parede = selecionar_parede()

        # Obter os parâmetros elétricos do usuário
        parametros_elec = obter_parametros_elec()
        if parametros_elec is None:
            forms.alert("Falha ao coletar parâmetros elétricos.", exitscript=True)

        # Obter o ponto de inserção
        ponto_insercao = obter_ponto_insercao(parede)

        # Inserir a tomada com os parâmetros elétricos
        inserir_tomada(parede, tomada_selecionada, ponto_insercao, parametros_elec)

    except Exception as e:
        tb = traceback.format_exc()
        forms.alert("Ocorreu um erro:\n{}".format(tb))
        script.get_output().print_md("### Erro Geral no Script:\n{}".format(tb))


# Executar o script
if __name__ == "__main__":
    main()
