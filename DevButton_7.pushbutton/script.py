# -*- coding: utf-8 -*-
__title__ = "Listar Parâmetros da Família"
__doc__ = """Script para listar todos os parâmetros de uma instância de família selecionada, incluindo parâmetros de instância e de tipo."""

import clr
import traceback

# Importar as classes necessárias do Revit API
from Autodesk.Revit.DB import FamilyInstance, BuiltInParameter, StorageType
from Autodesk.Revit.UI.Selection import ObjectType
from Autodesk.Revit.Exceptions import InvalidOperationException

# Importações do pyRevit
from pyrevit import revit, forms, script

# Importar System.Windows.Forms para salvar CSV e exibir caixas de diálogo
clr.AddReference('System.Windows.Forms')
from System.Windows.Forms import DialogResult, MessageBox, MessageBoxButtons, MessageBoxIcon


def listar_parametros():
    output = script.get_output()
    try:
        # Instruir o usuário a selecionar uma instância de família
        msg = "Selecione a instância da família para listar seus parâmetros."
        referencia = revit.uidoc.Selection.PickObject(ObjectType.Element, msg)
        elemento = revit.doc.GetElement(referencia.ElementId)

        # Verificar se o elemento selecionado é uma instância de família
        if not isinstance(elemento, FamilyInstance):
            forms.alert("O elemento selecionado não é uma instância de família.", exitscript=True)

        # Obter o nome da família
        family_name_param = elemento.Symbol.get_Parameter(BuiltInParameter.ALL_MODEL_FAMILY_NAME)
        family_name = family_name_param.AsString() if family_name_param and family_name_param.HasValue else "Sem Família"

        # Obter o nome do tipo da família
        type_name_param = elemento.Symbol.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME)
        type_name = type_name_param.AsString() if type_name_param and type_name_param.HasValue else "Sem Tipo"

        # Obter o tipo de elemento (classe)
        tipo_elemento = type(elemento).__name__

        output.print_md("## Parâmetros da Instância da Família Selecionada")
        output.print_md("### Família: {}".format(family_name))
        output.print_md("### Tipo: {}".format(type_name))
        output.print_md("### Tipo de Elemento: {}".format(tipo_elemento))
        output.print_md("### ID: {}".format(elemento.Id))

        # Obter todos os parâmetros da instância
        parametros_instancia = elemento.Parameters

        # Obter todos os parâmetros do tipo da família
        parametros_tipo = elemento.Symbol.Parameters

        # Preparar listas para armazenar os detalhes dos parâmetros
        lista_parametros_instancia = []
        lista_parametros_tipo = []

        # Contadores para depuração
        contador_instancia = 0
        contador_tipo = 0

        # Iterar sobre os parâmetros de instância
        for parametro in parametros_instancia:
            try:
                nome_parametro = parametro.Definition.Name
                tipo_parametro = parametro.StorageType
                valor = ""

                if parametro.StorageType == StorageType.Double:
                    valor = parametro.AsDouble()
                    # Converter de pés para metros se necessário
                    valor = valor * 0.3048  # Revit utiliza pés como unidade padrão
                    valor = round(valor, 3)  # Arredondar para melhor legibilidade
                elif parametro.StorageType == StorageType.Integer:
                    valor = parametro.AsInteger()
                elif parametro.StorageType == StorageType.String:
                    valor = parametro.AsString()
                elif parametro.StorageType == StorageType.ElementId:
                    valor = parametro.AsElementId().IntegerValue
                else:
                    valor = "Desconhecido"

                lista_parametros_instancia.append({
                    "Nome": nome_parametro,
                    "Tipo": tipo_parametro.ToString(),
                    "Valor": valor,
                    "Origem": "Instância"
                })
                contador_instancia += 1
            except Exception as e:
                lista_parametros_instancia.append({
                    "Nome": "Erro ao obter parâmetro",
                    "Tipo": "Erro",
                    "Valor": str(e),
                    "Origem": "Instância"
                })

        # Iterar sobre os parâmetros do tipo
        for parametro in parametros_tipo:
            try:
                nome_parametro = parametro.Definition.Name
                tipo_parametro = parametro.StorageType
                valor = ""

                if parametro.StorageType == StorageType.Double:
                    valor = parametro.AsDouble()
                    # Converter de pés para metros se necessário
                    valor = valor * 0.3048  # Revit utiliza pés como unidade padrão
                    valor = round(valor, 3)  # Arredondar para melhor legibilidade
                elif parametro.StorageType == StorageType.Integer:
                    valor = parametro.AsInteger()
                elif parametro.StorageType == StorageType.String:
                    valor = parametro.AsString()
                elif parametro.StorageType == StorageType.ElementId:
                    valor = parametro.AsElementId().IntegerValue
                else:
                    valor = "Desconhecido"

                lista_parametros_tipo.append({
                    "Nome": nome_parametro,
                    "Tipo": tipo_parametro.ToString(),
                    "Valor": valor,
                    "Origem": "Tipo"
                })
                contador_tipo += 1
            except Exception as e:
                lista_parametros_tipo.append({
                    "Nome": "Erro ao obter parâmetro",
                    "Tipo": "Erro",
                    "Valor": str(e),
                    "Origem": "Tipo"
                })

        # Combinar as listas
        lista_parametros_combined = lista_parametros_instancia + lista_parametros_tipo

        # Adicionar logs para depuração
        output.print_md("### Número de Parâmetros de Instância: {}".format(contador_instancia))
        output.print_md("### Número de Parâmetros de Tipo: {}".format(contador_tipo))

        # Verificar se existem parâmetros
        if not lista_parametros_combined:
            output.print_md("### Nenhum parâmetro encontrado na instância ou tipo da família.")
            forms.alert("Nenhum parâmetro encontrado na instância ou tipo da família.", exitscript=True)
        else:
            # Adicionar logs para depuração: mostrar alguns parâmetros
            if len(lista_parametros_combined) > 0:
                output.print_md("### Exemplos de Parâmetros:")
                for i, param in enumerate(
                        lista_parametros_combined[:5]):  # Mostrar apenas os 5 primeiros para não sobrecarregar
                    output.print_md("#### Parâmetro {}: {}".format(i + 1, param))

            # Preparar dados para a tabela como lista de listas
            table_rows = [
                [param["Nome"], param["Tipo"], param["Valor"], param["Origem"]]
                for param in lista_parametros_combined
            ]

            # Exibir os parâmetros na janela de output do pyRevit
            output.print_table(
                table_data=table_rows,
                title="Lista de Parâmetros",
                columns=["Nome", "Tipo", "Valor", "Origem"]
            )

            # Opcional: Salvar os parâmetros em um arquivo CSV usando MessageBox para confirmar
            resposta = MessageBox.Show(
                "Deseja salvar a lista de parâmetros em um arquivo CSV?",
                "Salvar CSV",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question  # Especifica o ícone da caixa de diálogo
            )

            if resposta == DialogResult.Yes:
                import csv
                from System.IO import StreamWriter
                import System.Text

                caminho = forms.save_file(
                    default_name="parametros_familia.csv",
                    file_types=[("CSV files", "*.csv")],
                    prompt="Salvar lista de parâmetros como CSV"
                )

                if caminho:
                    try:
                        with StreamWriter(caminho, False, System.Text.Encoding.UTF8) as sw:
                            writer = csv.writer(sw)
                            writer.writerow(["Nome", "Tipo", "Valor", "Origem"])
                            for param in lista_parametros_combined:
                                writer.writerow([param["Nome"], param["Tipo"], param["Valor"], param["Origem"]])
                        forms.alert("Parâmetros salvos com sucesso em:\n{}".format(caminho))
                    except Exception as e:
                        forms.alert("Erro ao salvar o arquivo CSV:\n{}".format(e))

    except InvalidOperationException:
        forms.alert("Nenhum elemento foi selecionado.")
    except Exception as e:
        tb = traceback.format_exc()
        forms.alert("Ocorreu um erro:\n{}".format(tb))
        output.print_md("### Erro no Script:\n{}".format(tb))


def main():
    listar_parametros()


if __name__ == "__main__":
    main()
