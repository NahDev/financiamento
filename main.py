from tkinter import *
import openpyxl

def calcular_financiamento():
    valor_carro = float(valor_carro_entry.get())
    entrada = float(entrada_entry.get())
    prazo = int(prazo_entry.get())
    taxa_juros = float(taxa_juros_entry.get())

    valor_financiado = valor_carro - entrada
    juros_mensais = taxa_juros / 100 / 12
    total_juros = 0

    # Cálculo das parcelas em modelo price decrescente
    parcelas = []
    valor_parcela = valor_financiado * ((1 + juros_mensais) ** prazo) * juros_mensais / (((1 + juros_mensais) ** prazo) - 1)

    for i in range(prazo):
        juros = valor_financiado * juros_mensais
        total_juros += juros
        valor_financiado -= valor_parcela - juros
        parcelas.append(valor_parcela - juros)

    valor_total = valor_carro + total_juros

    # Criar uma nova planilha no arquivo Excel
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Mensalidades"

    # Escrever o cabeçalho
    sheet.cell(row=1, column=1).value = "Mês"
    sheet.cell(row=1, column=2).value = "Parcela semm Juros"
    sheet.cell(row=1, column=3).value = "Parcela com Juros"
    sheet.cell(row=1, column=4).value = "Valor Restante sem Juros"
    sheet.cell(row=1, column=5).value = "Valor Restante com Juros"

    # Escrever as informações de cada mês
    for i, parcela in enumerate(parcelas):
        sheet.cell(row=i+2, column=1).value = i + 1
        sheet.cell(row=i+2, column=2).value = parcela
        sheet.cell(row=i+2, column=3).value = valor_parcela
        sheet.cell(row=i+2, column=4).value = valor_financiado + (prazo - i - 1) * parcela
        sheet.cell(row=i+2, column=5).value = valor_financiado + (prazo - i - 1) * valor_parcela

    resultado_label.config(text="Valor total do financiamento: R$ {:.2f}\nValor da primeira parcela: R$ {:.2f}\nTotal de juros pagos: R$ {:.2f}".format(valor_total, parcelas[0], total_juros))

    # Salvar o arquivo Excel
    filename = "mensalidades.xlsx"
    workbook.save(filename)
    print("Mensalidades salvas no arquivo:", filename)

# Configuração da interface gráfica
janela = Tk()
janela.title("Simulação de Financiamento de Carro")

# Campos de entrada
valor_carro_label = Label(janela, text="Valor do carro:")
valor_carro_label.pack()
valor_carro_entry = Entry(janela)
valor_carro_entry.pack()

entrada_label = Label(janela, text="Valor da entrada:")
entrada_label.pack()
entrada_entry = Entry(janela)
entrada_entry.pack()

prazo_label = Label(janela, text="Prazo (em meses):")
prazo_label.pack()
prazo_entry = Entry(janela)
prazo_entry.pack()

taxa_juros_label = Label(janela, text="Taxa de juros (% ao ano):")
taxa_juros_label.pack()
taxa_juros_entry = Entry(janela)
taxa_juros_entry.pack()

calcular_button = Button(janela, text="Calcular", command=calcular_financiamento)
calcular_button.pack()

resultado_label = Label(janela, text="")
resultado_label.pack()

janela.mainloop()
