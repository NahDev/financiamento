import tkinter as tk
from tkinter import messagebox
from openpyxl import Workbook


def calcular_financiamento():
    valor_carro = float(valor_carro_entry.get())
    taxa_juros = float(taxa_juros_entry.get()) / 100
    parcelas = int(parcelas_entry.get())

    juros_mensais = taxa_juros / 12
    parcela = (valor_carro * juros_mensais) / (1 - (1 + juros_mensais) ** -parcelas)

    wb = Workbook()
    planilha = wb.active
    planilha.title = "Financiamento"

    planilha["A1"] = "Mês"
    planilha["B1"] = "Parcela"
    planilha["C1"] = "Juros"
    planilha["D1"] = "Amortização"
    planilha["E1"] = "Saldo devedor"

    saldo_devedor = valor_carro

    for mes in range(1, parcelas + 1):
        juros = saldo_devedor * juros_mensais
        amortizacao = parcela - juros
        saldo_devedor -= amortizacao

        planilha["A" + str(mes + 1)] = mes
        planilha["B" + str(mes + 1)] = parcela
        planilha["C" + str(mes + 1)] = juros
        planilha["D" + str(mes + 1)] = amortizacao
        planilha["E" + str(mes + 1)] = saldo_devedor

    arquivo_excel = "financiamento_carro.xlsx"
    wb.save(arquivo_excel)

    messagebox.showinfo(
        "Concluído",
        "O financiamento foi calculado com sucesso e o arquivo Excel foi salvo.",
    )


janela = tk.Tk()
janela.title("Simulação de Financiamento de Carro")

valor_carro_label = tk.Label(janela, text="Valor do Carro:")
valor_carro_label.pack()
valor_carro_entry = tk.Entry(janela)
valor_carro_entry.pack()

taxa_juros_label = tk.Label(janela, text="Taxa de Juros (%):")
taxa_juros_label.pack()
taxa_juros_entry = tk.Entry(janela)
taxa_juros_entry.pack()

parcelas_label = tk.Label(janela, text="Número de Parcelas:")
parcelas_label.pack()
parcelas_entry = tk.Entry(janela)
parcelas_entry.pack()

calcular_button = tk.Button(janela, text="Calcular", command=calcular_financiamento)
calcular_button.pack()

janela.mainloop()
