from tkinter import *
import openpyxl


def calcular_financiamento():
    valor_carro = float(valor_carro_entry.get())
    entrada = float(entrada_entry.get())
    prazo = int(prazo_entry.get())
    taxa_juros = float(taxa_juros_entry.get())

    valor_financiado = valor_carro - entrada
    juros_mensais = taxa_juros / 100 / 12

    # Calcular o valor da parcela sem juros
    valor_parcela_sem_juros = valor_financiado / prazo

    # Criar uma nova planilha no arquivo Excel
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Mensalidades"

    # Escrever o cabeçalho
    sheet.cell(row=1, column=1).value = "Mês"
    sheet.cell(row=1, column=2).value = "Valor da Parcela com Juros"
    sheet.cell(row=1, column=3).value = "Valor da Parcela sem Juros"
    sheet.cell(row=1, column=4).value = "Restante a Pagar com Juros"
    sheet.cell(row=1, column=5).value = "Restante a Pagar sem Juros"

    # Calcular e escrever as mensalidades
    for i in range(prazo):
        valor_juros = valor_financiado * juros_mensais
        valor_parcela_com_juros = valor_parcela_sem_juros + valor_juros
        valor_financiado -= valor_parcela_sem_juros

        sheet.cell(row=i + 2, column=1).value = i + 1
        sheet.cell(row=i + 2, column=2).value = valor_parcela_com_juros
        sheet.cell(row=i + 2, column=3).value = valor_parcela_sem_juros
        sheet.cell(row=i + 2, column=4).value = valor_financiado
        sheet.cell(row=i + 2, column=5).value = valor_financiado + (
            valor_parcela_sem_juros * (prazo - i - 1)
        )

        valor_parcela_sem_juros -= valor_parcela_sem_juros / prazo

    valor_total = valor_carro + valor_financiado

    resultado_label.config(
        text="Valor total do financiamento: R$ {:.2f}\nValor da parcela mensal (última): R$ {:.2f}\n".format(
            valor_total, valor_parcela_sem_juros
        )
    )

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
