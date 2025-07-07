# fill_and_export.py  
# Python ≥ 3.8  |  pip install xlwings

import os
import xlwings as xw
from datetime import date

# 1. MAPEAMENTO: “Rótulo” → “Célula”
MAPPING_00001 = {
    "Nome do Cliente": "D8",
    "N° da Proposta": "D9",
    "Consultor": "D10",
    "Data": "D11",
    "Telefone": "D12",
    "Logradouro": "D13",
    "Endereço": "E13",
    "Bairro": "D14",
    "Cidade": "D15",
    "Estado": "J15",
    "Quantidade de Painéis": "G21",
    "Potência dos Painéis (W)": "I21",
    "Quantidade de Inversores": "G22",
    "Potência Inversor 1 (W)": "I22",
    "Estrutura Para": "G25",
    "Produção Média Mensal": "G26",
    "Preço": "G27",
}

MAPPING_00002 = {
    **MAPPING_00001,
    "Quantidade de Painéis 2":     "G33",
    "Potência dos Painéis (W) 2":  "I33",
    "Quantidade de Inversores 2":  "G34",
    "Potência Inversor 1 (W) 2":   "I34",
    "Estrutura Para 2":            "G37",
    "Produção Média Mensal 2":     "G38",
    "Preço 2":                     "G39",
}

MAPPING_00003 = {
    **MAPPING_00001,
    "Preço dos Equipamentos": "G27",
    "Preço da Mão de Obra": "G28",
    "Preço Total": "G29",
}

# 2. Pasta de saída
OUTPUT_DIR = "propostas"
os.makedirs(OUTPUT_DIR, exist_ok=True)


def get_next_proposal_number():
    # Busca o último número de proposta nos arquivos PDF
    last_number = 0
    if os.path.exists(OUTPUT_DIR):
        for filename in os.listdir(OUTPUT_DIR):
            if filename.endswith('.pdf'):
                try:
                    num = int(filename.split('PROPOSTA')[0])
                    last_number = max(last_number, num)
                except ValueError:
                    continue
    return str(last_number + 1)

def main():
    today = date.today().strftime("%d/%m/%Y")
    dados = {}
    numero_end = None
    next_number = get_next_proposal_number()

    # 3. Coleta dos valores com defaults e uppercase
    for label, cell in MAPPING.items():
        # Define default conforme o campo
        if label == "Data":
            default = today
        elif label == "N° da Proposta":
            default = next_number
        elif label == "Logradouro":
            default = "RUA"
        elif label == "Estado":
            default = "MG"
        elif label == "Cidade":
            default = "NOVA SERRANA"
        elif label == "Estrutura Para":
            default = "TELHADO METÁLICO"
        else:
            default = None

        # Prompt com default
        if default is not None:
            raw = input(f"{label} [{default}]: ").strip()
            valor = raw.upper() if raw else default.upper()
        else:
            raw = input(f"{label}: ").strip()
            valor = raw.upper()
        dados[label] = valor

        # Pergunta número do endereço logo após Logradouro
        if label == "Logradouro":
            raw_num = input("Nº [vazio se não houver]: ").strip()
            if raw_num:
                numero_end = raw_num.upper()

    # 4. Abre o template
    template = r"00001 - FAZER PROPOSTA PC.xlsx"
    app = xw.App(visible=False)
    wb  = app.books.open(template)
    sht = wb.sheets[0]

    MAPPING = MAPPING_00001

    if dados["Tipo de Proposta"] == "1- Proposta Simples":
        MAPPING = MAPPING_00001
    elif dados["Tipo de Proposta"] == "2- Proposta Dupla":
        MAPPING = MAPPING_00002
    elif dados["Tipo de Proposta"] == "3- Proposta com Mão de Obra":
        MAPPING = MAPPING_00003

    # 5. Preenche as células
    for label, cell in MAPPING.items():
        sht.range(cell).value = dados[label]

    # Preenche número do endereço, se informado
    if numero_end:
        sht.range("I13").value = "Nº"
        sht.range("J13").value = numero_end

    # 6. Gera nome do PDF
    numero = dados["N° da Proposta"]
    nome   = dados["Nome do Cliente"]
    filename = f"{numero}PROPOSTA {nome}.pdf"
    output_path = os.path.join(OUTPUT_DIR, filename)

    # 7. Exporta para PDF
    wb.to_pdf(output_path)

    # 8. Finaliza
    wb.close()
    app.quit()
    print(f"✔ PDF gerado: {output_path}")

if __name__ == "__main__":
    main()