import pandas as pd

# Ler planilha
df = pd.read_excel("Pan 02-09.xlsx")

# Adicionar traço
def padronizar_placa(placa):
    placa = str(placa).strip().upper().replace(" ", "")
    if "-" in placa:
        return placa
    if len(placa) >= 4:
        return placa[:3] + "-" + placa[3:]
    return placa

df["placa"] = df["placa"].apply(padronizar_placa)

# Salvar de volta
df.to_excel("Pendencia_formatado.xlsx", index=False)
print("✅ Planilha atualizada com traços!")
